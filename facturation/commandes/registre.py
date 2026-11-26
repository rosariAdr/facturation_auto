"""Commande `facturation registre` : registre des factures, envoi et paiement (F-25).

    facturation registre [--ecole X] [--mois AAAA-MM] [--echues] [--trous] [--profil CHEMIN]
    facturation registre initialiser [--reprendre-historique] [--oui] [--profil CHEMIN]
    facturation registre envoyee N° --le DATE [--remplacer] [--oui] [--profil CHEMIN]
    facturation registre payee N° --le DATE [--remplacer] [--oui] [--profil CHEMIN]
    facturation registre note N° "texte" [--remplacer] [--oui] [--profil CHEMIN]
    facturation registre manquant N° --raison "texte" [--oui] [--profil CHEMIN]

Sans action, la commande affiche le tableau du registre (factures échues en tête) et rend 3
s'il y a une facture échue ou une anomalie, 0 sinon (§ 1.6). Chaque action d'écriture affiche
l'avant et l'après, puis demande « oui » (ou `--oui`) ; la saisie dans Excel reste possible
(ADR031). Aucune écriture sans validation explicite.

Lectures retenues (`💡 [proposé — à valider]`) :
- `--mois` retient les factures dont la période désigne ce mois (« Novembre 2026 », « Nov 2026 »,
  « 11/2026 », « 2026-11 ») ; une période qui ne désigne pas un seul mois se replie sur la date
  d'émission ;
- avec un filtre (`--ecole`, `--mois`, `--echues`), seules les anomalies des factures affichées
  comptent, plus, sans `--mois` ni `--echues`, les N° de `Sessions` absents du registre (de
  l'école filtrée) ; les lignes illisibles de `Sessions` ne comptent que sans filtre ;
- `--trous` n'affiche que les numéros manquants sans raison notée, et rend 3 s'il y en a (une
  raison est à noter, ADR029), 0 sinon ;
- un MASTER ouvert est refusé avant la question ; une réponse autre que « oui » (ou la fin de
  l'entrée) n'écrit rien et rend 1 ; une écriture que le contrôle après écriture trouve non
  conforme rend 2.
"""

from __future__ import annotations

import argparse
import re
import sys
import unicodedata
from collections.abc import Mapping, Sequence
from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path
from typing import Any

from facturation import chemins
from facturation.commandes import ERREUR, ETIQUETTES, argument_date, argument_mois, charger_profil, demander

_AIDE_PROFIL = "Profil à utiliser (défaut : emplacement standard)"
_AIDE_OUI = "Écrire sans poser la question (validation donnée d'avance)"
_AIDE_REMPLACER = "Remplacer une valeur déjà saisie"
_QUESTION = "Écrire dans le MASTER ?"
_ENTETES = ("N°", "École", "Période", "Émission", "Échéance", "Total HT", "Envoi", "Paiement", "Statut", "Anomalies")
_ABREVIATIONS = ("janv", "fevr", "mars", "avr", "mai", "juin", "juil", "aout", "sept", "oct", "nov", "dec")
_PERIODE_MOIS_RE = re.compile(r"([a-z]+)\.?\s+(\d{4})")
_PERIODE_NUMERO_RE = re.compile(r"(\d{1,2})/(\d{4})")
_PERIODE_ISO_RE = re.compile(r"(\d{4})-(\d{2})")


def ajouter(sous) -> None:
    """Déclare `registre` et ses actions dans les sous-commandes de `facturation`."""
    registre = sous.add_parser(
        "registre", help="Registre des factures : échéances, envoi, paiement, numéros manquants",
        description="Sans action : tableau du registre (onglet Factures du MASTER), factures échues en tête.",
    )
    registre.add_argument("--ecole", default=None, help="Seulement les factures de cette école (code ou libellé)")
    registre.add_argument("--mois", type=argument_mois, default=None, help="Seulement les factures de ce mois (AAAA-MM)")
    registre.add_argument("--echues", action="store_true", help="Seulement les factures échues non payées")
    registre.add_argument("--trous", action="store_true", help="Lister les numéros manquants sans raison notée")
    registre.add_argument("--profil", type=Path, default=None, help=_AIDE_PROFIL)
    registre.set_defaults(func=_tableau)
    actions = registre.add_subparsers(dest="action", metavar="ACTION")

    initialiser = actions.add_parser("initialiser", help="Créer l'onglet Factures, et y reprendre l'historique")
    initialiser.add_argument(
        "--reprendre-historique", action="store_true",
        help="Ajouter une ligne par N° de Sessions absent du registre (dates à compléter, jamais inventées)",
    )
    _options_d_ecriture(initialiser)
    initialiser.set_defaults(func=_initialiser)

    for nom, aide, fonction in (
        ("envoyee", "Noter la date d'envoi d'une facture", _envoyee),
        ("payee", "Noter la date de paiement d'une facture", _payee),
    ):
        action = actions.add_parser(nom, help=aide)
        action.add_argument("numero", metavar="N°", help="N° de la facture, ex. R_ABC001")
        action.add_argument("--le", required=True, type=argument_date, metavar="DATE",
                            help="Date, JJ/MM/AAAA ou AAAA-MM-JJ")
        action.add_argument("--remplacer", action="store_true", help=_AIDE_REMPLACER)
        _options_d_ecriture(action)
        action.set_defaults(func=fonction)

    note = actions.add_parser("note", help="Écrire la note d'une facture")
    note.add_argument("numero", metavar="N°", help="N° de la facture, ex. R_ABC001")
    note.add_argument("texte", help="Texte de la note")
    note.add_argument("--remplacer", action="store_true", help=_AIDE_REMPLACER)
    _options_d_ecriture(note)
    note.set_defaults(func=_note)

    manquant = actions.add_parser("manquant", help="Noter la raison d'un numéro manquant (ADR029)")
    manquant.add_argument("numero", metavar="N°", help="N° manquant, ex. R_ABC002")
    manquant.add_argument("--raison", required=True, help="Pourquoi ce numéro n'a pas servi")
    _options_d_ecriture(manquant)
    manquant.set_defaults(func=_manquant)


def _options_d_ecriture(action: argparse.ArgumentParser) -> None:
    action.add_argument("--oui", action="store_true", help=_AIDE_OUI)
    # Sans valeur par défaut : un --profil donné avant l'action n'est pas effacé.
    action.add_argument("--profil", type=Path, default=argparse.SUPPRESS, help=_AIDE_PROFIL)


# ──────────────────────────────────────────────────────────────────────
# Mise en forme
# ──────────────────────────────────────────────────────────────────────


def _jj(valeur: Any) -> str:
    if isinstance(valeur, datetime):
        valeur = valeur.date() if valeur.time() == time() else valeur
        if isinstance(valeur, datetime):
            return valeur.strftime("%d/%m/%Y %H:%M")
    if isinstance(valeur, date):
        return valeur.strftime("%d/%m/%Y")
    return ""


def _euros(montant: Decimal | None) -> str:
    if montant is None:
        return ""
    from facturation import heures

    arrondi = heures.arrondi_centime(montant)
    entier, _, decimales = f"{abs(arrondi):.2f}".partition(".")
    milliers = f"{int(entier):,}".replace(",", " ")
    return f"{'-' if arrondi < 0 else ''}{milliers},{decimales} €"


def _valeur(valeur: Any) -> str:
    """Valeur d'une cellule dans l'avant / après."""
    if valeur is None or (isinstance(valeur, str) and not valeur.strip()):
        return "(vide)"
    if isinstance(valeur, (date, datetime)):
        return _jj(valeur)
    if isinstance(valeur, Decimal):
        return _euros(valeur)
    return str(valeur)


def _cle(numero: Any) -> str:
    return "".join(str(numero or "").split()).upper()


def _sans_accents(texte: str) -> str:
    decompose = unicodedata.normalize("NFD", texte)
    return "".join(c for c in decompose if not unicodedata.combining(c))


def _mois_de_la_periode(periode: str) -> tuple[int, int] | None:
    """(année, mois) que désigne une période : « Novembre 2026 », « Nov. 2026 », « 11/2026 », « 2026-11 »."""
    texte = " ".join(_sans_accents(periode.casefold()).split())
    trouve = _PERIODE_MOIS_RE.fullmatch(texte)
    if trouve:
        nom = trouve.group(1)
        for numero, (complet, abrege) in enumerate(zip(chemins.NOMS_MOIS, _ABREVIATIONS), start=1):
            if nom in (_sans_accents(complet), abrege):
                return int(trouve.group(2)), numero
        return None
    trouve = _PERIODE_NUMERO_RE.fullmatch(texte)
    if trouve and 1 <= int(trouve.group(1)) <= 12:
        return int(trouve.group(2)), int(trouve.group(1))
    trouve = _PERIODE_ISO_RE.fullmatch(texte)
    if trouve and 1 <= int(trouve.group(2)) <= 12:
        return int(trouve.group(1)), int(trouve.group(2))
    return None


def _dans_le_mois(ligne, mois: str) -> bool:
    debut, fin = chemins.mois_valide(mois)
    designe = _mois_de_la_periode(ligne.periode)
    if designe is not None:
        return designe == (debut.year, debut.month)
    return ligne.date_emission is not None and debut <= ligne.date_emission <= fin


def _tableau_texte(lignes: Sequence[Sequence[str]]) -> list[str]:
    largeurs = [max(len(ligne[i]) for ligne in lignes) for i in range(len(lignes[0]))]
    textes = ["  ".join(cellule.ljust(largeur) for cellule, largeur in zip(ligne, largeurs)).rstrip() for ligne in lignes]
    textes.insert(1, "  ".join("-" * largeur for largeur in largeurs))
    return textes


def _erreur(exc: BaseException) -> None:
    print(f"{ETIQUETTES[ERREUR]}Registre — {exc}", file=sys.stderr)


def _erreurs_attendues() -> tuple[type[BaseException], ...]:
    from facturation import master
    from facturation.fichiers import ClasseurVerrouille

    # ValueError couvre RegistreInvalide, MasterInvalide et FichierAbime ; OSError, FileNotFoundError.
    return (ValueError, master.MasterVerrouille, ClasseurVerrouille, OSError)


def _onglet_factures_present(chemin: Path) -> bool:
    from openpyxl import load_workbook

    from facturation.registre import ONGLET

    wb = load_workbook(chemin, read_only=True)
    try:
        return ONGLET in wb.sheetnames
    finally:
        wb.close()


# ──────────────────────────────────────────────────────────────────────
# Tableau
# ──────────────────────────────────────────────────────────────────────


def _codes_ecoles(m, code_ecole) -> list[str]:
    """Écoles de `Règles_facturation` et de `Écoles`, en code ; une valeur illisible est ignorée."""
    valeurs = [ligne.get("École") for ligne in m.regles_facturation]
    valeurs += [ligne.get("Code facturation") or ligne.get("École") for ligne in m.ecoles]
    codes: list[str] = []
    for valeur in valeurs:
        try:
            code = code_ecole(valeur, m.ecoles)
        except ValueError:
            continue
        if code and code not in codes:
            codes.append(code)
    return codes


def _trous(profil, m, lignes) -> int:
    from facturation import numerotation, registre
    from facturation.sessions import code_ecole

    codes = []
    for code in _codes_ecoles(m, code_ecole):
        try:
            codes.append(numerotation.valider_code_ecole(code))
        except ValueError:
            continue
    trous = registre.trous_sans_raison(
        lignes, numerotation.lire_numeros(profil.master), prefixe=profil.prefixe_facture, ecoles=codes,
    )
    print(f"NUMÉROS MANQUANTS SANS RAISON NOTÉE : {len(trous)}")
    for numero in trous:
        print(f"  {numero} — notez sa raison : facturation registre manquant {numero} --raison \"…\"")
    if not trous:
        print("  Aucun : chaque numéro manquant a sa raison au registre (ADR029).")
    return 3 if trous else 0


def _tableau(args: argparse.Namespace) -> int:
    from facturation import master, registre, sessions
    from facturation import pcf_format as fmt

    profil = charger_profil(args.profil)
    if profil is None:
        return 1
    fmt.set_data_root(profil.racine)
    try:
        m = master.lire_master(profil.master)
        if not _onglet_factures_present(profil.master):
            print(
                f"{ETIQUETTES[ERREUR]}Registre — onglet « Factures » absent du MASTER {profil.master.name} : ajoutez "
                f"l'onglet Factures avec `{registre.COMMANDE_INITIALISER}` (--reprendre-historique pour y reprendre "
                "les factures de Sessions).",
                file=sys.stderr,
            )
            return 1
        lignes = registre.lire_registre(m)
        seances, illisibles = sessions.lire_sessions(m)
        aujourd_hui = date.today()
        print(f"Registre des factures — entité {profil.entite} — aujourd'hui {aujourd_hui:%d/%m/%Y}")
        if args.trous:
            return _trous(profil, m, lignes)
        code_filtre = None
        if args.ecole is not None:
            try:
                code_filtre = sessions.code_ecole(args.ecole, m.ecoles)
            except ValueError:
                code_filtre = None
            code_filtre = code_filtre or args.ecole.strip().upper()
    except _erreurs_attendues() as exc:
        _erreur(exc)
        return 1

    filtres = []
    affichees = list(lignes)
    if code_filtre is not None:
        filtres.append(f"école {code_filtre}")
        affichees = [ligne for ligne in affichees if ligne.ecole.upper() == code_filtre]
    if args.mois is not None:
        filtres.append(f"mois {args.mois}")
        affichees = [ligne for ligne in affichees if _dans_le_mois(ligne, args.mois)]
    if args.echues:
        filtres.append("factures échues")
        affichees = [ligne for ligne in affichees if registre.statut_affiche(ligne, aujourd_hui=aujourd_hui)
                     == registre.ECHUE]
    if filtres:
        print(f"Filtre : {', '.join(filtres)}")

    par_numero: dict[str, list] = {}
    for seance in seances:
        if _cle(seance.n_facture):
            par_numero.setdefault(_cle(seance.n_facture), []).append(seance)
    echues = registre.echues(affichees, aujourd_hui=aujourd_hui)
    premieres = {id(ligne) for ligne in echues}
    ordre = echues + [ligne for ligne in affichees if id(ligne) not in premieres]

    constats: list[str] = []
    rangees = [_ENTETES]
    for ligne in ordre:
        propres = registre.anomalies([ligne], par_numero.get(_cle(ligne.n_facture), []))
        constats.extend(propres)
        statut = registre.statut_affiche(ligne, aujourd_hui=aujourd_hui)
        retard = registre.retard_jours(ligne, aujourd_hui=aujourd_hui)
        rangees.append((
            ligne.n_facture, ligne.ecole, ligne.periode, _jj(ligne.date_emission), _jj(ligne.echeance),
            _euros(ligne.total_ht), _jj(ligne.date_envoi), _jj(ligne.date_paiement),
            f"{statut} (+{retard} j)" if retard else statut, str(len(propres)) if propres else "—",
        ))

    if not args.echues and args.mois is None:
        au_registre = {_cle(ligne.n_facture) for ligne in lignes}
        orphelines = [s for cle, liste in par_numero.items() if cle not in au_registre for s in liste
                      if code_filtre is None or s.ecole.upper() == code_filtre]
        constats.extend(registre.anomalies([], orphelines))
        if not filtres:
            constats.extend(message for _, message in illisibles)

    print()
    if len(rangees) > 1:
        for texte in _tableau_texte(rangees):
            print(texte)
    else:
        print("Aucune facture ne correspond." if filtres else "Le registre ne contient encore aucune facture.")
    if constats:
        print(f"\nANOMALIES : {len(constats)}")
        for constat in constats:
            print(f"  - {constat}")
    print(f"\n{len(ordre)} facture(s) affichée(s), dont {len(echues)} échue(s) non payée(s) ; "
          f"{len(constats)} anomalie(s).")
    return 3 if echues or constats else 0


# ──────────────────────────────────────────────────────────────────────
# Écritures
# ──────────────────────────────────────────────────────────────────────


def _profil_pour_ecrire(args: argparse.Namespace):
    """Profil chargé, racine branchée et contrôlée, MASTER fermé ; None après avoir affiché l'erreur."""
    from facturation import master
    from facturation import pcf_format as fmt

    profil = charger_profil(getattr(args, "profil", None))
    if profil is None:
        return None
    fmt.set_data_root(profil.racine)
    try:
        chemins.refuser_racine_dans_un_depot(profil.racine)
        master.verifier_classeur_non_ouvert(profil.master)
    except _erreurs_attendues() as exc:
        _erreur(exc)
        return None
    return profil


def _valider(args: argparse.Namespace) -> bool:
    """Validation explicite : `--oui`, ou « oui » répondu à la question."""
    if args.oui:
        return True
    reponse = demander(_QUESTION, ("oui", "non"))
    if reponse != "oui":
        if reponse is None:
            print()
        print("Rien n'a été écrit.")
        return False
    return True


def _ecrire(ecriture) -> int:
    """Lance l'écriture validée ; affiche la sauvegarde, ou l'erreur avec son code."""
    from facturation import registre

    try:
        sauvegarde = ecriture()
    except registre.EcritureNonConforme as exc:
        _erreur(exc)
        return 2
    except _erreurs_attendues() as exc:
        _erreur(exc)
        return 1
    print(f"Écrit dans le MASTER. Copie d'avant l'écriture : {sauvegarde}")
    return 0


def _mettre_a_jour(args: argparse.Namespace, **champs: Any) -> int:
    from facturation import registre

    profil = _profil_pour_ecrire(args)
    if profil is None:
        return 1
    try:
        changements = registre.apercu_mise_a_jour(profil, args.numero, remplacer=args.remplacer, **champs)
    except _erreurs_attendues() as exc:
        _erreur(exc)
        return 1
    numero = _cle(args.numero)
    if not changements:
        print(f"Facture {numero} : rien ne change (valeurs déjà saisies) ; rien n'a été écrit.")
        return 0
    print(f"Registre — facture {numero} :")
    for colonne, (avant, apres) in changements.items():
        print(f"  {colonne} : {_valeur(avant)} → {_valeur(apres)}")
    if not _valider(args):
        return 1
    return _ecrire(lambda: registre.mettre_a_jour_registre(
        profil, args.numero, remplacer=args.remplacer, confirme=True, **champs))


def _envoyee(args: argparse.Namespace) -> int:
    return _mettre_a_jour(args, date_envoi=args.le)


def _payee(args: argparse.Namespace) -> int:
    return _mettre_a_jour(args, date_paiement=args.le)


def _note(args: argparse.Namespace) -> int:
    return _mettre_a_jour(args, note=args.texte)


def _manquant(args: argparse.Namespace) -> int:
    from facturation import registre

    profil = _profil_pour_ecrire(args)
    if profil is None:
        return 1
    try:
        ligne = registre.apercu_manquant(profil, args.numero, raison=args.raison)
    except _erreurs_attendues() as exc:
        _erreur(exc)
        return 1
    print("Registre — numéro manquant (ADR029 : il ne sera jamais réutilisé) :")
    print(f"  N° {ligne[registre.COL_NUMERO]} — école {ligne[registre.COL_ECOLE]} — statut "
          f"« {registre.ABANDONNEE} » — note : {ligne[registre.COL_NOTE]}")
    if not _valider(args):
        return 1
    return _ecrire(lambda: registre.noter_manquant(profil, args.numero, raison=args.raison, confirme=True))


def _ligne_reprise(ligne: Mapping[str, Any]) -> str:
    from facturation import registre

    morceaux = [ligne[registre.COL_NUMERO], ligne[registre.COL_ECOLE] or "(école vide)",
                ligne[registre.COL_PERIODE] or "(période vide)", _euros(ligne[registre.COL_TOTAL]) or "(total vide)",
                ligne[registre.COL_STATUT]]
    return "  " + " — ".join(str(m) for m in morceaux)


def _initialiser(args: argparse.Namespace) -> int:
    from facturation import registre

    profil = _profil_pour_ecrire(args)
    if profil is None:
        return 1
    try:
        cree, reprises = registre.apercu_initialisation(profil, reprendre_historique=args.reprendre_historique)
    except _erreurs_attendues() as exc:
        _erreur(exc)
        return 1
    if not cree and not reprises:
        print("Rien à reprendre : toutes les factures de Sessions sont déjà au registre ; rien n'a été écrit.")
        return 0
    if cree:
        print(f"L'onglet « {registre.ONGLET} » sera ajouté au MASTER, avec ses {len(registre.COLONNES)} colonnes.")
    if args.reprendre_historique:
        print(f"Factures reprises de Sessions : {len(reprises)} (statut « {registre.EMISE} » ; dates d'émission, "
              "échéances, envois et paiements à compléter : ils ne sont jamais inventés).")
        for ligne in reprises:
            print(_ligne_reprise(ligne))
    if not _valider(args):
        return 1
    return _ecrire(lambda: registre.initialiser_registre(
        profil, reprendre_historique=args.reprendre_historique, confirme=True))
