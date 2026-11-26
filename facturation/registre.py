"""registre.py — Registre des factures : échéances, envoi et paiement (F-25, ADR029, ADR031).

L'onglet `Factures` du MASTER porte une ligne par facture (MASTER_SCHEMA) : N°, école, contrat,
période, date d'émission, échéance, total HT, date d'envoi, date de paiement, statut et note.

- **Lecture** : `lire_registre` lit l'onglet sous une forme typée (`LigneRegistre`) et refuse
  toute cellule illisible (`RegistreInvalide`, qui nomme la ligne et la colonne) ;
  `statut_affiche`, `retard_jours` et `echues` calculent l'état d'une facture au jour donné ;
  « échue » et « à compléter » sont calculés à l'affichage, jamais écrits ; `anomalies` rapproche
  le registre de `Sessions` ; `trous_sans_raison` liste les numéros manquants sans raison notée
  (ADR029).
- **Écriture sur une feuille déjà ouverte** (`creer_onglet`, `ajouter_ligne`, `modifier_ligne`,
  `noter_numero_manquant`) : aucune ne prend de verrou, c'est l'appelant qui le tient (F-17).
  L'outil écrit des valeurs, jamais de formule, et n'écrase jamais une cellule à formule (§ 1.4).
- **Écritures de premier niveau** (`mettre_a_jour_registre`, `noter_manquant`,
  `initialiser_registre`) : validation explicite exigée ; l'opération est d'abord jouée à blanc
  (un refus ne laisse ni verrou ni sauvegarde), puis rejouée sous `master.ecriture_master`
  (refus d'un MASTER ouvert, verrou, sauvegarde), enregistrée de façon atomique et relue : seules
  les cellules annoncées ont changé. Chaque écriture ajoute un événement `registre` au journal
  des écritures et rend la sauvegarde. `apercu_mise_a_jour`, `apercu_manquant` et
  `apercu_initialisation` montrent, sans rien écrire, ce que ferait l'écriture (la commande
  affiche l'avant et l'après avant de demander « oui »).

Lectures retenues (`💡 [proposé — à valider]`) :
- `retard_jours` rend None pour une facture qui n'est pas échue (payée, abandonnée, en émission,
  sans échéance, ou le jour même de l'échéance) ;
- les dates lues sont des `date` ; les montants des `Decimal` du nombre tel qu'il s'écrit ; une
  cellule vide donne None (dates, montant, contrat) ou "" (statut, note) ; un statut s'écrit avec
  ou sans accents (« emise ») ; l'école est lue comme dans `Sessions` (code, ou libellé traduit
  par l'onglet `Écoles`, § 15) ; une ligne remplie sans N° est refusée ;
- contrôle du total (§ 1.2, aucune tolérance arbitraire) : le `Total HT` du registre, arrondi au
  centime, doit égaler la somme des `Total HT` des lignes `Sessions` du N°, ou, quand chaque ligne
  porte l'arrondi de ses propres heures × taux, l'arrondi de la somme exacte des heures × taux
  (c'est le total de la facture, qui arrondit une seule fois) ; une facture non abandonnée sans
  aucune ligne dans `Sessions` est une anomalie ; une ligne abandonnée n'est pas contrôlée ;
- « statut incohérent avec les dates » : « payée » sans date de paiement, « envoyée » sans date
  d'envoi ; une date saisie dans Excel sans changer le statut n'est pas une anomalie (le statut
  affiché suit les dates) ;
- `modifier_ligne` : une date se compare à celle déjà saisie (même date = rien ne change, sans
  refus) ; une date ou une note déjà saisie d'une autre valeur ne se remplace qu'avec
  `remplacer=True`, une cellule illisible aussi ; une formule n'est jamais écrasée, même avec
  `remplacer=True` ; envoi et paiement sont refusés sur une facture « en émission » (son contrôle
  n'est pas encore vert) ; `statut="émise"` vient de « en émission » (ou d'un statut vide),
  `statut="abandonnée"` de « en émission », « émise » ou d'un statut vide, et `statut` ne se
  combine pas avec une date ;
- `ajouter_ligne` refuse un texte qui commence par « = » (ce serait une formule) ;
- `noter_manquant` déduit l'école du N° et du préfixe du profil (« R_ABC002 » → ABC) et refuse un
  N° présent dans `Sessions` (ce n'est pas un numéro manquant) ;
- `lignes_historiques` écrit le N° en majuscules ; un N° dont les lignes `Sessions` portent deux
  écoles est refusé ; un total vide dans l'une de ses lignes laisse le total vide (jamais
  inventé) ; une ligne illisible de `Sessions` arrête la reprise ;
- avant d'écrire, `master.fidelite_aller_retour` est consulté (§ 15 : « garder le contrôle de
  fidélité à l'écriture ») : une perte autre que la valeur en cache des formules refuse
  l'écriture ;
- un événement `registre` par facture écrite ; la création de l'onglet a son propre événement,
  au N° nul ; rien à écrire (même date redonnée, rien à reprendre) lève `RegistreInvalide` sans
  rien écrire.
"""

from __future__ import annotations

import math
import re
import unicodedata
import zipfile
from collections.abc import Callable, Iterable, Mapping, Sequence
from dataclasses import dataclass, field
from datetime import date, datetime, time, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import TYPE_CHECKING, Any

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

from facturation import chemins, fichiers, heures, journal
from facturation.master import (
    ONGLETS_PREVUS,
    Master,
    MasterInvalide,
    MasterVerrouille,
    derniere_ligne_remplie,
    ecriture_master,
    fidelite_aller_retour,
    index_colonnes,
    lire_master,
    poste_courant,
)
from facturation.numerotation import numeros_manquants, valider_code_ecole
from facturation.sessions import LigneSessions, code_ecole, lire_sessions

if TYPE_CHECKING:
    from openpyxl import Workbook
    from openpyxl.worksheet.worksheet import Worksheet

    from facturation.config import Profil

ONGLET = "Factures"
COLONNES = ONGLETS_PREVUS["Factures"]
STATUTS_SAISIS = ("en émission", "émise", "envoyée", "payée", "abandonnée")   # 💡 « en émission », « abandonnée » (Q-B6)
STATUTS_AFFICHES = STATUTS_SAISIS + ("échue", "à compléter")                  # calculés, jamais écrits

EN_EMISSION, EMISE, ENVOYEE, PAYEE, ABANDONNEE = STATUTS_SAISIS
ECHUE, A_COMPLETER = STATUTS_AFFICHES[len(STATUTS_SAISIS):]

(
    COL_NUMERO, COL_ECOLE, COL_CONTRAT, COL_PERIODE, COL_EMISSION, COL_ECHEANCE, COL_TOTAL,
    COL_ENVOI, COL_PAIEMENT, COL_STATUT, COL_NOTE,
) = COLONNES
COLONNES_DATES = (COL_EMISSION, COL_ECHEANCE, COL_ENVOI, COL_PAIEMENT)
FORMAT_DATE = "DD/MM/YYYY"

# Champs de `modifier_ligne` → colonne du registre.
CHAMPS_MODIFIABLES = {"date_envoi": COL_ENVOI, "date_paiement": COL_PAIEMENT, "statut": COL_STATUT, "note": COL_NOTE}
STATUTS_DE_L_EMISSION = (EMISE, ABANDONNEE)   # seuls statuts que `modifier_ligne` écrit (F-17)

COMMANDE_INITIALISER = "facturation registre initialiser"
COMMANDE_REPRISE = "registre initialiser --reprendre-historique"

_LIBELLES_DATES = {COL_ENVOI: "date d'envoi", COL_PAIEMENT: "date de paiement"}
_DATE_JJ_MM_AAAA = re.compile(r"\d{1,2}/\d{1,2}/\d{4}")
_DATE_AAAA_MM_JJ = re.compile(r"\d{4}-\d{1,2}-\d{1,2}")
_ESPACES_MILLIERS = "   "
_MONTANT_RE = re.compile(rf"[+-]?(?:\d{{1,3}}(?:[{_ESPACES_MILLIERS}]\d{{3}})+|\d+)(?:[.,]\d+)?")
# Erreurs d'un classeur abîmé à l'ouverture (openpyxl, zipfile) ; PermissionError est traitée à part.
_ERREURS_OUVERTURE = (KeyError, SyntaxError, ValueError, TypeError, IndexError, EOFError, OSError,
                      zipfile.BadZipFile, InvalidFileException)
# Seule perte d'un aller-retour openpyxl tolérée à l'écriture : la valeur en cache des formules.
_PERTE_TOLEREE = "cellule(s) à formule"


class RegistreInvalide(ValueError):
    """Registre illisible (cellule, N° en double), ou écriture refusée : le message dit quoi corriger."""


class EcritureNonConforme(RuntimeError):
    """Le contrôle après écriture a trouvé dans le MASTER autre chose que les cellules annoncées."""


@dataclass(frozen=True)
class LigneRegistre:
    """Une ligne lue de l'onglet `Factures`."""

    rang: int                     # ligne Excel
    n_facture: str
    ecole: str                    # code (comme `Sessions`)
    contrat: str | None
    periode: str
    date_emission: date | None
    echeance: date | None
    total_ht: Decimal | None
    date_envoi: date | None
    date_paiement: date | None
    statut_saisi: str             # l'un des STATUTS_SAISIS, ou "" (vide, ou « échue » toléré)
    note: str


class _Illisible(ValueError):
    """Cellule illisible : le message nomme la colonne et cite la valeur."""


# ──────────────────────────────────────────────────────────────────────
# Textes et valeurs
# ──────────────────────────────────────────────────────────────────────


def _est_vide(valeur: Any) -> bool:
    return valeur is None or (isinstance(valeur, str) and not valeur.strip())


def _texte(valeur: Any) -> str:
    """Contenu d'une cellule en texte NFC, sans espaces autour ; "" si vide ; 3.0 s'écrit 3."""
    if valeur is None:
        return ""
    if isinstance(valeur, float) and valeur.is_integer():
        valeur = int(valeur)
    return unicodedata.normalize("NFC", str(valeur)).strip()


def _cle(numero: Any) -> str:
    """Forme de comparaison d'un N° : sans espaces, en majuscules (« r_abc001 » → « R_ABC001 »)."""
    return "".join(_texte(numero).split()).upper()


def _sans_accents(texte: str) -> str:
    decompose = unicodedata.normalize("NFD", texte)
    return "".join(c for c in decompose if not unicodedata.combining(c))


def _forme(texte: str) -> str:
    """Forme de comparaison d'un statut : sans casse, sans accents, espaces réduits."""
    return _sans_accents(" ".join(_texte(texte).casefold().split()))


def _jj(jour: date | None) -> str:
    return jour.strftime("%d/%m/%Y") if jour is not None else ""


def _lisible(valeur: Any) -> str:
    """Valeur d'une cellule telle qu'un message la cite."""
    if isinstance(valeur, datetime):
        return _jj(valeur.date()) if valeur.time() == time() else valeur.strftime("%d/%m/%Y %H:%M")
    if isinstance(valeur, date):
        return _jj(valeur)
    return _texte(valeur)


def _euros(montant: Decimal) -> str:
    """Montant à la française : « 1 234,56 € »."""
    arrondi = heures.arrondi_centime(montant)
    entier, _, decimales = f"{abs(arrondi):.2f}".partition(".")
    milliers = f"{int(entier):,}".replace(",", " ")
    return f"{'-' if arrondi < 0 else ''}{milliers},{decimales} €"


def _est_formule(cellule: Any) -> bool:
    valeur = cellule.value
    return (
        cellule.data_type == "f"
        or isinstance(valeur, (ArrayFormula, DataTableFormula))
        or (isinstance(valeur, str) and valeur.startswith("="))
    )


# ──────────────────────────────────────────────────────────────────────
# Lecture des cellules
# ──────────────────────────────────────────────────────────────────────


def _lire_date(valeur: Any, colonne: str) -> date | None:
    if _est_vide(valeur):
        return None
    if isinstance(valeur, datetime):
        return valeur.date()
    if isinstance(valeur, date):
        return valeur
    if isinstance(valeur, str):
        texte = valeur.strip()
        for motif, forme in ((_DATE_JJ_MM_AAAA, "%d/%m/%Y"), (_DATE_AAAA_MM_JJ, "%Y-%m-%d")):
            if motif.fullmatch(texte):
                try:
                    return datetime.strptime(texte, forme).date()
                except ValueError:
                    break
    raise _Illisible(f"colonne « {colonne} » : « {_lisible(valeur)} » n'est pas une date (JJ/MM/AAAA)")


def _lire_montant(valeur: Any, colonne: str) -> Decimal | None:
    if _est_vide(valeur):
        return None
    try:
        if isinstance(valeur, int) and not isinstance(valeur, bool):
            return Decimal(valeur)
        if isinstance(valeur, float) and math.isfinite(valeur):
            return Decimal(repr(valeur))
        if isinstance(valeur, Decimal) and valeur.is_finite():
            return valeur
        if isinstance(valeur, str) and _MONTANT_RE.fullmatch(valeur.strip()):
            texte = "".join(c for c in valeur.strip() if c not in _ESPACES_MILLIERS)
            return Decimal(texte.replace(",", "."))
    except InvalidOperation:
        pass
    raise _Illisible(f"colonne « {colonne} » : « {_lisible(valeur)} » n'est pas un montant (ex. 1 234,56)")


_STATUTS_PAR_FORME = {_forme(statut): statut for statut in STATUTS_SAISIS}


def _lire_statut(valeur: Any, colonne: str = COL_STATUT) -> str:
    """Statut saisi, écrit comme dans STATUTS_SAISIS ; "" pour une cellule vide ou « échue »."""
    if _est_vide(valeur):
        return ""
    forme = _forme(_texte(valeur))
    if forme == _forme(ECHUE):
        return ""
    if forme in _STATUTS_PAR_FORME:
        return _STATUTS_PAR_FORME[forme]
    raise _Illisible(
        f"colonne « {colonne} » : « {_lisible(valeur)} » n'est pas un statut du registre "
        f"(attendu : {', '.join(STATUTS_SAISIS)}, ou vide)"
    )


def _lire_ecole(valeur: Any, ecoles: Sequence[Mapping[str, Any]]) -> str:
    try:
        return code_ecole(valeur, ecoles) or ""
    except ValueError as exc:
        raise _Illisible(f"colonne « {COL_ECOLE} » : {exc}") from None


def _ligne_lue(rang: int, brute: Mapping[str, Any], ecoles: Sequence[Mapping[str, Any]]) -> LigneRegistre:
    """Une ligne de `Factures` ; `_Illisible` pour la première cellule qui ne se lit pas."""
    n_facture = _texte(brute.get(COL_NUMERO))
    if not n_facture:
        raise _Illisible(f"colonne « {COL_NUMERO} » : vide, alors que la ligne est remplie (une ligne par facture)")
    return LigneRegistre(
        rang=rang,
        n_facture=n_facture,
        ecole=_lire_ecole(brute.get(COL_ECOLE), ecoles),
        contrat=_texte(brute.get(COL_CONTRAT)) or None,
        periode=_texte(brute.get(COL_PERIODE)),
        date_emission=_lire_date(brute.get(COL_EMISSION), COL_EMISSION),
        echeance=_lire_date(brute.get(COL_ECHEANCE), COL_ECHEANCE),
        total_ht=_lire_montant(brute.get(COL_TOTAL), COL_TOTAL),
        date_envoi=_lire_date(brute.get(COL_ENVOI), COL_ENVOI),
        date_paiement=_lire_date(brute.get(COL_PAIEMENT), COL_PAIEMENT),
        statut_saisi=_lire_statut(brute.get(COL_STATUT)),
        note=_texte(brute.get(COL_NOTE)),
    )


def lire_registre(master: Master) -> list[LigneRegistre]:
    """Lignes de l'onglet `Factures`, dans l'ordre de la feuille ; onglet absent ou vide → [].

    Dates : `date`, `datetime`, texte JJ/MM/AAAA ou AAAA-MM-JJ, rendues en `date`. Montants : un
    nombre, ou un texte « 1 234,56 », rendus en `Decimal`. Statut : l'un des `STATUTS_SAISIS` ;
    « échue », toléré, se lit comme un statut vide.

    Raises:
        RegistreInvalide: cellule illisible (« onglet « Factures », ligne 7, colonne « Date de
            paiement » : « fin nov » n'est pas une date (JJ/MM/AAAA) »), statut inconnu, ligne
            remplie sans N°, ou N° en double (sans tenir compte de la casse).
    """
    brutes = list(master.factures)
    rangs = list((getattr(master, "rangs", None) or {}).get("factures") or [])
    if len(rangs) != len(brutes):  # Master construit à la main, sans rangs
        rangs = list(range(2, len(brutes) + 2))
    lignes: list[LigneRegistre] = []
    vus: dict[str, int] = {}
    for rang, brute in zip(rangs, brutes):
        try:
            ligne = _ligne_lue(rang, brute, master.ecoles)
        except _Illisible as exc:
            raise RegistreInvalide(f"onglet « {ONGLET} », ligne {rang}, {exc}.") from None
        cle = _cle(ligne.n_facture)
        if cle in vus:
            raise RegistreInvalide(
                f"onglet « {ONGLET} » : N° {cle} en double (lignes {vus[cle]} et {rang}). Gardez une seule "
                "ligne par facture."
            )
        vus[cle] = rang
        lignes.append(ligne)
    return lignes


# ──────────────────────────────────────────────────────────────────────
# Statut affiché, retard, échues
# ──────────────────────────────────────────────────────────────────────


def statut_affiche(ligne: LigneRegistre, *, aujourd_hui: date) -> str:
    """Statut à afficher, dans cet ordre : abandonnée ; date de paiement → payée ; « en émission »
    (jamais échue) ; échéance antérieure à aujourd'hui → échue (le jour de l'échéance n'est pas
    échu : les pénalités courent dès le lendemain) ; date d'envoi → envoyée ; date d'émission ou
    statut « émise » → émise ; sinon à compléter."""
    if ligne.statut_saisi == ABANDONNEE:
        return ABANDONNEE
    if ligne.date_paiement is not None:
        return PAYEE
    if ligne.statut_saisi == EN_EMISSION:
        return EN_EMISSION
    if ligne.echeance is not None and ligne.echeance < aujourd_hui:
        return ECHUE
    if ligne.date_envoi is not None:
        return ENVOYEE
    if ligne.date_emission is not None or ligne.statut_saisi == EMISE:
        return EMISE
    return A_COMPLETER


def retard_jours(ligne: LigneRegistre, *, aujourd_hui: date) -> int | None:
    """Jours écoulés depuis l'échéance d'une facture échue (1 le lendemain) ; None si elle ne
    l'est pas (payée, abandonnée, en émission, sans échéance, ou au jour de l'échéance)."""
    if statut_affiche(ligne, aujourd_hui=aujourd_hui) != ECHUE:
        return None
    return (aujourd_hui - ligne.echeance).days  # type: ignore[operator]  # échue ⇒ échéance présente


def echues(lignes: Sequence[LigneRegistre], *, aujourd_hui: date) -> list[LigneRegistre]:
    """Factures échues non payées, triées par échéance, puis par N°."""
    return sorted(
        (ligne for ligne in lignes if statut_affiche(ligne, aujourd_hui=aujourd_hui) == ECHUE),
        key=lambda ligne: (ligne.echeance, _cle(ligne.n_facture)),
    )


# ──────────────────────────────────────────────────────────────────────
# Anomalies et trous
# ──────────────────────────────────────────────────────────────────────


def _total_coherent(total: Decimal, seances: Sequence[LigneSessions]) -> bool:
    """Le total du registre égale-t-il, au centime, celui que donnent les lignes `Sessions` ?

    Deux calculs faits de la même façon que la facture (§ 1.2) : la somme des `Total HT` des
    lignes, ou, quand chaque ligne porte l'arrondi de ses heures × taux, l'arrondi de la somme
    exacte des heures × taux (le total de la facture n'arrondit qu'une fois).
    """
    attendu = heures.arrondi_centime(total)
    somme = sum((s.total_ht for s in seances if s.total_ht is not None), Decimal(0))
    if attendu == heures.arrondi_centime(somme):
        return True
    if any(s.total_ht is None or s.duree_h is None or s.taux is None for s in seances):
        return False
    exacts = [heures.montant(s.duree_h, s.taux) for s in seances]  # type: ignore[arg-type]
    if any(heures.arrondi_centime(e) != heures.arrondi_centime(s.total_ht)  # type: ignore[arg-type]
           for e, s in zip(exacts, seances)):
        return False
    return attendu == heures.arrondi_centime(sum(exacts, Decimal(0)))


def _anomalies_de_la_ligne(ligne: LigneRegistre, seances: Sequence[LigneSessions]) -> list[str]:
    if ligne.statut_saisi == ABANDONNEE:
        return []
    n = ligne.n_facture
    trouvees: list[str] = []
    emission = ligne.date_emission
    if emission is not None:
        for valeur, libelle in ((ligne.date_envoi, "date d'envoi"), (ligne.date_paiement, "date de paiement"),
                                (ligne.echeance, "échéance")):
            if valeur is not None and valeur < emission:
                trouvees.append(f"{n} : {libelle} ({_jj(valeur)}) antérieure à la date d'émission ({_jj(emission)}).")
    if ligne.date_paiement is not None and ligne.date_envoi is None:
        trouvees.append(
            f"{n} : paiement saisi ({_jj(ligne.date_paiement)}) sans date d'envoi (avertissement : notez la date "
            "d'envoi)."
        )
    if ligne.statut_saisi == PAYEE and ligne.date_paiement is None:
        trouvees.append(f"{n} : statut « {PAYEE} » sans date de paiement (colonne « {COL_PAIEMENT} »).")
    if ligne.statut_saisi == ENVOYEE and ligne.date_envoi is None:
        trouvees.append(f"{n} : statut « {ENVOYEE} » sans date d'envoi (colonne « {COL_ENVOI} »).")
    if not seances:
        trouvees.append(
            f"{n} : aucune ligne dans l'onglet « Sessions » pour ce N° : le Total HT ne peut pas être "
            "rapproché des séances facturées (ADR003)."
        )
    elif ligne.total_ht is None:
        somme = sum((s.total_ht for s in seances if s.total_ht is not None), Decimal(0))
        trouvees.append(f"{n} : Total HT vide, alors que les lignes Sessions de ce N° totalisent {_euros(somme)}.")
    elif not _total_coherent(ligne.total_ht, seances):
        somme = sum((s.total_ht for s in seances if s.total_ht is not None), Decimal(0))
        trouvees.append(
            f"{n} : Total HT {_euros(ligne.total_ht)} différent de la somme des lignes Sessions de ce N° "
            f"({_euros(somme)})."
        )
    return trouvees


def anomalies(lignes: Sequence[LigneRegistre], sessions: Sequence[LigneSessions]) -> list[str]:
    """Incohérences du registre, une phrase par incohérence, qui nomme le N° :

    - N° présent dans `Sessions` mais absent du registre (avec la proposition
      `registre initialiser --reprendre-historique`) ;
    - envoi, paiement ou échéance antérieurs à l'émission ;
    - paiement sans envoi (avertissement) ;
    - `Total HT` différent de celui des lignes `Sessions` du N° (au centime, voir
      `_total_coherent`), ou facture sans ligne `Sessions` ;
    - statut saisi incohérent avec les dates (« payée » sans paiement, « envoyée » sans envoi).

    Une ligne abandonnée n'est pas contrôlée. Ordre : celui du registre, puis les N° absents.
    """
    par_numero: dict[str, list[LigneSessions]] = {}
    for seance in sessions:
        cle = _cle(seance.n_facture)
        if cle:
            par_numero.setdefault(cle, []).append(seance)
    trouvees: list[str] = []
    au_registre: set[str] = set()
    for ligne in lignes:
        cle = _cle(ligne.n_facture)
        au_registre.add(cle)
        trouvees.extend(_anomalies_de_la_ligne(ligne, par_numero.get(cle, [])))
    for cle in sorted(par_numero.keys() - au_registre):
        trouvees.append(
            f"{cle} : présent dans l'onglet « Sessions » mais absent du registre ; reprenez-le avec "
            f"`facturation {COMMANDE_REPRISE}`."
        )
    return trouvees


def trous_sans_raison(lignes: Sequence[LigneRegistre], numeros: Iterable[str], *, prefixe: str,
                      ecoles: Iterable[str]) -> list[str]:
    """Numéros manquants de chaque école (`numerotation.numeros_manquants`), moins ceux qui ont une
    ligne au registre, où leur raison est notée (ADR029). Une école donnée deux fois ne compte
    qu'une fois ; une valeur vide est ignorée.

    Raises:
        ValueError: préfixe ou code école mal formé.
    """
    numeros = list(numeros)
    au_registre = {_cle(ligne.n_facture) for ligne in lignes}
    trous: list[str] = []
    vues: set[str] = set()
    for ecole in ecoles:
        code = _texte(ecole).upper()
        if not code or code in vues:
            continue
        vues.add(code)
        trous.extend(n for n in numeros_manquants(numeros, prefixe=prefixe, code_ecole=code) if n not in au_registre)
    return trous


# ──────────────────────────────────────────────────────────────────────
# Écriture sur une feuille ouverte (sans verrou : l'appelant le tient)
# ──────────────────────────────────────────────────────────────────────


def _date_donnee(valeur: Any, nom: str) -> date:
    if isinstance(valeur, datetime):
        return valeur.date()
    if isinstance(valeur, date):
        return valeur
    raise RegistreInvalide(f"{nom} : une date est attendue, reçu {valeur!r}.")


def ligne_pour_emission(*, n_facture: str, ecole: str, contrat: str | None, periode: str,
                        date_emission: date, echeance: date, total_ht: Decimal) -> dict[str, Any]:
    """Ligne du registre d'une facture qu'on émet (F-17) : les 11 colonnes, statut « en émission »,
    dates d'envoi et de paiement vides, total arrondi au centime.

    Raises:
        RegistreInvalide: N° vide, date qui n'en est pas une, échéance antérieure à l'émission.
        TypeError: total qui n'est pas un `Decimal` (un `float` compris).
    """
    numero = _texte(n_facture)
    if not numero:
        raise RegistreInvalide("Ligne du registre sans N° de facture.")
    emission = _date_donnee(date_emission, "date_emission")
    limite = _date_donnee(echeance, "echeance")
    if limite < emission:
        raise RegistreInvalide(
            f"N° {numero} : échéance {_jj(limite)} antérieure à la date d'émission {_jj(emission)}."
        )
    return {
        COL_NUMERO: numero,
        COL_ECOLE: _texte(ecole) or None,
        COL_CONTRAT: _texte(contrat) or None,
        COL_PERIODE: _texte(periode) or None,
        COL_EMISSION: emission,
        COL_ECHEANCE: limite,
        COL_TOTAL: heures.arrondi_centime(total_ht),
        COL_ENVOI: None,
        COL_PAIEMENT: None,
        COL_STATUT: EN_EMISSION,
        COL_NOTE: None,
    }


def creer_onglet(wb: Workbook) -> Worksheet:
    """Ajoute l'onglet `Factures` en fin de classeur, avec ses 11 en-têtes en ligne 1 (volets figés
    en A2, comme les autres onglets du MASTER) ; n'enregistre rien.

    Raises:
        RegistreInvalide: l'onglet existe déjà.
    """
    if ONGLET in wb.sheetnames:
        raise RegistreInvalide(f"L'onglet « {ONGLET} » existe déjà dans le classeur.")
    ws = wb.create_sheet(ONGLET)
    ws.append(list(COLONNES))
    ws.freeze_panes = "A2"
    return ws


def _index(ws: Worksheet) -> dict[str, int]:
    return index_colonnes(ws, COLONNES, onglet=ONGLET)


def _rangs_du_numero(ws: Worksheet, index: Mapping[str, int], numero: str) -> list[int]:
    """Lignes Excel qui portent ce N° (sans tenir compte de la casse ni des espaces)."""
    cle = _cle(numero)
    colonne = index[COL_NUMERO]
    return [
        rang for rang in range(2, derniere_ligne_remplie(ws) + 1)
        if _cle(ws.cell(row=rang, column=colonne).value) == cle
    ]


def _texte_a_ecrire(valeur: Any, colonne: str) -> str | None:
    texte = _texte(valeur)
    if texte.startswith("="):
        raise RegistreInvalide(
            f"Colonne « {colonne} » : « {texte} » commence par « = » : l'outil n'écrit jamais de formule."
        )
    return texte or None


def _valeur_a_ecrire(colonne: str, valeur: Any) -> Any:
    """Valeur d'une colonne prête à écrire ; None pour une cellule laissée vide."""
    if _est_vide(valeur):
        return None
    if colonne in COLONNES_DATES:
        return _date_donnee(valeur, f"Colonne « {colonne} »")
    if colonne == COL_TOTAL:
        if isinstance(valeur, str):
            raise RegistreInvalide(f"Colonne « {COL_TOTAL} » : un montant (Decimal) est attendu, reçu {valeur!r}.")
        return heures.arrondi_centime(valeur)
    if colonne == COL_STATUT:
        try:
            statut = _lire_statut(valeur)
        except _Illisible as exc:
            raise RegistreInvalide(f"{exc}.") from None
        if not statut:
            raise RegistreInvalide(f"Statut « {valeur} » : « {ECHUE} » est calculé à l'affichage, jamais écrit.")
        return statut
    return _texte_a_ecrire(valeur, colonne)


def _ecrire_cellule(cellule: Any, valeur: Any) -> Any:
    """Écrit une valeur (jamais une formule) ; rend la valeur telle qu'enregistrée."""
    if isinstance(valeur, Decimal):
        valeur = float(valeur)
    cellule.value = valeur
    if isinstance(valeur, date):
        cellule.number_format = FORMAT_DATE
    return valeur


def _refuser_formule(ws: Worksheet, cellule: Any, numero: str) -> None:
    if _est_formule(cellule):
        colonne = ws.cell(row=1, column=cellule.column).value
        raise RegistreInvalide(
            f"N° {numero} : la cellule {cellule.coordinate} (colonne « {colonne} ») de l'onglet « {ONGLET} » "
            "contient une formule : l'outil ne l'écrase jamais. Remplacez-la par une valeur dans Excel, puis "
            "relancez ; rien n'a été modifié."
        )


def _ajouter(ws: Worksheet, ligne: Mapping[str, Any]) -> tuple[int, dict[str, Any]]:
    """`ajouter_ligne`, qui rend aussi {coordonnée : valeur écrite}."""
    inconnues = [colonne for colonne in ligne if colonne not in COLONNES]
    if inconnues:
        raise RegistreInvalide(
            f"Colonne(s) inconnue(s) du registre : {', '.join(map(str, inconnues))} (colonnes : {', '.join(COLONNES)})."
        )
    valeurs = {colonne: _valeur_a_ecrire(colonne, ligne.get(colonne)) for colonne in COLONNES}
    numero = valeurs[COL_NUMERO]
    if not numero:
        raise RegistreInvalide("Ligne du registre sans N° de facture.")
    index = _index(ws)
    existants = _rangs_du_numero(ws, index, numero)
    if existants:
        raise RegistreInvalide(
            f"N° {_cle(numero)} déjà au registre (onglet « {ONGLET} », ligne {existants[0]}) : une seule ligne par "
            "facture ; rien n'a été ajouté."
        )
    rang = derniere_ligne_remplie(ws) + 1
    cibles = {colonne: ws.cell(row=rang, column=index[colonne]) for colonne, v in valeurs.items() if v is not None}
    for cellule in cibles.values():
        _refuser_formule(ws, cellule, numero)
    ecrites = {cellule.coordinate: _ecrire_cellule(cellule, valeurs[colonne]) for colonne, cellule in cibles.items()}
    return rang, ecrites


def ajouter_ligne(ws: Worksheet, ligne: Mapping[str, Any]) -> int:
    """Ajoute une ligne au registre, par nom de colonne, sous la dernière ligne remplie.

    `ligne` est un dictionnaire {colonne de `COLONNES` : valeur} (par exemple celui de
    `ligne_pour_emission`) ; une colonne absente ou vide laisse la cellule vide. Les dates
    prennent le format JJ/MM/AAAA, les montants sont écrits en nombre. Ne prend aucun verrou et
    n'enregistre rien.

    Returns:
        Le numéro de la ligne écrite.

    Raises:
        RegistreInvalide: N° vide ou déjà au registre (sans tenir compte de la casse), colonne
            inconnue, valeur qui ne convient pas à sa colonne (texte en « = », statut calculé),
            cellule cible à formule ; la feuille n'est pas modifiée.
        master.MasterInvalide: colonne du registre absente ou en double.
    """
    return _ajouter(ws, ligne)[0]


def _date_du_champ(valeur: Any, champ: str) -> date:
    if isinstance(valeur, datetime):
        return valeur.date()
    if isinstance(valeur, date):
        return valeur
    raise ValueError(f"Champ « {champ} » : une date est attendue, reçu {valeur!r}.")


def _champs_valides(champs: Mapping[str, Any]) -> dict[str, Any]:
    """Champs de `modifier_ligne` contrôlés et convertis ; `ValueError` qui nomme le champ fautif."""
    inconnus = sorted(set(champs) - set(CHAMPS_MODIFIABLES))
    if inconnus:
        pluriel = "s" if len(inconnus) > 1 else ""
        raise ValueError(
            f"Champ{pluriel} inconnu{pluriel} : {', '.join(inconnus)}. Champs modifiables du registre : "
            f"{', '.join(CHAMPS_MODIFIABLES)}."
        )
    if not champs:
        raise ValueError(f"Aucun champ à modifier (champs du registre : {', '.join(CHAMPS_MODIFIABLES)}).")
    valides: dict[str, Any] = {}
    for champ, valeur in champs.items():
        if champ in ("date_envoi", "date_paiement"):
            valides[champ] = _date_du_champ(valeur, champ)
        elif champ == "statut":
            statut = _STATUTS_PAR_FORME.get(_forme(valeur)) if isinstance(valeur, str) else None
            if statut not in STATUTS_DE_L_EMISSION:
                raise ValueError(
                    f"Statut « {valeur} » refusé : seuls « {EMISE} » et « {ABANDONNEE} » s'écrivent par le champ statut "
                    f"(émission, F-17) ; « {ENVOYEE} » et « {PAYEE} » suivent les dates d'envoi et de paiement, "
                    f"« {ECHUE} » est calculé à l'affichage, jamais écrit."
                )
            valides[champ] = statut
        else:  # note
            if not isinstance(valeur, str) or not valeur.strip():
                raise ValueError(f"Champ « note » : un texte non vide est attendu, reçu {valeur!r}.")
            valides[champ] = _texte_a_ecrire(valeur, COL_NOTE)
    if "statut" in valides and ({"date_envoi", "date_paiement"} & valides.keys()):
        raise ValueError("Le champ statut ne se combine pas avec une date d'envoi ou de paiement : un appel pour chacun.")
    return valides


def _modifier(ws: Worksheet, n_facture: str, *, remplacer: bool, champs: Mapping[str, Any]
              ) -> tuple[int, str, dict[str, tuple[Any, Any]], dict[str, Any]]:
    """`modifier_ligne`, qui rend aussi la ligne, le N° écrit et {coordonnée : valeur écrite}."""
    valides = _champs_valides(champs)
    index = _index(ws)
    numero = _texte(n_facture)
    rangs = _rangs_du_numero(ws, index, numero)
    if not rangs:
        raise RegistreInvalide(f"N° {_cle(numero) or '(vide)'} absent du registre (onglet « {ONGLET} »).")
    if len(rangs) > 1:
        raise RegistreInvalide(
            f"N° {_cle(numero)} ambigu : il figure aux lignes {', '.join(map(str, rangs))} de l'onglet « {ONGLET} ». "
            "Gardez une seule ligne par facture, puis relancez."
        )
    rang = rangs[0]

    def cellule(colonne: str) -> Any:
        return ws.cell(row=rang, column=index[colonne])

    numero = _texte(cellule(COL_NUMERO).value)
    for champ in valides:
        _refuser_formule(ws, cellule(CHAMPS_MODIFIABLES[champ]), numero)
    touche_statut = bool({"statut", "date_envoi", "date_paiement"} & valides.keys())
    if touche_statut:
        _refuser_formule(ws, cellule(COL_STATUT), numero)

    statut_brut = cellule(COL_STATUT).value
    try:
        statut = _lire_statut(statut_brut)
        emission = _lire_date(cellule(COL_EMISSION).value, COL_EMISSION)
    except _Illisible as exc:
        raise RegistreInvalide(f"onglet « {ONGLET} », ligne {rang}, {exc}.") from None

    if statut == ABANDONNEE and set(valides) - {"note"}:
        raise RegistreInvalide(f"N° {numero} : facture abandonnée ; seule sa note peut encore changer.")
    if statut == EN_EMISSION and {"date_envoi", "date_paiement"} & valides.keys():
        raise RegistreInvalide(
            f"N° {numero} : facture encore « {EN_EMISSION} » (son contrôle après écriture n'est pas vert) : "
            "terminez l'émission avant de noter son envoi ou son paiement."
        )

    changements: dict[str, tuple[Any, Any]] = {}
    for champ, colonne in (("date_envoi", COL_ENVOI), ("date_paiement", COL_PAIEMENT)):
        if champ not in valides:
            continue
        nouvelle = valides[champ]
        libelle = _LIBELLES_DATES[colonne]
        if emission is not None and nouvelle < emission:
            raise RegistreInvalide(
                f"N° {numero} : {libelle} {_jj(nouvelle)} antérieure à la date d'émission {_jj(emission)} ; "
                "rien n'a été modifié."
            )
        brute = cellule(colonne).value
        try:
            actuelle = _lire_date(brute, colonne)
        except _Illisible as exc:
            if not remplacer:
                raise RegistreInvalide(
                    f"onglet « {ONGLET} », ligne {rang}, {exc} : ajoutez --remplacer pour la remplacer."
                ) from None
            actuelle = None
        if actuelle == nouvelle:
            continue
        if not _est_vide(brute) and not remplacer:
            raise RegistreInvalide(
                f"N° {numero} : {libelle} déjà saisie ({_lisible(brute)}) : ajoutez --remplacer pour la remplacer "
                f"par le {_jj(nouvelle)} ; rien n'a été modifié."
            )
        changements[colonne] = (brute, nouvelle)

    if "note" in valides:
        brute = cellule(COL_NOTE).value
        actuelle = _texte(brute)
        nouvelle_note = valides["note"]
        if actuelle != nouvelle_note:
            if actuelle and not remplacer:
                raise RegistreInvalide(
                    f"N° {numero} : note déjà saisie (« {actuelle} ») : ajoutez --remplacer pour la remplacer ; "
                    "rien n'a été modifié."
                )
            changements[COL_NOTE] = (brute, nouvelle_note)

    nouveau = statut
    if "statut" in valides:
        demande = valides["statut"]
        if demande == EMISE and statut in (EN_EMISSION, "", EMISE):
            nouveau = EMISE
        elif demande == ABANDONNEE and statut in (EN_EMISSION, EMISE, ""):
            nouveau = ABANDONNEE
        else:
            raise RegistreInvalide(
                f"N° {numero} : le statut « {statut} » ne passe pas à « {demande} » (une facture envoyée ou payée "
                "ne s'abandonne pas : faites un avoir)."
            )
    if "date_paiement" in valides:
        nouveau = PAYEE
    elif "date_envoi" in valides and nouveau in (EMISE, ""):
        nouveau = ENVOYEE
    if nouveau != statut:
        changements[COL_STATUT] = (statut_brut, nouveau)

    ecrites = {
        cellule(colonne).coordinate: _ecrire_cellule(cellule(colonne), apres)
        for colonne, (_, apres) in changements.items()
    }
    return rang, numero, changements, ecrites


def modifier_ligne(ws: Worksheet, n_facture: str, *, remplacer: bool = False, **champs: Any
                   ) -> dict[str, tuple[Any, Any]]:
    """Modifie la ligne d'une facture du registre. Ne prend aucun verrou et n'enregistre rien.

    Champs admis : `date_envoi`, `date_paiement`, `statut` (réservé à l'émission, F-17 : « émise »
    ou « abandonnée ») et `note`. Une date d'envoi fait passer le statut « émise » à « envoyée » ;
    une date de paiement fait passer le statut à « payée ». « Échue » n'est jamais écrit. Une
    date égale à celle déjà saisie ne change rien.

    Returns:
        {colonne : (valeur avant, valeur après)} des seules cellules changées.

    Raises:
        ValueError: champ inconnu ou de mauvais type, statut autre que « émise » ou « abandonnée »,
            statut combiné à une date.
        RegistreInvalide: N° absent ou ambigu, ligne abandonnée (sauf la note), facture encore
            « en émission » (envoi, paiement), date antérieure à l'émission, cellule déjà remplie
            d'une autre valeur sans `remplacer=True`, cellule cible à formule (même avec
            `remplacer=True`), passage de statut impossible. Sur refus, la feuille n'est pas
            modifiée.
        master.MasterInvalide: colonne du registre absente ou en double.
    """
    return _modifier(ws, n_facture, remplacer=remplacer, champs=champs)[2]


def noter_numero_manquant(ws: Worksheet, n_facture: str, *, ecole: str, raison: str) -> int:
    """Ajoute la ligne d'un numéro manquant (ADR029, Q-G2) : le N°, l'école, le statut
    « abandonnée » et la raison en note, sans date ni montant. `numerotation.lire_numeros` voit
    cette ligne : le numéro n'est jamais réutilisé. Ne prend aucun verrou et n'enregistre rien.

    Returns:
        Le numéro de la ligne écrite.

    Raises:
        RegistreInvalide: raison vide, N° vide ou déjà au registre.
    """
    if not isinstance(raison, str) or not raison.strip():
        raise RegistreInvalide(f"N° {_texte(n_facture)} : la raison du numéro manquant est obligatoire (ADR029).")
    return ajouter_ligne(ws, {COL_NUMERO: n_facture, COL_ECOLE: ecole, COL_STATUT: ABANDONNEE, COL_NOTE: raison})


# ──────────────────────────────────────────────────────────────────────
# Reprise de l'historique
# ──────────────────────────────────────────────────────────────────────


def lignes_historiques(master: Master) -> list[dict[str, Any]]:
    """Une ligne proposée par N° de `Sessions` absent du registre, triées par N° : école (code),
    `Période_facture`, Σ `Total HT` exacte (arrondie une fois au centime), statut « émise », note
    « repris de Sessions le JJ/MM/AAAA : dates à compléter ». La date d'émission, l'échéance et les
    dates d'envoi et de paiement restent vides : elles ne sont jamais inventées.

    Returns:
        Des dictionnaires {colonne de `COLONNES` : valeur}, à passer à `ajouter_ligne`.

    Raises:
        RegistreInvalide: ligne illisible dans `Sessions` ou dans le registre ; N° dont les
            lignes `Sessions` portent deux écoles.
    """
    seances, illisibles = lire_sessions(master)
    if illisibles:
        detail = " ; ".join(message for _, message in illisibles[:3])
        reste = f" (et {len(illisibles) - 3} autre(s))" if len(illisibles) > 3 else ""
        raise RegistreInvalide(
            f"Reprise impossible : l'onglet « Sessions » a des lignes illisibles{reste} : {detail}. Corrigez-les, "
            "puis relancez : un total repris doit compter toutes les séances."
        )
    au_registre = {_cle(ligne.n_facture) for ligne in lire_registre(master)}
    groupes: dict[str, list[LigneSessions]] = {}
    for seance in seances:
        cle = _cle(seance.n_facture)
        if cle and cle not in au_registre:
            groupes.setdefault(cle, []).append(seance)
    note = f"repris de Sessions le {_jj(date.today())} : dates à compléter"
    proposees: list[dict[str, Any]] = []
    for cle in sorted(groupes):
        lignes = groupes[cle]
        ecoles = list(dict.fromkeys(s.ecole for s in lignes if s.ecole))
        if len(ecoles) > 1:
            rangs = ", ".join(str(s.rang) for s in lignes)
            raise RegistreInvalide(
                f"N° {cle} : ses lignes de l'onglet « Sessions » (lignes {rangs}) portent plusieurs écoles "
                f"({', '.join(ecoles)}). Corrigez Sessions avant de reprendre l'historique."
            )
        periodes = list(dict.fromkeys(s.periode for s in lignes if s.periode))
        if any(s.total_ht is None for s in lignes):
            total = None
        else:
            total = heures.arrondi_centime(sum((s.total_ht for s in lignes), Decimal(0)))  # type: ignore[misc]
        proposees.append({
            COL_NUMERO: cle,
            COL_ECOLE: ecoles[0] if ecoles else None,
            COL_CONTRAT: None,
            COL_PERIODE: ", ".join(periodes) or None,
            COL_EMISSION: None,
            COL_ECHEANCE: None,
            COL_TOTAL: total,
            COL_ENVOI: None,
            COL_PAIEMENT: None,
            COL_STATUT: EMISE,
            COL_NOTE: note,
        })
    return proposees


# ──────────────────────────────────────────────────────────────────────
# Écritures de premier niveau
# ──────────────────────────────────────────────────────────────────────


@dataclass
class _Resultat:
    """Ce qu'une opération a fait sur le classeur ouvert."""

    cellules: dict[tuple[str, str], Any] = field(default_factory=dict)   # (onglet, coordonnée) → valeur écrite
    evenements: list[tuple[str | None, Any, Any]] = field(default_factory=list)   # (N°, avant, après)
    onglets_crees: list[str] = field(default_factory=list)
    apercu: Any = None
    rien: str = "rien à écrire"


Operation = Callable[["Workbook"], _Resultat]


def _confirmation(confirme: bool) -> None:
    if confirme is not True:
        raise ValueError("validation explicite requise")


def _charger(chemin: Path) -> Workbook:
    """Le MASTER chargé en entier, formules comprises (`data_only=False`)."""
    if not chemin.is_file():
        raise FileNotFoundError(f"MASTER introuvable : {chemin}")
    try:
        return load_workbook(chemin, data_only=False)
    except PermissionError as exc:
        raise MasterVerrouille(
            f"Le MASTER {chemin.name} ne s'ouvre pas (accès refusé) : fermez le classeur s'il est ouvert dans "
            "Excel, ou vérifiez vos droits sur le fichier, puis relancez."
        ) from exc
    except _ERREURS_OUVERTURE as exc:
        raise MasterInvalide(
            f"MASTER {chemin.name} : classeur abîmé ou illisible ({type(exc).__name__} : {exc})."
        ) from exc


def _onglet(wb: Workbook) -> Worksheet:
    if ONGLET not in wb.sheetnames:
        raise RegistreInvalide(
            f"Onglet « {ONGLET} » absent du MASTER : ajoutez l'onglet {ONGLET} avec `{COMMANDE_INITIALISER}` "
            "(--reprendre-historique pour y reprendre les factures de Sessions)."
        )
    return wb[ONGLET]


def _a_blanc(profil: Profil, operation: Operation) -> _Resultat:
    """Joue l'opération sur le MASTER chargé en mémoire, sans rien enregistrer."""
    wb = _charger(profil.master)
    try:
        return operation(wb)
    finally:
        wb.close()


def _comparable(valeur: Any) -> Any:
    """Valeur d'une cellule comparable d'un chargement à l'autre (dates à la seconde près)."""
    if isinstance(valeur, datetime):
        return (valeur + timedelta(microseconds=500_000)).replace(microsecond=0, tzinfo=None)
    if isinstance(valeur, date):
        return datetime.combine(valeur, time())
    if isinstance(valeur, time):
        return _comparable(datetime.combine(date(2000, 1, 1), valeur)).time()
    if isinstance(valeur, Decimal):
        return float(valeur)
    if isinstance(valeur, ArrayFormula):
        return ("formule matricielle", str(valeur.ref), str(valeur.text))
    if isinstance(valeur, DataTableFormula):
        return ("table de données", tuple(sorted((cle, str(v)) for cle, v in valeur)))
    return valeur


def _instantane(wb: Workbook) -> dict[tuple[str, str], Any]:
    """Toutes les cellules non vides du classeur, sans en créer de nouvelles."""
    resultat: dict[tuple[str, str], Any] = {}
    for ws in wb.worksheets:
        cellules = getattr(ws, "_cells", None)
        iterable = list(cellules.values()) if cellules is not None else (c for ligne in ws.iter_rows() for c in ligne)
        for c in iterable:
            if c.value is not None:
                resultat[(ws.title, c.coordinate)] = _comparable(c.value)
    return resultat


def _verifier_relecture(chemin: Path, avant: dict[tuple[str, str], Any], onglets_avant: list[str],
                        resultat: _Resultat, sauvegarde: Path) -> None:
    """Relit le MASTER enregistré : mêmes onglets (plus ceux créés), et seules les cellules annoncées
    ont changé, avec la valeur annoncée."""
    wb = load_workbook(chemin, data_only=False)
    try:
        onglets_apres = list(wb.sheetnames)
        apres = _instantane(wb)
    finally:
        wb.close()
    ecarts: list[str] = []
    attendus = onglets_avant + [o for o in resultat.onglets_crees if o not in onglets_avant]
    if onglets_apres != attendus:
        ecarts.append(f"onglets {', '.join(onglets_apres)} au lieu de {', '.join(attendus)}")
    annoncees = {cle: _comparable(valeur) for cle, valeur in resultat.cellules.items()}
    changees = {cle for cle in avant.keys() | apres.keys() if avant.get(cle) != apres.get(cle)}
    imprevues = sorted(changees - annoncees.keys())
    if imprevues:
        ecarts.append("cellule(s) changée(s) sans être annoncée(s) : " + ", ".join(f"{o}!{c}" for o, c in imprevues[:5]))
    fausses = sorted(cle for cle, valeur in annoncees.items() if apres.get(cle) != valeur)
    if fausses:
        ecarts.append("cellule(s) relue(s) avec une autre valeur que celle écrite : "
                      + ", ".join(f"{o}!{c}" for o, c in fausses[:5]))
    if ecarts:
        raise EcritureNonConforme(
            f"Contrôle après écriture du MASTER {chemin.name} non conforme : {' ; '.join(ecarts)}. La sauvegarde "
            f"{sauvegarde} contient le MASTER d'avant l'écriture : comparez, et restaurez-la si besoin."
        )


def _refuser_si_perte(chemin: Path) -> None:
    """Refuse d'écrire un MASTER qu'un enregistrement par openpyxl abîmerait (§ 15)."""
    pertes = [perte for perte in fidelite_aller_retour(chemin) if _PERTE_TOLEREE not in perte]
    if pertes:
        raise RegistreInvalide(
            f"Écriture refusée : un enregistrement du MASTER {chemin.name} par l'outil perdrait : "
            f"{' ; '.join(pertes)}. Faites la saisie dans Excel ; rien n'a été modifié."
        )


def _ecrire(profil: Profil, operation: Operation) -> Path:
    """Joue l'opération à blanc (un refus ne laisse aucune trace), puis sous `ecriture_master` :
    enregistrement atomique, événements `registre`, relecture. Rend la sauvegarde."""
    chemins.refuser_racine_dans_un_depot(profil.racine)
    essai = _a_blanc(profil, operation)
    if not essai.cellules:
        raise RegistreInvalide(f"{essai.rien[:1].upper()}{essai.rien[1:]} : rien n'a été écrit.")
    _refuser_si_perte(profil.master)
    poste = poste_courant()
    with ecriture_master(profil.master, poste=poste) as sauvegarde:
        wb = _charger(profil.master)
        try:
            onglets_avant = list(wb.sheetnames)
            avant = _instantane(wb)
            resultat = operation(wb)
            if not resultat.cellules:
                raise RegistreInvalide(f"{resultat.rien[:1].upper()}{resultat.rien[1:]} : rien n'a été écrit.")
            fichiers.enregistrer_classeur_atomique(wb, profil.master)
        finally:
            wb.close()
        chemin_journal = chemins.journal_ecritures(profil)
        for numero, valeurs_avant, valeurs_apres in resultat.evenements:
            journal.ajouter_evenement(
                chemin_journal, "registre", poste=poste, entite=profil.entite,
                n_facture=numero, avant=valeurs_avant, apres=valeurs_apres, sauvegarde=sauvegarde,
            )
        _verifier_relecture(profil.master, avant, onglets_avant, resultat, sauvegarde)
    return sauvegarde


def _operation_mise_a_jour(n_facture: str, champs: Mapping[str, Any], remplacer: bool) -> Operation:
    def operation(wb: Workbook) -> _Resultat:
        ws = _onglet(wb)
        _, numero, changements, ecrites = _modifier(ws, n_facture, remplacer=remplacer, champs=champs)
        resultat = _Resultat(apercu=changements, rien=f"N° {numero} : rien ne change (valeurs déjà saisies)")
        resultat.cellules = {(ONGLET, coordonnee): valeur for coordonnee, valeur in ecrites.items()}
        if changements:
            resultat.evenements.append((
                numero,
                {colonne: avant for colonne, (avant, _) in changements.items()},
                {colonne: apres for colonne, (_, apres) in changements.items()},
            ))
        return resultat

    return operation


def _champs_mise_a_jour(date_envoi: date | None, date_paiement: date | None, note: str | None) -> dict[str, Any]:
    champs = {nom: valeur for nom, valeur in (("date_envoi", date_envoi), ("date_paiement", date_paiement),
                                              ("note", note)) if valeur is not None}
    if not champs:
        raise ValueError("Rien à modifier : donnez une date d'envoi, une date de paiement ou une note.")
    return champs


def apercu_mise_a_jour(profil: Profil, n_facture: str, *, date_envoi: date | None = None,
                       date_paiement: date | None = None, note: str | None = None,
                       remplacer: bool = False) -> dict[str, tuple[Any, Any]]:
    """Ce que ferait `mettre_a_jour_registre`, sans rien écrire : {colonne : (avant, après)}
    (vide si rien ne change). Mêmes refus que `modifier_ligne`."""
    champs = _champs_mise_a_jour(date_envoi, date_paiement, note)
    return _a_blanc(profil, _operation_mise_a_jour(n_facture, champs, remplacer)).apercu


def mettre_a_jour_registre(profil: Profil, n_facture: str, *, date_envoi: date | None = None,
                           date_paiement: date | None = None, note: str | None = None,
                           remplacer: bool = False, confirme: bool) -> Path:
    """Note la date d'envoi, la date de paiement ou une note d'une facture du registre.

    Sous `ecriture_master` (refus d'un MASTER ouvert, verrou, sauvegarde), enregistrement atomique
    et relecture ; un événement `registre` au journal des écritures.

    Returns:
        La sauvegarde du MASTER faite avant l'écriture.

    Raises:
        ValueError: « validation explicite requise » sans `confirme=True` ; aucun champ donné.
        RegistreInvalide: refus de `modifier_ligne`, onglet absent, rien ne change ; rien n'est
            écrit, ni sauvegardé, ni journalisé.
        master.MasterVerrouille: MASTER ouvert ou en cours d'écriture ailleurs ; rien n'est écrit.
        EcritureNonConforme: la relecture trouve autre chose que les cellules annoncées.
    """
    _confirmation(confirme)
    champs = _champs_mise_a_jour(date_envoi, date_paiement, note)
    return _ecrire(profil, _operation_mise_a_jour(n_facture, champs, remplacer))


def _numero_et_ecole(n_facture: str, profil: Profil) -> tuple[str, str]:
    """N° en majuscules et code école qu'il porte (« R_ABC002 » → ABC), au préfixe du profil."""
    cle = _cle(n_facture)
    prefixe = profil.prefixe_facture
    trouve = re.fullmatch(rf"{re.escape(prefixe)}([A-Z][A-Z0-9]*?)_?(\d{{3}})", cle)
    if not trouve:
        raise RegistreInvalide(
            f"N° « {_texte(n_facture)} » : attendu {prefixe}<CODE ÉCOLE><NNN> (ex. {prefixe}ABC002), au préfixe "
            f"de l'entité {profil.entite}."
        )
    return cle, valider_code_ecole(trouve.group(1))


def _numeros_sessions(wb: Workbook) -> set[str]:
    """N° de la colonne `N°_Facture` de l'onglet `Sessions`, sous leur forme de comparaison."""
    ws = wb["Sessions"]
    colonne = index_colonnes(ws, ("N°_Facture",), onglet="Sessions")["N°_Facture"]
    return {
        _cle(ws.cell(row=rang, column=colonne).value)
        for rang in range(2, derniere_ligne_remplie(ws) + 1)
    } - {""}


def _operation_manquant(numero: str, ecole: str, raison: str) -> Operation:
    def operation(wb: Workbook) -> _Resultat:
        ws = _onglet(wb)
        if numero in _numeros_sessions(wb):
            raise RegistreInvalide(
                f"N° {numero} : il figure dans l'onglet « Sessions », il a donc servi ; ce n'est pas un numéro "
                f"manquant (reprenez-le avec `facturation {COMMANDE_REPRISE}`)."
            )
        if not isinstance(raison, str) or not raison.strip():
            raise RegistreInvalide(f"N° {numero} : la raison du numéro manquant est obligatoire (ADR029).")
        ligne = {COL_NUMERO: numero, COL_ECOLE: ecole, COL_STATUT: ABANDONNEE, COL_NOTE: raison.strip()}
        _, ecrites = _ajouter(ws, ligne)
        resultat = _Resultat(apercu=ligne)
        resultat.cellules = {(ONGLET, coordonnee): valeur for coordonnee, valeur in ecrites.items()}
        resultat.evenements.append((numero, None, ligne))
        return resultat

    return operation


def apercu_manquant(profil: Profil, n_facture: str, *, raison: str) -> dict[str, Any]:
    """Ce que ferait `noter_manquant`, sans rien écrire : la ligne ajoutée {colonne : valeur}.
    Mêmes refus que `noter_manquant`."""
    numero, ecole = _numero_et_ecole(n_facture, profil)
    return _a_blanc(profil, _operation_manquant(numero, ecole, raison)).apercu


def noter_manquant(profil: Profil, n_facture: str, *, raison: str, confirme: bool) -> Path:
    """Note au registre la raison d'un numéro manquant (ADR029) : une ligne au statut
    « abandonnée », école déduite du N° et du préfixe du profil (« R_ABC002 » → ABC).

    Returns:
        La sauvegarde du MASTER faite avant l'écriture.

    Raises:
        ValueError: « validation explicite requise » sans `confirme=True`.
        RegistreInvalide: N° mal formé ou d'un autre préfixe, déjà au registre, présent dans
            `Sessions` ; raison vide ; onglet absent. Rien n'est écrit.
        master.MasterVerrouille: MASTER ouvert ou en cours d'écriture ailleurs ; rien n'est écrit.
        EcritureNonConforme: la relecture trouve autre chose que les cellules annoncées.
    """
    _confirmation(confirme)
    numero, ecole = _numero_et_ecole(n_facture, profil)
    return _ecrire(profil, _operation_manquant(numero, ecole, raison))


def _operation_initialiser(profil: Profil, reprendre_historique: bool) -> Operation:
    def operation(wb: Workbook) -> _Resultat:
        resultat = _Resultat(rien="rien à reprendre : toutes les factures de Sessions sont au registre")
        cree = ONGLET not in wb.sheetnames
        if cree:
            ws = creer_onglet(wb)
            resultat.onglets_crees.append(ONGLET)
            resultat.cellules.update({(ONGLET, c.coordinate): c.value for c in ws[1] if c.value is not None})
            resultat.evenements.append((None, None, {"onglet": ONGLET, "colonnes": list(COLONNES)}))
        elif not reprendre_historique:
            raise RegistreInvalide(
                f"L'onglet « {ONGLET} » existe déjà : rien à initialiser. Pour y reprendre les factures de Sessions "
                "absentes du registre, ajoutez --reprendre-historique."
            )
        else:
            ws = wb[ONGLET]
        reprises: list[dict[str, Any]] = []
        if reprendre_historique:
            for ligne in lignes_historiques(lire_master(profil.master)):
                _, ecrites = _ajouter(ws, ligne)
                resultat.cellules.update({(ONGLET, coordonnee): valeur for coordonnee, valeur in ecrites.items()})
                resultat.evenements.append((ligne[COL_NUMERO], None, ligne))
                reprises.append(ligne)
        resultat.apercu = (cree, reprises)
        return resultat

    return operation


def apercu_initialisation(profil: Profil, *, reprendre_historique: bool) -> tuple[bool, list[dict[str, Any]]]:
    """Ce que ferait `initialiser_registre`, sans rien écrire : (onglet à créer ?, lignes reprises).
    Mêmes refus que `initialiser_registre`."""
    return _a_blanc(profil, _operation_initialiser(profil, reprendre_historique)).apercu


def initialiser_registre(profil: Profil, *, reprendre_historique: bool, confirme: bool) -> Path:
    """Crée l'onglet `Factures` s'il manque (en fin de classeur, l'ordre des autres onglets ne
    change pas) et, avec `reprendre_historique`, y ajoute les lignes de `lignes_historiques` (sur
    un registre existant, seulement les N° de `Sessions` qui y manquent, sans toucher aux autres).

    Returns:
        La sauvegarde du MASTER faite avant l'écriture.

    Raises:
        ValueError: « validation explicite requise » sans `confirme=True`.
        RegistreInvalide: onglet déjà présent sans reprise, rien à reprendre, `Sessions` ou
            registre illisible. Rien n'est écrit.
        master.MasterVerrouille: MASTER ouvert ou en cours d'écriture ailleurs ; rien n'est écrit.
        EcritureNonConforme: la relecture trouve autre chose que les cellules annoncées.
    """
    _confirmation(confirme)
    return _ecrire(profil, _operation_initialiser(profil, reprendre_historique))


# ──────────────────────────────────────────────────────────────────────
# Aperçu
# ──────────────────────────────────────────────────────────────────────


def section_apercu(profil: Profil, *, aujourd_hui: date) -> tuple[str, list[str]] | None:
    """Section de l'aperçu : (« FACTURES ÉCHUES NON PAYÉES : N », une ligne par facture échue, avec
    son N° et son retard en jours) ; None s'il n'y en a pas, ou sans onglet `Factures`.

    Raises:
        RegistreInvalide, master.MasterInvalide: MASTER ou registre illisible.
    """
    lignes = echues(lire_registre(lire_master(profil.master)), aujourd_hui=aujourd_hui)
    if not lignes:
        return None
    textes = []
    for ligne in lignes:
        retard = retard_jours(ligne, aujourd_hui=aujourd_hui)
        ecole = f" ({ligne.ecole})" if ligne.ecole else ""
        textes.append(
            f"{ligne.n_facture}{ecole} — échéance {_jj(ligne.echeance)} — {retard} jour{'s' if retard != 1 else ''} "
            "de retard"
        )
    return f"FACTURES ÉCHUES NON PAYÉES : {len(lignes)}", textes
