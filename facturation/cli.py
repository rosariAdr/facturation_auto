"""cli.py — Commande `facturation`.

Lot 0 :
    facturation profil chemin                      # où le profil est cherché
    facturation profil init --nom … --entite … --racine …
    facturation profil verifier [--profil CHEMIN]
    facturation profil agenda [--profil CHEMIN]    # enregistre l'adresse iCal secrète

Lot 1 :
    facturation agenda lire --mois AAAA-MM [--profil CHEMIN] [--fichier AGENDA.ics]

Les commandes `preparer`, `emettre`, `audit` et `classer` arrivent au lot 1.
"""

from __future__ import annotations

import argparse
import calendar
import getpass
import re
import sys
from datetime import date
from pathlib import Path

from facturation import __version__, config
from facturation.verification import ATTENTION, ERREUR, verifier

_ETIQUETTES = {"OK": "[OK]        ", ATTENTION: "[ATTENTION] ", ERREUR: "[ERREUR]    "}
_MOIS = ("janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août",
         "septembre", "octobre", "novembre", "décembre")
_JOURS = ("lun.", "mar.", "mer.", "jeu.", "ven.", "sam.", "dim.")


def _mois(valeur: str) -> tuple[date, date]:
    """`AAAA-MM` → premier et dernier jour du mois."""
    trouve = re.fullmatch(r"(\d{4})-(\d{2})", valeur.strip())
    if not trouve or not 1 <= int(trouve.group(2)) <= 12:
        raise argparse.ArgumentTypeError(f"mois invalide : {valeur!r}, attendu AAAA-MM (ex. 2026-11)")
    annee, mois = int(trouve.group(1)), int(trouve.group(2))
    return date(annee, mois, 1), date(annee, mois, calendar.monthrange(annee, mois)[1])


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="facturation",
        description="Factures mensuelles à partir de l'agenda de cours.",
    )
    parser.add_argument("--version", action="version", version=f"facturation {__version__}")
    sous = parser.add_subparsers(dest="commande", required=True)

    profil = sous.add_parser("profil", help="Créer, situer ou vérifier le profil utilisateur")
    actions = profil.add_subparsers(dest="action", required=True)

    actions.add_parser("chemin", help="Afficher l'emplacement du profil")

    init = actions.add_parser("init", help="Créer un profil commenté")
    init.add_argument("--nom", required=True, help="Nom de la personne qui facture")
    init.add_argument("--entite", required=True, help="Code de l'entité, ex. ROSARI")
    init.add_argument("--racine", required=True, help="Dossier qui contient <ENTITE>/<ENTITE>_MASTER.xlsx")
    init.add_argument("--emplacement", type=Path, default=None, help="Fichier à créer (défaut : emplacement standard)")
    init.add_argument("--force", action="store_true", help="Remplacer un profil existant")

    verif = actions.add_parser("verifier", help="Contrôler le profil, les données et les outils")
    verif.add_argument("--profil", type=Path, default=None, help="Profil à utiliser (défaut : emplacement standard)")

    agenda = actions.add_parser("agenda", help="Enregistrer l'adresse iCal secrète de l'agenda de cours")
    agenda.add_argument("--profil", type=Path, default=None, help="Profil à utiliser (défaut : emplacement standard)")
    agenda.add_argument(
        "--adresse", default=None,
        help="Adresse iCal secrète (déconseillé : elle reste dans l'historique du terminal ; "
             "sans cette option, elle est demandée sans s'afficher)",
    )

    agenda = sous.add_parser("agenda", help="Lire l'agenda des cours")
    actions_agenda = agenda.add_subparsers(dest="action", required=True)
    lire = actions_agenda.add_parser("lire", help="Lister les séances d'un mois")
    lire.add_argument("--mois", required=True, type=_mois, help="Mois à lire, au format AAAA-MM (ex. 2026-11)")
    lire.add_argument("--profil", type=Path, default=None, help="Profil à utiliser (défaut : emplacement standard)")
    lire.add_argument(
        "--fichier", type=Path, default=None,
        help="Lire un fichier .ics local au lieu de télécharger l'agenda (sans profil : fuseau Europe/Paris)",
    )
    return parser


def _profil_chemin() -> int:
    print(config.emplacement_par_defaut())
    return 0


def _profil_init(args: argparse.Namespace) -> int:
    cible = args.emplacement or config.emplacement_par_defaut()
    if cible.exists() and not args.force:
        print(f"Un profil existe déjà : {cible}\nAjoutez --force pour le remplacer.", file=sys.stderr)
        return 1
    try:
        contenu = config.rediger(args.nom, args.entite, args.racine)
    except ValueError as exc:
        print(f"Erreur : {exc}", file=sys.stderr)
        return 1
    cible.parent.mkdir(parents=True, exist_ok=True)
    cible.write_text(contenu, encoding="utf-8")
    print(f"Profil créé : {cible}\nVérifiez-le avec : facturation profil verifier")
    return 0


def _profil_verifier(args: argparse.Namespace) -> int:
    try:
        profil = config.charger(args.profil)
    except config.ProfilInvalide as exc:
        print(f"{_ETIQUETTES[ERREUR]}Profil — {exc}")
        return 1
    constats = verifier(profil)
    for c in constats:
        print(f"{_ETIQUETTES[c.niveau]}{c.sujet} — {c.message}")
    n_erreurs = sum(c.niveau == ERREUR for c in constats)
    n_attention = sum(c.niveau == ATTENTION for c in constats)
    print(f"\n{n_erreurs} erreur(s), {n_attention} point(s) d'attention.")
    return 1 if n_erreurs else 0


def _profil_agenda(args: argparse.Namespace) -> int:
    try:
        profil = config.charger(args.profil)
    except config.ProfilInvalide as exc:
        print(f"{_ETIQUETTES[ERREUR]}Profil — {exc}", file=sys.stderr)
        return 1
    adresse = args.adresse
    if adresse is None:
        print(
            "Dans Google Agenda (navigateur) : survolez l'agenda de cours dans la colonne de gauche,\n"
            "cliquez sur les trois points, puis « Paramètres et partage », et copiez\n"
            "« Adresse secrète au format iCal ».",
        )
        adresse = getpass.getpass("Collez l'adresse (elle ne s'affiche pas), puis Entrée : ")
    try:
        secret = config.enregistrer_adresse_ical(profil, adresse)
    except ValueError as exc:
        print(f"{_ETIQUETTES[ERREUR]}Agenda — {exc}", file=sys.stderr)
        return 1
    print(f"Adresse enregistrée dans {secret} (hors dépôt) et déclarée dans {profil.chemin}.")
    print("Contrôlez avec : facturation profil verifier")
    return 0


def _heures(duree_h: float) -> str:
    """Heures décimales à la française : 3,5 h ; 4 h ; 1,75 h."""
    return f"{duree_h:.2f}".rstrip("0").rstrip(".").replace(".", ",") + " h"


def _agenda_lire(args: argparse.Namespace) -> int:
    # Import différé : les commandes `profil` restent utilisables sans les bibliothèques iCal.
    from facturation import agenda

    debut, fin = args.mois
    fuseau = agenda.FUSEAU_PAR_DEFAUT
    profil = None
    if args.profil is not None or args.fichier is None:
        try:
            profil = config.charger(args.profil)
        except config.ProfilInvalide as exc:
            print(f"{_ETIQUETTES[ERREUR]}Profil — {exc}", file=sys.stderr)
            return 1
        fuseau = profil.fuseau

    try:
        if args.fichier is not None:
            try:
                contenu = args.fichier.read_bytes()
            except OSError as exc:
                raison = "fichier introuvable" if isinstance(exc, FileNotFoundError) else (exc.strerror or str(exc))
                print(f"{_ETIQUETTES[ERREUR]}Agenda — impossible de lire {args.fichier} : {raison}", file=sys.stderr)
                return 1
        else:
            contenu = agenda.telecharger_ics(profil)
        seances = agenda.lire_ics(contenu, debut=debut, fin=fin, fuseau=fuseau)
    except (agenda.AgendaIndisponible, ValueError) as exc:
        print(f"{_ETIQUETTES[ERREUR]}Agenda — {exc}", file=sys.stderr)
        return 1

    titre = f"{_MOIS[debut.month - 1]} {debut.year}"
    if not seances:
        print(f"Aucune séance en {titre} (fuseau {fuseau}).")
        return 0
    print(f"Séances de {titre} (fuseau {fuseau}) :")
    for s in seances:
        print(f"  {_JOURS[s.date.weekday()]} {s.date:%d/%m/%Y}  {s.debut:%H:%M}–{s.fin:%H:%M}  "
              f"{_heures(s.duree_h):>7}  {s.intitule or '(sans intitulé)'}")
    total = sum(s.duree_h for s in seances)
    print(f"\n{len(seances)} séance(s), {_heures(total)} au total.")
    return 0


def main(argv: list[str] | None = None) -> int:
    # Une console Windows redirigée (tâche planifiée) peut ne pas être en UTF-8.
    for flux in (sys.stdout, sys.stderr):
        if hasattr(flux, "reconfigure"):
            flux.reconfigure(errors="replace")
    args = _parser().parse_args(argv)
    if args.commande == "profil":
        if args.action == "chemin":
            return _profil_chemin()
        if args.action == "init":
            return _profil_init(args)
        if args.action == "verifier":
            return _profil_verifier(args)
        if args.action == "agenda":
            return _profil_agenda(args)
    if args.commande == "agenda" and args.action == "lire":
        return _agenda_lire(args)
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
