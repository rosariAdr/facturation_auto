"""Commande `facturation agenda` : lire l'agenda des cours (F-10).

    facturation agenda lire --mois AAAA-MM [--profil CHEMIN] [--fichier AGENDA.ics]
"""

from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

from facturation import chemins, config
from facturation.commandes import ERREUR, ETIQUETTES, argument_mois

_JOURS = ("lun.", "mar.", "mer.", "jeu.", "ven.", "sam.", "dim.")


def _mois(valeur: str) -> tuple[date, date]:
    """`AAAA-MM` → premier et dernier jour du mois (refus : `argparse.ArgumentTypeError`)."""
    return chemins.mois_valide(argument_mois(valeur))


def ajouter(sous) -> None:
    """Déclare `agenda` et son action `lire` dans les sous-commandes de `facturation`."""
    agenda = sous.add_parser("agenda", help="Lire l'agenda des cours")
    actions = agenda.add_subparsers(dest="action", required=True)
    lire = actions.add_parser("lire", help="Lister les séances d'un mois")
    lire.add_argument("--mois", required=True, type=_mois, help="Mois à lire, au format AAAA-MM (ex. 2026-11)")
    lire.add_argument("--profil", type=Path, default=None, help="Profil à utiliser (défaut : emplacement standard)")
    lire.add_argument(
        "--fichier", type=Path, default=None,
        help="Lire un fichier .ics local au lieu de télécharger l'agenda (sans profil : fuseau Europe/Paris)",
    )
    lire.set_defaults(func=_lire)


def _heures(duree_h: float) -> str:
    """Heures décimales à la française : 3,5 h ; 4 h ; 1,75 h."""
    return f"{duree_h:.2f}".rstrip("0").rstrip(".").replace(".", ",") + " h"


def _lire(args: argparse.Namespace) -> int:
    # Import différé : les commandes `profil` restent utilisables sans les bibliothèques iCal.
    from facturation import agenda

    debut, fin = args.mois
    fuseau = agenda.FUSEAU_PAR_DEFAUT
    profil = None
    if args.profil is not None or args.fichier is None:
        try:
            profil = config.charger(args.profil)
        except config.ProfilInvalide as exc:
            print(f"{ETIQUETTES[ERREUR]}Profil — {exc}", file=sys.stderr)
            return 1
        fuseau = profil.fuseau

    try:
        if args.fichier is not None:
            try:
                contenu = args.fichier.read_bytes()
            except OSError as exc:
                raison = "fichier introuvable" if isinstance(exc, FileNotFoundError) else (exc.strerror or str(exc))
                print(f"{ETIQUETTES[ERREUR]}Agenda — impossible de lire {args.fichier} : {raison}", file=sys.stderr)
                return 1
        else:
            contenu = agenda.telecharger_ics(profil)
        seances = agenda.lire_ics(contenu, debut=debut, fin=fin, fuseau=fuseau)
    except (agenda.AgendaIndisponible, ValueError) as exc:
        print(f"{ETIQUETTES[ERREUR]}Agenda — {exc}", file=sys.stderr)
        return 1

    titre = f"{chemins.NOMS_MOIS[debut.month - 1]} {debut.year}"
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
