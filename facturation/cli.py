"""cli.py — Commande `facturation`.

Lot 0 :
    facturation profil chemin                      # où le profil est cherché
    facturation profil init --nom … --entite … --racine …
    facturation profil verifier [--profil CHEMIN]

Les commandes `preparer`, `emettre`, `audit` et `classer` arrivent au lot 1.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from facturation import __version__, config
from facturation.verification import ATTENTION, ERREUR, verifier

_ETIQUETTES = {"OK": "[OK]        ", ATTENTION: "[ATTENTION] ", ERREUR: "[ERREUR]    "}


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
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
