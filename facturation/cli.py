"""cli.py — Commande `facturation`.

Lot 0 :
    facturation profil chemin                      # où le profil est cherché
    facturation profil init --nom … --entite … --racine …
    facturation profil verifier [--profil CHEMIN]
    facturation profil agenda [--profil CHEMIN]    # enregistre l'adresse iCal secrète

Lot 1 :
    facturation agenda lire --mois AAAA-MM [--profil CHEMIN] [--fichier AGENDA.ics]
    facturation audit [--mois AAAA-MM | --du JJ/MM/AAAA --au JJ/MM/AAAA] [--ecoles ABC,ECO]   # F-19

Les commandes `preparer`, `emettre`, `audit` et `classer` arrivent au lot 1.

Répartition (P-0b) : chaque sous-commande vit dans un module de `facturation.commandes` qui
expose `ajouter(sous)` ; `_COMMANDES` les liste, un par ligne (ajouter en fin de liste), et
`main` appelle la fonction que l'analyseur a retenue (`args.func`).
"""

from __future__ import annotations

import argparse
import sys

from facturation import __version__
from facturation.commandes import agenda, preparer, profil
from facturation.commandes import agenda, profil
from facturation.commandes import audit

_COMMANDES = [
    profil,
    agenda,
    preparer,
    audit,
]


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="facturation",
        description="Factures mensuelles à partir de l'agenda de cours.",
    )
    parser.add_argument("--version", action="version", version=f"facturation {__version__}")
    sous = parser.add_subparsers(dest="commande", required=True)
    for module in _COMMANDES:
        module.ajouter(sous)
    return parser


def main(argv: list[str] | None = None) -> int:
    # Une console Windows redirigée (tâche planifiée) peut ne pas être en UTF-8.
    for flux in (sys.stdout, sys.stderr):
        if hasattr(flux, "reconfigure"):
            flux.reconfigure(errors="replace")
    args = _parser().parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
