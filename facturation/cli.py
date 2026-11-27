"""cli.py — Commande `facturation`.

Lot 0 :
    facturation profil chemin                      # où le profil est cherché
    facturation profil init --nom … --entite … --racine …
    facturation profil verifier [--profil CHEMIN]
    facturation profil agenda [--profil CHEMIN]    # enregistre l'adresse iCal secrète

Lot 1 :
    facturation agenda lire --mois AAAA-MM [--profil CHEMIN] [--fichier AGENDA.ics]
    facturation preparer [--mois AAAA-MM] [--ecoles ABC,ECO] [--fichier AGENDA.ics]   # F-16a
    facturation apercu [--mois AAAA-MM] [--facture ID|F<n>] [--ouvrir]   # F-16a, lecture seule
    facturation audit [--mois AAAA-MM | --du JJ/MM/AAAA --au JJ/MM/AAAA] [--ecoles ABC,ECO]   # F-19
    facturation registre [--ecole X] [--mois AAAA-MM] [--echues] [--trous]   # F-25
    facturation registre initialiser | envoyee | payee | note | manquant …   # F-25, écritures
Toutes acceptent `--profil CHEMIN`.

`preparer` écrit les brouillons et les aperçus sous `_sorties/<AAAA-MM>/`, `audit` son rapport
sous `_sorties/audits/` : ni l'une ni l'autre n'écrit dans le MASTER. Les actions de `registre`
écrivent l'onglet `Factures` du MASTER, après validation (« oui » ou `--oui`), avec sauvegarde et
verrou. Les commandes `classer` (F-16b) et `emettre` (F-17) arrivent plus tard dans le lot 1.

Répartition (P-0b) : chaque sous-commande vit dans un module de `facturation.commandes` qui
expose `ajouter(sous)` ; `_COMMANDES` les liste, un par ligne (ajouter en fin de liste), et
`main` appelle la fonction que l'analyseur a retenue (`args.func`).
"""

from __future__ import annotations

import argparse
import sys

from facturation import __version__
from facturation.commandes import agenda, preparer, profil
from facturation.commandes import audit
from facturation.commandes import registre

_COMMANDES = [
    profil,
    agenda,
    preparer,
    audit,
    registre,
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
