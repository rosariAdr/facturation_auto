"""Commande `facturation profil` : situer, créer, vérifier le profil, ranger l'adresse iCal (lot 0).

    facturation profil chemin
    facturation profil init --nom … --entite … --racine … [--emplacement F] [--force]
    facturation profil verifier [--profil CHEMIN]
    facturation profil agenda [--profil CHEMIN] [--adresse …]
"""

from __future__ import annotations

import argparse
import getpass
import sys
from pathlib import Path

from facturation import config
from facturation.commandes import ERREUR, ETIQUETTES


def ajouter(sous) -> None:
    """Déclare `profil` et ses actions dans les sous-commandes de `facturation`."""
    profil = sous.add_parser("profil", help="Créer, situer ou vérifier le profil utilisateur")
    actions = profil.add_subparsers(dest="action", required=True)

    chemin = actions.add_parser("chemin", help="Afficher l'emplacement du profil")
    chemin.set_defaults(func=_chemin)

    init = actions.add_parser("init", help="Créer un profil commenté")
    init.add_argument("--nom", required=True, help="Nom de la personne qui facture")
    init.add_argument("--entite", required=True, help="Code de l'entité, ex. ROSARI")
    init.add_argument("--racine", required=True, help="Dossier qui contient <ENTITE>/<ENTITE>_MASTER.xlsx")
    init.add_argument("--emplacement", type=Path, default=None, help="Fichier à créer (défaut : emplacement standard)")
    init.add_argument("--force", action="store_true", help="Remplacer un profil existant")
    init.set_defaults(func=_init)

    verif = actions.add_parser("verifier", help="Contrôler le profil, les données et les outils")
    verif.add_argument("--profil", type=Path, default=None, help="Profil à utiliser (défaut : emplacement standard)")
    verif.set_defaults(func=_verifier)

    agenda = actions.add_parser("agenda", help="Enregistrer l'adresse iCal secrète de l'agenda de cours")
    agenda.add_argument("--profil", type=Path, default=None, help="Profil à utiliser (défaut : emplacement standard)")
    agenda.add_argument(
        "--adresse", default=None,
        help="Adresse iCal secrète (déconseillé : elle reste dans l'historique du terminal ; "
             "sans cette option, elle est demandée sans s'afficher)",
    )
    agenda.set_defaults(func=_agenda)


def _chemin(args: argparse.Namespace) -> int:
    print(config.emplacement_par_defaut())
    return 0


def _init(args: argparse.Namespace) -> int:
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


def _verifier(args: argparse.Namespace) -> int:
    from facturation.verification import ATTENTION, verifier

    try:
        profil = config.charger(args.profil)
    except config.ProfilInvalide as exc:
        print(f"{ETIQUETTES[ERREUR]}Profil — {exc}")
        return 1
    constats = verifier(profil)
    for c in constats:
        print(f"{ETIQUETTES[c.niveau]}{c.sujet} — {c.message}")
    n_erreurs = sum(c.niveau == ERREUR for c in constats)
    n_attention = sum(c.niveau == ATTENTION for c in constats)
    print(f"\n{n_erreurs} erreur(s), {n_attention} point(s) d'attention.")
    return 1 if n_erreurs else 0


def _agenda(args: argparse.Namespace) -> int:
    try:
        profil = config.charger(args.profil)
    except config.ProfilInvalide as exc:
        print(f"{ETIQUETTES[ERREUR]}Profil — {exc}", file=sys.stderr)
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
        print(f"{ETIQUETTES[ERREUR]}Agenda — {exc}", file=sys.stderr)
        return 1
    print(f"Adresse enregistrée dans {secret} (hors dépôt) et déclarée dans {profil.chemin}.")
    print("Contrôlez avec : facturation profil verifier")
    return 0
