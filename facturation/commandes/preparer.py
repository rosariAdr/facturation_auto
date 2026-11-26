"""Commandes `facturation preparer` et `facturation apercu` : préparation du mois (F-16a).

    facturation preparer [--mois AAAA-MM] [--ecoles ABC,ECO] [--fichier AGENDA.ics] [--profil CHEMIN]
    facturation apercu [--mois AAAA-MM] [--facture ID|F<n>] [--ouvrir] [--profil CHEMIN]

`preparer` lit l'agenda et le MASTER, écrit les brouillons et les aperçus sous
`_sorties/<AAAA-MM>/` (jamais dans le MASTER), puis affiche l'aperçu global et le dossier.
Codes : 3 si une facture est à valider, à revalider, à corriger ou bloquée, ou si un créneau est à
classer ; 0 sinon ; 1 si la préparation est impossible ou le profil invalide.

`apercu` relit `preparation.json` et le journal du mois, sans relire l'agenda ni rien écrire.
Codes : 1 sans préparation du mois (« lancez `facturation preparer --mois …` »), 0 sinon.

Lectures retenues (`💡 [proposé — à valider]`) : le mois par défaut se calcule sur la date du jour
de la machine ; `--ecoles` accepte des codes séparés par des virgules (casse et espaces ignorés) ;
`apercu --ouvrir` ouvre le dossier du mois sans afficher l'aperçu, et affiche seulement son chemin
hors de Windows ; une facture inconnue de `--facture` est une erreur (code 1).
"""

from __future__ import annotations

import argparse
import os
import sys
from datetime import date
from pathlib import Path

from facturation import chemins
from facturation.commandes import ERREUR, ETIQUETTES, argument_mois, charger_profil, entete, mois_par_defaut


def ajouter(sous) -> None:
    """Déclare `preparer` et `apercu` dans les sous-commandes de `facturation`."""
    preparer = sous.add_parser("preparer", help="Préparer les brouillons de facture du mois")
    preparer.add_argument("--mois", type=argument_mois, default=None,
                          help="Mois à préparer, AAAA-MM (défaut : selon le jour de lancement du profil)")
    preparer.add_argument("--ecoles", default=None, help="Codes des écoles à préparer, séparés par des virgules")
    preparer.add_argument("--fichier", type=Path, default=None,
                          help="Lire un fichier .ics local au lieu de télécharger l'agenda")
    preparer.add_argument("--profil", type=Path, default=None, help="Profil à utiliser (défaut : emplacement standard)")
    preparer.set_defaults(func=_preparer)

    apercu = sous.add_parser("apercu", help="Relire l'aperçu du mois préparé, sans relire l'agenda")
    apercu.add_argument("--mois", type=argument_mois, default=None, help="Mois préparé, AAAA-MM")
    apercu.add_argument("--facture", default=None, help="Une facture : son identifiant ou son numéro court (F1…)")
    apercu.add_argument("--ouvrir", action="store_true", help="Ouvrir le dossier du mois")
    apercu.add_argument("--profil", type=Path, default=None, help="Profil à utiliser (défaut : emplacement standard)")
    apercu.set_defaults(func=_apercu)


def _erreur(message: str) -> None:
    print(f"{ETIQUETTES[ERREUR]}{message}", file=sys.stderr)


def _code_de_retour(preparation, etats) -> int:
    from facturation import preparation as prp

    if preparation.a_classer or any(etats.get(b.identifiant) in prp.ETATS_A_TRAITER for b in preparation.brouillons):
        return 3
    return 0


def _preparer(args: argparse.Namespace) -> int:
    from facturation import preparation as prp

    profil = charger_profil(args.profil)
    if profil is None:
        return 1
    mois = args.mois or mois_par_defaut(profil, date.today())
    ecoles = None
    if args.ecoles is not None:
        ecoles = [code.strip().upper() for code in args.ecoles.split(",") if code.strip()]
    contenu = None
    if args.fichier is not None:
        try:
            contenu = args.fichier.read_bytes()
        except OSError as exc:
            print(entete(profil, mois))
            raison = "fichier introuvable" if isinstance(exc, FileNotFoundError) else (exc.strerror or str(exc))
            _erreur(f"Agenda — impossible de lire {args.fichier} : {raison}")
            return 1
    try:
        preparation = prp.lancer_preparation(profil, mois, contenu_ics=contenu, ecoles=ecoles)
        etats = prp.etats_de_validation(profil, preparation)
        sections = prp.sections_annexes(profil, aujourd_hui=preparation.prepare_le.date())
    except (prp.PreparationImpossible, ValueError, OSError) as exc:
        print(entete(profil, mois))
        _erreur(f"Préparation — {exc}")
        return 1
    print(prp.apercu_global(preparation, etats=etats, sections=sections), end="")
    print(f"\nDossier : {chemins.dossier_mois(profil, mois)}")
    return _code_de_retour(preparation, etats)


def _apercu(args: argparse.Namespace) -> int:
    from facturation import preparation as prp

    profil = charger_profil(args.profil)
    if profil is None:
        return 1
    mois = args.mois or mois_par_defaut(profil, date.today())
    dossier = chemins.dossier_mois(profil, mois)
    try:
        preparation = prp.lire_preparation(profil, mois)
    except ValueError as exc:
        print(entete(profil, mois))
        _erreur(f"Aperçu — {exc}")
        return 1
    if preparation is None:
        print(entete(profil, mois))
        _erreur(f"Aperçu — aucune préparation pour ce mois : lancez `facturation preparer --mois {mois}`")
        return 1
    if args.facture is not None or args.ouvrir:
        print(entete(profil, mois))
    if args.ouvrir:
        ouvrir = getattr(os, "startfile", None)
        if ouvrir is None:
            print(f"Dossier : {dossier}")
        else:
            ouvrir(str(dossier))
            print(f"Dossier ouvert : {dossier}")
        return 0
    try:
        etats = prp.etats_de_validation(profil, preparation)
    except (ValueError, OSError) as exc:
        _erreur(f"Aperçu — {exc}")
        return 1
    numeros = prp.numeros_courts(preparation)
    if args.facture is not None:
        demande = args.facture.strip()
        identifiant = numeros.get(demande.upper(), demande.upper())
        brouillon = next((b for b in preparation.brouillons if b.identifiant == identifiant), None)
        if brouillon is None:
            connues = ", ".join(f"{k} = {v}" for k, v in numeros.items() if k.startswith("F")) or "aucune"
            _erreur(f"Aperçu — facture « {demande} » inconnue dans la préparation du mois (factures : {connues})")
            return 1
        propres = [d for d in preparation.differences if d.identifiant == identifiant]
        print(prp.apercu_facture(brouillon, differences=None if preparation.precedente_le is None else propres,
                                 etat=etats.get(identifiant), numeros=numeros))
        return 0
    sections = prp.sections_annexes(profil, aujourd_hui=date.today())
    print(prp.apercu_global(preparation, etats=etats, sections=sections), end="")
    print(f"\nDossier : {dossier}")
    return 0
