"""Commande `facturation audit` : rapprocher l'agenda et les séances facturées (F-19).

    facturation audit [--mois AAAA-MM | --du JJ/MM/AAAA --au JJ/MM/AAAA] [--ecoles ABC,ECO]
                      [--fichier AGENDA.ics] [--profil CHEMIN]

Lecture seule du MASTER et de l'agenda ; le rapport (.txt et .csv) est écrit sous
`<racine>/<ENTITE>/_sorties/audits/`.

Codes de sortie (contrats des vagues 3 à 5, § 1.6) : 0 sans écart, 3 avec écarts ou avec des lignes de
`Sessions` illisibles ou sans date de séance (💡, T-AUD-61), 1 en cas d'erreur (profil, MASTER, agenda, racine, écriture), 2 pour une erreur d'usage.

Lectures retenues (`💡 [proposé — à valider]`) : période par défaut du 1er du mois M−3 à hier
(Q-I2), « aujourd'hui » étant la date du poste ; `--mois` ne se combine pas avec `--du` / `--au`,
qui vont ensemble ; une période inversée est une erreur d'usage (code 2) ; les avertissements
levés pendant l'audit (règle écartée…) s'affichent en « [ATTENTION] ».
"""

from __future__ import annotations

import argparse
import sys
import warnings
from datetime import date, timedelta
from pathlib import Path

from facturation import chemins
from facturation.commandes import ATTENTION, ERREUR, ETIQUETTES, argument_date, argument_mois, charger_profil
from facturation.fichiers import masquer_secrets


def _argument_ecoles(valeur: str) -> list[str]:
    """Type argparse de `--ecoles` : codes séparés par des virgules (« ABC,ECO »)."""
    ecoles = [code.strip().upper() for code in valeur.split(",") if code.strip()]
    if not ecoles:
        raise argparse.ArgumentTypeError(f"liste d'écoles vide : {valeur!r}, attendu des codes séparés par des virgules (ex. ABC,ECO)")
    return ecoles


def ajouter(sous) -> None:
    """Déclare `audit` dans les sous-commandes de `facturation`."""
    commande = sous.add_parser(
        "audit",
        help="Rapprocher l'agenda et les séances facturées (onglet Sessions)",
        description=(
            "Repère les séances tenues non facturées, les lignes de Sessions sans séance, les doublons et "
            "les écarts d'heures, sans rien modifier. Période par défaut : du 1er du mois M-3 à hier."
        ),
    )
    commande.add_argument("--mois", type=argument_mois, default=None, help="Mois audité, au format AAAA-MM (ex. 2026-11)")
    commande.add_argument("--du", type=argument_date, default=None, help="Premier jour audité (JJ/MM/AAAA ou AAAA-MM-JJ), avec --au")
    commande.add_argument("--au", type=argument_date, default=None, help="Dernier jour audité (JJ/MM/AAAA ou AAAA-MM-JJ), avec --du")
    commande.add_argument("--ecoles", type=_argument_ecoles, default=None, help="Limiter l'audit à ces écoles (ex. ABC,ECO)")
    commande.add_argument(
        "--fichier", type=Path, default=None,
        help="Lire un fichier .ics local au lieu de télécharger l'agenda",
    )
    commande.add_argument("--profil", type=Path, default=None, help="Profil à utiliser (défaut : emplacement standard)")
    commande.set_defaults(func=_audit, analyseur_audit=commande)


def _periode_par_defaut(aujourd_hui: date) -> tuple[date, date]:
    """Du 1er du mois M−3 à hier 💡 (Q-I2)."""
    annee, mois = aujourd_hui.year, aujourd_hui.month - 3
    if mois < 1:
        annee, mois = annee - 1, mois + 12
    return date(annee, mois, 1), aujourd_hui - timedelta(days=1)


def _erreur(sujet: str, exc: BaseException | str) -> int:
    print(f"{ETIQUETTES[ERREUR]}{sujet} — {masquer_secrets(str(exc))}", file=sys.stderr)
    return 1


def _audit(args: argparse.Namespace) -> int:
    analyseur: argparse.ArgumentParser = args.analyseur_audit
    if args.mois is not None and (args.du is not None or args.au is not None):
        analyseur.error("--mois ne se combine pas avec --du / --au : choisissez un mois, ou une période")
    if (args.du is None) != (args.au is None):
        manquante, donnee = ("--au", "--du") if args.au is None else ("--du", "--au")
        analyseur.error(f"{manquante} manque : {donnee} et {manquante} s'emploient ensemble")
    if args.du is not None and args.du > args.au:
        analyseur.error(f"période inversée : du {args.du:%d/%m/%Y} au {args.au:%d/%m/%Y}")

    profil = charger_profil(args.profil)
    if profil is None:
        return 1

    # Imports différés : les commandes `profil` restent utilisables sans les bibliothèques iCal.
    from facturation import agenda, audit, master
    from facturation import pcf_format as fmt

    if args.mois is not None:
        du, au = chemins.mois_valide(args.mois)
    elif args.du is not None:
        du, au = args.du, args.au
    else:
        du, au = _periode_par_defaut(date.today())

    contenu = None
    if args.fichier is not None:
        try:
            contenu = args.fichier.read_bytes()
        except OSError as exc:
            raison = "fichier introuvable" if isinstance(exc, FileNotFoundError) else (exc.strerror or str(exc))
            return _erreur("Agenda", f"impossible de lire {args.fichier} : {raison}")

    fmt.set_data_root(profil.racine)  # § 1.1 : chaque commande vide le cache des coordonnées
    with warnings.catch_warnings(record=True) as avertissements:
        warnings.simplefilter("always")
        try:
            rapport, chemin = audit.lancer_audit(profil, du, au, contenu_ics=contenu, ecoles=args.ecoles)
        except agenda.AgendaIndisponible as exc:
            return _erreur("Agenda", exc)
        except (ValueError, OSError, master.MasterVerrouille) as exc:
            return _erreur("Audit", exc)
        finally:
            for avertissement in avertissements:
                print(f"{ETIQUETTES[ATTENTION]}{masquer_secrets(str(avertissement.message))}", file=sys.stderr)

    print(audit.rapport_texte(rapport), end="")
    print(f"\nRapport écrit : {chemin}")
    print(f"Tableau (Excel) : {chemin.with_suffix('.csv')}")
    return rapport.code_retour
