"""commandes — Sous-commandes de `facturation` et outils qu'elles partagent (P-0b).

Chaque sous-commande vit dans un module de ce paquet qui expose `ajouter(sous) -> None` :
il y déclare son analyseur et branche sa fonction par `set_defaults(func=…)`. `cli._COMMANDES`
liste ces modules, un par ligne.

Outils communs (contrats des vagues 3 à 5, § 1.6 et 3.7) :
- `argument_mois`, `argument_date` : types d'argument pour argparse ;
- `mois_par_defaut` : le mois en cours à partir du jour de lancement du profil, sinon le
  mois précédent ;
- `charger_profil` : le profil, ou None après avoir affiché l'erreur ;
- `demander` : une réponse choisie parmi des mots, saisie en entier, sans défaut ;
- `entete` : la première ligne de chaque commande, « Mois : décembre 2026 — entité ROSARI ».

Lectures retenues (`💡 [proposé — à valider]`) : un refus d'`argument_mois` ou d'`argument_date`
est une `argparse.ArgumentTypeError` qui cite la valeur ; `demander` repose la question tant que
la réponse n'est pas l'un des mots proposés (espaces autour et casse ignorés) ; `charger_profil`
affiche l'erreur sur la sortie d'erreur.
"""

from __future__ import annotations

import argparse
import re
import sys
from collections.abc import Callable, Sequence
from datetime import date, datetime
from pathlib import Path

from facturation import chemins, config
from facturation.config import Profil

OK, ATTENTION, ERREUR = "OK", "ATTENTION", "ERREUR"
ETIQUETTES = {OK: "[OK]        ", ATTENTION: "[ATTENTION] ", ERREUR: "[ERREUR]    "}

_DATE_JJ_MM_AAAA = re.compile(r"\d{2}/\d{2}/\d{4}")
_DATE_AAAA_MM_JJ = re.compile(r"\d{4}-\d{2}-\d{2}")


def argument_mois(valeur: str) -> str:
    """Type argparse d'un mois : « AAAA-MM » (espaces autour retirés).

    Raises:
        argparse.ArgumentTypeError: autre forme, ou mois hors de 01 à 12 ; le message cite la valeur.
    """
    mois = valeur.strip()
    try:
        chemins.mois_valide(mois)
    except ValueError:
        raise argparse.ArgumentTypeError(f"mois invalide : {valeur!r}, attendu AAAA-MM (ex. 2026-11)") from None
    return mois


def argument_date(valeur: str) -> date:
    """Type argparse d'une date : JJ/MM/AAAA ou AAAA-MM-JJ.

    Raises:
        argparse.ArgumentTypeError: autre forme ou date impossible ; le message cite la valeur.
    """
    texte = valeur.strip()
    for motif, forme in ((_DATE_JJ_MM_AAAA, "%d/%m/%Y"), (_DATE_AAAA_MM_JJ, "%Y-%m-%d")):
        if motif.fullmatch(texte):
            try:
                return datetime.strptime(texte, forme).date()
            except ValueError:
                break
    raise argparse.ArgumentTypeError(f"date invalide : {valeur!r}, attendu JJ/MM/AAAA ou AAAA-MM-JJ (ex. 01/12/2026)")


def mois_par_defaut(profil: Profil, aujourd_hui: date) -> str:
    """Mois traité par défaut : le mois en cours à partir du jour de lancement du profil, sinon
    le mois précédent (changement d'année compris)."""
    mois = f"{aujourd_hui.year}-{aujourd_hui.month:02d}"
    return mois if aujourd_hui.day >= profil.lancement_jour else chemins.mois_precedent(mois)


def charger_profil(chemin: Path | None) -> Profil | None:
    """Le profil (emplacement par défaut si `chemin` est None), ou None après avoir affiché
    l'erreur sur la sortie d'erreur."""
    try:
        return config.charger(chemin)
    except config.ProfilInvalide as exc:
        print(f"{ETIQUETTES[ERREUR]}Profil — {exc}", file=sys.stderr)
        return None


def demander(question: str, reponses: Sequence[str], *, entree: Callable[[str], str] = input) -> str | None:
    """Pose la question jusqu'à obtenir l'un des mots proposés, saisi en entier.

    Aucun défaut : une réponse vide, un début de mot ou un autre mot font reposer la question.
    Les espaces autour et la casse ne comptent pas. La fin de l'entrée (EOF, Ctrl+Z) rend None.

    Returns:
        Le mot proposé choisi, écrit comme dans `reponses` ; None en fin d'entrée.
    """
    invite = f"{question} ({' / '.join(reponses)}) : "
    proposes = {reponse.casefold(): reponse for reponse in reponses}
    while True:
        try:
            saisie = entree(invite)
        except EOFError:
            return None
        choisi = proposes.get(saisie.strip().casefold())
        if choisi is not None:
            return choisi


def entete(profil: Profil, mois: str) -> str:
    """Première ligne de chaque commande : « Mois : décembre 2026 — entité ROSARI »."""
    debut, _ = chemins.mois_valide(mois)
    return f"Mois : {chemins.NOMS_MOIS[debut.month - 1]} {debut.year} — entité {profil.entite}"
