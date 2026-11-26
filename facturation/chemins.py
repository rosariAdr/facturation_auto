"""chemins.py — Mois, noms d'onglets et arborescence des sorties (P-0b, socle du lot 1).

Tous les chemins se calculent sous `<racine>/<ENTITE>/` (ADR001), à partir du profil :

    <racine>/<ENTITE>/
    ├── <ECOLE>/<ENTITE>_<ECOLE>.xlsx · <ECOLE>/<onglet>.pdf
    └── _sorties/
        ├── journal.jsonl                 journal des écritures
        ├── <AAAA-MM>/journal.jsonl       journal du mois
        ├── emissions/<N°>.json · <N°>.final.json · <N°>.pdf · abandonnees/
        ├── audits/
        └── signalements.jsonl
    <dossier du profil>/journal_tache.jsonl

Les fonctions de chemin ne créent aucun dossier. Les noms d'onglet créés par l'outil prennent
les abréviations Janv, Févr, Mars, Avr, Mai, Juin, Juil, Août, Sept, Oct, Nov, Déc (constat du
MASTER réel, contrats des vagues 3 à 5, § 15).

Lectures retenues (`💡 [proposé — à valider]`) : `.git` peut être un dossier (clone) ou un
fichier (copie de travail `git worktree`) ; un n° de facture ou un code école qui contient un
séparateur de chemin est refusé ; un mois écrit avec des chiffres non ASCII (« ２０２６-12 ») est
refusé ; la racine est comparée après résolution (« .. », liens et jonctions suivis).
"""

from __future__ import annotations

import calendar
import os
import re
from datetime import date
from pathlib import Path
from typing import TYPE_CHECKING

from facturation import config

if TYPE_CHECKING:  # agenda importe les bibliothèques iCal : pas à l'exécution
    from facturation.agenda import Seance
    from facturation.config import Profil

ABREVIATIONS_MOIS = ("Janv", "Févr", "Mars", "Avr", "Mai", "Juin", "Juil", "Août", "Sept", "Oct", "Nov", "Déc")
NOMS_MOIS = (
    "janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre",
    "octobre", "novembre", "décembre",
)

LONGUEUR_MAX_ONGLET = 31  # limite d'Excel
CARACTERES_INTERDITS_ONGLET = "[]:*?/\\"

DOSSIER_SORTIES = "_sorties"
NOM_JOURNAL = "journal.jsonl"
NOM_JOURNAL_TACHE = "journal_tache.jsonl"

# Chiffres ASCII seulement : « ２０２６-12 » ferait un second dossier `_sorties/` pour le même mois.
_MOIS_RE = re.compile(r"([0-9]{4})-(0[1-9]|1[0-2])")


# ──────────────────────────────────────────────────────────────────────
# Mois et libellés
# ──────────────────────────────────────────────────────────────────────


def _annee_mois(mois: str) -> tuple[int, int]:
    trouve = _MOIS_RE.fullmatch(mois) if isinstance(mois, str) else None
    if not trouve:
        raise ValueError(f"Mois invalide : « {mois} ». Attendu : AAAA-MM, mois de 01 à 12 (ex. 2026-12).")
    return int(trouve.group(1)), int(trouve.group(2))


def mois_valide(mois: str) -> tuple[date, date]:
    """« AAAA-MM » → (premier jour, dernier jour) du mois.

    Raises:
        ValueError: le texte n'est pas « AAAA-MM » avec un mois de 01 à 12 ; le message le cite.
    """
    annee, numero = _annee_mois(mois)
    return date(annee, numero, 1), date(annee, numero, calendar.monthrange(annee, numero)[1])


def mois_precedent(mois: str) -> str:
    """Mois qui précède « AAAA-MM » (changement d'année compris)."""
    annee, numero = _annee_mois(mois)
    return f"{annee - 1}-12" if numero == 1 else f"{annee}-{numero - 1:02d}"


def mois_suivant(mois: str) -> str:
    """Mois qui suit « AAAA-MM » (changement d'année compris)."""
    annee, numero = _annee_mois(mois)
    return f"{annee + 1}-01" if numero == 12 else f"{annee}-{numero + 1:02d}"


def libelle_periode(mois: str) -> str:
    """Libellé de la période dans un nom d'onglet : « 2026-12 » → « Déc 2026 »."""
    annee, numero = _annee_mois(mois)
    return f"{ABREVIATIONS_MOIS[numero - 1]} {annee}"


def onglet_facture(n_facture: str, periode_label: str) -> str:
    """Nom de l'onglet d'une facture : « <N°> <Période> » (ex. « R_ABC001 Déc 2026 »).

    Raises:
        ValueError: nom vide, de plus de 31 caractères (limite d'Excel), ou qui contient l'un des
            caractères interdits par Excel `[]:*?/\\`.
    """
    nom = f"{str(n_facture).strip()} {str(periode_label).strip()}".strip()
    if not str(n_facture).strip():
        raise ValueError(f"Nom d'onglet sans n° de facture : « {nom} ».")
    if len(nom) > LONGUEUR_MAX_ONGLET:
        raise ValueError(
            f"Nom d'onglet trop long : « {nom} » ({len(nom)} caractères, {LONGUEUR_MAX_ONGLET} au plus "
            "dans Excel). Raccourcissez la période."
        )
    interdits = sorted({c for c in nom if c in CARACTERES_INTERDITS_ONGLET})
    if interdits:
        raise ValueError(
            f"Nom d'onglet « {nom} » : caractère{'s' if len(interdits) > 1 else ''} interdit"
            f"{'s' if len(interdits) > 1 else ''} par Excel : {' '.join(interdits)}."
        )
    return nom


def cle_seance(seance: "Seance") -> str:
    """Identifiant stable d'une occurrence de séance : `<uid>@<début AAAA-MM-JJTHH:MM±hhmm>`."""
    return f"{seance.uid}@{seance.debut:%Y-%m-%dT%H:%M%z}"


# ──────────────────────────────────────────────────────────────────────
# Racine hors dépôt
# ──────────────────────────────────────────────────────────────────────


def refuser_racine_dans_un_depot(racine: Path) -> None:
    """Refuse une racine de données placée dans un dépôt git : aucune donnée réelle dans le dépôt.

    La racine est comparée après résolution (`resolve`) : `<dépôt>/../donnees` est hors du dépôt,
    et une jonction ou un lien symbolique qui mène dans un dépôt y est.

    Raises:
        ValueError: la racine, ou l'un de ses parents, contient `.git` (dossier d'un clone ou
            fichier d'une copie de travail).
    """
    ecrite = Path(racine).expanduser()
    try:
        chemin = ecrite.resolve()
    except (OSError, RuntimeError):  # lecteur absent, boucle de liens : « .. » au moins est résolu
        chemin = Path(os.path.abspath(ecrite))
    for dossier in (chemin, *chemin.parents):
        if (dossier / ".git").exists():
            designation = str(chemin) if chemin == ecrite.absolute() else f"{ecrite} (soit {chemin})"
            raise ValueError(
                f"Racine de données refusée : {designation} est dans le dépôt git {dossier}. Aucune donnée "
                "réelle ne vit dans un dépôt : placez la racine ailleurs (par exemple sur le lecteur "
                "Google Drive) et corrigez [donnees] racine dans le profil."
            )


# ──────────────────────────────────────────────────────────────────────
# Arborescence des sorties
# ──────────────────────────────────────────────────────────────────────


def _nom_sur(valeur: str, quoi: str) -> str:
    """Un élément de chemin (n° de facture, code école, onglet) : non vide, sans séparateur."""
    texte = str(valeur).strip()
    if not texte or texte in (".", "..") or any(c in texte for c in "/\\:"):
        raise ValueError(f"{quoi} invalide pour un nom de fichier : « {valeur} ».")
    return texte


def dossier_sorties(profil: "Profil") -> Path:
    """`<racine>/<ENTITE>/_sorties/` (celui du profil)."""
    return profil.dossier_sorties


def dossier_mois(profil: "Profil", mois: str) -> Path:
    """`_sorties/<AAAA-MM>/` ; un mois invalide lève `ValueError`."""
    mois_valide(mois)
    return dossier_sorties(profil) / mois


def journal_mois(profil: "Profil", mois: str) -> Path:
    """Journal du mois : préparations, décisions, validations."""
    return dossier_mois(profil, mois) / NOM_JOURNAL


def journal_ecritures(profil: "Profil") -> Path:
    """Journal des écritures : émission, abandon, registre, correspondances."""
    return dossier_sorties(profil) / NOM_JOURNAL


def dossier_emissions(profil: "Profil") -> Path:
    return dossier_sorties(profil) / "emissions"


def piece_emission(profil: "Profil", n_facture: str) -> Path:
    return dossier_emissions(profil) / f"{_nom_sur(n_facture, 'N° de facture')}.json"


def piece_finale(profil: "Profil", n_facture: str) -> Path:
    return dossier_emissions(profil) / f"{_nom_sur(n_facture, 'N° de facture')}.final.json"


def pdf_en_attente(profil: "Profil", n_facture: str) -> Path:
    """PDF produit, en attente du contrôle après écriture."""
    return dossier_emissions(profil) / f"{_nom_sur(n_facture, 'N° de facture')}.pdf"


def dossier_abandonnees(profil: "Profil") -> Path:
    return dossier_emissions(profil) / "abandonnees"


def classeur_ecole(profil: "Profil", ecole: str) -> Path:
    """`<racine>/<E>/<ECOLE>/<E>_<ECOLE>.xlsx`."""
    code = _nom_sur(ecole, "Code école")
    return profil.dossier_entite / code / f"{profil.entite}_{code}.xlsx"


def pdf_facture(profil: "Profil", ecole: str, onglet: str) -> Path:
    """`<racine>/<E>/<ECOLE>/<onglet>.pdf`."""
    code = _nom_sur(ecole, "Code école")
    nom = _nom_sur(onglet, "Nom d'onglet")
    return profil.dossier_entite / code / f"{nom}.pdf"


def dossier_audits(profil: "Profil") -> Path:
    return dossier_sorties(profil) / "audits"


def fichier_signalements(profil: "Profil") -> Path:
    return dossier_sorties(profil) / "signalements.jsonl"


def journal_tache(chemin_profil: Path | None) -> Path:
    """Journal de la tâche planifiée, à côté du profil (hors racine : il survit à un lecteur absent)."""
    profil = Path(chemin_profil) if chemin_profil else config.emplacement_par_defaut()
    return profil.parent / NOM_JOURNAL_TACHE
