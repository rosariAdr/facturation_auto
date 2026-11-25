"""master.py — Lecture contrôlée et écriture sûre du MASTER (F-11, ADR013, ADR032).

Lecture : `lire_master(chemin)` vérifie que les onglets obligatoires (`Coordonnées`,
`Sessions`) existent et portent toutes les colonnes de `references/MASTER_SCHEMA.md`,
puis lit chaque colonne par son nom, quel que soit son ordre. Les onglets prévus
(`Règles_facturation`, `Correspondances`, `Factures`) sont lus s'ils existent ; absents,
ils donnent une liste vide. Un MASTER mal formé lève `MasterInvalide`, qui nomme
l'onglet et la colonne fautifs.

Écriture : `ecriture_master(chemin, poste=...)` encadre toute écriture dans le MASTER.
À l'entrée du bloc, dans cet ordre :
1. refus si le classeur est ouvert (fichier `~$<nom>` d'Excel ou `.~lock.<nom>#` de
   LibreOffice à côté, ou fichier inaccessible en écriture) ;
2. pose du verrou `<ENTITE>_MASTER.verrou`, qui porte le nom du poste et l'heure ; le verrou
   d'un autre poste bloque (un seul écrivain à la fois, ADR032) ; un verrou laissé par le
   même poste, après un arrêt brutal, est remplacé avec un avertissement ;
3. copie intacte du MASTER dans `sauvegardes/<nom>_AAAAMMJJ-HHMMSS.xlsx`, avant toute
   écriture (ADR013) ; deux sauvegardes dans la même seconde prennent un suffixe `-1`, `-2`…
À la sortie, même sur erreur, le verrou est retiré.

Ce module ne remplace pas encore le chargement des coordonnées de `pcf_format`.
"""

from __future__ import annotations

import os
import platform
import shutil
import unicodedata
import warnings
import zipfile
from contextlib import closing, contextmanager
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Iterator

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

# Colonnes de MASTER_SCHEMA, onglet par onglet. Toutes sont exigées quand l'onglet est lu.
ONGLETS_OBLIGATOIRES: dict[str, tuple[str, ...]] = {
    "Coordonnées": (
        "Type", "Code", "Nom / Raison sociale", "Adresse L1", "Adresse L2", "Adresse L3",
        "Email", "Téléphone",
    ),
    "Sessions": (
        "Date_log", "École", "Période_facture", "N°_Facture", "Date_séance", "Début", "Fin",
        "Durée_h", "Promotion", "Module / UE", "Taux", "Total HT", "Summary calendrier verbatim",
    ),
}
ONGLETS_PREVUS: dict[str, tuple[str, ...]] = {
    "Règles_facturation": (
        "École", "Statut lot", "Rythme", "Début de cycle", "Date limite d'envoi",
        "Délai de paiement", "N° de contrat", "Une facture par contrat", "Bon de commande requis",
        "Colonne Jour", "Regroupement", "Canal d'envoi", "Pièces jointes", "Notes",
    ),
    "Correspondances": (
        "Clé", "Type de clé", "École", "Module", "Promotion", "Clé tarif", "Mode heures", "Valeur",
        "Gardée le", "Notes",
    ),
    "Factures": (
        "N° facture", "École", "Contrat", "Période", "Date d'émission", "Échéance", "Total HT",
        "Date d'envoi", "Date de paiement", "Statut", "Note",
    ),
}
# Onglet → attribut de `Master`.
ATTRIBUTS = {
    "Coordonnées": "coordonnees",
    "Sessions": "sessions",
    "Règles_facturation": "regles_facturation",
    "Correspondances": "correspondances",
    "Factures": "factures",
}

DOSSIER_SAUVEGARDES = "sauvegardes"
SUFFIXE_VERROU = ".verrou"
_FORMAT_HORODATAGE = "%Y%m%d-%H%M%S"

Ligne = dict[str, Any]


class MasterInvalide(ValueError):
    """Le MASTER ne respecte pas MASTER_SCHEMA : onglet ou colonne manquant, en double, etc."""


class MasterVerrouille(RuntimeError):
    """Écriture refusée : le classeur est ouvert, ou un autre poste écrit dans le MASTER."""


@dataclass
class Master:
    """Contenu lu du MASTER : une liste de lignes {en-tête: valeur} par onglet, lignes vides ignorées."""

    coordonnees: list[Ligne] = field(default_factory=list)
    sessions: list[Ligne] = field(default_factory=list)
    regles_facturation: list[Ligne] = field(default_factory=list)
    correspondances: list[Ligne] = field(default_factory=list)
    factures: list[Ligne] = field(default_factory=list)


# ──────────────────────────────────────────────────────────────────────
# Lecture
# ──────────────────────────────────────────────────────────────────────


def _nom(valeur: Any) -> str:
    """En-tête normalisé : texte NFC sans espaces autour ; '' pour une cellule vide."""
    if valeur is None:
        return ""
    return unicodedata.normalize("NFC", str(valeur)).strip()


def _est_vide(valeur: Any) -> bool:
    return valeur is None or (isinstance(valeur, str) and not valeur.strip())


def _liste(noms: list[str] | tuple[str, ...]) -> str:
    return ", ".join(f"« {n} »" for n in noms)


def _lire_onglet(ws, onglet: str, attendues: tuple[str, ...], nom_fichier: str) -> list[Ligne]:
    """Lit un onglet par le nom de ses colonnes (en-têtes en ligne 1)."""
    # Parcours fermé même sur erreur : en lecture seule, il garde le fichier ouvert (Windows).
    with closing(ws.iter_rows(values_only=True)) as lignes:
        entetes = [_nom(v) for v in next(lignes, ())]

        vues: set[str] = set()
        doublons = []
        for entete in entetes:
            if entete and entete in vues and entete not in doublons:
                doublons.append(entete)
            vues.add(entete)
        if doublons:
            pluriel = "s" if len(doublons) > 1 else ""
            raise MasterInvalide(
                f"MASTER {nom_fichier} : onglet « {onglet} », colonne{pluriel} en double : "
                f"{_liste(doublons)}. Gardez une seule colonne par en-tête (ligne 1)."
            )

        manquantes = [c for c in attendues if unicodedata.normalize("NFC", c) not in vues]
        if manquantes:
            pluriel = "s" if len(manquantes) > 1 else ""
            raise MasterInvalide(
                f"MASTER {nom_fichier} : onglet « {onglet} », colonne{pluriel} manquante{pluriel} : "
                f"{_liste(manquantes)}. Les en-têtes se lisent en ligne 1, écrits comme dans "
                "references/MASTER_SCHEMA.md."
            )

        colonnes = [(i, entete) for i, entete in enumerate(entetes) if entete]
        resultat = []
        for valeurs in lignes:
            ligne = {entete: valeurs[i] if i < len(valeurs) else None for i, entete in colonnes}
            if not all(_est_vide(v) for v in ligne.values()):
                resultat.append(ligne)
        return resultat


def lire_master(chemin: str | os.PathLike) -> Master:
    """Lit le MASTER en contrôlant sa forme (MASTER_SCHEMA).

    Raises:
        FileNotFoundError: le fichier n'existe pas.
        MasterInvalide: fichier illisible, onglet obligatoire absent, colonne manquante ou en
            double dans un onglet lu ; le message nomme l'onglet et la colonne.
    """
    chemin = Path(chemin)
    if not chemin.is_file():
        raise FileNotFoundError(f"MASTER introuvable : {chemin}")
    try:
        wb = load_workbook(chemin, read_only=True, data_only=True)
    except (zipfile.BadZipFile, InvalidFileException) as exc:
        raise MasterInvalide(f"MASTER {chemin.name} : ce n'est pas un classeur Excel lisible ({exc}).") from exc

    try:
        absents = [o for o in ONGLETS_OBLIGATOIRES if o not in wb.sheetnames]
        if absents:
            pluriel = "s" if len(absents) > 1 else ""
            raise MasterInvalide(
                f"MASTER {chemin.name} : onglet{pluriel} obligatoire{pluriel} absent{pluriel} : "
                f"{_liste(absents)} (voir references/MASTER_SCHEMA.md)."
            )
        contenu = {}
        for onglet, attendues in {**ONGLETS_OBLIGATOIRES, **ONGLETS_PREVUS}.items():
            if onglet in wb.sheetnames:
                contenu[ATTRIBUTS[onglet]] = _lire_onglet(wb[onglet], onglet, attendues, chemin.name)
    finally:
        wb.close()
    return Master(**contenu)


# ──────────────────────────────────────────────────────────────────────
# Écriture
# ──────────────────────────────────────────────────────────────────────

# Verrous posés par ce processus : chemin du verrou → contenu écrit.
_VERROUS_TENUS: dict[str, str] = {}


def poste_courant() -> str:
    """Nom de ce poste, à passer à `ecriture_master` (même valeur pour la commande et l'appli)."""
    return os.environ.get("COMPUTERNAME") or platform.node() or "poste inconnu"


def _cle(chemin: Path) -> str:
    return os.path.normcase(os.path.abspath(chemin))


def _lire_verrou(verrou: Path) -> dict[str, str]:
    """Champs du verrou (`clé = valeur` par ligne) ; vide s'il est illisible."""
    try:
        texte = verrou.read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError):
        return {}
    champs = {}
    for ligne in texte.splitlines():
        cle, signe, valeur = ligne.partition("=")
        if signe and not ligne.lstrip().startswith("#"):
            champs[cle.strip()] = valeur.strip()
    return champs


def _depuis_lisible(depuis: str) -> str:
    try:
        return datetime.fromisoformat(depuis).strftime("le %d/%m/%Y à %H:%M:%S")
    except ValueError:
        return depuis or "une heure inconnue"


def _refuser_si_ouvert(chemin: Path):
    """Refuse si Excel ou LibreOffice ont le classeur ouvert, ou s'il n'est pas modifiable."""
    for proprietaire, logiciel in (
        (chemin.parent / f"~${chemin.name}", "Excel"),
        (chemin.parent / f".~lock.{chemin.name}#", "LibreOffice"),
    ):
        if proprietaire.exists():
            raise MasterVerrouille(
                f"Le MASTER {chemin.name} est ouvert dans {logiciel} (fichier « {proprietaire.name} » "
                "présent) : fermez le classeur, puis relancez. S'il n'est ouvert sur aucun poste, ce "
                f"fichier est resté après une fermeture brutale de {logiciel} : supprimez-le."
            )
    try:
        os.close(os.open(chemin, os.O_RDWR))
    except PermissionError as exc:
        raise MasterVerrouille(
            f"Le MASTER {chemin.name} n'est pas modifiable : fermez le classeur s'il est ouvert, "
            "ou retirez sa lecture seule, puis relancez."
        ) from exc


def _poser_verrou(chemin: Path, poste: str) -> tuple[Path, str]:
    """Pose le verrou de façon atomique ; renvoie son chemin et le contenu écrit."""
    verrou = chemin.with_suffix(SUFFIXE_VERROU)
    for _ in range(3):
        contenu = (
            "# Verrou d'écriture du MASTER (facturation_auto) : retiré à la fin de l'écriture.\n"
            f"poste = {poste}\n"
            f"depuis = {datetime.now().isoformat(sep=' ', timespec='seconds')}\n"
            f"pid = {os.getpid()}\n"
        )
        try:
            descripteur = os.open(verrou, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            pass
        else:
            with os.fdopen(descripteur, "w", encoding="utf-8", newline="\n") as f:
                f.write(contenu)
            _VERROUS_TENUS[_cle(verrou)] = contenu
            return verrou, contenu

        champs = _lire_verrou(verrou)
        if not verrou.exists():  # libéré entre-temps : nouvel essai
            continue
        autre = champs.get("poste", "")
        depuis = _depuis_lisible(champs.get("depuis", ""))
        if not autre:
            raise MasterVerrouille(
                f"Le MASTER {chemin.name} est verrouillé (fichier « {verrou.name} »), sans indication du "
                "poste qui écrit. Si aucune écriture n'est en cours, supprimez ce fichier, puis relancez."
            )
        if autre.casefold() != poste.casefold():
            raise MasterVerrouille(
                f"Le MASTER {chemin.name} est en cours d'écriture sur le poste « {autre} », depuis "
                f"{depuis} (verrou « {verrou.name} »). Attendez la fin de cette écriture, puis relancez. "
                "Si rien n'est en cours sur ce poste (arrêt brutal), supprimez le fichier verrou."
            )
        if _cle(verrou) in _VERROUS_TENUS:
            raise MasterVerrouille(
                f"Le MASTER {chemin.name} est déjà en cours d'écriture par ce programme sur le poste "
                f"« {autre} », depuis {depuis} : terminez cette écriture avant d'en commencer une autre."
            )
        warnings.warn(
            f"Verrou laissé par ce poste (« {autre} ») {depuis}, sans doute après un arrêt brutal : "
            f"il est remplacé (« {verrou.name} »).",
            stacklevel=4,
        )
        verrou.unlink(missing_ok=True)
    raise MasterVerrouille(
        f"Impossible de poser le verrou « {verrou.name} » du MASTER {chemin.name} : réessayez."
    )


def _retirer_verrou(verrou: Path, contenu: str):
    """Retire le verrou s'il est toujours le nôtre."""
    _VERROUS_TENUS.pop(_cle(verrou), None)
    try:
        actuel = verrou.read_text(encoding="utf-8")
    except FileNotFoundError:
        return
    if actuel == contenu:
        verrou.unlink(missing_ok=True)
    else:
        warnings.warn(
            f"Le verrou « {verrou.name} » a été remplacé pendant l'écriture : il est laissé en place.",
            stacklevel=4,
        )


def _sauvegarder(chemin: Path) -> Path:
    """Copie intacte du MASTER dans `sauvegardes/`, sans jamais écraser une sauvegarde."""
    dossier = chemin.parent / DOSSIER_SAUVEGARDES
    dossier.mkdir(exist_ok=True)
    base = f"{chemin.stem}_{datetime.now().strftime(_FORMAT_HORODATAGE)}"
    rang = 0
    while True:
        copie = dossier / f"{base}{f'-{rang}' if rang else ''}{chemin.suffix}"
        try:
            destination = open(copie, "xb")
        except FileExistsError:
            rang += 1
            continue
        try:
            with destination, open(chemin, "rb") as source:
                shutil.copyfileobj(source, destination)
                destination.flush()
                os.fsync(destination.fileno())
        except BaseException:
            copie.unlink(missing_ok=True)
            raise
        return copie


@contextmanager
def ecriture_master(chemin: str | os.PathLike, *, poste: str) -> Iterator[Path]:
    """Encadre une écriture dans le MASTER : refus si ouvert ou verrouillé, verrou, sauvegarde.

    Usage :
        with ecriture_master(chemin, poste=poste_courant()) as sauvegarde:
            ...  # écrire dans le MASTER (openpyxl)

    Args:
        chemin: chemin du MASTER (`<ENTITE>_MASTER.xlsx`).
        poste: nom du poste qui écrit, inscrit dans le verrou.

    Yields:
        Le chemin de la sauvegarde faite à l'entrée du bloc.

    Raises:
        FileNotFoundError: le MASTER n'existe pas.
        MasterVerrouille: classeur ouvert, verrou d'un autre poste, ou écriture déjà en cours
            dans ce programme ; rien n'est sauvegardé ni verrouillé.
    """
    chemin = Path(chemin)
    poste = str(poste).strip()
    if not poste or "\n" in poste or "\r" in poste:
        raise ValueError(f"Nom de poste invalide : {poste!r}.")
    if not chemin.is_file():
        raise FileNotFoundError(f"MASTER introuvable : {chemin}")

    _refuser_si_ouvert(chemin)
    verrou, contenu = _poser_verrou(chemin, poste)
    try:
        sauvegarde = _sauvegarder(chemin)
        yield sauvegarde
    finally:
        _retirer_verrou(verrou, contenu)
