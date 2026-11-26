"""master.py — Lecture contrôlée et écriture sûre du MASTER (F-11, ADR013, ADR032).

Lecture : `lire_master(chemin)` vérifie que les onglets obligatoires (`Coordonnées`,
`Sessions`) existent et portent toutes les colonnes de `references/MASTER_SCHEMA.md`,
puis lit chaque colonne par son nom, quel que soit son ordre. Les onglets prévus
(`Règles_facturation`, `Correspondances`, `Factures`) sont lus s'ils existent ; absents,
ils donnent une liste vide. Un MASTER mal formé lève `MasterInvalide`, qui nomme
l'onglet et la colonne fautifs.

Les onglets de suivi (`Tarifs`, `Écoles`, `Cours`, P-0b) sont lus de façon tolérante, car un
onglet de suivi « ne bloque pas » (MASTER_SCHEMA) : absent, il donne une liste vide ; une
colonne exigée manquante, ou une colonne en double, donne une liste vide et un message dans
`Master.suivi_invalide[onglet]`, sans lever. Un fichier abîmé lève toujours `MasterInvalide`.
`Master.rangs[attribut]` donne la ligne Excel de chaque élément de chaque liste.

Écriture : `ecriture_master(chemin, poste=...)` encadre toute écriture dans le MASTER.
`verrou_ecriture(chemin, poste=...)` ne pose que le verrou, pour les écritures sous `_sorties/`.
`verifier_classeur_non_ouvert(chemin)` refuse un classeur d'école ouvert ou en lecture seule.
`fidelite_aller_retour(chemin)` dit ce qu'un enregistrement par openpyxl perdrait.
À l'entrée du bloc d'`ecriture_master`, dans cet ordre :
1. refus si le classeur est ouvert (fichier `~$<nom>` d'Excel ou `.~lock.<nom>#` de
   LibreOffice à côté, ou fichier inaccessible en écriture) ;
2. pose du verrou `<ENTITE>_MASTER.verrou`, qui porte le nom du poste, l'heure et le numéro
   du processus ; le verrou d'un autre poste bloque (un seul écrivain à la fois, ADR032), comme
   celui d'une autre instance de l'outil qui tourne encore sur ce poste ; un verrou laissé par
   le même poste, après un arrêt brutal, est remplacé avec un avertissement ;
3. copie intacte du MASTER dans `sauvegardes/<nom>_AAAAMMJJ-HHMMSS.xlsx`, avant toute
   écriture (ADR013) ; deux sauvegardes dans la même seconde prennent un suffixe `-1`, `-2`…
À la sortie, même sur erreur, le verrou est retiré.

Ce module ne remplace pas encore le chargement des coordonnées de `pcf_format`.
"""

from __future__ import annotations

import os
import platform
import posixpath
import re
import shutil
import sys
import tempfile
import unicodedata
import warnings
import zipfile
import zlib
from collections.abc import Sequence
from contextlib import closing, contextmanager
from dataclasses import dataclass, field
from datetime import datetime, time, timedelta
from pathlib import Path
from typing import Any, Iterator
from xml.etree import ElementTree

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet.formula import ArrayFormula, DataTableFormula

from facturation import fichiers
from facturation.fichiers import ClasseurVerrouille

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
# Onglets de suivi lus par l'outil (P-0b) : colonnes exigées. Les autres colonnes (`Notes` et
# `Réf. contrat` de `Tarifs`, par exemple) sont lues si elles existent.
ONGLETS_SUIVI: dict[str, tuple[str, ...]] = {
    "Tarifs": ("École", "Module / Matière", "Classe / Programme", "Taux HT/h"),
    "Écoles": ("École", "Code facturation"),
    "Cours": (
        "École", "Module / Matière", "Classe / Groupe", "Volume contractuel (h)",
        "Heures dispensées (h)", "Heures facturées (h)", "Heures à facturer (h)",
        "N° Facture liée", "N° Devis lié", "Statut",
    ),
}
COLONNES_FACULTATIVES: dict[str, tuple[str, ...]] = {"Tarifs": ("Notes", "Réf. contrat")}
# Onglet → attribut de `Master`.
ATTRIBUTS = {
    "Coordonnées": "coordonnees",
    "Sessions": "sessions",
    "Règles_facturation": "regles_facturation",
    "Correspondances": "correspondances",
    "Factures": "factures",
    "Tarifs": "tarifs",
    "Écoles": "ecoles",
    "Cours": "cours",
}
# Onglets que lit l'outil : leurs formules se liraient vides après un enregistrement par openpyxl.
ONGLETS_LUS = (*ONGLETS_OBLIGATOIRES, *ONGLETS_PREVUS, *ONGLETS_SUIVI)

DOSSIER_SAUVEGARDES = "sauvegardes"
SUFFIXE_VERROU = ".verrou"

# Erreurs d'un xlsx abîmé, levées par openpyxl, zipfile ou zlib à l'ouverture comme pendant la
# lecture des lignes : partie absente de l'archive (KeyError), XML corrompu (SyntaxError, dont
# héritent ParseError d'ElementTree et XMLSyntaxError de lxml), archive ou compression abîmée,
# valeur ou attribut invalide (ValueError, TypeError), fichier qui ne se lit plus (OSError).
# Elles ne sont interceptées qu'autour des appels à openpyxl, jamais autour du code de ce module.
_ERREURS_DE_LECTURE = (
    KeyError, SyntaxError, ValueError, TypeError, IndexError, EOFError, OSError,
    zipfile.BadZipFile, zipfile.LargeZipFile, zlib.error, InvalidFileException,
)

Ligne = dict[str, Any]


class MasterInvalide(ValueError):
    """Le MASTER ne respecte pas MASTER_SCHEMA (onglet ou colonne manquant, en double…) ou est abîmé."""


class MasterVerrouille(RuntimeError):
    """Classeur inaccessible : ouvert dans Excel ou LibreOffice, ou un autre poste ou une autre
    instance de l'outil écrit dans le MASTER."""


@dataclass
class Master:
    """Contenu lu du MASTER : une liste de lignes {en-tête: valeur} par onglet, lignes vides ignorées.

    `rangs[attribut]` donne la ligne Excel de chaque élément de la liste du même nom (lignes
    vides comprises dans le compte) ; `suivi_invalide[onglet]` dit pourquoi un onglet de suivi
    n'a pas été lu.
    """

    coordonnees: list[Ligne] = field(default_factory=list)
    sessions: list[Ligne] = field(default_factory=list)
    regles_facturation: list[Ligne] = field(default_factory=list)
    correspondances: list[Ligne] = field(default_factory=list)
    factures: list[Ligne] = field(default_factory=list)
    tarifs: list[Ligne] = field(default_factory=list)
    ecoles: list[Ligne] = field(default_factory=list)
    cours: list[Ligne] = field(default_factory=list)
    rangs: dict[str, list[int]] = field(default_factory=dict)
    suivi_invalide: dict[str, str] = field(default_factory=dict)


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


def _abime(nom_fichier: str, exc: BaseException, onglet: str | None = None) -> MasterInvalide:
    """Erreur qui nomme le fichier abîmé (et l'onglet, si l'erreur vient de sa lecture)."""
    ou = f", onglet « {onglet} »" if onglet else ""
    return MasterInvalide(
        f"MASTER {nom_fichier}{ou} : classeur abîmé ou illisible ({type(exc).__name__} : {exc}). "
        "Ouvrez-le dans Excel, qui propose de le réparer, ou reprenez sa dernière copie du dossier "
        f"« {DOSSIER_SAUVEGARDES} »."
    )


def _lignes(ws, onglet: str, nom_fichier: str) -> Iterator[tuple]:
    """Valeurs des lignes de l'onglet ; une erreur de lecture du fichier devient `MasterInvalide`."""
    # En lecture seule, openpyxl se fie à la dimension déclarée par le fichier, qui peut être
    # fausse (A1:A1, ligne d'en-tête seule) : il refuserait un MASTER valide ou perdrait des lignes.
    ws.reset_dimensions()
    lignes = ws.iter_rows(values_only=True)
    try:
        while True:
            try:
                valeurs = next(lignes)
            except StopIteration:
                return
            except _ERREURS_DE_LECTURE as exc:  # feuille abîmée plus loin que son début
                raise _abime(nom_fichier, exc, onglet) from exc
            yield valeurs
    finally:
        lignes.close()


def _erreur_entetes(entetes: list[str], attendues: Sequence[str], onglet: str, nom_fichier: str) -> str | None:
    """Message d'en-têtes invalides (colonne en double, colonne attendue manquante) ; None s'ils vont."""
    vues: set[str] = set()
    doublons = []
    for entete in entetes:
        if entete and entete in vues and entete not in doublons:
            doublons.append(entete)
        vues.add(entete)
    if doublons:
        pluriel = "s" if len(doublons) > 1 else ""
        return (
            f"MASTER {nom_fichier} : onglet « {onglet} », colonne{pluriel} en double : "
            f"{_liste(doublons)}. Gardez une seule colonne par en-tête (ligne 1)."
        )
    manquantes = [c for c in attendues if unicodedata.normalize("NFC", c) not in vues]
    if manquantes:
        pluriel = "s" if len(manquantes) > 1 else ""
        return (
            f"MASTER {nom_fichier} : onglet « {onglet} », colonne{pluriel} manquante{pluriel} : "
            f"{_liste(manquantes)}. Les en-têtes se lisent en ligne 1, écrits comme dans "
            "references/MASTER_SCHEMA.md."
        )
    return None


def _lire_onglet(
    ws, onglet: str, attendues: tuple[str, ...], nom_fichier: str, *, tolerant: bool = False,
) -> tuple[list[Ligne], list[int], str | None]:
    """Lit un onglet par le nom de ses colonnes (en-têtes en ligne 1).

    Returns:
        (lignes non vides, ligne Excel de chacune, None). Avec `tolerant`, des en-têtes invalides
        donnent ([], [], message) au lieu de lever.

    Raises:
        MasterInvalide: en-têtes invalides (sans `tolerant`) ; feuille abîmée (toujours).
    """
    # Parcours fermé même sur erreur : en lecture seule, il garde le fichier ouvert (Windows).
    with closing(_lignes(ws, onglet, nom_fichier)) as lignes:
        entetes = [_nom(v) for v in next(lignes, ())]
        erreur = _erreur_entetes(entetes, attendues, onglet, nom_fichier)
        if erreur is not None:
            if tolerant:
                return [], [], erreur
            raise MasterInvalide(erreur)

        colonnes = [(i, entete) for i, entete in enumerate(entetes) if entete]
        resultat = []
        rangs = []
        # Les lignes absentes du fichier sont rendues vides par openpyxl : le compte suit Excel.
        for rang, valeurs in enumerate(lignes, start=2):
            ligne = {entete: valeurs[i] if i < len(valeurs) else None for i, entete in colonnes}
            if not all(_est_vide(v) for v in ligne.values()):
                resultat.append(ligne)
                rangs.append(rang)
        return resultat, rangs, None


def lire_master(chemin: str | os.PathLike) -> Master:
    """Lit le MASTER en contrôlant sa forme (MASTER_SCHEMA).

    Les onglets de suivi (`Tarifs`, `Écoles`, `Cours`) sont lus de façon tolérante : absents,
    ils donnent une liste vide ; une colonne exigée manquante ou une colonne en double donne une
    liste vide et un message dans `suivi_invalide`, sans lever.

    Raises:
        FileNotFoundError: le fichier n'existe pas.
        MasterVerrouille: le système refuse l'ouverture (classeur ouvert dans Excel, droits).
        MasterInvalide: fichier abîmé ou illisible (le message nomme le fichier), onglet
            obligatoire absent, colonne manquante ou en double dans un onglet obligatoire ou
            prévu (le message nomme l'onglet et la colonne).
    """
    chemin = Path(chemin)
    if not chemin.is_file():
        raise FileNotFoundError(f"MASTER introuvable : {chemin}")
    try:
        wb = load_workbook(chemin, read_only=True, data_only=True)
    except PermissionError as exc:  # avant OSError, qui figure dans _ERREURS_DE_LECTURE
        raise MasterVerrouille(
            f"Le MASTER {chemin.name} ne s'ouvre pas (accès refusé) : fermez le classeur s'il est "
            "ouvert dans Excel, ou vérifiez vos droits sur le fichier, puis relancez."
        ) from exc
    except _ERREURS_DE_LECTURE as exc:
        raise _abime(chemin.name, exc) from exc

    try:
        absents = [o for o in ONGLETS_OBLIGATOIRES if o not in wb.sheetnames]
        if absents:
            pluriel = "s" if len(absents) > 1 else ""
            raise MasterInvalide(
                f"MASTER {chemin.name} : onglet{pluriel} obligatoire{pluriel} absent{pluriel} : "
                f"{_liste(absents)} (voir references/MASTER_SCHEMA.md)."
            )
        contenu: dict[str, Any] = {}
        rangs: dict[str, list[int]] = {attribut: [] for attribut in ATTRIBUTS.values()}
        suivi_invalide: dict[str, str] = {}
        for onglet, attendues in {**ONGLETS_OBLIGATOIRES, **ONGLETS_PREVUS, **ONGLETS_SUIVI}.items():
            if onglet not in wb.sheetnames:
                continue
            attribut = ATTRIBUTS[onglet]
            lignes, rangs[attribut], erreur = _lire_onglet(
                wb[onglet], onglet, attendues, chemin.name, tolerant=onglet in ONGLETS_SUIVI,
            )
            contenu[attribut] = lignes
            if erreur is not None:
                suivi_invalide[onglet] = erreur
    finally:
        wb.close()
    return Master(**contenu, rangs=rangs, suivi_invalide=suivi_invalide)


def index_colonnes(ws, attendues: Sequence[str], *, onglet: str) -> dict[str, int]:
    """Numéro de colonne (openpyxl, à partir de 1) de chaque en-tête de la ligne 1.

    Les en-têtes sont lus comme par `lire_master` (NFC, sans espaces autour). Un en-tête en
    double qui n'est pas attendu garde sa première colonne.

    Raises:
        MasterInvalide: colonne attendue absente ou en double ; le message nomme l'onglet et la
            colonne.
    """
    premiere = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
    index: dict[str, int] = {}
    doublons: list[str] = []
    cherchees = {unicodedata.normalize("NFC", c).strip() for c in attendues}
    for numero, valeur in enumerate(premiere, start=1):
        entete = _nom(valeur)
        if not entete:
            continue
        if entete in index:
            if entete in cherchees and entete not in doublons:
                doublons.append(entete)
            continue
        index[entete] = numero
    if doublons:
        pluriel = "s" if len(doublons) > 1 else ""
        raise MasterInvalide(
            f"Onglet « {onglet} » : colonne{pluriel} en double : {_liste(doublons)}. Gardez une seule "
            "colonne par en-tête (ligne 1)."
        )
    manquantes = [c for c in attendues if unicodedata.normalize("NFC", c).strip() not in index]
    if manquantes:
        pluriel = "s" if len(manquantes) > 1 else ""
        raise MasterInvalide(
            f"Onglet « {onglet} » : colonne{pluriel} manquante{pluriel} : {_liste(manquantes)}. Les "
            "en-têtes se lisent en ligne 1, écrits comme dans references/MASTER_SCHEMA.md."
        )
    return index


def derniere_ligne_remplie(ws) -> int:
    """Numéro de la dernière ligne qui porte une valeur (0 pour une feuille vide).

    Une cellule seulement mise en forme, ou faite d'espaces, ne compte pas.
    """
    cellules = getattr(ws, "_cells", None)
    if cellules is not None:  # feuille ordinaire : ses cellules, sans en créer de nouvelles
        return max((ligne for (ligne, _), cellule in cellules.items() if not _est_vide(cellule.value)), default=0)
    derniere = 0  # feuille en lecture seule
    for numero, valeurs in enumerate(ws.iter_rows(values_only=True), start=1):
        if any(not _est_vide(v) for v in valeurs):
            derniere = numero
    return derniere


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


def _pid_du_verrou(champs: dict[str, str]) -> int | None:
    """Numéro du processus noté dans le verrou ; None s'il est absent ou illisible (tenu pour mort)."""
    texte = champs.get("pid", "")
    if not (texte.isascii() and texte.isdigit()):
        return None
    return int(texte) or None


# Windows : droit d'accès minimal pour lire l'état d'un processus, code de sortie « encore actif ».
_PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
_STILL_ACTIVE = 259
_ERROR_ACCESS_DENIED = 5


def _processus_vivant(pid: int) -> bool:
    """Le processus `pid` tourne-t-il encore ? Un accès refusé compte comme vivant (prudence)."""
    if sys.platform == "win32":
        return _processus_vivant_windows(pid)
    try:
        os.kill(pid, 0)  # signal 0 : teste l'existence du processus sans rien lui envoyer
    except ProcessLookupError:
        return False
    except PermissionError:  # il existe, mais appartient à un autre utilisateur
        return True
    except (OverflowError, ValueError):  # hors des valeurs possibles d'un pid
        return False
    return True


def _processus_vivant_windows(pid: int) -> bool:
    # Jamais `os.kill` sous Windows : quel que soit le signal, il termine le processus visé.
    import ctypes
    from ctypes import wintypes

    if pid > 0xFFFFFFFF:  # un pid Windows tient sur 32 bits
        return False
    noyau = ctypes.WinDLL("kernel32", use_last_error=True)
    noyau.OpenProcess.argtypes = (wintypes.DWORD, wintypes.BOOL, wintypes.DWORD)
    noyau.OpenProcess.restype = wintypes.HANDLE
    noyau.GetExitCodeProcess.argtypes = (wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD))
    noyau.GetExitCodeProcess.restype = wintypes.BOOL
    noyau.CloseHandle.argtypes = (wintypes.HANDLE,)
    noyau.CloseHandle.restype = wintypes.BOOL

    poignee = noyau.OpenProcess(_PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not poignee:
        # Accès refusé : le processus existe. Paramètre invalide : aucun processus ne porte ce pid.
        return ctypes.get_last_error() == _ERROR_ACCESS_DENIED
    try:
        code = wintypes.DWORD()
        if not noyau.GetExitCodeProcess(poignee, ctypes.byref(code)):
            return True  # état inconnu : prudence, le verrou n'est pas remplacé
        # Un processus terminé reste ouvrable tant qu'une poignée le retient : seul son code de
        # sortie dit s'il tourne encore.
        return code.value == _STILL_ACTIVE
    finally:
        noyau.CloseHandle(poignee)


def _depuis_lisible(depuis: str) -> str:
    try:
        return datetime.fromisoformat(depuis).strftime("le %d/%m/%Y à %H:%M:%S")
    except ValueError:
        return depuis or "une heure inconnue"


def _proprietaire(chemin: Path) -> tuple[Path, str] | None:
    """Fichier laissé par Excel ou LibreOffice à côté d'un classeur ouvert, et le logiciel."""
    for proprietaire, logiciel in (
        (chemin.parent / f"~${chemin.name}", "Excel"),
        (chemin.parent / f".~lock.{chemin.name}#", "LibreOffice"),
    ):
        if proprietaire.exists():
            return proprietaire, logiciel
    return None


def _modifiable(chemin: Path) -> bool:
    """Le système laisse-t-il ouvrir le fichier en écriture (ni verrouillé, ni en lecture seule) ?"""
    try:
        os.close(os.open(chemin, os.O_RDWR))
    except PermissionError:
        return False
    return True


def _refuser_si_ouvert(chemin: Path):
    """Refuse si Excel ou LibreOffice ont le classeur ouvert, ou s'il n'est pas modifiable."""
    ouvert = _proprietaire(chemin)
    if ouvert is not None:
        proprietaire, logiciel = ouvert
        raise MasterVerrouille(
            f"Le MASTER {chemin.name} est ouvert dans {logiciel} (fichier « {proprietaire.name} » "
            "présent) : fermez le classeur, puis relancez. S'il n'est ouvert sur aucun poste, ce "
            f"fichier est resté après une fermeture brutale de {logiciel} : supprimez-le."
        )
    if not _modifiable(chemin):
        raise MasterVerrouille(
            f"Le MASTER {chemin.name} n'est pas modifiable : fermez le classeur s'il est ouvert, "
            "ou retirez sa lecture seule, puis relancez."
        )


def verifier_classeur_non_ouvert(chemin: str | os.PathLike) -> None:
    """Refuse un classeur d'école ouvert dans Excel ou LibreOffice, ou en lecture seule.

    Un classeur qui n'existe pas encore est accepté (il sera créé).

    Raises:
        fichiers.ClasseurVerrouille: le message nomme le classeur et dit quoi faire.
    """
    chemin = Path(chemin)
    if not chemin.exists():
        return None
    ouvert = _proprietaire(chemin)
    if ouvert is not None:
        proprietaire, logiciel = ouvert
        raise ClasseurVerrouille(
            f"Le classeur {chemin.name} est ouvert dans {logiciel} (fichier « {proprietaire.name} » "
            "présent) : fermez-le, puis relancez ; rien n'a été modifié. S'il n'est ouvert sur aucun "
            f"poste, ce fichier est resté après une fermeture brutale de {logiciel} : supprimez-le."
        )
    if not _modifiable(chemin):
        raise ClasseurVerrouille(
            f"Le classeur {chemin.name} n'est pas modifiable : fermez-le s'il est ouvert, ou retirez "
            "sa lecture seule, puis relancez ; rien n'a été modifié."
        )
    return None


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
        pid = _pid_du_verrou(champs)
        if pid is not None and pid != os.getpid() and _processus_vivant(pid):
            raise MasterVerrouille(
                f"Le MASTER {chemin.name} est verrouillé (« {verrou.name} ») : une autre instance de "
                f"l'outil écrit déjà sur ce poste (« {autre} », processus {pid}, depuis {depuis}). "
                "Attendez la fin de cette écriture, puis relancez. Si aucune autre instance n'est "
                "ouverte, supprimez le fichier verrou."
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
    return fichiers.sauvegarder_fichier(chemin, chemin.parent / DOSSIER_SAUVEGARDES, horodatage=datetime.now())


def _poste_valide(poste: Any) -> str:
    poste = str(poste).strip()
    if not poste or "\n" in poste or "\r" in poste:
        raise ValueError(f"Nom de poste invalide : {poste!r}.")
    return poste


@contextmanager
def verrou_ecriture(chemin_master: str | os.PathLike, *, poste: str) -> Iterator[Path]:
    """Pose le verrou d'écriture de l'entité (`<ENTITE>_MASTER.verrou`), sans autre contrôle.

    Pour les écritures sous `_sorties/` (préparation, décisions, validations, signalements) :
    ni refus d'un MASTER ouvert dans Excel, ni sauvegarde. Mêmes règles que le verrou
    d'`ecriture_master` : le verrou d'un autre poste, ou d'une autre instance de l'outil sur ce
    poste, bloque ; celui laissé par ce poste après un arrêt brutal est remplacé avec un
    avertissement ; il n'est pas réentrant. Il est retiré à la sortie, même sur erreur.

    Yields:
        Le chemin du verrou.

    Raises:
        MasterVerrouille: verrou d'un autre poste ou d'une autre instance, ou écriture déjà en
            cours dans ce programme.
    """
    chemin = Path(chemin_master)
    poste = _poste_valide(poste)
    verrou, contenu = _poser_verrou(chemin, poste)
    try:
        yield verrou
    finally:
        _retirer_verrou(verrou, contenu)


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
        MasterVerrouille: classeur ouvert, verrou d'un autre poste ou d'une autre instance de
            l'outil qui tourne encore sur ce poste, ou écriture déjà en cours dans ce programme ;
            rien n'est sauvegardé ni verrouillé.
    """
    chemin = Path(chemin)
    poste = _poste_valide(poste)
    if not chemin.is_file():
        raise FileNotFoundError(f"MASTER introuvable : {chemin}")

    _refuser_si_ouvert(chemin)
    with verrou_ecriture(chemin, poste=poste):
        yield _sauvegarder(chemin)


# ──────────────────────────────────────────────────────────────────────
# Fidélité de l'aller-retour openpyxl
# ──────────────────────────────────────────────────────────────────────

# Parties de l'archive dont la disparition à l'enregistrement est une perte (§ 3.5 et 17).
# Les dessins sont comptés par leurs ancres : un dessin sans ancre (export Google) ne compte pas.
_PARTIES_SURVEILLEES: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("graphique(s)", re.compile(r"^xl/charts/chart[^/]*\.xml$")),
    ("image(s) (xl/media)", re.compile(r"^xl/media/[^/]+$")),
    ("tableau(x) croisé(s)", re.compile(r"^xl/pivot(?:Tables|Cache)/[^/]+\.xml$")),
    ("liaison(s) externe(s)", re.compile(r"^xl/externalLinks/[^/]+\.xml$")),
    ("commentaire(s)", re.compile(r"^xl/comments[^/]*\.xml$|^xl/comments/[^/]+\.xml$")),
    ("commentaire(s) en fil de discussion", re.compile(r"^xl/threadedComments/[^/]+\.xml$")),
    ("segment(s) (slicers)", re.compile(r"^xl/slicer(?:s|Caches)/[^/]+\.xml$")),
    ("macro(s) VBA", re.compile(r"(?:^|/)vbaProject\.bin$")),
)
_ANCRE_RE = re.compile(rb"<(?:[\w.-]+:)?(?:twoCellAnchor|oneCellAnchor|absoluteAnchor)\b")
_VALIDATION_RE = re.compile(rb"<(?:[\w.-]+:)?dataValidation\b")
# Règles de mise en forme conditionnelle (y compris celles des extensions x14) : compter les règles
# plutôt que les blocs évite de signaler deux blocs de même plage qu'openpyxl réunit en un seul.
_REGLE_MFC_RE = re.compile(rb"<(?:[\w.-]+:)?cfRule\b")
_NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
_NS_SSML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _cible(source: str, cible: str) -> str:
    """Chemin, dans l'archive, d'une cible de relation relative à la partie `source`."""
    if cible.startswith("/"):
        return cible.lstrip("/")
    return posixpath.normpath(posixpath.join(posixpath.dirname(source), cible))


def _relations(archive: zipfile.ZipFile, partie: str) -> dict[str, str]:
    """Relations d'une partie : identifiant → chemin de la cible dans l'archive."""
    rels = posixpath.join(posixpath.dirname(partie), "_rels", posixpath.basename(partie) + ".rels")
    if rels not in archive.namelist():
        return {}
    racine = ElementTree.fromstring(archive.read(rels))
    return {
        r.get("Id", ""): _cible(partie, r.get("Target", ""))
        for r in racine.iter(f"{{{_NS_RELS}}}Relationship")
        if r.get("TargetMode") != "External"
    }


def _parties_des_feuilles(archive: zipfile.ZipFile) -> dict[str, str]:
    """Nom de chaque onglet → partie XML de sa feuille."""
    classeur = "xl/workbook.xml"
    racine_rels = ElementTree.fromstring(archive.read("_rels/.rels"))
    for relation in racine_rels.iter(f"{{{_NS_RELS}}}Relationship"):
        if relation.get("Type", "").endswith("/officeDocument"):
            classeur = _cible("", relation.get("Target", classeur))
    cibles = _relations(archive, classeur)
    feuilles = {}
    for feuille in ElementTree.fromstring(archive.read(classeur)).iter(f"{{{_NS_SSML}}}sheet"):
        cible = cibles.get(feuille.get(f"{{{_NS_REL_DOC}}}id", ""))
        if cible:
            feuilles[feuille.get("name", "")] = cible
    return feuilles


def _inventaire(chemin: Path) -> dict[str, Any]:
    """Ce que l'archive contient de surveillé : parties, ancres de dessin, validations, règles."""
    with zipfile.ZipFile(chemin) as archive:
        noms = archive.namelist()
        parties = {libelle: sum(1 for n in noms if motif.search(n)) for libelle, motif in _PARTIES_SURVEILLEES}
        ancres = sum(
            len(_ANCRE_RE.findall(archive.read(n)))
            for n in noms if re.fullmatch(r"xl/drawings/[^/]+\.xml", n)
        )
        par_feuille = {}
        for onglet, partie in _parties_des_feuilles(archive).items():
            if partie in noms:
                contenu = archive.read(partie)
                par_feuille[onglet] = (len(_VALIDATION_RE.findall(contenu)), len(_REGLE_MFC_RE.findall(contenu)))
    return {"parties": parties, "ancres": ancres, "feuilles": par_feuille}


def _valeur_comparable(valeur: Any) -> Any:
    """Valeur d'une cellule, comparable d'un chargement à l'autre (dates à la seconde près)."""
    if isinstance(valeur, datetime):
        return (valeur + timedelta(microseconds=500_000)).replace(microsecond=0)
    if isinstance(valeur, time):
        return _valeur_comparable(datetime.combine(datetime(2000, 1, 1), valeur)).time()
    if isinstance(valeur, timedelta):
        return round(valeur.total_seconds())
    if isinstance(valeur, ArrayFormula):
        return ("formule matricielle", str(valeur.ref), str(valeur.text))
    if isinstance(valeur, DataTableFormula):
        return ("table de données", tuple(sorted((cle, str(v)) for cle, v in valeur)))
    return valeur


def _toutes_les_cellules(ws) -> Iterator[Any]:
    """Cellules existantes d'une feuille chargée en entier, sans en créer de nouvelles."""
    cellules = getattr(ws, "_cells", None)
    if cellules is not None:
        return iter(list(cellules.values()))
    return (cellule for ligne in ws.iter_rows() for cellule in ligne)


def _cellules(ws) -> dict[str, Any]:
    return {
        cellule.coordinate: _valeur_comparable(cellule.value)
        for cellule in _toutes_les_cellules(ws) if cellule.value is not None
    }


def _masquees(ws) -> tuple[str, set[int], set[int]]:
    lignes = {numero for numero, dimension in ws.row_dimensions.items() if dimension.hidden}
    colonnes: set[int] = set()
    for lettre, dimension in ws.column_dimensions.items():
        if dimension.hidden:
            debut = dimension.min or column_index_from_string(lettre)
            fin = dimension.max or debut
            colonnes.update(range(debut, fin + 1))
    return ws.sheet_state, lignes, colonnes


def _noms_definis(wb) -> set[tuple[str, str, str]]:
    noms = {("classeur", nom, str(defini.attr_text)) for nom, defini in wb.defined_names.items()}
    for ws in wb.worksheets:
        noms |= {(ws.title, nom, str(defini.attr_text)) for nom, defini in ws.defined_names.items()}
        for nom, valeur in (("zone d'impression", ws.print_area), ("lignes à répéter", ws.print_title_rows),
                            ("colonnes à répéter", ws.print_title_cols)):
            if valeur:
                noms.add((ws.title, nom, str(valeur)))
    return noms


def _comparer_classeurs(avant, apres, pertes: list[str]) -> None:
    """Onglets et leur ordre, valeurs et formules, fusions, noms définis, états masqués."""
    if avant.sheetnames != apres.sheetnames:
        pertes.append(
            f"onglets ou leur ordre changés : {', '.join(avant.sheetnames)} → {', '.join(apres.sheetnames)}"
        )
    titres = [ws.title for ws in avant.worksheets if ws.title in apres.sheetnames]
    for titre in titres:
        ws_avant, ws_apres = avant[titre], apres[titre]
        if not hasattr(ws_apres, "iter_rows"):
            pertes.append(f"onglet « {titre} » : n'est plus une feuille de calcul")
            continue
        cellules_avant, cellules_apres = _cellules(ws_avant), _cellules(ws_apres)
        changees = sorted(
            {c for c in cellules_avant.keys() | cellules_apres.keys() if cellules_avant.get(c) != cellules_apres.get(c)},
            key=lambda c: (int(re.sub(r"\D", "", c)), c),
        )
        if changees:
            exemple = ", ".join(changees[:3])
            pertes.append(
                f"onglet « {titre} » : {len(changees)} cellule(s) changent de valeur ou de formule (ex. {exemple})"
            )
        fusions_avant = {str(plage) for plage in ws_avant.merged_cells.ranges}
        fusions_apres = {str(plage) for plage in ws_apres.merged_cells.ranges}
        if fusions_avant != fusions_apres:
            pertes.append(f"onglet « {titre} » : cellules fusionnées changées ({len(fusions_avant)} avant, "
                          f"{len(fusions_apres)} après)")
        etat_avant, lignes_avant, colonnes_avant = _masquees(ws_avant)
        etat_apres, lignes_apres, colonnes_apres = _masquees(ws_apres)
        if etat_avant != etat_apres:
            pertes.append(f"onglet « {titre} » : état « {etat_avant} » devenu « {etat_apres} »")
        if lignes_avant != lignes_apres or colonnes_avant != colonnes_apres:
            pertes.append(f"onglet « {titre} » : lignes ou colonnes masquées changées")
    noms_avant, noms_apres = _noms_definis(avant), _noms_definis(apres)
    if noms_avant - noms_apres:
        perdus = ", ".join(sorted(f"{nom} ({portee})" for portee, nom, _ in noms_avant - noms_apres))
        pertes.append(f"nom(s) défini(s) perdu(s) ou changé(s) : {perdus}")


def _formules(wb) -> list[str]:
    """Une ligne par onglet lu par l'outil qui porte des formules : leur valeur en cache se perd."""
    lignes = []
    for onglet in ONGLETS_LUS:
        if onglet not in wb.sheetnames or not hasattr(wb[onglet], "iter_rows"):
            continue
        nombre = sum(
            1 for cellule in _toutes_les_cellules(wb[onglet])
            if cellule.data_type == "f" or isinstance(cellule.value, (ArrayFormula, DataTableFormula))
        )
        if nombre:
            lignes.append(
                f"onglet « {onglet} » : {nombre} cellule(s) à formule ; un enregistrement par l'outil n'en "
                "garde pas la valeur calculée : elles se liront vides pour l'outil jusqu'au prochain "
                "enregistrement par Excel"
            )
    return lignes


def fidelite_aller_retour(chemin: str | os.PathLike) -> list[str]:
    """Ce qu'un enregistrement par openpyxl perdrait dans le classeur ; liste vide si rien.

    Le classeur est copié dans un dossier temporaire, hors de son dossier ; l'original n'est
    jamais touché. La copie est chargée (`data_only=False`), enregistrée, puis comparée :
    - dans l'archive : graphiques, images, tableaux croisés, liaisons externes, commentaires,
      segments, macros ; dessins avec ancre (un dessin vide, sans ancre, ne compte pas, comme
      `xl/metadata`, parties vides des exports Google Sheets) ; nombre de validations de
      données et de règles de mise en forme conditionnelle par feuille ;
    - dans les classeurs : onglets et leur ordre, valeurs et formules, fusions, noms définis
      (zones d'impression comprises), états masqués ;
    - formules des onglets lus par l'outil (MASTER_SCHEMA) : leur valeur en cache se perd.

    Returns:
        Une ligne par perte, en français.

    Raises:
        FileNotFoundError: le classeur n'existe pas.
        MasterInvalide: classeur abîmé ou illisible.
    """
    chemin = Path(chemin)
    if not chemin.is_file():
        raise FileNotFoundError(f"Classeur introuvable : {chemin}")
    with tempfile.TemporaryDirectory(prefix="facturation_fidelite_", ignore_cleanup_errors=True) as dossier:
        copie = Path(dossier) / f"copie{chemin.suffix or '.xlsx'}"
        enregistree = Path(dossier) / f"enregistree{chemin.suffix or '.xlsx'}"
        shutil.copyfile(chemin, copie)
        try:
            with warnings.catch_warnings():
                # openpyxl prévient de ce qu'il abandonne : c'est justement ce qui est mesuré ici.
                warnings.simplefilter("ignore")
                avant = load_workbook(copie, data_only=False)
                avant.save(enregistree)
                apres = load_workbook(enregistree, data_only=False)
            inventaire_avant, inventaire_apres = _inventaire(copie), _inventaire(enregistree)
        except _ERREURS_DE_LECTURE as exc:
            raise _abime(chemin.name, exc) from exc

        pertes: list[str] = []
        for libelle, nombre in inventaire_avant["parties"].items():
            reste = inventaire_apres["parties"].get(libelle, 0)
            if reste < nombre:
                pertes.append(f"{libelle} perdu(s) à l'enregistrement : {nombre} avant, {reste} après")
        if inventaire_apres["ancres"] < inventaire_avant["ancres"]:
            pertes.append(
                f"dessin(s) avec ancre (formes, images ou graphiques posés sur la feuille) perdu(s) : "
                f"{inventaire_avant['ancres']} avant, {inventaire_apres['ancres']} après"
            )
        for onglet, (validations, regles) in inventaire_avant["feuilles"].items():
            validations_apres, regles_apres = inventaire_apres["feuilles"].get(onglet, (0, 0))
            if validations_apres < validations:
                pertes.append(f"onglet « {onglet} » : validation(s) de données perdue(s) : "
                              f"{validations} avant, {validations_apres} après")
            if regles_apres < regles:
                pertes.append(f"onglet « {onglet} » : règle(s) de mise en forme conditionnelle perdue(s) : "
                              f"{regles} avant, {regles_apres} après")
        _comparer_classeurs(avant, apres, pertes)
        pertes.extend(_formules(avant))
        avant.close()
        apres.close()
    return pertes
