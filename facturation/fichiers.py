"""fichiers.py — Écritures sûres, journaux JSON Lines et empreintes (P-0b, socle du lot 1).

Primitives partagées par la préparation, l'émission, le registre et les signalements :

- **écritures atomiques** (`ecrire_texte_atomique`, `ecrire_json_atomique`,
  `enregistrer_classeur_atomique`) : le contenu est écrit dans un fichier temporaire
  `<nom>.tmp-<pid>` du même dossier, mis sur disque (`fsync`), puis substitué à l'original par
  `os.replace`. Sur erreur, le temporaire est supprimé et l'original reste intact ;
- **création exclusive** (`creer_json_exclusif`) : un fichier qui existe déjà n'est jamais
  remplacé ;
- **JSON Lines** (`ajouter_jsonl`, `lire_jsonl`) : ajout seul, une ligne compacte par appel ;
  une dernière ligne tronquée (écriture interrompue) est ignorée avec un avertissement ;
- **empreintes** (`empreinte`, `empreinte_fichier`) : sha256 du JSON canonique (clés triées,
  séparateurs `(",", ":")`, `ensure_ascii=False`, encodé en UTF-8) ; aucun `float` n'y entre ;
- **secrets** (`masquer_secrets`) : toute adresse `schéma://…` devient « <adresse masquée> ».

Lectures retenues (`💡 [proposé — à valider]`) : `en_json` refuse un `set` (non figé, donc sans
ordre stable) et un nombre non fini (`float` ou `Decimal`) ; un `datetime` sans fuseau s'écrit
en ISO sans décalage ; les clés d'un dictionnaire doivent être des textes.
"""

from __future__ import annotations

import dataclasses
import hashlib
import json
import math
import os
import re
import shutil
import warnings
from collections.abc import Mapping
from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path, PurePath
from typing import Any

DOSSIER_SAUVEGARDES = "sauvegardes"
_FORMAT_HORODATAGE = "%Y%m%d-%H%M%S"
_ADRESSE_RE = re.compile(r"\w+://\S+")
ADRESSE_MASQUEE = "<adresse masquée>"


class FichierAbime(ValueError):
    """Fichier d'état illisible : le message nomme le fichier et, s'il y a lieu, la ligne."""


class ClasseurVerrouille(RuntimeError):
    """Classeur d'école ouvert dans Excel ou LibreOffice, ou en lecture seule."""


# ──────────────────────────────────────────────────────────────────────
# Conversions
# ──────────────────────────────────────────────────────────────────────


def en_json(valeur: Any) -> Any:
    """Valeur convertie en types JSON, récursivement.

    - `Decimal` → texte (`"1800.00"`) ; `date` → `"AAAA-MM-JJ"` ; `datetime` → ISO (avec son
      décalage) ; `time` → `"HH:MM"` ; `Path` → texte POSIX ;
    - tuple et liste → liste ; frozenset → liste triée ; dataclass → dictionnaire ;
      dictionnaire → dictionnaire (clés textes) ;
    - texte, entier, booléen et None inchangés ; `float` fini inchangé.

    Raises:
        TypeError: tout autre type (un `set` compris), un nombre non fini, une clé qui n'est pas
            un texte.
    """
    if valeur is None or isinstance(valeur, (str, bool, int)):
        return valeur
    if isinstance(valeur, float):
        if not math.isfinite(valeur):
            raise TypeError(f"Nombre non fini refusé dans un fichier d'état : {valeur!r}.")
        return valeur
    if isinstance(valeur, Decimal):
        if not valeur.is_finite():
            raise TypeError(f"Montant non fini refusé dans un fichier d'état : {valeur!r}.")
        return str(valeur)
    if isinstance(valeur, datetime):
        return valeur.isoformat()
    if isinstance(valeur, date):
        return valeur.isoformat()
    if isinstance(valeur, time):
        return valeur.strftime("%H:%M")
    if isinstance(valeur, PurePath):
        return valeur.as_posix()
    if dataclasses.is_dataclass(valeur) and not isinstance(valeur, type):
        return {champ.name: en_json(getattr(valeur, champ.name)) for champ in dataclasses.fields(valeur)}
    if isinstance(valeur, Mapping):
        resultat = {}
        for cle, contenu in valeur.items():
            if not isinstance(cle, str):
                raise TypeError(f"Clé {cle!r} refusée dans un fichier d'état : les clés sont des textes.")
            resultat[cle] = en_json(contenu)
        return resultat
    if isinstance(valeur, (list, tuple)):
        return [en_json(element) for element in valeur]
    if isinstance(valeur, frozenset):
        return sorted(en_json(element) for element in valeur)
    raise TypeError(
        f"Type {type(valeur).__name__} refusé dans un fichier d'état ({valeur!r}) : convertissez-le "
        "d'abord (texte, nombre, date, liste ou dictionnaire)."
    )


def _sans_float(valeur: Any) -> None:
    """Lève `TypeError` si un `float` reste dans une valeur déjà passée par `en_json`."""
    if isinstance(valeur, float):
        raise TypeError(
            f"Nombre décimal {valeur!r} dans une empreinte : écrivez les heures et les taux en texte "
            "(repr), pour qu'une empreinte ne dépende jamais de l'écriture binaire d'un nombre."
        )
    if isinstance(valeur, dict):
        for contenu in valeur.values():
            _sans_float(contenu)
    elif isinstance(valeur, list):
        for element in valeur:
            _sans_float(element)


def _json_canonique(valeur: Any) -> str:
    return json.dumps(valeur, sort_keys=True, separators=(",", ":"), ensure_ascii=False, allow_nan=False)


def empreinte(donnees: Any) -> str:
    """sha256 du JSON canonique des données (après `en_json`) : stable quel que soit l'ordre des clés.

    Raises:
        TypeError: type refusé par `en_json`, ou `float` dans les données.
    """
    convertie = en_json(donnees)
    _sans_float(convertie)
    return hashlib.sha256(_json_canonique(convertie).encode("utf-8")).hexdigest()


def empreinte_fichier(chemin: Path) -> str:
    """sha256 des octets du fichier."""
    condensat = hashlib.sha256()
    with open(chemin, "rb") as source:
        for bloc in iter(lambda: source.read(1 << 20), b""):
            condensat.update(bloc)
    return condensat.hexdigest()


def masquer_secrets(texte: str) -> str:
    """Remplace toute adresse (`https://…`, `webcal://…`) par « <adresse masquée> »."""
    return _ADRESSE_RE.sub(ADRESSE_MASQUEE, texte)


# ──────────────────────────────────────────────────────────────────────
# Écritures atomiques
# ──────────────────────────────────────────────────────────────────────


def _temporaire(chemin: Path) -> Path:
    """Fichier temporaire du même dossier : `<nom>.tmp-<pid>`."""
    return chemin.with_name(f"{chemin.name}.tmp-{os.getpid()}")


def _supprimer(chemin: Path) -> None:
    try:
        chemin.unlink(missing_ok=True)
    except OSError:  # nettoyage au mieux : l'erreur d'origine compte davantage
        pass


def _ecrire_octets_atomique(chemin: Path, octets: bytes) -> None:
    chemin = Path(chemin)
    chemin.parent.mkdir(parents=True, exist_ok=True)
    temporaire = _temporaire(chemin)
    try:
        with open(temporaire, "wb") as destination:
            destination.write(octets)
            destination.flush()
            os.fsync(destination.fileno())
        os.replace(temporaire, chemin)
    except BaseException:
        _supprimer(temporaire)
        raise


def _texte_json(donnees: Any) -> str:
    return json.dumps(en_json(donnees), ensure_ascii=False, sort_keys=True, indent=2, allow_nan=False) + "\n"


def ecrire_texte_atomique(chemin: Path, texte: str) -> None:
    """Écrit le texte en UTF-8, de façon atomique (dossiers parents créés au besoin).

    Les fins de ligne sont écrites telles quelles. Sur erreur, l'original reste intact, le
    temporaire est supprimé et l'exception remonte.
    """
    _ecrire_octets_atomique(chemin, texte.encode("utf-8"))


def ecrire_json_atomique(chemin: Path, donnees: Any) -> None:
    """Écrit les données en JSON (UTF-8, accents non échappés, clés triées, indentation de 2),
    de façon atomique ; les données passent d'abord par `en_json`."""
    _ecrire_octets_atomique(chemin, _texte_json(donnees).encode("utf-8"))


def lire_json(chemin: Path) -> Any:
    """Contenu d'un fichier JSON.

    Raises:
        FileNotFoundError: le fichier n'existe pas.
        FichierAbime: le fichier n'est pas un JSON lisible ; le message nomme le fichier et la ligne.
    """
    chemin = Path(chemin)
    try:
        texte = chemin.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError as exc:
        raise FichierAbime(f"Fichier abîmé : {chemin} n'est pas du texte UTF-8 ({exc}).") from exc
    try:
        return json.loads(texte)
    except json.JSONDecodeError as exc:
        raise FichierAbime(
            f"Fichier abîmé : {chemin}, ligne {exc.lineno} : {exc.msg}. Reprenez-le depuis sa dernière "
            "copie, ou supprimez-le pour qu'il soit recréé."
        ) from exc


def creer_json_exclusif(chemin: Path, donnees: Any) -> bool:
    """Crée le fichier JSON s'il n'existe pas (`O_CREAT | O_EXCL`) ; rend False, sans rien écrire,
    s'il existe déjà. Le JSON a la forme de `ecrire_json_atomique`."""
    chemin = Path(chemin)
    octets = _texte_json(donnees).encode("utf-8")  # avant la création : un refus ne laisse rien
    chemin.parent.mkdir(parents=True, exist_ok=True)
    drapeaux = os.O_CREAT | os.O_EXCL | os.O_WRONLY | getattr(os, "O_BINARY", 0)
    try:
        descripteur = os.open(chemin, drapeaux)
    except FileExistsError:
        return False
    try:
        with os.fdopen(descripteur, "wb") as destination:
            destination.write(octets)
            destination.flush()
            os.fsync(destination.fileno())
    except BaseException:
        _supprimer(chemin)
        raise
    return True


def enregistrer_classeur_atomique(wb: Any, chemin: Path) -> None:
    """Enregistre un classeur openpyxl de façon atomique.

    Le classeur est enregistré dans un fichier temporaire du même dossier, relu
    (`load_workbook(read_only=True)`) pour s'assurer qu'il s'ouvre, puis substitué à l'original.

    Raises:
        ClasseurVerrouille: le système refuse le remplacement (classeur ouvert dans Excel, ou en
            lecture seule) ; l'original est intact et rien n'a été modifié.
    """
    from openpyxl import load_workbook

    chemin = Path(chemin)
    temporaire = _temporaire(chemin)
    try:
        wb.save(temporaire)
        # Par un fichier ouvert : openpyxl refuse un nom qui ne finit pas par .xlsx.
        with open(temporaire, "rb") as source:
            relu = load_workbook(source, read_only=True)
            relu.close()
        try:
            os.replace(temporaire, chemin)
        except PermissionError as exc:
            raise ClasseurVerrouille(
                f"Le classeur {chemin.name} est ouvert ou en lecture seule : fermez-le, puis relancez ; "
                "rien n'a été modifié."
            ) from exc
    except BaseException:
        _supprimer(temporaire)
        raise


def sauvegarder_fichier(chemin: Path, dossier: Path | None = None, *, horodatage: datetime | None = None) -> Path:
    """Copie intacte du fichier dans `dossier` (par défaut `<dossier du fichier>/sauvegardes/`).

    Le nom est `<nom>_AAAAMMJJ-HHMMSS<extension>` (heure `horodatage`, maintenant par défaut) ;
    deux sauvegardes dans la même seconde prennent un suffixe `-1`, `-2`… : aucune sauvegarde
    n'est jamais écrasée.

    Returns:
        Le chemin de la copie.
    """
    chemin = Path(chemin)
    dossier = Path(dossier) if dossier is not None else chemin.parent / DOSSIER_SAUVEGARDES
    dossier.mkdir(parents=True, exist_ok=True)
    base = f"{chemin.stem}_{(horodatage or datetime.now()).strftime(_FORMAT_HORODATAGE)}"
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
            _supprimer(copie)
            raise
        return copie


# ──────────────────────────────────────────────────────────────────────
# JSON Lines
# ──────────────────────────────────────────────────────────────────────


def ajouter_jsonl(chemin: Path, enregistrement: Mapping[str, Any]) -> None:
    """Ajoute une ligne JSON compacte au fichier (créé avec ses dossiers au besoin), puis
    `flush` et `fsync`. L'enregistrement passe d'abord par `en_json`.

    Raises:
        TypeError: l'enregistrement n'est pas un `Mapping`, ou contient un nombre non fini ou un
            type refusé par `en_json` ; rien n'est écrit.
    """
    if not isinstance(enregistrement, Mapping):
        raise TypeError(
            f"Enregistrement de journal refusé : un dictionnaire est attendu, reçu {type(enregistrement).__name__}."
        )
    ligne = json.dumps(en_json(enregistrement), ensure_ascii=False, separators=(",", ":"), allow_nan=False)
    chemin = Path(chemin)
    chemin.parent.mkdir(parents=True, exist_ok=True)
    with open(chemin, "ab") as destination:
        destination.write(ligne.encode("utf-8") + b"\n")
        destination.flush()
        os.fsync(destination.fileno())


def lire_jsonl(chemin: Path) -> list[dict[str, Any]]:
    """Enregistrements d'un fichier JSON Lines, dans l'ordre ; fichier absent → [].

    Une dernière ligne sans fin de ligne et illisible (écriture interrompue) est ignorée, avec
    un avertissement. Les lignes vides sont ignorées.

    Raises:
        FichierAbime: ligne illisible ailleurs, ou qui n'est pas un objet JSON ; le message nomme
            le fichier et le numéro de la ligne.
    """
    chemin = Path(chemin)
    try:
        octets = chemin.read_bytes()
    except FileNotFoundError:
        return []
    lignes = octets.split(b"\n")
    tronquee = lignes[-1] if lignes[-1].strip() else None  # dernière ligne sans « \n »
    if not lignes[-1].strip():
        lignes = lignes[:-1]
    enregistrements: list[dict[str, Any]] = []
    for numero, brute in enumerate(lignes, start=1):
        if not brute.strip():
            continue
        try:
            valeur = json.loads(brute.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            if tronquee is not None and numero == len(lignes):
                warnings.warn(
                    f"{chemin}, ligne {numero} : dernière ligne incomplète (écriture interrompue), ignorée.",
                    UserWarning,
                    stacklevel=2,
                )
                continue
            raise FichierAbime(f"Fichier abîmé : {chemin}, ligne {numero} : {exc}.") from exc
        if not isinstance(valeur, dict):
            raise FichierAbime(f"Fichier abîmé : {chemin}, ligne {numero} : un objet JSON est attendu.")
        enregistrements.append(valeur)
    return enregistrements
