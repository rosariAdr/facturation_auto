"""numerotation.py — Numérotation automatique des factures (F-15, ADR002, ADR029).

Format d'un numéro : `<préfixe d'entité><CODE ÉCOLE><NNN>`, par exemple `R_ABC008`
avec le préfixe `R_`. La séquence de 3 chiffres est propre à chaque école et à
chaque entité : `R_ABC`, `K_ABC` et `R_ABCD` sont trois séries distinctes.

Règles (ADR029) :
- le numéro suivant est le plus grand numéro de la série plus un ;
- un numéro manquant n'est jamais réutilisé : il est signalé pour que sa raison
  soit notée, et la série continue après le plus grand numéro ;
- les codes école existants sont conservés quelle que soit leur longueur ; une
  nouvelle école reçoit un code de 3 lettres.

La variante historique avec un tiret bas (`R_ABC_007`) compte pour trouver le plus
grand numéro, mais le numéro proposé est toujours au format courant (`R_ABC008`).

Ce module ne fait que lire : il n'attribue ni n'écrit aucun numéro.
"""

from __future__ import annotations

import re
from collections.abc import Iterable
from pathlib import Path

from openpyxl import load_workbook

# Dernier numéro possible d'une série : la séquence tient sur 3 chiffres (ADR002).
NUMERO_MAX = 999

# Code école existant : majuscules et chiffres, longueur libre (ADR029).
_CODE_RE = re.compile(r"[A-Z0-9]+")
# Code d'une nouvelle école : exactement 3 lettres (ADR029).
_CODE_NOUVEAU_RE = re.compile(r"[A-Z]{3}")
# Préfixe d'entité : même règle que `prefixe_facture` dans le profil (1 à 3 majuscules + « _ »).
_PREFIXE_RE = re.compile(r"[A-Z]{1,3}_")

# Où lire les numéros déjà attribués dans le MASTER (MASTER_SCHEMA) : onglet, colonne.
_ONGLET_SESSIONS = ("Sessions", "N°_Facture")
_ONGLET_FACTURES = ("Factures", "N° facture")


def valider_code_ecole(code: str, *, nouveau: bool = False) -> str:
    """Vérifie le format d'un code école et le renvoie.

    Args:
        code: code école, par exemple `ABC`.
        nouveau: True pour le code d'une nouvelle école, qui doit compter
            exactement 3 lettres (ADR029).

    Raises:
        ValueError: code vide ou contenant autre chose que des majuscules et des
            chiffres ; avec `nouveau=True`, code qui n'est pas fait de 3 lettres.
    """
    if not isinstance(code, str) or not _CODE_RE.fullmatch(code):
        raise ValueError(
            f"Code école invalide : {code!r}. Attendu : majuscules et chiffres, "
            "sans espace ni tiret (ex. ABC)."
        )
    if nouveau and not _CODE_NOUVEAU_RE.fullmatch(code):
        raise ValueError(
            f"Code de nouvelle école invalide : {code!r}. Attendu : exactement "
            "3 lettres majuscules (ex. ABC), unique dans l'entité (ADR029)."
        )
    return code


def _valider_prefixe(prefixe: str) -> str:
    if not isinstance(prefixe, str) or not _PREFIXE_RE.fullmatch(prefixe):
        raise ValueError(
            f"Préfixe d'entité invalide : {prefixe!r}. Attendu : 1 à 3 majuscules "
            "suivies de « _ » (ex. R_)."
        )
    return prefixe


def _rangs(numeros: Iterable, prefixe: str, code_ecole: str) -> set[int]:
    """Rangs (partie NNN) des numéros de la série `prefixe` + `code_ecole`.

    Les numéros d'autres écoles ou d'autres préfixes, les valeurs qui ne sont pas
    du texte et les numéros mal formés sont ignorés.
    """
    _valider_prefixe(prefixe)
    valider_code_ecole(code_ecole)
    motif = re.compile(rf"{re.escape(prefixe)}{re.escape(code_ecole)}_?(\d{{3}})")
    rangs = set()
    for numero in numeros:
        if not isinstance(numero, str):
            continue
        trouve = motif.fullmatch(numero.strip())
        if trouve:
            rangs.add(int(trouve.group(1)))
    return rangs


def _format(prefixe: str, code_ecole: str, rang: int) -> str:
    return f"{prefixe}{code_ecole}{rang:03d}"


def prochain_numero(numeros: Iterable, *, prefixe: str, code_ecole: str) -> str:
    """Numéro suivant de la série : le plus grand numéro plus un, sur 3 chiffres.

    Une série vide commence à 001. Les trous ne sont jamais comblés (ADR029).

    Args:
        numeros: numéros déjà attribués, toutes écoles et entités confondues
            (par exemple le résultat de `lire_numeros`).
        prefixe: préfixe d'entité, par exemple `R_`.
        code_ecole: code de l'école, par exemple `ABC`.

    Raises:
        ValueError: préfixe ou code école mal formé, ou série déjà arrivée à 999.
    """
    rangs = _rangs(numeros, prefixe, code_ecole)
    plus_grand = max(rangs, default=0)
    if plus_grand >= NUMERO_MAX:
        raise ValueError(
            f"Série {prefixe}{code_ecole} complète : {_format(prefixe, code_ecole, plus_grand)} "
            f"est le dernier numéro possible sur 3 chiffres. Aucun numéro suivant à proposer."
        )
    return _format(prefixe, code_ecole, plus_grand + 1)


def numeros_manquants(numeros: Iterable, *, prefixe: str, code_ecole: str) -> list[str]:
    """Trous de la série, au format courant et dans l'ordre (ADR029).

    Tous les numéros absents entre 001 et le plus grand numéro de la série, y
    compris ceux qui précèdent le plus petit numéro présent. Ils sont à noter,
    jamais à réutiliser. Série vide : aucun trou.

    Raises:
        ValueError: préfixe ou code école mal formé.
    """
    rangs = _rangs(numeros, prefixe, code_ecole)
    plus_grand = max(rangs, default=0)
    return [_format(prefixe, code_ecole, rang) for rang in range(1, plus_grand + 1) if rang not in rangs]


def _texte(valeur) -> str | None:
    """Contenu d'une cellule en texte ; None pour une cellule vide."""
    if valeur is None:
        return None
    if isinstance(valeur, float) and valeur.is_integer():
        valeur = int(valeur)
    texte = str(valeur).strip()
    return texte or None


def _lire_colonne(classeur, onglet: str, colonne: str, chemin: Path) -> list[str]:
    """Valeurs non vides d'une colonne, trouvée par son en-tête en première ligne."""
    ws = classeur[onglet]
    # Ne pas se fier aux dimensions déclarées dans le fichier : un numéro oublié
    # ferait proposer un numéro déjà attribué.
    ws.reset_dimensions()
    entetes = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
    noms = [_texte(e) for e in entetes]
    if colonne not in noms:
        raise ValueError(
            f"MASTER {chemin.name} : colonne « {colonne} » introuvable dans l'onglet « {onglet} »."
        )
    indice = noms.index(colonne) + 1
    valeurs = []
    for (valeur,) in ws.iter_rows(min_row=2, min_col=indice, max_col=indice, values_only=True):
        texte = _texte(valeur)
        if texte is not None:
            valeurs.append(texte)
    return valeurs


def lire_numeros(chemin_master) -> list[str]:
    """Numéros de facture déjà présents dans le MASTER, sans doublon et triés.

    Lit la colonne `N°_Facture` de l'onglet `Sessions`, et la colonne
    `N° facture` de l'onglet `Factures` s'il existe. Les colonnes sont trouvées
    par leur en-tête ; les cellules vides sont ignorées. Lecture seule : le
    MASTER n'est jamais modifié.

    Raises:
        FileNotFoundError: MASTER introuvable.
        ValueError: onglet `Sessions` absent, ou colonne attendue introuvable.
    """
    chemin = Path(chemin_master)
    if not chemin.is_file():
        raise FileNotFoundError(f"MASTER introuvable : {chemin}")
    classeur = load_workbook(chemin, read_only=True, data_only=True)
    try:
        onglet, colonne = _ONGLET_SESSIONS
        if onglet not in classeur.sheetnames:
            raise ValueError(f"MASTER {chemin.name} : onglet « {onglet} » absent.")
        numeros = set(_lire_colonne(classeur, onglet, colonne, chemin))
        onglet, colonne = _ONGLET_FACTURES
        if onglet in classeur.sheetnames:
            numeros.update(_lire_colonne(classeur, onglet, colonne, chemin))
    finally:
        classeur.close()
    return sorted(numeros)
