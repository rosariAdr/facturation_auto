"""sessions.py — Lecture de l'onglet `Sessions` du MASTER (P-0b, socle du lot 1 ; lecture seule).

`Sessions` est le journal d'audit des séances facturées (ADR003). Ce module le lit sous une
forme typée (`LigneSessions`) et y cherche les lignes d'une séance de l'agenda, pour que la
préparation (F-16a), l'audit (F-19) et les signalements (F-23) partagent la même lecture.

Formes lues (constat du MASTER réel, contrats des vagues 3 à 5, § 15) :
- `École` : le code, ou le libellé de l'école (colonne `École` de l'onglet `Écoles`) ;
- `Date_séance` : une date, ou un texte JJ/MM/AAAA ou AAAA-MM-JJ ;
- `Début` et `Fin` : une heure, ou un texte aux formes historiques « 09:00 », « 9h », « 9h30 »,
  « 09h30 » ; vides, ils valent None ;
- `Durée_h`, `Taux`, `Total HT` : un nombre, ou un texte « 3,5 » ;
- `Période_facture` : un texte libre, lu tel quel (intervalles, « Mois-Mois AAAA »,
  parenthèses) ; on ne s'y fie pas pour repérer une période facturée.

Une cellule illisible place la ligne dans la seconde liste de `lire_sessions` (rang Excel et
message qui nomme la colonne et la valeur), sans plantage.

Lectures retenues (`💡 [proposé — à valider]`) : un libellé d'école inconnu de `Écoles` est
refusé (§ 15), sauf s'il a la forme d'un code école (majuscules et chiffres, ADR029), rendu
alors nettoyé, en majuscules (§ 3.3) ; un libellé qui désigne deux codes différents est refusé ;
`Total HT` devient le `Decimal` du nombre tel qu'il s'écrit ; le statut « abandonnée » de
`Factures` se reconnaît sans tenir compte de la casse, des espaces ni de l'accent
(« abandonnee ») ; le n° abandonné se compare à celui de `Sessions` sans tenir compte de la
casse ni des espaces ; une ligne sans école a le code "" ; dans `chercher`, une ligne du même
jour à laquelle il manque une heure (début ou fin) va dans la seconde liste, jamais dans les
exactes (revue de P-0b).
"""

from __future__ import annotations

import math
import re
import unicodedata
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal, InvalidOperation
from typing import TYPE_CHECKING, Any

from facturation.numerotation import valider_code_ecole

if TYPE_CHECKING:
    from facturation.master import Master

ONGLET = "Sessions"
ONGLET_ECOLES = "Écoles"
ONGLET_FACTURES = "Factures"
STATUT_ABANDONNEE = "abandonnée"

_HEURE_H_RE = re.compile(r"(\d{1,2})\s*[hH]\s*(\d{2})?")
_HEURE_DEUX_POINTS_RE = re.compile(r"(\d{1,2}):(\d{2})(?::(\d{2}))?")
_NOMBRE_RE = re.compile(r"[+-]?\d+(?:[.,]\d+)?")


@dataclass(frozen=True)
class LigneSessions:
    """Une ligne lue de l'onglet `Sessions`."""

    rang: int                 # ligne Excel
    ecole: str                # code (code_ecole)
    n_facture: str
    periode: str
    date_seance: date | None
    debut: time | None
    fin: time | None
    duree_h: float | None
    promotion: str
    module: str
    taux: float | None
    total_ht: Decimal | None
    summary: str
    abandonnee: bool          # N° au statut « abandonnée » dans Factures


class _Illisible(ValueError):
    """Cellule illisible : le message nomme la colonne et la valeur."""


# ──────────────────────────────────────────────────────────────────────
# Écoles
# ──────────────────────────────────────────────────────────────────────


def _forme(valeur: Any) -> str:
    """Forme de comparaison : NFC, sans casse, espaces réduits et retirés autour."""
    return " ".join(unicodedata.normalize("NFC", str(valeur)).casefold().split())


def _est_vide(valeur: Any) -> bool:
    return valeur is None or (isinstance(valeur, str) and not valeur.strip())


def _texte(valeur: Any) -> str:
    """Contenu d'une cellule en texte NFC, sans espaces autour ; "" si vide ; 3.0 s'écrit 3."""
    if valeur is None:
        return ""
    if isinstance(valeur, float) and valeur.is_integer():
        valeur = int(valeur)
    return unicodedata.normalize("NFC", str(valeur)).strip()


def code_ecole(valeur: Any, ecoles: Sequence[Mapping[str, Any]]) -> str | None:
    """Code facturation d'une école désignée par son code ou par son libellé.

    - cellule vide → None ;
    - un `Code facturation` de `Écoles` (sans tenir compte de la casse, des espaces ni de la
      forme Unicode) → ce code ;
    - un libellé de la colonne `École` de `Écoles` → son code ;
    - sinon, une valeur qui a la forme d'un code école (majuscules et chiffres) → la valeur
      nettoyée, en majuscules.

    Raises:
        ValueError: libellé inconnu de l'onglet `Écoles`, libellé qui désigne plusieurs codes,
            ou libellé dont la ligne n'a pas de code ; le message cite la valeur et l'onglet.
    """
    if _est_vide(valeur):
        return None
    texte = _texte(valeur)
    cherche = _forme(texte)
    for ecole in ecoles:
        code = _texte(ecole.get("Code facturation"))
        if code and _forme(code) == cherche:
            return code.upper()
    codes = []
    for ecole in ecoles:
        if not _est_vide(ecole.get("École")) and _forme(ecole.get("École")) == cherche:
            code = _texte(ecole.get("Code facturation")).upper()
            if not code:
                raise ValueError(
                    f"École « {texte} » : sa ligne de l'onglet « {ONGLET_ECOLES} » n'a pas de « Code "
                    "facturation ». Renseignez-le : l'école n'est jamais devinée."
                )
            if code not in codes:
                codes.append(code)
    if len(codes) == 1:
        return codes[0]
    if len(codes) > 1:
        raise ValueError(
            f"École « {texte} » : plusieurs lignes de l'onglet « {ONGLET_ECOLES} » portent ce libellé, "
            f"avec des codes différents ({', '.join(codes)}). Gardez un libellé par école."
        )
    candidat = " ".join(texte.split()).upper()
    try:
        return valider_code_ecole(candidat)
    except ValueError:
        pass
    raise ValueError(
        f"École « {texte} » inconnue de l'onglet « {ONGLET_ECOLES} » (ni « Code facturation », ni libellé "
        f"de la colonne « École »). Ajoutez l'école à l'onglet {ONGLET_ECOLES}, ou écrivez son code : "
        "l'école n'est jamais devinée."
    )


def _sans_accents(texte: str) -> str:
    """Texte sans signes diacritiques (« abandonnée » → « abandonnee »)."""
    decompose = unicodedata.normalize("NFD", texte)
    return unicodedata.normalize("NFC", "".join(c for c in decompose if not unicodedata.combining(c)))


def _cle_numero(valeur: Any) -> str:
    """Forme de comparaison d'un n° de facture : sans espaces, en majuscules (« r_abc001 » → « R_ABC001 »)."""
    return "".join(_texte(valeur).split()).upper()


def numeros_abandonnes(factures: Sequence[Mapping[str, Any]]) -> set[str]:
    """N° des factures au statut « abandonnée » dans `Factures`.

    Le statut se reconnaît sans tenir compte de la casse, des espaces ni de l'accent
    (« Abandonnée », « abandonnee ») ; les n° sont rendus sans espaces et en majuscules, la forme
    sous laquelle `lire_sessions` compare le n° de chaque ligne de `Sessions`.
    """
    statut = _sans_accents(STATUT_ABANDONNEE)
    return {
        _cle_numero(facture.get("N° facture"))
        for facture in factures
        if _sans_accents(_forme(_texte(facture.get("Statut")))) == statut and _cle_numero(facture.get("N° facture"))
    }


# ──────────────────────────────────────────────────────────────────────
# Lecture des cellules
# ──────────────────────────────────────────────────────────────────────


def _lisible(valeur: Any) -> str:
    return _texte(valeur) if not isinstance(valeur, (date, time)) else str(valeur)


def _date(valeur: Any, colonne: str) -> date | None:
    if _est_vide(valeur):
        return None
    if isinstance(valeur, datetime):
        return valeur.date()
    if isinstance(valeur, date):
        return valeur
    if isinstance(valeur, str):
        texte = valeur.strip()
        for forme in ("%d/%m/%Y", "%Y-%m-%d"):
            try:
                return datetime.strptime(texte, forme).date()
            except ValueError:
                continue
    raise _Illisible(f"colonne « {colonne} » : « {_lisible(valeur)} » illisible (attendu une date JJ/MM/AAAA)")


def _heure(valeur: Any, colonne: str) -> time | None:
    if _est_vide(valeur):
        return None
    if isinstance(valeur, datetime):
        return valeur.time().replace(tzinfo=None)
    if isinstance(valeur, time):
        return valeur.replace(tzinfo=None)
    if isinstance(valeur, str):
        texte = valeur.strip()
        trouve = _HEURE_DEUX_POINTS_RE.fullmatch(texte) or _HEURE_H_RE.fullmatch(texte)
        if trouve:
            heures, minutes = int(trouve.group(1)), int(trouve.group(2) or 0)
            secondes = int(trouve.group(3) or 0) if trouve.re is _HEURE_DEUX_POINTS_RE else 0
            if heures <= 23 and minutes <= 59 and secondes <= 59:
                return time(heures, minutes, secondes)
    raise _Illisible(
        f"colonne « {colonne} » : « {_lisible(valeur)} » illisible (attendu une heure : 09:00, 9h ou 9h30)"
    )


def _nombre(valeur: Any, colonne: str) -> float | None:
    if _est_vide(valeur):
        return None
    if isinstance(valeur, (int, float)) and not isinstance(valeur, bool):
        if math.isfinite(valeur):
            return float(valeur)
    elif isinstance(valeur, str) and _NOMBRE_RE.fullmatch(valeur.strip()):
        return float(valeur.strip().replace(",", "."))
    raise _Illisible(f"colonne « {colonne} » : « {_lisible(valeur)} » illisible (attendu un nombre, ex. 3,5)")


def _montant(valeur: Any, colonne: str) -> Decimal | None:
    if _est_vide(valeur):
        return None
    try:
        if isinstance(valeur, int) and not isinstance(valeur, bool):
            return Decimal(valeur)
        if isinstance(valeur, float) and math.isfinite(valeur):
            return Decimal(repr(valeur))
        if isinstance(valeur, str) and _NOMBRE_RE.fullmatch(valeur.strip()):
            return Decimal(valeur.strip().replace(",", "."))
    except InvalidOperation:
        pass
    raise _Illisible(f"colonne « {colonne} » : « {_lisible(valeur)} » illisible (attendu un montant, ex. 157,50)")


def _ligne(rang: int, brute: Mapping[str, Any], ecoles: Sequence[Mapping[str, Any]],
           abandonnes: set[str]) -> LigneSessions:
    """Une ligne de `Sessions` ; `_Illisible` pour la première cellule qui ne se lit pas."""
    try:
        ecole = code_ecole(brute.get("École"), ecoles) or ""
    except ValueError as exc:
        raise _Illisible(f"colonne « École » : {exc}") from exc
    n_facture = _texte(brute.get("N°_Facture"))
    return LigneSessions(
        rang=rang,
        ecole=ecole,
        n_facture=n_facture,
        periode=_texte(brute.get("Période_facture")),
        date_seance=_date(brute.get("Date_séance"), "Date_séance"),
        debut=_heure(brute.get("Début"), "Début"),
        fin=_heure(brute.get("Fin"), "Fin"),
        duree_h=_nombre(brute.get("Durée_h"), "Durée_h"),
        promotion=_texte(brute.get("Promotion")),
        module=_texte(brute.get("Module / UE")),
        taux=_nombre(brute.get("Taux"), "Taux"),
        total_ht=_montant(brute.get("Total HT"), "Total HT"),
        summary=_texte(brute.get("Summary calendrier verbatim")),
        abandonnee=bool(n_facture) and _cle_numero(n_facture) in abandonnes,
    )


def lire_sessions(master: "Master") -> tuple[list[LigneSessions], list[tuple[int, str]]]:
    """Lignes de `Sessions`, et lignes illisibles.

    Returns:
        (lignes lues, [(rang Excel, message)] des lignes dont une cellule ne se lit pas). Le
        message nomme la colonne et cite la valeur ; aucune ligne illisible ne fait planter.
    """
    abandonnes = numeros_abandonnes(master.factures)
    rangs = list(getattr(master, "rangs", {}).get("sessions") or [])
    if len(rangs) != len(master.sessions):  # Master construit à la main, sans rangs
        rangs = list(range(2, len(master.sessions) + 2))
    lignes: list[LigneSessions] = []
    erreurs: list[tuple[int, str]] = []
    for rang, brute in zip(rangs, master.sessions):
        try:
            lignes.append(_ligne(rang, brute, master.ecoles, abandonnes))
        except _Illisible as exc:
            erreurs.append((rang, f"onglet « {ONGLET} », ligne {rang}, {exc}"))
    return lignes, erreurs


# ──────────────────────────────────────────────────────────────────────
# Recherche
# ──────────────────────────────────────────────────────────────────────


def _a_la_minute(heure: time) -> time:
    return heure.replace(second=0, microsecond=0, tzinfo=None)


def chercher(ecole: str, debut: datetime, fin: datetime, lignes: Sequence[LigneSessions]
             ) -> tuple[list[LigneSessions], list[LigneSessions]]:
    """Lignes de `Sessions` d'une séance, hors lignes d'un N° abandonné.

    Returns:
        (lignes exactes : même école, même date, même début et même fin à la minute ;
         lignes du même jour et de la même école auxquelles il manque une heure, début ou fin,
         à montrer comme des doublons possibles : une séance déjà facturée ne disparaît jamais
         de la recherche parce qu'une de ses heures n'a pas été saisie).
    """
    code = _texte(ecole).upper()
    jour = debut.date()
    heure_debut, heure_fin = _a_la_minute(debut.time()), _a_la_minute(fin.time())
    exactes: list[LigneSessions] = []
    sans_heure: list[LigneSessions] = []
    for ligne in lignes:
        if ligne.abandonnee or ligne.ecole.upper() != code or ligne.date_seance != jour:
            continue
        if ligne.debut is None or ligne.fin is None:
            sans_heure.append(ligne)
        elif _a_la_minute(ligne.debut) == heure_debut and _a_la_minute(ligne.fin) == heure_fin:
            exactes.append(ligne)
    return exactes, sans_heure
