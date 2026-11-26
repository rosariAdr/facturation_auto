"""heures.py — Heures facturables et taux horaire d'une séance (F-13, ADR009, ADR016).

Deux calculs, à partir d'une séance de l'agenda (`agenda.Seance`) et des lignes du MASTER
(dictionnaires {en-tête : valeur}, forme de `master.py`) :

- `heures_facturables(seance, correspondance)` applique la règle d'heures de la ligne de
  `Correspondances` retenue pour la séance, selon la colonne `Mode heures` :
  `agenda` (défaut, cellule vide comprise) → la durée de la séance ; `coefficient` → durée
  × `Valeur` ; `forfait` → `Valeur`. Le résultat est arrondi à 2 décimales, sans autre
  arrondi (ni à l'heure, ni au quart d'heure) ;
- `taux_horaire(cle_tarif, tarifs)` lit, dans l'onglet `Tarifs`, le taux HT/h de la ligne
  désignée par la clé `ÉCOLE|Module|Classe` (colonne `Clé tarif` de `Correspondances`).

Aucune valeur n'est devinée : un mode inconnu, une `Valeur` manquante, une clé tarif
introuvable ou un taux vide lèvent une `ValueError` qui nomme la cause et dit quoi
corriger dans le MASTER (« ne jamais inventer un taux », AGENTS.md §7).

Ce module ne fait que calculer : il n'écrit rien, ni dans le MASTER ni dans les lignes
qu'il reçoit.
"""

from __future__ import annotations

import math
import unicodedata
from collections.abc import Iterable, Mapping
from decimal import ROUND_HALF_UP, Decimal
from typing import Any

from facturation.agenda import Seance

# Modes de la colonne `Mode heures` (MASTER_SCHEMA, onglet `Correspondances`).
MODE_AGENDA = "agenda"
MODE_COEFFICIENT = "coefficient"
MODE_FORFAIT = "forfait"
MODES = (MODE_AGENDA, MODE_COEFFICIENT, MODE_FORFAIT)

COLONNE_MODE = "Mode heures"
COLONNE_VALEUR = "Valeur"

# Colonnes de l'onglet `Tarifs` (MASTER_SCHEMA) : les trois segments de la clé, puis le taux.
COLONNES_CLE_TARIF = ("École", "Module / Matière", "Classe / Programme")
COLONNE_TAUX = "Taux HT/h"
ONGLET_TARIFS = "Tarifs"
SEPARATEUR_CLE = "|"

# Arrondi des heures facturables : 2 décimales, au plus proche, la moitié vers le haut.
_CENTIEME = Decimal("0.01")


# ──────────────────────────────────────────────────────────────────────
# Outils communs
# ──────────────────────────────────────────────────────────────────────


def _est_vide(valeur: Any) -> bool:
    """Cellule vide : None, ou texte fait seulement d'espaces."""
    return valeur is None or (isinstance(valeur, str) and not valeur.strip())


def _nombre(valeur: Any) -> float | None:
    """Valeur numérique d'une cellule ; None si elle n'est pas un nombre fini.

    Seuls les nombres saisis comme nombres comptent : un texte, même `"0.875"`, est
    refusé plutôt qu'interprété (`"1,250"` peut valoir 1,25 ou 1 250). Un booléen
    (VRAI/FAUX dans Excel) n'est pas un nombre, ni NaN ou l'infini.
    """
    if isinstance(valeur, bool) or not isinstance(valeur, (int, float)):
        return None
    nombre = float(valeur)
    return nombre if math.isfinite(nombre) else None


def _decimal(nombre: float) -> Decimal:
    """Décimal écrit comme le nombre s'affiche (`0.875`, pas sa valeur binaire exacte)."""
    return Decimal(repr(float(nombre)))


def _arrondi(valeur: Decimal) -> float:
    """Arrondi à 2 décimales, la moitié vers le haut (1,125 → 1,13), rendu en float."""
    return float(valeur.quantize(_CENTIEME, rounding=ROUND_HALF_UP))


# ──────────────────────────────────────────────────────────────────────
# Heures facturables
# ──────────────────────────────────────────────────────────────────────


def _seance_lisible(seance: Seance) -> str:
    """La séance telle qu'Adrian la retrouve dans son agenda : intitulé et date."""
    return f"Séance « {seance.intitule} » du {seance.date:%d/%m/%Y}"


def _mode(seance: Seance, correspondance: Mapping[str, Any]) -> str:
    """Mode d'heures de la ligne, en minuscules ; `agenda` si la cellule est vide ou absente."""
    brut = correspondance.get(COLONNE_MODE)
    if _est_vide(brut):
        return MODE_AGENDA
    saisi = str(brut).strip()
    mode = saisi.casefold()
    if mode not in MODES:
        raise ValueError(
            f"{_seance_lisible(seance)} : mode d'heures « {saisi} » inconnu (colonne "
            f"« {COLONNE_MODE} » de Correspondances). Modes reconnus : "
            f"{', '.join(MODES)} (cellule vide = {MODE_AGENDA})."
        )
    return mode


def _valeur(seance: Seance, correspondance: Mapping[str, Any], mode: str) -> float:
    """`Valeur` exigée par le mode `coefficient` ou `forfait` : un nombre positif ou nul."""
    if mode == MODE_COEFFICIENT:
        quoi = "le coefficient (ex. 0.875 pour 7/8)"
    else:
        quoi = "le nombre d'heures du forfait"
    brut = correspondance.get(COLONNE_VALEUR)
    if _est_vide(brut):
        raise ValueError(
            f"{_seance_lisible(seance)} : mode « {mode} » sans valeur (colonne "
            f"« {COLONNE_VALEUR} » vide dans Correspondances). Renseignez {quoi}."
        )
    nombre = _nombre(brut)
    if nombre is None:
        raise ValueError(
            f"{_seance_lisible(seance)} : mode « {mode} », valeur « {brut} » qui n'est pas un "
            f"nombre (colonne « {COLONNE_VALEUR} » de Correspondances). Saisissez {quoi} comme un "
            "nombre, pas comme du texte (dans Excel en français : 0,875)."
        )
    if nombre < 0:
        raise ValueError(
            f"{_seance_lisible(seance)} : mode « {mode} », valeur négative ({brut}) dans la "
            f"colonne « {COLONNE_VALEUR} » de Correspondances. Renseignez {quoi}, positif ou nul."
        )
    return nombre


def heures_facturables(seance: Seance, correspondance: Mapping[str, Any]) -> float:
    """Heures à facturer pour une séance, selon la règle de sa ligne de `Correspondances`.

    - `agenda` (défaut : cellule `Mode heures` vide ou colonne absente) : la durée de la
      séance, `seance.duree_h` (durée réelle, changement d'heure compris) ; une `Valeur`
      restée dans la ligne est ignorée ;
    - `coefficient` : durée × `Valeur` (0.875 pour une journée de 8 h facturée 7 h) ;
    - `forfait` : `Valeur`, quelle que soit la durée (un jury de 7 h facturé 7,5 h).

    Le mode est reconnu sans tenir compte de la casse ni des espaces autour. Le résultat
    est un float arrondi à 2 décimales, la moitié vers le haut, sans autre arrondi.
    Une `Valeur` nulle est acceptée (elle donne 0 h). La ligne n'est pas modifiée.

    Args:
        seance: séance de l'agenda (`agenda.lire_ics`).
        correspondance: ligne de `Correspondances` retenue pour la séance (classement F-12),
            dictionnaire {en-tête : valeur}.

    Raises:
        TypeError: aucune correspondance (`None`) : la séance n'est pas classée.
        ValueError: mode inconnu ; `Valeur` vide, non numérique ou négative en
            `coefficient` ou `forfait`. Le message nomme la séance (intitulé et date)
            et le mode.
    """
    if correspondance is None:
        raise TypeError(
            f"{_seance_lisible(seance)} : aucune ligne de Correspondances. Classez la séance "
            "avant de calculer ses heures : un créneau à classer n'est jamais facturé par défaut."
        )
    mode = _mode(seance, correspondance)
    duree = _decimal(seance.duree_h)
    if mode == MODE_AGENDA:
        return _arrondi(duree)
    valeur = _decimal(_valeur(seance, correspondance, mode))
    if mode == MODE_COEFFICIENT:
        return _arrondi(duree * valeur)
    return _arrondi(valeur)


# ──────────────────────────────────────────────────────────────────────
# Taux horaire
# ──────────────────────────────────────────────────────────────────────


def _texte(valeur: Any) -> str:
    """Contenu d'une cellule en texte, sans espaces autour ; '' pour une cellule vide.

    Un nombre entier lu en float (classe `3` saisie comme nombre) s'écrit sans décimale.
    """
    if valeur is None:
        return ""
    if isinstance(valeur, float) and valeur.is_integer():
        valeur = int(valeur)
    return str(valeur).strip()


def _comparable(texte: str) -> str:
    """Forme de comparaison : NFC (`É` composé ou décomposé se valent), sans casse, accents compris."""
    return unicodedata.normalize("NFC", texte).casefold()


def _segments_cle(cle_tarif: Any) -> tuple[str, str, str]:
    """Les trois segments de la clé `ÉCOLE|Module|Classe`, sans espaces autour.

    Raises:
        ValueError: clé vide, ou qui n'a pas exactement trois segments non vides.
    """
    if _est_vide(cle_tarif) or not isinstance(cle_tarif, str):
        raise ValueError(
            f"Clé tarif absente ou qui n'est pas du texte ({cle_tarif!r}) : renseignez la "
            "colonne « Clé tarif » de Correspondances sous la forme ÉCOLE|Module|Classe "
            "(ex. ABC|Statistiques|M1)."
        )
    segments = tuple(s.strip() for s in cle_tarif.split(SEPARATEUR_CLE))
    if len(segments) != len(COLONNES_CLE_TARIF) or not all(segments):
        raise ValueError(
            f"Clé tarif mal formée : « {cle_tarif.strip()} ». Attendu : trois parties non vides "
            f"séparées par « {SEPARATEUR_CLE} », ÉCOLE|Module|Classe (ex. ABC|Statistiques|M1), "
            f"comme les colonnes {', '.join(f'« {c} »' for c in COLONNES_CLE_TARIF)} de "
            f"l'onglet {ONGLET_TARIFS}."
        )
    return segments  # type: ignore[return-value]


def _taux_de_la_ligne(cle: str, ligne: Mapping[str, Any]) -> float:
    """Taux HT/h d'une ligne de `Tarifs` : un nombre strictement positif, jamais deviné."""
    brut = ligne.get(COLONNE_TAUX)
    ou = f"onglet {ONGLET_TARIFS}, colonne « {COLONNE_TAUX} »"
    if _est_vide(brut):
        raise ValueError(
            f"Clé tarif « {cle} » : taux vide ({ou}). Renseignez le taux du contrat dans le "
            "MASTER : aucun taux n'est deviné."
        )
    taux = _nombre(brut)
    if taux is None:
        raise ValueError(
            f"Clé tarif « {cle} » : taux « {brut} » qui n'est pas un nombre ({ou}). Saisissez "
            "le taux du contrat comme un nombre : aucun taux n'est deviné."
        )
    if taux <= 0:
        raise ValueError(
            f"Clé tarif « {cle} » : taux nul ou négatif ({brut}) ({ou}). Renseignez le taux du "
            "contrat, strictement positif : aucun taux n'est deviné."
        )
    return taux


def taux_horaire(cle_tarif: str, tarifs: Iterable[Mapping[str, Any]]) -> float:
    """Taux HT/h de la ligne de `Tarifs` désignée par la clé `ÉCOLE|Module|Classe`.

    Chaque segment de la clé est comparé à la colonne correspondante (`École`,
    `Module / Matière`, `Classe / Programme`) sans tenir compte de la casse (accents
    compris) ni des espaces autour du segment ou de la cellule. Aucun repli : ni
    préfixe (`M` ne désigne pas `M1`), ni ligne voisine (autre classe, autre école).
    Les lignes incomplètes et les colonnes en plus sont ignorées.

    Plusieurs lignes pour la même clé ne sont acceptées que si elles portent le même
    taux ; des taux différents lèvent une erreur plutôt que de choisir l'un d'eux.

    Args:
        cle_tarif: valeur de la colonne `Clé tarif` de `Correspondances`.
        tarifs: lignes de l'onglet `Tarifs`, dictionnaires {en-tête : valeur}.

    Returns:
        Le taux, en float, tel qu'il est saisi (aucun arrondi).

    Raises:
        ValueError: clé absente ou mal formée ; clé introuvable dans `Tarifs` ; taux
            vide, non numérique, nul ou négatif ; même clé avec des taux différents.
            Le message nomme la clé.
    """
    segments = _segments_cle(cle_tarif)
    cle = SEPARATEUR_CLE.join(segments)
    cherche = tuple(_comparable(s) for s in segments)

    lignes = list(tarifs)
    retenues = [
        ligne for ligne in lignes
        if tuple(_comparable(_texte(ligne.get(c))) for c in COLONNES_CLE_TARIF) == cherche
    ]
    if not retenues:
        cause = f"l'onglet {ONGLET_TARIFS} est vide" if not lignes else (
            f"aucune ligne de l'onglet {ONGLET_TARIFS} ne porte cette école, ce module et cette classe"
        )
        raise ValueError(
            f"Clé tarif « {cle} » introuvable : {cause} (colonnes "
            f"{', '.join(f'« {c} »' for c in COLONNES_CLE_TARIF)}). Ajoutez la ligne et le taux du "
            "contrat dans le MASTER, ou corrigez la clé dans Correspondances : aucun taux n'est deviné."
        )

    taux = {_taux_de_la_ligne(cle, ligne) for ligne in retenues}
    if len(taux) > 1:
        valeurs = ", ".join(f"{t:g}" for t in sorted(taux))
        raise ValueError(
            f"Clé tarif « {cle} » : {len(retenues)} lignes de l'onglet {ONGLET_TARIFS} avec des taux "
            f"différents ({valeurs}). Gardez une seule ligne par école, module et classe : aucun "
            "taux n'est choisi à votre place."
        )
    return taux.pop()
