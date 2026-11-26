"""regles.py — Règles de facturation par école (F-14, ADR015, ADR026, ADR030, ADR031).

Lit les lignes de l'onglet `Règles_facturation` du MASTER, sous la forme que rend
`master.lire_master` (une liste de dictionnaires {en-tête : valeur}), et en tire :

- `regles_par_ecole(lignes)` : une `RegleEcole` par école, indexée par son code ;
- `periode_a_facturer(regle, annee, mois)` : les dates couvertes par la facture du mois,
  selon le rythme de l'école (`mensuel`, `cycle` ; `fin de module` n'est pas encore pris
  en charge) ;
- `echeance_selon_regle(date_emission, delai_paiement)` : la date d'échéance (ADR030) ;
- `ecoles_du_lot(regles, lot)` : les écoles d'un lot ;
- `blocages(regle, ref_bon_commande=...)` : ce qui empêche d'émettre la facture.

Lecture d'une ligne :
- chaque colonne de MASTER_SCHEMA doit figurer dans la ligne, écrite comme dans le schéma ;
  une colonne en plus est permise, comme dans `lire_master` ;
- un texte est lu sans les espaces de début et de fin ; une cellule faite d'espaces est vide ;
- le code école est la clé, sans espaces et en majuscules (`abc` → `ABC`, comme pour la
  numérotation) ; une école en double, même à la casse près, est refusée, comme une ligne
  remplie sans code école ;
- `Statut lot` vaut `Lot 1`, `Plus tard` ou `Hors facturation` (la liste de MASTER_SCHEMA),
  sans tenir compte de la casse ni des espaces ; vide, l'école n'est dans aucun lot ;
- `Rythme`, `Colonne Jour` et `Regroupement` sont lus sans tenir compte de la casse et rangés
  en minuscules ; vides, ils valent `mensuel`, `affichée` et `séance` ;
- `Une facture par contrat` et `Bon de commande requis` : `oui`/`non` ou un booléen Excel ;
  vide = non ;
- `Début de cycle` est un jour du mois entier, de 2 à 28, exigé quand le rythme est `cycle` ;
  pour un autre rythme, la colonne ne sert pas mais doit rester un jour du mois (1 à 31) ;
- `Délai de paiement` est contrôlé dès la lecture : vide, « N jours », « N jours fin de mois »
  ou « fin de mois », N au plus 365 ;
- `Date limite d'envoi` est un texte libre (« 15 à 12 h ») : une date ou une heure qu'Excel a
  convertie est refusée.
Une valeur qui ne se lit pas lève `ValueError`, qui nomme l'école, la colonne et la valeur :
une règle n'est jamais devinée. Toutes les erreurs de l'onglet sont rapportées dans une seule
`ValueError`, pour tout corriger en une fois. Seule exception, dans l'esprit d'ADR031 (une
facture bloquée ne retient pas les autres) : la ligne illisible d'une école `Plus tard` ou
`Hors facturation` est écartée du résultat, avec un avertissement `RegleEcartee`.

`💡 [proposé — à valider]` : liste fermée des statuts (un `Lot 2` est refusé tant que
MASTER_SCHEMA ne le prévoit pas), plafond de 365 jours du délai, contrôle du début de cycle
hors du rythme `cycle`, école hors lot écartée avec un avertissement, date limite convertie
par Excel refusée, une facture par contrat sans n° de contrat bloquée (`blocages`).

Ce module calcule : il n'écrit rien dans le MASTER.
"""

from __future__ import annotations

import calendar
import re
import unicodedata
import warnings
from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from typing import Any, Callable

from facturation.master import ONGLETS_PREVUS
from facturation.numerotation import valider_code_ecole

ONGLET = "Règles_facturation"

# Colonnes de l'onglet, dans l'ordre de MASTER_SCHEMA (une seule liste, celle de `master`).
COLONNES = ONGLETS_PREVUS[ONGLET]

STATUTS = ("Lot 1", "Plus tard", "Hors facturation")
STATUTS_HORS_LOT = ("Plus tard", "Hors facturation")
RYTHMES = ("mensuel", "cycle", "fin de module")
COLONNE_JOUR = ("affichée", "masquée")
REGROUPEMENTS = ("séance", "jour")
RYTHME_DEFAUT = "mensuel"
COLONNE_JOUR_DEFAUT = "affichée"
REGROUPEMENT_DEFAUT = "séance"
LOT_DEFAUT = "Lot 1"

# Délai de paiement quand la règle n'en donne pas : date d'émission + 30 jours (ADR030).
DELAI_DEFAUT_JOURS = 30
# `💡 [proposé — à valider]` Délai le plus long accepté : au-delà d'un an, c'est sans doute une
# faute de saisie, et un délai démesuré mènerait hors du calendrier.
DELAI_MAX_JOURS = 365

# Plage du début de cycle : un jour présent dans tous les mois. Un cycle qui commence le 1er
# est le rythme mensuel ; au-delà du 28, le jour manque certains mois.
DEBUT_CYCLE_MIN = 2
DEBUT_CYCLE_MAX = 28
# Un jour du mois, quel que soit le rythme.
JOUR_MIN = 1
JOUR_MAX = 31

# Colonnes de `Règles_facturation` (MASTER_SCHEMA).
_ECOLE = "École"
_STATUT = "Statut lot"
_RYTHME = "Rythme"
_DEBUT_CYCLE = "Début de cycle"
_DATE_LIMITE = "Date limite d'envoi"
_DELAI = "Délai de paiement"
_CONTRAT = "N° de contrat"
_PAR_CONTRAT = "Une facture par contrat"
_BC = "Bon de commande requis"
_JOUR = "Colonne Jour"
_REGROUPEMENT = "Regroupement"
_CANAL = "Canal d'envoi"
_PJ = "Pièces jointes"
_NOTES = "Notes"

# Formes reconnues du délai de paiement, comparées après `_forme` (minuscules, espaces réduits).
_DELAI_JOURS_RE = re.compile(r"([0-9]+) jours?")
_DELAI_JOURS_FIN_DE_MOIS_RE = re.compile(r"([0-9]+) jours? fin de mois")
_FIN_DE_MOIS = "fin de mois"
_FORMES_DELAI = (
    "vide (date d'émission + 30 jours, ADR030), « N jours », « N jours fin de mois » "
    f"ou « fin de mois », N au plus {DELAI_MAX_JOURS}"
)


class RegleEcartee(UserWarning):
    """Ligne illisible d'une école `Plus tard` ou `Hors facturation`, écartée du résultat de
    `regles_par_ecole` sans retenir les autres écoles. À corriger avant de facturer l'école."""


@dataclass(frozen=True)
class RegleEcole:
    """Règle de facturation d'une école : une ligne de l'onglet `Règles_facturation`.

    Les textes sans valeur valent "" ; `debut_cycle` vaut None quand la cellule est vide.
    Les valeurs par défaut sont celles d'une ligne vide (rythme `mensuel`, colonne Jour
    `affichée`, regroupement `séance`, oui/non à faux).
    """

    ecole: str
    statut_lot: str = ""
    rythme: str = RYTHME_DEFAUT
    debut_cycle: int | None = None
    date_limite: str = ""
    delai_paiement: str = ""
    n_contrat: str = ""
    une_facture_par_contrat: bool = False
    bc_requis: bool = False
    colonne_jour: str = COLONNE_JOUR_DEFAUT
    regroupement: str = REGROUPEMENT_DEFAUT
    canal: str = ""
    pieces_jointes: str = ""
    notes: str = ""


# ──────────────────────────────────────────────────────────────────────
# Lecture des cellules
# ──────────────────────────────────────────────────────────────────────


def _texte(valeur: Any) -> str:
    """Contenu d'une cellule en texte, sans espaces autour ; "" pour une cellule vide."""
    if valeur is None:
        return ""
    if isinstance(valeur, float) and valeur.is_integer():
        valeur = int(valeur)
    return str(valeur).strip()


def _forme(texte: str) -> str:
    """Forme de comparaison : NFC, minuscules, espaces intérieurs réduits à un seul."""
    return " ".join(unicodedata.normalize("NFC", texte).lower().split())


def _est_vide(valeur: Any) -> bool:
    return valeur is None or (isinstance(valeur, str) and not valeur.strip())


def _liste(valeurs: Iterable[str]) -> str:
    return ", ".join(f"« {v} »" for v in valeurs)


def _pluriel(nombre: int) -> str:
    return "s" if nombre > 1 else ""


def _statut(ligne: Mapping[str, Any], ecole: str) -> str:
    """`Statut lot` tel qu'écrit, sans espaces autour ; "" si vide ; refusé hors de la liste."""
    texte = _texte(ligne.get(_STATUT))
    if texte and _forme(texte) not in {_forme(s) for s in STATUTS}:
        raise ValueError(
            f"{ONGLET}, école {ecole} : « {_STATUT} » vaut « {texte} », valeur inconnue. "
            f"Valeurs possibles : {_liste(STATUTS)} ; vide = dans aucun lot."
        )
    return texte


def _valeur_de_liste(ligne: Mapping[str, Any], colonne: str, ecole: str,
                     permises: tuple[str, ...], defaut: str) -> str:
    """Valeur d'une colonne à liste fermée, en minuscules ; `defaut` pour une cellule vide."""
    texte = _texte(ligne.get(colonne))
    if not texte:
        return defaut
    forme = _forme(texte)
    if forme not in permises:
        raise ValueError(
            f"{ONGLET}, école {ecole} : « {colonne} » vaut « {texte} », valeur inconnue. "
            f"Valeurs possibles : {_liste(permises)} ; vide = « {defaut} »."
        )
    return forme


def _oui_non(ligne: Mapping[str, Any], colonne: str, ecole: str) -> bool:
    """`oui`/`non` dans toute casse, ou un booléen Excel ; vide = non."""
    valeur = ligne.get(colonne)
    if isinstance(valeur, bool):
        return valeur
    texte = _texte(valeur)
    if not texte:
        return False
    forme = _forme(texte)
    if forme in ("oui", "non"):
        return forme == "oui"
    raise ValueError(
        f"{ONGLET}, école {ecole} : « {colonne} » vaut « {texte} ». "
        "Attendu : « oui » ou « non » (vide = non)."
    )


def _debut_cycle(ligne: Mapping[str, Any], ecole: str) -> int | None:
    """Jour du début de cycle en entier (15, 15.0 et "15" donnent 15) ; None si vide.

    Quel que soit le rythme, la cellule est lue de la même façon, qu'elle soit un nombre ou un
    texte : un jour du mois entier, de 1 à 31. La plage 2 à 28 du rythme `cycle` est contrôlée
    par `_verifier_cycle`.
    """
    valeur = ligne.get(_DEBUT_CYCLE)
    jour: int | None = None
    if isinstance(valeur, bool):
        jour = None
    elif isinstance(valeur, int):
        jour = valeur
    elif isinstance(valeur, float) and valeur.is_integer():
        jour = int(valeur)
    else:
        texte = _texte(valeur)
        if not texte:
            return None
        if re.fullmatch(r"[0-9]{1,9}", texte):
            jour = int(texte)
    if jour is None or not JOUR_MIN <= jour <= JOUR_MAX:
        raise ValueError(
            f"{ONGLET}, école {ecole} : « {_DEBUT_CYCLE} » vaut « {_texte(valeur)} ». Attendu : un "
            f"jour du mois entier, de {DEBUT_CYCLE_MIN} à {DEBUT_CYCLE_MAX} pour le rythme « cycle », "
            "le seul qui s'en sert ; la cellule peut rester vide pour les autres rythmes."
        )
    return jour


def _verifier_cycle(ecole: str, debut: Any) -> int:
    """Début de cycle d'une règle au rythme `cycle` : présent, entier, de 2 à 28."""
    if debut is None or debut == "":
        raise ValueError(
            f"Règle de l'école {ecole} : rythme « cycle » sans début de cycle. Indiquez dans la "
            f"colonne « {_DEBUT_CYCLE} » le jour du mois où commence le cycle, "
            f"de {DEBUT_CYCLE_MIN} à {DEBUT_CYCLE_MAX}."
        )
    if isinstance(debut, bool) or not isinstance(debut, int):
        raise ValueError(
            f"Règle de l'école {ecole} : début de cycle illisible : « {debut} ». "
            f"Attendu : un jour du mois entier, de {DEBUT_CYCLE_MIN} à {DEBUT_CYCLE_MAX}."
        )
    if not DEBUT_CYCLE_MIN <= debut <= DEBUT_CYCLE_MAX:
        conseil = " Un cycle qui commence le 1er est le rythme « mensuel »." if debut == 1 else ""
        raise ValueError(
            f"Règle de l'école {ecole} : début de cycle {debut} hors de la plage "
            f"{DEBUT_CYCLE_MIN} à {DEBUT_CYCLE_MAX} (un jour présent dans tous les mois).{conseil}"
        )
    return debut


def _lire_delai(delai: Any) -> tuple[int, bool]:
    """Délai de paiement → (nombre de jours, aller à la fin du mois).

    Raises:
        ValueError: texte qui n'est aucune des formes reconnues, ou délai de plus de
            `DELAI_MAX_JOURS` jours ; le message nomme la colonne et cite le texte.
    """
    texte = _texte(delai)
    if not texte:
        return DELAI_DEFAUT_JOURS, False
    forme = _forme(texte)
    if forme == _FIN_DE_MOIS:
        return 0, True
    for motif, fin_de_mois in ((_DELAI_JOURS_FIN_DE_MOIS_RE, True), (_DELAI_JOURS_RE, False)):
        trouve = motif.fullmatch(forme)
        if not trouve:
            continue
        chiffres = trouve.group(1).lstrip("0") or "0"
        # Longueur vérifiée avant la conversion : un nombre de milliers de chiffres ne se
        # convertit pas (limite de Python) et ne serait de toute façon pas un délai.
        if len(chiffres) > len(str(DELAI_MAX_JOURS)) or int(chiffres) > DELAI_MAX_JOURS:
            raise ValueError(
                f"« {_DELAI} » démesuré : « {texte} ». Au plus {DELAI_MAX_JOURS} jours : "
                "au-delà, c'est sans doute une faute de saisie."
            )
        return int(chiffres), fin_de_mois
    raise ValueError(f"« {_DELAI} » illisible : « {texte} ». Formes reconnues : {_FORMES_DELAI}.")


def _horodatage_lisible(valeur: date | time) -> str:
    """Date ou heure Excel, écrite comme Excel l'affiche en français."""
    if isinstance(valeur, datetime):
        jour = f"{valeur:%d/%m/%Y}"
        return jour if valeur.time() == time(0) else f"{jour} {valeur:%H:%M}"
    if isinstance(valeur, date):
        return f"{valeur:%d/%m/%Y}"
    return f"{valeur:%H:%M}"


def _date_limite(ligne: Mapping[str, Any], ecole: str) -> str:
    """`Date limite d'envoi` : un texte libre ; une date ou une heure convertie par Excel est refusée."""
    valeur = ligne.get(_DATE_LIMITE)
    if isinstance(valeur, (date, time)):  # datetime compris
        nature = "heure" if isinstance(valeur, time) else "date"
        raise ValueError(
            f"{ONGLET}, école {ecole} : « {_DATE_LIMITE} » a été convertie par Excel en {nature} "
            f"(« {_horodatage_lisible(valeur)} »). C'est un texte libre, par exemple « 15 à 12 h » : "
            "saisissez-le dans une cellule au format Texte, ou précédé d'une apostrophe."
        )
    return _texte(valeur)


def _dernier_jour(annee: int, mois: int) -> date:
    return date(annee, mois, calendar.monthrange(annee, mois)[1])


# ──────────────────────────────────────────────────────────────────────
# Règles par école
# ──────────────────────────────────────────────────────────────────────


def _lire_regle(ligne: Mapping[str, Any], ecole: str) -> tuple[RegleEcole | None, list[str]]:
    """Règle d'une ligne dont le code école est déjà lu et vérifié.

    Returns:
        (la règle, []) si toute la ligne se lit ; (None, les erreurs de la ligne) sinon :
        chaque colonne est lue, pour rapporter toutes les erreurs en une fois.
    """
    erreurs: list[str] = []

    def lire(fonction: Callable[..., Any], *args: Any) -> Any:
        try:
            return fonction(*args)
        except ValueError as exc:
            erreurs.append(str(exc))
            return None

    statut = lire(_statut, ligne, ecole)
    rythme = lire(_valeur_de_liste, ligne, _RYTHME, ecole, RYTHMES, RYTHME_DEFAUT)
    try:
        debut = _debut_cycle(ligne, ecole)
        if rythme == "cycle":
            _verifier_cycle(ecole, debut)
    except ValueError as exc:
        erreurs.append(str(exc))
        debut = None
    delai = _texte(ligne.get(_DELAI))
    try:
        _lire_delai(delai)  # contrôlé dès la lecture, pour que le message nomme l'école
    except ValueError as exc:
        erreurs.append(f"{ONGLET}, école {ecole} : {exc}")
    date_limite = lire(_date_limite, ligne, ecole)
    par_contrat = lire(_oui_non, ligne, _PAR_CONTRAT, ecole)
    bc_requis = lire(_oui_non, ligne, _BC, ecole)
    colonne_jour = lire(_valeur_de_liste, ligne, _JOUR, ecole, COLONNE_JOUR, COLONNE_JOUR_DEFAUT)
    regroupement = lire(_valeur_de_liste, ligne, _REGROUPEMENT, ecole, REGROUPEMENTS, REGROUPEMENT_DEFAUT)
    if erreurs:
        return None, erreurs
    return RegleEcole(
        ecole=ecole,
        statut_lot=statut,
        rythme=rythme,
        debut_cycle=debut,
        date_limite=date_limite,
        delai_paiement=delai,
        n_contrat=_texte(ligne.get(_CONTRAT)),
        une_facture_par_contrat=par_contrat,
        bc_requis=bc_requis,
        colonne_jour=colonne_jour,
        regroupement=regroupement,
        canal=_texte(ligne.get(_CANAL)),
        pieces_jointes=_texte(ligne.get(_PJ)),
        notes=_texte(ligne.get(_NOTES)),
    ), []


def _verifier_colonnes(ligne: Mapping[str, Any]):
    """Chaque colonne de MASTER_SCHEMA figure dans la ligne ; une colonne en plus est permise."""
    manquantes = [c for c in COLONNES if c not in ligne]
    if not manquantes:
        return
    s = _pluriel(len(manquantes))
    inconnues = [c for c in ligne if c not in COLONNES]
    indice = (
        f" Colonne{_pluriel(len(inconnues))} hors du schéma, peut-être mal orthographiée"
        f"{_pluriel(len(inconnues))} : {_liste(inconnues)}."
        if inconnues else ""
    )
    raise ValueError(
        f"{ONGLET} : colonne{s} manquante{s} : {_liste(manquantes)}. Les en-têtes s'écrivent "
        f"comme dans references/MASTER_SCHEMA.md ; une colonne absente n'est jamais lue comme "
        f"vide.{indice}"
    )


def _ligne_sans_ecole(ligne: Mapping[str, Any], precedente: str) -> str:
    """Message d'une ligne remplie sans code école, qui permet de la retrouver dans le MASTER."""
    contenu = " ; ".join(
        f"{colonne} « {_texte(valeur)} »" for colonne, valeur in ligne.items() if not _est_vide(valeur)
    )
    ou = f"après celle de l'école {precedente}" if precedente else "première règle de l'onglet"
    return (
        f"{ONGLET} : une règle est remplie mais sa colonne « {_ECOLE} » est vide ({ou} ; "
        f"contenu : {contenu}). Indiquez le code de l'école, ou videz la ligne si elle ne sert pas."
    )


def _une_erreur(erreurs: list[str]) -> str:
    """Toutes les erreurs de l'onglet dans un seul message."""
    if len(erreurs) == 1:
        return erreurs[0]
    return f"{ONGLET} : {len(erreurs)} erreurs à corriger.\n" + "\n".join(f"- {e}" for e in erreurs)


def regles_par_ecole(lignes: Iterable[Mapping[str, Any]]) -> dict[str, RegleEcole]:
    """Une règle par école, indexée par son code en majuscules.

    Args:
        lignes: lignes de l'onglet `Règles_facturation`, par exemple
            `lire_master(chemin).regles_facturation`. Une ligne entièrement vide est ignorée.
            Chaque ligne remplie porte toutes les colonnes de MASTER_SCHEMA (une colonne en
            plus est permise) ; une cellule vide vaut None ou "".

    Toutes les lignes sont lues avant de conclure : les erreurs de l'onglet sont rapportées
    ensemble. La ligne illisible d'une école `Plus tard` ou `Hors facturation` ne retient pas
    les autres : elle est absente du résultat et signalée par un avertissement `RegleEcartee`.

    Raises:
        ValueError: colonne de MASTER_SCHEMA absente (aussitôt) ; ligne remplie sans code
            école, code école mal formé, école en double (même à la casse ou aux espaces près),
            ou valeur illisible sur la ligne d'une école qui n'est pas hors lot. Un seul
            message rapporte toutes ces erreurs ; chacune nomme l'école, la colonne et la valeur.
    """
    regles: dict[str, RegleEcole] = {}
    erreurs: list[str] = []
    ecartees: list[str] = []
    contrats: dict[str, str] = {}  # école → n° de contrat de sa première ligne
    precedente = ""
    for brute in lignes:
        ligne = {unicodedata.normalize("NFC", str(cle)).strip(): valeur for cle, valeur in brute.items()}
        if all(_est_vide(v) for v in ligne.values()):
            continue
        _verifier_colonnes(ligne)
        ecole = _texte(ligne.get(_ECOLE)).upper()
        if not ecole:
            erreurs.append(_ligne_sans_ecole(ligne, precedente))
            continue
        precedente = ecole
        try:
            valider_code_ecole(ecole)
        except ValueError as exc:
            erreurs.append(f"{ONGLET}, colonne « {_ECOLE} » : {exc}")
            continue
        if ecole in contrats:
            numeros = sorted({contrats[ecole], _texte(ligne.get(_CONTRAT))} - {""})
            precision = (
                f" Plusieurs contrats pour une même école ({_liste(numeros)}) ne sont pas encore "
                "pris en charge."
                if len(numeros) > 1 else ""
            )
            erreurs.append(f"{ONGLET} : école {ecole} en double. Une seule ligne par école.{precision}")
            continue
        contrats[ecole] = _texte(ligne.get(_CONTRAT))
        regle, erreurs_ligne = _lire_regle(ligne, ecole)
        if regle is not None:
            regles[ecole] = regle
            continue
        statut = _texte(ligne.get(_STATUT))
        if _forme(statut) in {_forme(s) for s in STATUTS_HORS_LOT}:
            ecartees.append(
                f"{ONGLET}, école {ecole} (« {statut} ») : règle illisible, écartée sans retenir "
                "les autres écoles ; à corriger avant de facturer cette école.\n"
                + "\n".join(f"- {e}" for e in erreurs_ligne)
            )
        else:
            erreurs.extend(erreurs_ligne)
    for message in ecartees:
        warnings.warn(message, RegleEcartee, stacklevel=2)
    if erreurs:
        raise ValueError(_une_erreur(erreurs))
    return regles


# ──────────────────────────────────────────────────────────────────────
# Calculs
# ──────────────────────────────────────────────────────────────────────


def periode_a_facturer(regle: RegleEcole, annee: int, mois: int) -> tuple[date, date]:
    """Première et dernière dates couvertes par la facture du mois `mois` de `annee`.

    - `mensuel` (ou rythme vide) : du 1er au dernier jour du mois ;
    - `cycle` de début d : du d du mois précédent au d-1 du mois, c'est-à-dire le cycle qui
      se termine dans le mois ; les périodes de deux mois qui se suivent se touchent, sans
      trou ni chevauchement.

    Raises:
        ValueError: mois hors de 1 à 12 ; rythme inconnu ; cycle sans début de cycle valide.
        NotImplementedError: rythme `fin de module`, pas encore pris en charge.
    """
    if isinstance(mois, bool) or not isinstance(mois, int) or not 1 <= mois <= 12:
        raise ValueError(f"Mois invalide : {mois!r}. Attendu : un entier de 1 à 12.")
    if isinstance(annee, bool) or not isinstance(annee, int):
        raise ValueError(f"Année invalide : {annee!r}. Attendu : un entier, par exemple 2026.")
    rythme = _forme(_texte(regle.rythme)) or RYTHME_DEFAUT
    if rythme == "mensuel":
        return date(annee, mois, 1), _dernier_jour(annee, mois)
    if rythme == "cycle":
        debut = _verifier_cycle(regle.ecole, regle.debut_cycle)
        annee_prec, mois_prec = (annee - 1, 12) if mois == 1 else (annee, mois - 1)
        return date(annee_prec, mois_prec, debut), date(annee, mois, debut) - timedelta(days=1)
    if rythme == "fin de module":
        raise NotImplementedError(
            f"Règle de l'école {regle.ecole} : rythme « fin de module ». La facture couvre le "
            "module entier et part à sa fin : ce calcul n'est pas encore pris en charge."
        )
    raise ValueError(
        f"Règle de l'école {regle.ecole} : rythme « {regle.rythme} » inconnu. "
        f"Valeurs possibles : {_liste(RYTHMES)}."
    )


def echeance_selon_regle(date_emission: date, delai_paiement: str | None) -> date:
    """Date d'échéance d'une facture selon le délai de paiement de la règle (ADR030).

    - vide : date d'émission + 30 jours ;
    - « N jours » : date d'émission + N jours ;
    - « N jours fin de mois » : dernier jour du mois de (date d'émission + N jours) ;
    - « fin de mois » : dernier jour du mois d'émission.
    La casse et les espaces ne comptent pas ; N vaut au plus `DELAI_MAX_JOURS` (365).

    Raises:
        TypeError: `date_emission` n'est pas une date.
        ValueError: tout autre texte, délai démesuré ou échéance hors du calendrier (le message
            cite le délai) : jamais d'échéance devinée.
    """
    if isinstance(date_emission, datetime):
        date_emission = date_emission.date()
    if not isinstance(date_emission, date):
        raise TypeError(f"Date d'émission attendue (datetime.date), reçu : {date_emission!r}.")
    jours, fin_de_mois = _lire_delai(delai_paiement)
    try:
        echeance = date_emission + timedelta(days=jours)
    except OverflowError as exc:
        raise ValueError(
            f"Échéance hors du calendrier : « {_DELAI} » « {_texte(delai_paiement)} » après une "
            f"émission le {date_emission:%d/%m/%Y}."
        ) from exc
    if fin_de_mois:
        echeance = _dernier_jour(echeance.year, echeance.month)
    return echeance


def ecoles_du_lot(regles: Mapping[str, RegleEcole] | Iterable[RegleEcole],
                  lot: str = LOT_DEFAUT) -> list[str]:
    """Codes des écoles dont le `Statut lot` vaut `lot`, triés.

    Le statut est comparé sans tenir compte de la casse ni des espaces ; une école au statut
    vide n'appartient à aucun lot.

    Args:
        regles: le dictionnaire rendu par `regles_par_ecole` (ou une suite de `RegleEcole`).
        lot: `Lot 1` par défaut, `Plus tard`, `Hors facturation`…
    """
    valeurs = regles.values() if isinstance(regles, Mapping) else regles
    cible = _forme(_texte(lot))
    if not cible:  # un statut vide ne désigne aucun lot
        return []
    return sorted({regle.ecole for regle in valeurs if _forme(_texte(regle.statut_lot)) == cible})


def blocages(regle: RegleEcole, *, ref_bon_commande: str | None) -> list[str]:
    """Raisons qui empêchent d'émettre la facture de l'école ; liste vide si rien ne bloque.

    La facture reste préparée : seul son envoi attend que la raison soit levée (ADR031).
    - bon de commande requis et absent ;
    - `💡 [proposé — à valider]` une facture par contrat demandée sans n° de contrat : la facture
      ne pourrait pas citer son contrat (ADR030 §3).

    Args:
        regle: règle de l'école.
        ref_bon_commande: référence du bon de commande reçu pour cette facture ; None, "" ou
            des espaces s'il n'y en a pas.
    """
    raisons = []
    if regle.bc_requis and not _texte(ref_bon_commande):
        raisons.append(
            f"École {regle.ecole} : bon de commande requis et absent. La facture est préparée, "
            "mais n'est pas émise sans la référence du bon de commande."
        )
    if regle.une_facture_par_contrat and not _texte(regle.n_contrat):
        raisons.append(
            f"École {regle.ecole} : une facture par contrat demandée, mais aucun « {_CONTRAT} ». "
            "La facture est préparée, mais n'est pas émise sans la référence de son contrat."
        )
    return raisons
