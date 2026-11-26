"""classement.py — Classement des séances de l'agenda (F-12, ADR016, ADR028).

`classer(seances, correspondances)` range chaque séance soit parmi les séances classées
(école, module, promotion et clé tarif lus sur une ligne de l'onglet `Correspondances` du
MASTER), soit dans la file à classer. Une séance à classer n'est jamais facturée par défaut
(ADR028).

Ordre de recherche (ADR016). Le premier niveau qui trouve des candidats tranche :
1. le code d'événement `#ÉCOLE-PROMO-MODULE` écrit dans la description (le premier code
   trouvé compte), cherché parmi les lignes de type `code`. Un code absent de
   `Correspondances` met la séance à classer, raison « code inconnu », sans repli sur
   l'intitulé ;
2. sinon l'intitulé exact, parmi les lignes de type `exact` ;
3. sinon le plus long préfixe de l'intitulé, parmi les lignes de type `préfixe`. Un préfixe
   ne vaut que s'il est suivi, dans l'intitulé, de la fin du texte ou d'un caractère qui n'est
   ni une lettre ni un chiffre : « Stat » ne reconnaît pas « Statut de la réunion », ni
   « Statistiques M1 » l'intitulé « Statistiques M12 » (💡 [proposé — à valider]) ;
4. sinon la file à classer, raison « non reconnue ».
Deux lignes qui conviennent au même niveau (au plus long préfixe, pour le niveau 3) mettent
la séance à classer, raison « ambiguë », sans repli sur le niveau suivant. Deux telles lignes
portent forcément la même clé ; 💡 [proposé — à valider, trou du contrat] si l'une d'elles a
été gardée plus récemment que toutes les autres (colonne « Gardée le », une ligne sans date
passant pour la plus ancienne), c'est elle qui classe la séance : la dernière décision
d'Adrian sur une clé l'emporte, et une règle gardée pour corriger une séance s'applique au
lancement suivant (ADR028). Même date, ou aucune date : « ambiguë ».

Clés, intitulés et types de clé se comparent sans tenir compte de la casse ni des espaces
en début et en fin, après normalisation Unicode NFC (comme les en-têtes dans `master.py`) :
un `É` composé ou décomposé se valent.

Un code s'écrit `#` suivi d'au moins trois segments séparés par des tirets (ÉCOLE, PROMO,
MODULE…), faits de lettres majuscules et de chiffres, le premier commençant par une lettre.
Le dièse ne colle pas à une lettre, un chiffre ou un souligné qui le précède
(`page#ABC-M1-STATS` n'en est pas un). Le code se lit en entier, jusqu'au premier caractère
qui n'est ni une lettre, ni un chiffre, ni un souligné, ni un tiret ; un tiret final est de la
ponctuation. `#visio`, `#abc-m1-stats`, `#VISIO`, `#ABC-M1` ou `#2026-11-03` ne sont pas des
codes : la séance est rapprochée par son intitulé.
Un code n'est jamais tronqué ni ignoré (💡 [proposé — à valider]) : un jeton qui s'ouvre sur
un code école en majuscules et compte au moins trois segments, mais qui contient autre chose
que des majuscules et des chiffres (`#ABC-M1-STATSbis`, `#ABC-M1-STATS_2`, `#DEF-M2-Eco`) est
un code mal formé. La séance est mise à classer, raison « code inconnu », sans repli sur un
code écrit plus loin ni sur l'intitulé, jusqu'à ce que le code soit corrigé dans l'agenda.

Une ligne de `Correspondances` est refusée (`ValueError`) si elle n'a pas de clé, pas de type
de clé connu, pas d'école ou pas de clé tarif, si sa clé de type `code` n'a pas la forme d'un
code, ou si sa date « Gardée le » n'est pas une date.

`ligne_correspondance(...)` construit la ligne à ajouter à `Correspondances` quand Adrian
garde une règle (ADR028). Ce module ne fait que calculer : il n'écrit rien dans le MASTER.
"""

from __future__ import annotations

import math
import re
import unicodedata
from collections.abc import Iterable, Mapping
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any

from facturation.agenda import Seance
from facturation.master import ONGLETS_PREVUS

# Colonnes de l'onglet Correspondances (MASTER_SCHEMA), dans l'ordre du schéma.
COLONNES_CORRESPONDANCES: tuple[str, ...] = ONGLETS_PREVUS["Correspondances"]
COL_CLE = "Clé"
COL_TYPE = "Type de clé"
COL_ECOLE = "École"
COL_MODULE = "Module"
COL_PROMOTION = "Promotion"
COL_CLE_TARIF = "Clé tarif"
COL_MODE = "Mode heures"
COL_VALEUR = "Valeur"
COL_GARDEE_LE = "Gardée le"
COL_NOTES = "Notes"

# Niveaux de reconnaissance (valeurs de `SeanceClassee.par` et de la colonne « Type de clé »).
PAR_CODE = "code"
PAR_EXACT = "exact"
PAR_PREFIXE = "préfixe"

# Raisons de la mise en file à classer (`SeanceAClasser.raison`).
NON_RECONNUE = "non reconnue"
AMBIGUE = "ambiguë"
CODE_INCONNU = "code inconnu"

# Modes d'heures (colonne « Mode heures », calcul en F-13).
MODE_AGENDA = "agenda"
MODE_COEFFICIENT = "coefficient"
MODE_FORFAIT = "forfait"
MODES_HEURES = (MODE_AGENDA, MODE_COEFFICIENT, MODE_FORFAIT)

# Type de clé saisi (forme de comparaison) → forme du contrat. « prefixe » sans accent est
# accepté : la faute de frappe ne laisse aucun doute sur l'intention.
_TYPES_DE_CLE = {
    "code": PAR_CODE,
    "exact": PAR_EXACT,
    "préfixe": PAR_PREFIXE,
    "prefixe": PAR_PREFIXE,
}
_TYPES_ATTENDUS = "code, exact ou préfixe"
_MODES_ATTENDUS = "agenda, coefficient ou forfait"
_FORME_DU_CODE = "#ÉCOLE-PROMO-MODULE"

# Dièse qui ne colle pas à une lettre, un chiffre ou un souligné, suivi de tout ce qui peut
# appartenir à un code : le jeton entier est ensuite lu par `_lire_code`, pour que
# `#ABC-M1-STATSbis` soit un code mal formé, et non un code tronqué ou ignoré.
_DIESE_RE = re.compile(r"(?<!\w)#([\w-]+)")
_SEGMENTS_MIN = 3  # ÉCOLE-PROMO-MODULE


@dataclass(frozen=True)
class SeanceClassee:
    """Une séance reconnue par une ligne de `Correspondances`.

    `ecole`, `module`, `promotion` et `cle_tarif` sont les valeurs de la ligne retenue, sans
    espaces autour (cellule vide : chaîne vide ; école et clé tarif ne sont jamais vides).
    `correspondance` est cette ligne telle qu'elle a été lue (mode d'heures et valeur compris,
    pour F-13). `par` dit le niveau qui l'a reconnue : « code », « exact » ou « préfixe ».
    """

    seance: Seance
    ecole: str
    module: str
    promotion: str
    cle_tarif: str
    correspondance: dict[str, Any] = field(hash=False)
    par: str


@dataclass(frozen=True)
class SeanceAClasser:
    """Une séance de la file à classer, jamais facturée par défaut (ADR028).

    `raison` : « non reconnue », « ambiguë » ou « code inconnu » (code absent de
    `Correspondances`, ou code mal formé).
    """

    seance: Seance
    raison: str


@dataclass(frozen=True)
class _Regle:
    """Une ligne de `Correspondances` contrôlée : type du contrat, clé sous forme comparable et
    date à laquelle la règle a été gardée (None si la cellule est vide)."""

    ligne: Mapping[str, Any]
    type_cle: str
    cle: str
    gardee_le: date | None


@dataclass(frozen=True)
class _Code:
    """Un code d'événement lu dans une description : tel qu'écrit, dièse compris, sans tiret
    final ; `valide` est faux pour un code mal formé (minuscules, souligné, segment vide)."""

    texte: str
    valide: bool


# ──────────────────────────────────────────────────────────────────────
# Normalisation
# ──────────────────────────────────────────────────────────────────────


def _texte(valeur: Any) -> str:
    """Contenu d'une cellule en texte, sans espaces autour ; '' pour une cellule vide."""
    if valeur is None:
        return ""
    if isinstance(valeur, float) and valeur.is_integer():
        valeur = int(valeur)
    return str(valeur).strip()


def _comparable(valeur: Any) -> str:
    """Forme de comparaison : texte NFC, sans espaces autour, sans casse."""
    return unicodedata.normalize("NFC", _texte(valeur)).casefold()


def _type_de_cle(valeur: Any) -> str | None:
    """Type de clé sous la forme du contrat ; None s'il est vide ou inconnu."""
    return _TYPES_DE_CLE.get(_comparable(valeur))


def _segment_de_code(segment: str) -> bool:
    """Vrai si `segment` est non vide et fait de lettres majuscules et de chiffres."""
    return bool(segment) and all(c.isupper() or c.isdecimal() for c in segment)


def _est_code(corps: str) -> bool:
    """Vrai si `corps` (ce qui suit le dièse) a exactement la forme ÉCOLE-PROMO-MODULE."""
    segments = corps.split("-")
    return (
        len(segments) >= _SEGMENTS_MIN
        and all(_segment_de_code(s) for s in segments)
        and segments[0][0].isalpha()
    )


def _lire_code(jeton: str) -> _Code | None:
    """Le code porté par un jeton (ce qui suit le dièse) ; None s'il n'en a pas l'allure.

    Un jeton a l'allure d'un code s'il compte au moins trois segments et s'ouvre sur un code
    école en majuscules (lettre, puis lettres majuscules et chiffres). Il est alors lu en
    entier : valide s'il a exactement la forme ÉCOLE-PROMO-MODULE, mal formé sinon.
    """
    corps = jeton.rstrip("-")  # tiret final : ponctuation, pas un segment
    segments = corps.split("-")
    ecole = segments[0]
    if len(segments) < _SEGMENTS_MIN or not (_segment_de_code(ecole) and ecole[0].isalpha()):
        return None
    return _Code(texte=f"#{corps}", valide=_est_code(corps))


def _premier_code(description: str | None) -> _Code | None:
    """Premier code d'événement de la description, valide ou mal formé ; None s'il n'y en a
    pas."""
    texte = unicodedata.normalize("NFC", description or "")
    for trouve in _DIESE_RE.finditer(texte):
        code = _lire_code(trouve.group(1))
        if code is not None:
            return code
    return None


def _a_forme_de_code(cle: str) -> bool:
    """Vrai si la clé d'une ligne `code` a la forme #ÉCOLE-PROMO-MODULE, casse ignorée (une
    clé saisie en minuscules reste admise)."""
    cle = unicodedata.normalize("NFC", cle).upper()
    return cle.startswith("#") and _est_code(cle[1:])


def _commence_par(intitule: str, prefixe: str) -> bool:
    """Vrai si l'intitulé commence par le préfixe, suivi de la fin du texte ou d'un caractère
    qui n'est ni une lettre ni un chiffre (sans objet si le préfixe finit lui-même par un tel
    caractère). Les deux textes sont sous forme comparable, le préfixe non vide."""
    if not intitule.startswith(prefixe):
        return False
    suite = intitule[len(prefixe):]
    return not suite or not suite[0].isalnum() or not prefixe[-1].isalnum()


# ──────────────────────────────────────────────────────────────────────
# Classement
# ──────────────────────────────────────────────────────────────────────


def _date_gardee(valeur: Any, cle: str) -> date | None:
    """Date de la colonne « Gardée le » ; None pour une cellule vide.

    Raises:
        ValueError: valeur qui n'est pas une date (texte, nombre…) ; le message cite la clé.
    """
    if _est_absente(valeur):
        return None
    if isinstance(valeur, datetime):  # openpyxl lit une cellule de date en datetime
        return valeur.date()
    if isinstance(valeur, date):
        return valeur
    raise ValueError(
        f"Correspondances : « {COL_GARDEE_LE} » vaut {valeur!r} pour la clé « {cle} », ce qui "
        "n'est pas une date. Saisissez une date (JJ/MM/AAAA) ou videz la cellule."
    )


def _regle(ligne: Mapping[str, Any]) -> _Regle:
    """Contrôle une ligne de `Correspondances` et la rend sous forme comparable.

    Raises:
        ValueError: clé vide (une clé vide de type `préfixe` reconnaîtrait toutes les
            séances ; le message cite ce qui permet de retrouver la ligne), type de clé vide
            ou inconnu, école ou clé tarif vide (ne jamais inventer un taux), clé de type
            `code` sans la forme #ÉCOLE-PROMO-MODULE, « Gardée le » qui n'est pas une date.
            Sauf pour une clé vide, le message cite la clé de la ligne.
    """
    cle = _texte(ligne.get(COL_CLE))
    type_saisi = _texte(ligne.get(COL_TYPE))
    ecole = _texte(ligne.get(COL_ECOLE))
    cle_tarif = _texte(ligne.get(COL_CLE_TARIF))
    if not cle:
        reperes = ", ".join(
            f"{nom} « {valeur} »"
            for nom, valeur in (
                ("type de clé", type_saisi),
                ("école", ecole),
                ("module", _texte(ligne.get(COL_MODULE))),
                ("promotion", _texte(ligne.get(COL_PROMOTION))),
                ("clé tarif", cle_tarif),
                ("notes", _texte(ligne.get(COL_NOTES))),
            )
            if valeur
        )
        raise ValueError(
            f"Correspondances : ligne sans clé{f' ({reperes})' if reperes else ''}. Chaque "
            "ligne doit porter sa clé, code d'événement ou intitulé : une clé vide de type "
            "préfixe reconnaîtrait toutes les séances."
        )
    type_cle = _type_de_cle(type_saisi)
    if type_cle is None:
        cause = f"type de clé inconnu « {type_saisi} »" if type_saisi else "type de clé manquant"
        raise ValueError(
            f"Correspondances : {cause} pour la clé « {cle} » (attendu : {_TYPES_ATTENDUS})."
        )
    if not ecole:
        raise ValueError(
            f"Correspondances : école manquante pour la clé « {cle} ». Chaque ligne nomme "
            "l'école à facturer."
        )
    if not cle_tarif:
        raise ValueError(
            f"Correspondances : clé tarif manquante pour la clé « {cle} ». Chaque ligne désigne "
            "la ligne de Tarifs à appliquer (ne jamais inventer un taux)."
        )
    if type_cle == PAR_CODE and not _a_forme_de_code(cle):
        raise ValueError(
            f"Correspondances : la clé « {cle} », de type « {PAR_CODE} », n'a pas la forme d'un "
            f"code d'événement {_FORME_DU_CODE} (un dièse, puis au moins trois segments de "
            "lettres et de chiffres séparés par des tirets) : elle ne reconnaîtrait aucune "
            "séance."
        )
    return _Regle(
        ligne=ligne,
        type_cle=type_cle,
        cle=_comparable(cle),
        gardee_le=_date_gardee(ligne.get(COL_GARDEE_LE), cle),
    )


def _departager(candidates: list[_Regle]) -> _Regle | None:
    """L'unique candidate, ou celle gardée seule le plus récemment ; None si l'ambiguïté
    demeure (même date, ou aucune date).

    💡 [proposé — à valider, trou du contrat] : une ligne sans date « Gardée le » passe pour
    plus ancienne que toute ligne datée."""
    if len(candidates) == 1:
        return candidates[0]
    datees = [r for r in candidates if r.gardee_le is not None]
    if not datees:
        return None
    plus_recente = max(r.gardee_le for r in datees)
    retenues = [r for r in datees if r.gardee_le == plus_recente]
    return retenues[0] if len(retenues) == 1 else None


def _trancher(seance: Seance, candidates: list[_Regle], par: str) -> SeanceClassee | SeanceAClasser:
    """Classe la séance par la ligne candidate retenue ; sinon : « ambiguë »."""
    retenue = _departager(candidates)
    if retenue is None:
        return SeanceAClasser(seance=seance, raison=AMBIGUE)
    ligne = retenue.ligne
    return SeanceClassee(
        seance=seance,
        ecole=_texte(ligne.get(COL_ECOLE)),
        module=_texte(ligne.get(COL_MODULE)),
        promotion=_texte(ligne.get(COL_PROMOTION)),
        cle_tarif=_texte(ligne.get(COL_CLE_TARIF)),
        correspondance=dict(ligne),
        par=par,
    )


def classer(
    seances: Iterable[Seance],
    correspondances: Iterable[Mapping[str, Any]],
) -> tuple[list[SeanceClassee], list[SeanceAClasser]]:
    """Classe les séances d'après les lignes de `Correspondances` (ADR016, ADR028).

    Toutes les lignes sont contrôlées avant le classement, même si aucune séance ne les
    utilise. Les entrées ne sont pas modifiées.

    Args:
        seances: séances de la période (`facturation.agenda.Seance`).
        correspondances: lignes de l'onglet `Correspondances`, sous forme de dictionnaires
            {en-tête : valeur} (forme rendue par `facturation.master.lire_master`).

    Returns:
        (séances classées, séances à classer) : chaque séance figure une seule fois, dans
        l'une des deux listes, et chaque liste garde l'ordre des séances reçues.

    Raises:
        ValueError: ligne de `Correspondances` invalide : sans clé, sans type de clé ou d'un
            type inconnu, sans école ou sans clé tarif, clé de type `code` sans la forme
            #ÉCOLE-PROMO-MODULE, « Gardée le » qui n'est pas une date. Le message cite la clé
            de la ligne, ou pour une ligne sans clé ce qui permet de la retrouver.
    """
    codes: dict[str, list[_Regle]] = {}
    exacts: dict[str, list[_Regle]] = {}
    prefixes: list[_Regle] = []
    for ligne in correspondances:
        regle = _regle(ligne)
        if regle.type_cle == PAR_CODE:
            codes.setdefault(regle.cle, []).append(regle)
        elif regle.type_cle == PAR_EXACT:
            exacts.setdefault(regle.cle, []).append(regle)
        else:
            prefixes.append(regle)

    classees: list[SeanceClassee] = []
    a_classer: list[SeanceAClasser] = []
    for seance in seances:
        resultat = _classer_seance(seance, codes, exacts, prefixes)
        if isinstance(resultat, SeanceClassee):
            classees.append(resultat)
        else:
            a_classer.append(resultat)
    return classees, a_classer


def _classer_seance(
    seance: Seance,
    codes: dict[str, list[_Regle]],
    exacts: dict[str, list[_Regle]],
    prefixes: list[_Regle],
) -> SeanceClassee | SeanceAClasser:
    """Une séance, niveau par niveau : le premier niveau qui a des candidates tranche."""
    code = _premier_code(seance.description)
    if code is not None:
        # Un code mal formé n'est dans aucune ligne (`_regle` refuse une clé sans forme de code).
        candidates = codes.get(_comparable(code.texte), []) if code.valide else []
        if not candidates:
            return SeanceAClasser(seance=seance, raison=CODE_INCONNU)
        return _trancher(seance, candidates, PAR_CODE)

    intitule = _comparable(seance.intitule)
    candidates = exacts.get(intitule, [])
    if candidates:
        return _trancher(seance, candidates, PAR_EXACT)

    convenables = [r for r in prefixes if _commence_par(intitule, r.cle)]
    if convenables:
        plus_long = max(len(r.cle) for r in convenables)
        return _trancher(seance, [r for r in convenables if len(r.cle) == plus_long], PAR_PREFIXE)

    return SeanceAClasser(seance=seance, raison=NON_RECONNUE)


# ──────────────────────────────────────────────────────────────────────
# Règle gardée (ADR028)
# ──────────────────────────────────────────────────────────────────────


def _est_absente(valeur: Any) -> bool:
    return valeur is None or (isinstance(valeur, str) and not valeur.strip())


def _controler_valeur(valeur: Any, mode: str) -> None:
    """Valeur d'un mode d'heures : obligatoire, nombre positif ou nul, en coefficient et en
    forfait ; absente en mode agenda, qui l'ignorerait (F-13)."""
    if mode == MODE_AGENDA:
        if not _est_absente(valeur):
            raise ValueError(
                f"Valeur {valeur!r} donnée en mode d'heures « {MODE_AGENDA} », qui l'ignore : "
                f"choisissez le mode « {MODE_COEFFICIENT} » ou « {MODE_FORFAIT} », ou retirez "
                "la valeur."
            )
        return
    quoi = "le coefficient" if mode == MODE_COEFFICIENT else "le nombre d'heures du forfait"
    if _est_absente(valeur):
        raise ValueError(
            f"Mode d'heures « {mode} » sans valeur : indiquez {quoi} (ne jamais inventer un "
            "volume horaire)."
        )
    if (
        isinstance(valeur, bool)
        or not isinstance(valeur, (int, float))
        or not math.isfinite(valeur)
        or valeur < 0
    ):
        raise ValueError(
            f"Mode d'heures « {mode} » : valeur {valeur!r} invalide, {quoi} doit être un nombre "
            "positif ou nul."
        )


def ligne_correspondance(
    seance: Seance,
    *,
    ecole: str,
    module: str,
    promotion: str,
    cle_tarif: str,
    mode_heures: str = MODE_AGENDA,
    valeur: float | None = None,
    type_cle: str = PAR_EXACT,
    garde_le: date,
) -> dict[str, Any]:
    """La ligne de `Correspondances` à ajouter quand Adrian garde une règle (ADR028).

    Rien n'est écrit ici : la ligne rendue est à montrer, puis à écrire après validation.
    Elle porte toutes les colonnes de `Correspondances` (MASTER_SCHEMA). La clé vient de la
    séance : son intitulé sans espaces autour pour les types `exact` et `préfixe` (en
    `préfixe`, l'intitulé entier sert de préfixe), le premier code de sa description pour le
    type `code`. Ainsi la règle reconnaît la séance d'où elle vient.

    Une séance dont la description porte un code est classée par ce code, jamais par son
    intitulé : une règle `exact` ou `préfixe` ne s'y appliquerait pas, elle est donc refusée.
    Une séance au code mal formé reste à classer tant que le code n'est pas corrigé dans
    l'agenda : aucune règle n'est acceptée (💡 [proposé — à valider]).

    La ligne s'ajoute aux lignes existantes, sans les voir. Si une ligne de même type et de
    même clé existe déjà, la ligne gardée le plus récemment l'emporte au classement (voir
    `classer`).

    Args:
        seance: la séance classée par Adrian.
        ecole, module, promotion, cle_tarif: classement choisi ; école et clé tarif
            obligatoires.
        mode_heures: `agenda` (défaut, vide compris), `coefficient` ou `forfait`.
        valeur: coefficient ou nombre d'heures du forfait, obligatoire dans ces deux modes et
            seulement dans eux.
        type_cle: `exact` (défaut), `code` ou `préfixe`, sans tenir compte de la casse.
        garde_le: date à laquelle Adrian garde la règle (colonne « Gardée le »).

    Raises:
        ValueError: type de clé ou mode d'heures inconnu, valeur absente, négative ou en trop,
            école ou clé tarif vide, séance sans code (type `code`) ou sans intitulé (types
            `exact` et `préfixe`), séance qui porte un code (types `exact` et `préfixe`, le
            message cite le code), code mal formé. Le message nomme la cause.
        TypeError: `garde_le` n'est pas une date.
    """
    type_normalise = _type_de_cle(type_cle)
    if type_normalise is None:
        raise ValueError(f"Type de clé inconnu « {_texte(type_cle)} » (attendu : {_TYPES_ATTENDUS}).")

    mode = _comparable(mode_heures) or MODE_AGENDA
    if mode not in MODES_HEURES:
        raise ValueError(f"Mode d'heures inconnu « {_texte(mode_heures)} » (attendu : {_MODES_ATTENDUS}).")
    _controler_valeur(valeur, mode)

    ecole_texte = _texte(ecole)
    if not ecole_texte:
        raise ValueError("École manquante : une règle de Correspondances nomme l'école à facturer.")
    cle_tarif_texte = _texte(cle_tarif)
    if not cle_tarif_texte:
        raise ValueError(
            "Clé tarif manquante : une règle de Correspondances désigne la ligne de Tarifs à "
            "appliquer (ne jamais inventer un taux)."
        )
    if not isinstance(garde_le, date):
        raise TypeError(
            f"« {COL_GARDEE_LE} » doit être une date, et non {garde_le!r} : la colonne du MASTER "
            "ne reçoit ni texte ni cellule vide."
        )

    intitule = _texte(seance.intitule)
    seance_lisible = f"« {intitule} » du {seance.date:%d/%m/%Y}"
    code = _premier_code(seance.description)
    if type_normalise == PAR_CODE:
        if code is None:
            raise ValueError(
                f"Aucun code d'événement ({_FORME_DU_CODE}) dans la description de la séance "
                f"{seance_lisible} : impossible de garder une règle de type « {PAR_CODE} »."
            )
        cle = code.texte
    else:
        cle = intitule
        if not cle:
            raise ValueError(
                f"Intitulé vide (séance du {seance.date:%d/%m/%Y}) : impossible de garder une "
                f"règle de type « {type_normalise} »."
            )
    if code is not None and not code.valide:
        raise ValueError(
            f"Code d'événement mal formé, {code.texte}, dans la description de la séance "
            f"{seance_lisible} (attendu : {_FORME_DU_CODE}, en lettres majuscules, chiffres et "
            "tirets) : la séance reste à classer tant que le code n'est pas corrigé dans "
            "l'agenda, aucune règle ne s'y appliquerait."
        )
    if code is not None and type_normalise != PAR_CODE:
        raise ValueError(
            f"La description de la séance {seance_lisible} porte le code d'événement "
            f"{code.texte} : la séance est classée par ce code, jamais par son intitulé, et une "
            f"règle de type « {type_normalise} » ne s'y appliquerait pas. Gardez une règle de "
            f"type « {PAR_CODE} », ou corrigez la ligne de ce code dans Correspondances."
        )

    valeurs = {
        COL_CLE: cle,
        COL_TYPE: type_normalise,
        COL_ECOLE: ecole_texte,
        COL_MODULE: _texte(module),
        COL_PROMOTION: _texte(promotion),
        COL_CLE_TARIF: cle_tarif_texte,
        COL_MODE: mode,
        COL_VALEUR: None if mode == MODE_AGENDA else valeur,
        COL_GARDEE_LE: garde_le,
        COL_NOTES: None,
    }
    return {colonne: valeurs[colonne] for colonne in COLONNES_CORRESPONDANCES}
