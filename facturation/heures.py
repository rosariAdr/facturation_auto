"""heures.py — Heures facturables et taux horaire d'une séance (F-13, ADR009, ADR016).

Deux calculs, à partir d'une séance de l'agenda (`agenda.Seance`) et des lignes du MASTER
(dictionnaires {en-tête : valeur}, forme de `master.py`) :

- `heures_facturables(seance, correspondance)` applique la règle d'heures de la ligne de
  `Correspondances` retenue pour la séance, selon la colonne `Mode heures` :
  `agenda` (défaut, cellule vide comprise) → la durée de la séance ; `coefficient` → durée
  × `Valeur` ; `forfait` → `Valeur`. Le résultat est arrondi à 2 décimales, sans autre
  arrondi (ni à l'heure, ni au quart d'heure) ;
- `taux_horaire(cle_tarif, tarifs)` lit, dans l'onglet `Tarifs`, le taux HT/h de la ligne
  désignée par la clé `ÉCOLE|Module|Classe` (colonne `Clé tarif` de `Correspondances`) ;
  `ligne_tarif(cle_tarif, tarifs)` rend la ligne elle-même, `Réf. contrat` comprise (P-0b).

Montants (P-0b, contrats des vagues 3 à 5, § 1.2) : `montant(h, t)` donne la valeur exacte
`Decimal(repr(h)) * Decimal(repr(t))` ; un total est la somme exacte, arrondie une seule fois au
centime par `arrondi_centime` (la moitié vers le haut).

Le `Taux HT/h` du MASTER réel est souvent du texte (« 60 €/h », § 15) : un taux se lit en
nombre, ou en texte de la forme « 60 », « 60 € », « 60 €/h », « 60,5 €/h » ou « 60.50 »
(espaces, insécables compris, tolérés, y compris autour de « € » et de « / »). La colonne
`École` de `Tarifs` porte parfois le libellé de l'école : avec les lignes de l'onglet `Écoles`
(`ecoles=`), la clé trouve la ligne par le code ou par le libellé.

Aucune valeur n'est devinée : un mode inconnu, une `Valeur` manquante, une durée d'agenda
nulle ou négative, une colonne de `Tarifs` absente, une clé tarif introuvable ou un taux
vide ou illisible lèvent une `ValueError` qui nomme la cause et dit quoi corriger dans le
MASTER ou l'agenda (« ne jamais inventer un taux ni un volume horaire », AGENTS.md §7).

💡 [proposé — à valider] Une durée d'agenda nulle (événement sans fin ni durée, que F-10 lit
comme durant zéro) ou négative est refusée en `agenda` et en `coefficient` ; le `forfait`, qui
ne lit pas la durée, n'est pas concerné.

Ce module ne fait que calculer : il n'écrit rien, ni dans le MASTER ni dans les lignes
qu'il reçoit.
"""

from __future__ import annotations

import math
import re
import unicodedata
from collections.abc import Iterable, Mapping, Sequence
from datetime import date, time, timedelta
from decimal import ROUND_HALF_UP, Decimal
from typing import Any

from facturation.agenda import Seance
from facturation.sessions import code_ecole

# Modes de la colonne `Mode heures` (MASTER_SCHEMA, onglet `Correspondances`).
MODE_AGENDA = "agenda"
MODE_COEFFICIENT = "coefficient"
MODE_FORFAIT = "forfait"
MODES = (MODE_AGENDA, MODE_COEFFICIENT, MODE_FORFAIT)

COLONNE_MODE = "Mode heures"
COLONNE_VALEUR = "Valeur"

# Ce que la `Valeur` doit contenir dans chaque mode, et un exemple écrit comme dans Excel en français.
_VALEUR_ATTENDUE = {
    MODE_COEFFICIENT: ("le coefficient", "0,875 pour une journée de 8 h facturée 7 h"),
    MODE_FORFAIT: ("le nombre d'heures du forfait", "7,5 pour 7 h 30"),
}

# Colonnes de l'onglet `Tarifs` (MASTER_SCHEMA) : les trois segments de la clé, puis le taux.
COLONNES_CLE_TARIF = ("École", "Module / Matière", "Classe / Programme")
COLONNE_TAUX = "Taux HT/h"
COLONNE_CONTRAT = "Réf. contrat"
ONGLET_TARIFS = "Tarifs"
ONGLET_ECOLES = "Écoles"
SEPARATEUR_CLE = "|"

# Arrondi des heures facturables et des montants : 2 décimales, la moitié vers le haut.
_CENTIEME = Decimal("0.01")

# Taux saisi en texte (§ 15) : « 60 », « 60 € », « 60 €/h », « 60,5 €/h », « 60.50 », espaces
# (insécables compris) tolérés autour du nombre, de « € » et de « / ».
_TAUX_TEXTE_RE = re.compile(r"(\d+(?:[.,]\d+)?)(?:\s*€(?:\s*/\s*[hH])?)?")
_FORMES_TAUX = "60, 60 €, 60 €/h, 60,5 €/h ou 60.50"


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


def _affiche(nombre: int | float) -> str:
    """Nombre écrit à la française, sans rien perdre : 45, 52,5, 45,0000001."""
    if isinstance(nombre, int):
        return str(nombre)
    texte = repr(float(nombre))
    if texte.endswith(".0"):
        texte = texte[:-2]
    return texte.replace(".", ",")


def _cellule_lisible(valeur: Any) -> str:
    """Contenu d'une cellule tel qu'Adrian le voit dans Excel : vide, VRAI, 45, « texte »."""
    if _est_vide(valeur):
        return "vide"
    if isinstance(valeur, bool):
        return "VRAI" if valeur else "FAUX"
    if isinstance(valeur, (int, float)):
        return _affiche(valeur)
    return f"« {str(valeur).strip()} »"


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


def _duree(seance: Seance, mode: str) -> Decimal:
    """Durée de la séance, exigée par les modes `agenda` et `coefficient` : strictement positive.

    Une durée nulle (événement sans fin ni durée) ou négative (fin avant le début) n'est
    jamais facturée en silence : ce serait inventer un volume horaire.
    """
    duree = _nombre(seance.duree_h)
    if duree is None or duree <= 0:
        cause = "illisible" if duree is None else "nulle ou négative"
        raise ValueError(
            f"{_seance_lisible(seance)} : mode « {mode} », durée {cause} dans l'agenda "
            f"({_cellule_lisible(seance.duree_h)} h, de {seance.debut:%H:%M} à "
            f"{seance.fin:%H:%M}). Corrigez l'heure de fin de l'événement dans l'agenda ou, si la "
            f"séance se facture au forfait, passez sa ligne de Correspondances en « {MODE_FORFAIT} » "
            "avec le nombre d'heures du contrat : aucun volume horaire n'est inventé."
        )
    return _decimal(duree)


def _pas_un_nombre(brut: Any) -> str:
    """Pourquoi une cellule non vide n'est pas lue comme un nombre, dit comme Excel l'a saisie."""
    if isinstance(brut, bool):
        return f"valeur logique {_cellule_lisible(brut)} au lieu d'un nombre"
    if isinstance(brut, (int, float)):
        return f"valeur « {brut} » qui n'est pas un nombre fini"
    if isinstance(brut, date):  # datetime compris
        return (
            f"valeur lue comme une date ({brut:%d/%m/%Y}), Excel ayant converti la saisie "
            "(une fraction comme 7/8 devient le 7 août)"
        )
    if isinstance(brut, time):
        return (
            f"valeur lue comme une heure ({brut:%H:%M}), Excel ayant converti la saisie "
            "(7:30 devient une heure de la journée, pas 7,5 h)"
        )
    if isinstance(brut, timedelta):
        return f"valeur lue comme une durée ({brut}), Excel ayant converti la saisie"
    if isinstance(brut, str):
        return f"valeur « {brut.strip()} » saisie comme du texte, pas comme un nombre"
    return f"valeur « {brut} » qui n'est pas un nombre"


def _valeur(seance: Seance, correspondance: Mapping[str, Any], mode: str) -> float:
    """`Valeur` exigée par le mode `coefficient` ou `forfait` : un nombre positif ou nul."""
    quoi, exemple = _VALEUR_ATTENDUE[mode]
    attendu = f"{quoi} (ex. {exemple})"
    ou = f"colonne « {COLONNE_VALEUR} » de Correspondances"
    brut = correspondance.get(COLONNE_VALEUR)
    if _est_vide(brut):
        raise ValueError(
            f"{_seance_lisible(seance)} : mode « {mode} » sans valeur ({ou} vide). "
            f"Renseignez {attendu}."
        )
    nombre = _nombre(brut)
    if nombre is None:
        raise ValueError(
            f"{_seance_lisible(seance)} : mode « {mode} », {ou} : {_pas_un_nombre(brut)}. "
            f"Saisissez {attendu} comme un nombre décimal, dans une cellule au format Nombre "
            "ou Standard."
        )
    if nombre < 0:
        raise ValueError(
            f"{_seance_lisible(seance)} : mode « {mode} », valeur négative ({_affiche(brut)}) dans "
            f"la {ou}. Renseignez {attendu}, positif ou nul."
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
        ValueError: mode inconnu ; durée de la séance nulle, négative ou illisible en
            `agenda` ou `coefficient` (💡 [proposé — à valider]) ; `Valeur` vide, non
            numérique ou négative en `coefficient` ou `forfait`. Le message nomme la
            séance (intitulé et date) et le mode.
    """
    if correspondance is None:
        raise TypeError(
            f"{_seance_lisible(seance)} : aucune ligne de Correspondances. Classez la séance "
            "avant de calculer ses heures : un créneau à classer n'est jamais facturé par défaut."
        )
    mode = _mode(seance, correspondance)
    if mode == MODE_FORFAIT:
        return _arrondi(_decimal(_valeur(seance, correspondance, mode)))
    duree = _duree(seance, mode)
    if mode == MODE_AGENDA:
        return _arrondi(duree)
    return _arrondi(duree * _decimal(_valeur(seance, correspondance, mode)))


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


def _en_tete(nom: Any) -> Any:
    """En-tête normalisé comme le fait `master.py` : NFC, sans espaces autour."""
    return unicodedata.normalize("NFC", nom).strip() if isinstance(nom, str) else nom


def _par_en_tete(ligne: Mapping[str, Any]) -> dict[Any, Any]:
    """La ligne, ses en-têtes normalisés (un `École` collé en forme décomposée est reconnu)."""
    normalisee: dict[Any, Any] = {}
    for nom, valeur in ligne.items():
        normalisee.setdefault(_en_tete(nom), valeur)
    return normalisee


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


def _verifier_colonnes(cle: str, lignes: list[dict[Any, Any]]) -> None:
    """Les colonnes de `Tarifs` que lit la recherche sont présentes, sous leur nom exact.

    Sans ce contrôle, une colonne mal nommée (`Classe / Groupe`, `Ecole`) se lirait comme
    une colonne vide : « clé introuvable » ou « taux vide », alors que la ligne et le taux
    existent bel et bien.

    Raises:
        ValueError: une colonne attendue n'apparaît dans aucune ligne.
    """
    lues: list[Any] = []
    for ligne in lignes:
        for nom in ligne:
            if nom not in lues:
                lues.append(nom)
    manquantes = [c for c in (*COLONNES_CLE_TARIF, COLONNE_TAUX) if c not in lues]
    if not manquantes:
        return
    pluriel = "s" if len(manquantes) > 1 else ""
    designation = "ces colonnes" if pluriel else "cette colonne"
    en_tetes = ", ".join(f"« {nom} »" for nom in lues if isinstance(nom, str) and nom)
    raise ValueError(
        f"Onglet {ONGLET_TARIFS} : colonne{pluriel} "
        f"{', '.join(f'« {c} »' for c in manquantes)} absente{pluriel} "
        f"(en-têtes lus : {en_tetes or 'aucun'}). Nommez les en-têtes de la ligne 1 comme dans "
        f"references/MASTER_SCHEMA.md : sans {designation}, le taux de la clé « {cle} » ne peut "
        "pas être lu, et aucun taux n'est deviné."
    )


def _taux_lu(brut: Any) -> float | None:
    """Taux d'une cellule : un nombre fini, ou un texte d'une forme reconnue (« 60 €/h ») ; sinon None."""
    nombre = _nombre(brut)
    if nombre is not None:
        return nombre
    if isinstance(brut, str):
        trouve = _TAUX_TEXTE_RE.fullmatch(brut.strip())
        if trouve:
            return float(trouve.group(1).replace(",", "."))
    return None


def _taux_valide(brut: Any) -> float | None:
    """Taux d'une cellule s'il est utilisable (nombre fini strictement positif), sinon None."""
    taux = _taux_lu(brut)
    return taux if taux is not None and taux > 0 else None


def _taux_de_la_ligne(cle: str, ligne: Mapping[str, Any]) -> float:
    """Taux HT/h d'une ligne de `Tarifs` : un nombre strictement positif, jamais deviné."""
    brut = ligne.get(COLONNE_TAUX)
    ou = f"onglet {ONGLET_TARIFS}, colonne « {COLONNE_TAUX} »"
    if _est_vide(brut):
        raise ValueError(
            f"Clé tarif « {cle} » : taux vide ({ou}). Renseignez le taux du contrat dans le "
            "MASTER : aucun taux n'est deviné."
        )
    taux = _taux_lu(brut)
    if taux is None:
        raise ValueError(
            f"Clé tarif « {cle} » : taux {_cellule_lisible(brut)} qui n'est pas un nombre ({ou}). "
            f"Saisissez le taux du contrat comme un nombre, ou sous l'une des formes {_FORMES_TAUX} : "
            "aucun taux n'est deviné."
        )
    if taux <= 0:
        raise ValueError(
            f"Clé tarif « {cle} » : taux nul ou négatif ({_cellule_lisible(brut)}) ({ou}). Renseignez le "
            "taux du contrat, strictement positif : aucun taux n'est deviné."
        )
    return taux


def _ligne_correspond(ligne: Mapping[str, Any], cherche: tuple[str, str, str],
                      ecoles: Sequence[Mapping[str, Any]] | None) -> bool:
    """La ligne de `Tarifs` porte-t-elle l'école, le module et la classe de la clé ?

    Sans `ecoles`, la colonne `École` se compare au segment de la clé. Avec les lignes de
    l'onglet `Écoles`, elle peut porter le code ou le libellé de l'école (§ 15).

    Raises:
        ValueError: avec `ecoles`, ligne au module et à la classe de la clé dont l'école est un
            libellé inconnu de `Écoles` : le taux n'est ni deviné ni déclaré introuvable.
    """
    if tuple(_comparable(_texte(ligne.get(c))) for c in COLONNES_CLE_TARIF[1:]) != cherche[1:]:
        return False
    brut = ligne.get(COLONNES_CLE_TARIF[0])
    if ecoles is None:
        return _comparable(_texte(brut)) == cherche[0]
    try:
        code = code_ecole(brut, ecoles)
    except ValueError as exc:
        raise ValueError(
            f"Onglet {ONGLET_TARIFS}, ligne « {_texte(ligne.get(COLONNES_CLE_TARIF[1]))} » "
            f"({_texte(ligne.get(COLONNES_CLE_TARIF[2]))}), colonne « {COLONNES_CLE_TARIF[0]} » : {exc}"
        ) from exc
    return code is not None and _comparable(code) == cherche[0]


def ligne_tarif(cle_tarif: str, tarifs: Iterable[Mapping[str, Any]], *,
                ecoles: Sequence[Mapping[str, Any]] | None = None) -> dict[str, Any]:
    """Ligne de `Tarifs` désignée par la clé `ÉCOLE|Module|Classe`, `Réf. contrat` comprise.

    Chaque segment de la clé est comparé à la colonne correspondante (`École`,
    `Module / Matière`, `Classe / Programme`) sans tenir compte de la casse (accents
    compris) ni des espaces autour du segment ou de la cellule. Aucun repli : ni
    préfixe (`M` ne désigne pas `M1`), ni ligne voisine (autre classe, autre école).
    Les lignes incomplètes et les colonnes en plus sont ignorées ; les en-têtes sont
    lus comme `master.py` les lit (NFC, sans espaces autour), sous leur nom exact.
    Avec `ecoles` (lignes de l'onglet `Écoles`), la colonne `École` peut porter le code
    ou le libellé de l'école.

    Le taux de la ligne est contrôlé (un nombre strictement positif, ou un texte d'une forme
    reconnue). Plusieurs lignes pour la même clé ne sont acceptées que si elles portent
    toutes le même taux, valide, et la même `Réf. contrat` ; sinon l'erreur dit combien de
    lignes portent la clé et cite leurs taux ou leurs références, plutôt que de choisir.

    Args:
        cle_tarif: valeur de la colonne `Clé tarif` de `Correspondances`.
        tarifs: lignes de l'onglet `Tarifs`, dictionnaires {en-tête : valeur}.
        ecoles: lignes de l'onglet `Écoles`, pour lire une école désignée par son libellé.

    Returns:
        Une copie de la ligne (en-têtes normalisés) ; la première, s'il y en a plusieurs.

    Raises:
        ValueError: clé absente ou mal formée ; colonne attendue absente de `Tarifs` ;
            clé introuvable dans `Tarifs` ; taux vide, illisible, nul ou négatif ;
            même clé sur plusieurs lignes aux taux différents ou invalides, ou aux
            références de contrat différentes ; libellé d'école inconnu (avec `ecoles`).
            Le message nomme la clé (ou la colonne absente).
    """
    segments = _segments_cle(cle_tarif)
    cle = SEPARATEUR_CLE.join(segments)
    cherche = tuple(_comparable(s) for s in segments)

    lignes = [_par_en_tete(ligne) for ligne in tarifs]
    if lignes:
        _verifier_colonnes(cle, lignes)
    retenues = [ligne for ligne in lignes if _ligne_correspond(ligne, cherche, ecoles)]
    if not retenues:
        cause = f"l'onglet {ONGLET_TARIFS} est vide" if not lignes else (
            f"aucune ligne de l'onglet {ONGLET_TARIFS} ne porte cette école, ce module et cette classe"
        )
        raise ValueError(
            f"Clé tarif « {cle} » introuvable : {cause} (colonnes "
            f"{', '.join(f'« {c} »' for c in COLONNES_CLE_TARIF)}). Ajoutez la ligne et le taux du "
            "contrat dans le MASTER, ou corrigez la clé dans Correspondances : aucun taux n'est deviné."
        )
    if len(retenues) == 1:
        _taux_de_la_ligne(cle, retenues[0])
        return dict(retenues[0])

    bruts = [ligne.get(COLONNE_TAUX) for ligne in retenues]
    taux = [_taux_valide(brut) for brut in bruts]
    if None not in taux and len(set(taux)) == 1:
        references = [ligne.get(COLONNE_CONTRAT) for ligne in retenues]
        if len({_comparable(_texte(r)) for r in references}) > 1:
            raise ValueError(
                f"Clé tarif « {cle} » : {len(retenues)} lignes de l'onglet {ONGLET_TARIFS} portent cette "
                f"école, ce module et cette classe, au même taux mais avec des « {COLONNE_CONTRAT} » "
                f"différentes (dans l'ordre des lignes : {' ; '.join(_cellule_lisible(r) for r in references)}). "
                "La facture ne peut pas citer son contrat : gardez une seule ligne par école, module et "
                "classe, avec la référence du contrat signé."
            )
        return dict(retenues[0])
    if None in taux:
        nature = "dont au moins une au taux vide, non numérique, nul ou négatif"
    else:
        nature = "avec des taux différents"
    raise ValueError(
        f"Clé tarif « {cle} » : {len(retenues)} lignes de l'onglet {ONGLET_TARIFS} portent cette "
        f"école, ce module et cette classe, {nature} (taux lus, dans l'ordre des lignes : "
        f"{' ; '.join(_cellule_lisible(brut) for brut in bruts)}). Gardez une seule ligne par "
        "école, module et classe, avec le taux du contrat : aucun taux n'est choisi à votre place."
    )


def taux_horaire(cle_tarif: str, tarifs: Iterable[Mapping[str, Any]], *,
                 ecoles: Sequence[Mapping[str, Any]] | None = None) -> float:
    """Taux HT/h de la ligne de `Tarifs` désignée par la clé `ÉCOLE|Module|Classe`.

    Même recherche et mêmes refus que `ligne_tarif`, sur laquelle elle s'appuie. Le taux se lit
    en nombre, ou en texte de la forme « 60 », « 60 € », « 60 €/h », « 60,5 €/h » ou « 60.50 ».

    Returns:
        Le taux, en float, tel qu'il est saisi (aucun arrondi).

    Raises:
        ValueError: comme `ligne_tarif`.
    """
    ligne = ligne_tarif(cle_tarif, tarifs, ecoles=ecoles)
    return _taux_lu(ligne.get(COLONNE_TAUX))  # type: ignore[return-value]  # contrôlé par ligne_tarif


# ──────────────────────────────────────────────────────────────────────
# Montants
# ──────────────────────────────────────────────────────────────────────


def _nombre_de_calcul(valeur: Any, nom: str) -> Decimal:
    if isinstance(valeur, bool) or not isinstance(valeur, (int, float)):
        raise TypeError(f"{nom} attendu en nombre, reçu {valeur!r}.")
    if not math.isfinite(valeur):
        raise ValueError(f"{nom} non fini : {valeur!r}.")
    return Decimal(repr(valeur))


def montant(heures: float, taux: float) -> Decimal:
    """Montant exact d'une ligne : `Decimal(repr(heures)) * Decimal(repr(taux))`, sans arrondi.

    Raises:
        TypeError: heures ou taux qui ne sont pas des nombres (un booléen compris).
        ValueError: heures ou taux non finis.
    """
    return _nombre_de_calcul(heures, "Heures") * _nombre_de_calcul(taux, "Taux")


def arrondi_centime(valeur: Decimal) -> Decimal:
    """Arrondi au centime, la moitié vers le haut (59,325 → 59,33), deux décimales écrites."""
    return Decimal(valeur).quantize(_CENTIEME, rounding=ROUND_HALF_UP)
