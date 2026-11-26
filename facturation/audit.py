"""audit.py — Audit agenda ↔ facturé (F-19 ; lecture seule du MASTER et de l'agenda).

Rapproche, sur une période [du, au], les séances de l'agenda et les lignes de l'onglet `Sessions`
du MASTER (le journal des séances facturées, ADR003), pour repérer ce qui a été tenu sans être
facturé, facturé sans figurer dans l'agenda, facturé deux fois, ou facturé pour d'autres heures.
Rien n'est corrigé : le rapport dit quoi vérifier, et où.

- `apparier(seances, lignes)` rapproche séances classées et lignes de `Sessions`, par (école,
  date), niveau par niveau sur toute la journée de l'école :
  1. « exact » : même début et même fin, à la minute ; toutes les lignes exactes sont rattachées
     à la séance, pour qu'une séance facturée deux fois se voie ;
  2. « début » : même début ;
  3. « intitulé » : même intitulé (Summary verbatim, casse et espaces normalisés) ;
  4. « cours » : même (promotion, module) ;
  5. « jour sans heure » : lignes sans heure, une pour une, par ordre chronologique, quand il
     reste autant de séances que de lignes ce jour-là ; sinon chaque séance restante porte la
     raison « N séance(s) pour M ligne(s) sans heure le JJ/MM ».
- `auditer(...)` (pure) classe les séances (`classement.classer`), lit `Sessions`
  (`sessions.lire_sessions`), apparie et produit les écarts (`NATURES_ECART`).
- `rapport_texte`, `ecrire_rapport` : le rapport en texte, et ses fichiers .txt et .csv.
- `lancer_audit(profil, du, au)` : lit le MASTER, l'agenda et les préparations des mois
  concernés, puis écrit le rapport sous `_sorties/audits/`. Aucun verrou, aucune écriture dans
  le MASTER.

Lectures retenues (`💡 [proposé — à valider]`, là où le contrat se tait) :
- `heures_agenda` porte les heures facturables recalculées depuis l'agenda (7 h pour une journée
  de 8 h au coefficient 7/8), `heures_facturees` la `Durée_h` de `Sessions` ; pour un écart de
  volume, les deux sommes ;
- `nb_seances_tenues` compte les séances « à facturer » (tenues, dans la période, d'une école
  auditée), `nb_rapprochees` celles d'entre elles qui ont trouvé leur ligne ;
- une séance à venir est appariée comme les autres (sa ligne, facture produite en avance, n'est
  pas « facturée absente ») mais ne donne jamais d'écart d'heures ; seul le doublon de ses lignes
  (double facturation déjà faite) est signalé ;
- une école est `Hors facturation` quand toutes ses lignes de `Règles_facturation` le disent ;
  une école sans règle est auditée ; une école est « à bon de commande » dès qu'une de ses
  lignes le demande ; les lignes `Sessions` d'une école non auditée sont ignorées comme ses
  séances ;
- les niveaux « début », « intitulé » et « cours » ne prennent que des lignes qui portent leurs
  deux heures ; une ligne à laquelle il manque le début ou la fin est « sans heure » (comme dans
  `sessions.chercher`) et n'entre qu'au niveau 5. Le niveau « cours » exige une promotion ou un
  module non vide ;
- quand deux séances visent une même ligne (ou une séance deux lignes) au même niveau, la séance
  la plus tôt choisit la première, et elle prend la ligne la plus proche (fin la plus proche au
  niveau « début », début le plus proche ensuite), puis la ligne la plus haute de `Sessions` :
  le résultat ne dépend jamais de l'ordre des entrées ;
- une même situation de doublon (plusieurs lignes pour une séance, lignes identiques entre elles)
  donne un seul écart « doublon », qui porte toutes ses lignes, et jamais en plus une « tenue non
  facturée », une « facturée absente » ou un écart d'heures ;
- une séance et une ligne de la même école, le même jour, à d'autres heures (constat de la revue
  de P-0b) : si elles sont rapprochées (niveaux 2 à 4), l'écart dit « à vérifier : même jour,
  autres heures » et nomme la ligne (« écart d'heures » si les heures diffèrent, « heures
  incalculables », sinon la huitième nature « même jour, autres heures ») ; si elles restent
  seules, la « tenue non facturée » et la « facturée absente » le disent toutes deux ;
- un créneau à classer tenu est signalé même quand l'audit est limité à des écoles : son école
  est inconnue ; une école demandée qui n'existe nulle part (`Règles_facturation`, `Écoles`,
  `Correspondances`, `Sessions`) lève `ValueError` ;
- l'écart de volume est daté du dernier jour de la période ; un groupe dont une séance a des
  heures incalculables ou une ligne une `Durée_h` vide n'est pas comparé ;
- les lignes illisibles de `Sessions` sont toutes rapportées, quelle que soit leur date (elle peut
  être la cellule illisible) ;
- `lancer_audit` lit aussi la préparation du mois qui suit la période (une école à cycle range sa
  fin de période dans le mois suivant) ; un brouillon est « non émis » tant qu'aucune pièce
  d'émission « émise » ne porte son identifiant ; des pièces illisibles retirent toute annotation ;
  une préparation illisible ou d'une autre entité est ignorée et notée (`RapportAudit.notes`) ;
- le .txt est en UTF-8 sans BOM, comme les autres textes de l'outil ; le .csv en UTF-8 avec BOM,
  séparateur « ; », décimales à la virgule (Excel en français) ; une cellule de texte qui
  commencerait par = + - ou @ est préfixée d'une apostrophe, pour qu'Excel ne la lise jamais
  comme une formule.
"""

from __future__ import annotations

import csv
import dataclasses
import io
import os
import unicodedata
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass, field
from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path
from typing import TYPE_CHECKING, Any, Callable
from zoneinfo import ZoneInfo

from facturation import agenda, chemins, fichiers, journal
from facturation.agenda import Seance
from facturation.classement import SeanceAClasser, SeanceClassee, classer
from facturation.heures import heures_facturables
from facturation.master import Master, lire_master
from facturation.regles import RegleEcole, regles_par_ecole
from facturation.sessions import LigneSessions, code_ecole, lire_sessions

if TYPE_CHECKING:
    from facturation.config import Profil

# Natures d'écart : les sept du contrat, puis 💡 « même jour, autres heures ».
TENUE_NON_FACTUREE = "tenue non facturée"
FACTUREE_ABSENTE = "facturée absente de l'agenda"
ECART_HEURES = "écart d'heures"
DOUBLON = "doublon"
TENUE_NON_RECONNUE = "tenue non reconnue"
HEURES_INCALCULABLES = "heures incalculables"
ECART_VOLUME = "écart de volume"
MEME_JOUR_AUTRES_HEURES = "même jour, autres heures"
NATURES_ECART = (
    TENUE_NON_FACTUREE, FACTUREE_ABSENTE, ECART_HEURES, DOUBLON, TENUE_NON_RECONNUE,
    HEURES_INCALCULABLES, ECART_VOLUME, MEME_JOUR_AUTRES_HEURES,
)

# Niveaux d'appariement, dans l'ordre où ils s'appliquent.
NIVEAU_EXACT = "exact"
NIVEAU_DEBUT = "début"
NIVEAU_INTITULE = "intitulé"
NIVEAU_COURS = "cours"
NIVEAU_SANS_HEURE = "jour sans heure"
NIVEAUX = (NIVEAU_EXACT, NIVEAU_DEBUT, NIVEAU_INTITULE, NIVEAU_COURS, NIVEAU_SANS_HEURE)

STATUT_HORS_FACTURATION = "Hors facturation"
# Écart toléré entre heures facturables recalculées et `Durée_h` : un demi-centième d'heure.
TOLERANCE_HEURES = Decimal("0.005")
NOM_PREPARATION = "preparation.json"
FORMAT_PREPARATION = 1

_A_VERIFIER = "à vérifier : même jour, autres heures"
# Raison d'une séance à classer (`classement.SeanceAClasser.raison`), dite pour Adrian.
_RAISONS_A_CLASSER = {
    "non reconnue": "aucune ligne de l'onglet Correspondances ne reconnaît son intitulé",
    "ambiguë": "plusieurs lignes de l'onglet Correspondances la reconnaissent",
    "code inconnu": "son code d'événement est absent de l'onglet Correspondances, ou mal formé",
}
_ONGLET_SESSIONS = "Sessions"
_COLONNES_CSV = (
    "Nature", "École", "Date", "Début", "Fin", "Intitulé (agenda)", "Lignes Sessions", "N° facture",
    "Heures agenda", "Heures facturées", "Brouillon non émis", "Détail",
)


# ──────────────────────────────────────────────────────────────────────
# Types
# ──────────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class Appariement:
    """Une séance classée et la ou les lignes de `Sessions` qui la journalisent."""

    seance: SeanceClassee
    lignes: tuple[LigneSessions, ...]
    niveau: str        # "exact" | "début" | "intitulé" | "cours" | "jour sans heure"


@dataclass(frozen=True)
class Rapprochement:
    """Résultat d'`apparier` : paires, séances seules, lignes seules, et raison d'une séance restée
    seule au niveau « jour sans heure » (clé : `chemins.cle_seance`)."""

    appariements: tuple[Appariement, ...]
    agenda_seul: tuple[SeanceClassee, ...]
    sessions_seules: tuple[LigneSessions, ...]
    raisons: dict[str, str] = field(hash=False)


@dataclass(frozen=True)
class Ecart:
    """Un écart entre l'agenda et `Sessions` (nature : l'une de `NATURES_ECART`)."""

    nature: str
    ecole: str | None
    date: date
    seance: Seance | None
    lignes: tuple[LigneSessions, ...]
    heures_agenda: float | None
    heures_facturees: float | None
    detail: str
    dans_brouillon: str | None = None   # identifiant d'un brouillon non émis qui contient la séance


@dataclass(frozen=True)
class RapportAudit:
    """Rapport d'un audit agenda ↔ facturé.

    `notes` (💡, hors contrat, en dernier) : ce que l'audit a ignoré en le notant (préparation
    illisible, école demandée mais hors facturation…).
    """

    entite: str
    du: date
    au: date
    etabli_le: datetime
    ecoles: tuple[str, ...]
    nb_seances_tenues: int
    nb_rapprochees: int
    ecarts: tuple[Ecart, ...]
    a_venir: tuple[SeanceClassee, ...]
    hors_facturation: tuple[SeanceClassee, ...]
    lignes_non_datees: tuple[LigneSessions, ...]
    lignes_illisibles: tuple[tuple[int, str], ...]
    notes: tuple[str, ...] = ()

    def sans_ecart(self, ecole: str) -> bool:
        """Vrai si aucun écart ne porte sur cette école (code, sans tenir compte de la casse)."""
        code = _code(ecole)
        return not any(ecart.ecole == code for ecart in self.ecarts)

    @property
    def code_retour(self) -> int:
        """0 sans écart, 3 avec écarts (une action d'Adrian est requise, § 1.6)."""
        return 3 if self.ecarts else 0


@dataclass(frozen=True)
class _LigneLue:
    """Ligne d'un brouillon lu dans `preparation.json` : seule sa clé de séance sert à l'audit."""

    cle_seance: str


@dataclass(frozen=True)
class _BrouillonLu:
    """Brouillon lu dans `preparation.json`, réduit à ce qu'un audit lit."""

    identifiant: str
    lignes: tuple[_LigneLue, ...]


# ──────────────────────────────────────────────────────────────────────
# Outils
# ──────────────────────────────────────────────────────────────────────


def _forme(valeur: Any) -> str:
    """Forme de comparaison d'un texte : NFC, sans casse, espaces réduits et retirés autour."""
    return " ".join(unicodedata.normalize("NFC", "" if valeur is None else str(valeur)).casefold().split())


def _code(valeur: Any) -> str:
    """Code école sous forme comparable : NFC, espaces réduits, majuscules."""
    return " ".join(unicodedata.normalize("NFC", "" if valeur is None else str(valeur)).split()).upper()


def _hm(moment: time | datetime) -> tuple[int, int]:
    """Heure et minute d'une heure, ou d'un moment (heure locale de son fuseau)."""
    return moment.hour, moment.minute


def _minutes(moment: time | datetime) -> int:
    return moment.hour * 60 + moment.minute


def _a_des_heures(ligne: LigneSessions) -> bool:
    """Vrai si la ligne porte son début et sa fin (sinon elle est « sans heure »)."""
    return ligne.debut is not None and ligne.fin is not None


def _autres_heures(classee: SeanceClassee, ligne: LigneSessions) -> bool:
    """Vrai si la ligne porte ses deux heures, et qu'elles diffèrent de celles de la séance."""
    if not _a_des_heures(ligne):
        return False
    seance = classee.seance
    return (_hm(ligne.debut), _hm(ligne.fin)) != (_hm(seance.debut), _hm(seance.fin))


def _ordre_seance(classee: SeanceClassee) -> tuple:
    """Ordre chronologique d'une séance (heure locale), puis tout ce qui la distingue."""
    s = classee.seance
    return (s.date, _hm(s.debut), _hm(s.fin), s.uid, s.intitule, s.description, classee.ecole,
            classee.module, classee.promotion, classee.cle_tarif, classee.par)


def _texte_heure(heure: time | None) -> str:
    return "" if heure is None else f"{heure:%H:%M}"


def _ordre_ligne(ligne: LigneSessions) -> tuple:
    """Ordre d'une ligne de `Sessions` : son rang Excel, puis tout ce qui la distingue."""
    return (ligne.rang, str(ligne.date_seance), ligne.ecole, ligne.n_facture, _texte_heure(ligne.debut),
            _texte_heure(ligne.fin), repr(ligne.duree_h), ligne.summary, ligne.promotion, ligne.module)


def _decimal(nombre: float) -> Decimal:
    return Decimal(repr(float(nombre)))


def _somme(valeurs: Iterable[float]) -> float:
    """Somme exacte de nombres écrits comme ils s'affichent, rendue en float."""
    return float(sum((_decimal(v) for v in valeurs), Decimal(0)))


def _differe(a: float, b: float) -> bool:
    """Vrai si |a − b| dépasse la tolérance d'un demi-centième d'heure (comparaison exacte)."""
    return abs(_decimal(a) - _decimal(b)) > TOLERANCE_HEURES


def _h(nombre: float | None) -> str:
    """Heures à la française, sans rien perdre : 3,5 ; 7 ; 3,504 ; « — » si inconnues."""
    if nombre is None:
        return "—"
    texte = repr(float(nombre))
    if texte.endswith(".0"):
        texte = texte[:-2]
    return texte.replace(".", ",")


def _plage(debut: time | datetime | None, fin: time | datetime | None) -> str:
    if debut is None and fin is None:
        return "sans heure"
    return f"{_texte_heure_lisible(debut)}–{_texte_heure_lisible(fin)}"


def _texte_heure_lisible(moment: time | datetime | None) -> str:
    return "?" if moment is None else f"{moment:%H:%M}"


def _desc_seance(seance: Seance) -> str:
    return f"séance {_plage(seance.debut, seance.fin)} « {seance.intitule} »"


def _desc_ligne(ligne: LigneSessions, *, cours: bool = False) -> str:
    numero = ligne.n_facture or "sans N°"
    suite = f", {ligne.promotion or '—'} / {ligne.module or '—'}" if cours else ""
    return (f"ligne {ligne.rang} de {_ONGLET_SESSIONS} ({numero}, {_plage(ligne.debut, ligne.fin)}, "
            f"Durée_h {_h(ligne.duree_h)}{suite})")


def _rangs(lignes: Sequence[LigneSessions]) -> str:
    return ", ".join(str(ligne.rang) for ligne in lignes)


def _jour(valeur: date) -> date:
    return valeur.date() if isinstance(valeur, datetime) else valeur


def _controler_periode(du: date, au: date) -> tuple[date, date]:
    """(du, au) en dates ; `ValueError` si la période est inversée."""
    du, au = _jour(du), _jour(au)
    if du > au:
        raise ValueError(f"période inversée : du {du:%d/%m/%Y} au {au:%d/%m/%Y} (« du » doit précéder « au »)")
    return du, au


def _champ(objet: Any, nom: str) -> Any:
    """Champ d'un objet (attribut) ou d'un dictionnaire (clé) : lecture des brouillons sans en
    dépendre (F-16a)."""
    if isinstance(objet, Mapping):
        return objet.get(nom)
    return getattr(objet, nom, None)


# ──────────────────────────────────────────────────────────────────────
# Appariement
# ──────────────────────────────────────────────────────────────────────


def _apparier_jour(
    seances: list[SeanceClassee], lignes: list[LigneSessions], jour: date,
) -> tuple[list[Appariement], list[SeanceClassee], list[LigneSessions], dict[str, str]]:
    """Appariement d'une journée d'une école ; séances et lignes déjà triées (`_ordre_*`)."""
    paires: dict[int, tuple[list[LigneSessions], str]] = {}
    prises: set[int] = set()  # id() des lignes prises

    # Niveau 1 : même début et même fin. Toutes les lignes exactes vont à la séance ; deux séances
    # au même créneau (doublon d'agenda) se partagent les lignes une à une, la dernière gardant le reste.
    lignes_par_horaire: dict[tuple, list[LigneSessions]] = {}
    for ligne in lignes:
        if _a_des_heures(ligne):
            lignes_par_horaire.setdefault((_hm(ligne.debut), _hm(ligne.fin)), []).append(ligne)
    seances_par_horaire: dict[tuple, list[int]] = {}
    for i, classee in enumerate(seances):
        seances_par_horaire.setdefault((_hm(classee.seance.debut), _hm(classee.seance.fin)), []).append(i)
    for horaire, indices in seances_par_horaire.items():
        exactes = lignes_par_horaire.get(horaire, [])
        for n, i in enumerate(indices[:len(exactes)]):
            part = exactes[n:] if n == len(indices) - 1 else [exactes[n]]
            paires[i] = (part, NIVEAU_EXACT)
            prises.update(id(ligne) for ligne in part)

    def niveau(nom: str, convient: Callable[[SeanceClassee, LigneSessions], bool],
               distance: Callable[[SeanceClassee, LigneSessions], int]) -> None:
        for i, classee in enumerate(seances):
            if i in paires:
                continue
            candidates = [
                ligne for ligne in lignes
                if id(ligne) not in prises and _a_des_heures(ligne) and convient(classee, ligne)
            ]
            if candidates:
                choisie = min(candidates, key=lambda ligne: (distance(classee, ligne), _ordre_ligne(ligne)))
                paires[i] = ([choisie], nom)
                prises.add(id(choisie))

    def ecart_debut(classee: SeanceClassee, ligne: LigneSessions) -> int:
        return abs(_minutes(ligne.debut) - _minutes(classee.seance.debut))

    # Niveau 2 : même début (la fin la plus proche l'emporte).
    niveau(
        NIVEAU_DEBUT,
        lambda c, ligne: _hm(ligne.debut) == _hm(c.seance.debut),
        lambda c, ligne: abs(_minutes(ligne.fin) - _minutes(c.seance.fin)),
    )
    # Niveau 3 : même intitulé, casse et espaces normalisés.
    niveau(
        NIVEAU_INTITULE,
        lambda c, ligne: bool(_forme(c.seance.intitule)) and _forme(ligne.summary) == _forme(c.seance.intitule),
        ecart_debut,
    )
    # Niveau 4 : même (promotion, module), l'un des deux au moins renseigné.
    niveau(
        NIVEAU_COURS,
        lambda c, ligne: (_forme(c.promotion) or _forme(c.module)) != ""
        and (_forme(ligne.promotion), _forme(ligne.module)) == (_forme(c.promotion), _forme(c.module)),
        ecart_debut,
    )

    # Niveau 5 : lignes sans heure, une pour une si les nombres sont égaux.
    raisons: dict[str, str] = {}
    restantes = [i for i in range(len(seances)) if i not in paires]
    sans_heure = [ligne for ligne in lignes if id(ligne) not in prises and not _a_des_heures(ligne)]
    if restantes and sans_heure:
        if len(restantes) == len(sans_heure):
            for i, ligne in zip(restantes, sans_heure):
                paires[i] = ([ligne], NIVEAU_SANS_HEURE)
                prises.add(id(ligne))
        else:
            raison = f"{len(restantes)} séance(s) pour {len(sans_heure)} ligne(s) sans heure le {jour:%d/%m}"
            for i in restantes:
                raisons[chemins.cle_seance(seances[i].seance)] = raison

    appariements = [
        Appariement(seance=seances[i], lignes=tuple(sorted(paires[i][0], key=_ordre_ligne)), niveau=paires[i][1])
        for i in sorted(paires)
    ]
    seules = [classee for i, classee in enumerate(seances) if i not in paires]
    lignes_seules = [ligne for ligne in lignes if id(ligne) not in prises]
    return appariements, seules, lignes_seules, raisons


def apparier(seances: Sequence[SeanceClassee], lignes: Sequence[LigneSessions]) -> Rapprochement:
    """Rapproche les séances classées et les lignes de `Sessions`, par (école, date).

    Les niveaux (exact, début, intitulé, cours, jour sans heure) s'appliquent dans cet ordre à toute
    la journée de l'école. Le résultat ne dépend pas de l'ordre des entrées. Une ligne sans
    `Date_séance` n'est jamais appariée. Les entrées sont prises telles quelles : c'est à
    l'appelant d'écarter les lignes d'un N° abandonné.
    """
    jours: dict[tuple[str, date], tuple[list[SeanceClassee], list[LigneSessions]]] = {}
    sessions_seules: list[LigneSessions] = []
    for classee in seances:
        jours.setdefault((_code(classee.ecole), classee.seance.date), ([], []))[0].append(classee)
    for ligne in lignes:
        if ligne.date_seance is None:
            sessions_seules.append(ligne)
        else:
            jours.setdefault((_code(ligne.ecole), ligne.date_seance), ([], []))[1].append(ligne)

    appariements: list[Appariement] = []
    agenda_seul: list[SeanceClassee] = []
    raisons: dict[str, str] = {}
    for cle in sorted(jours):
        seances_jour, lignes_jour = jours[cle]
        paires, seules, lignes_seules, raisons_jour = _apparier_jour(
            sorted(seances_jour, key=_ordre_seance), sorted(lignes_jour, key=_ordre_ligne), cle[1],
        )
        appariements.extend(paires)
        agenda_seul.extend(seules)
        sessions_seules.extend(lignes_seules)
        raisons.update(raisons_jour)
    return Rapprochement(
        appariements=tuple(appariements),
        agenda_seul=tuple(agenda_seul),
        sessions_seules=tuple(sorted(sessions_seules, key=_ordre_ligne)),
        raisons=raisons,
    )


# ──────────────────────────────────────────────────────────────────────
# Audit
# ──────────────────────────────────────────────────────────────────────


def _code_tolerant(valeur: Any, ecoles: Sequence[Mapping[str, Any]]) -> str:
    """Code d'une école désignée par son code ou son libellé (onglet `Écoles`) ; à défaut, la
    valeur nettoyée en majuscules."""
    try:
        return code_ecole(valeur, ecoles) or ""
    except ValueError:
        return _code(valeur)


def _avec_code(classee: SeanceClassee, ecoles: Sequence[Mapping[str, Any]]) -> SeanceClassee:
    """La séance classée, son école ramenée au code (la colonne `École` de `Correspondances` peut
    porter le libellé)."""
    code = _code_tolerant(classee.ecole, ecoles)
    return classee if code == classee.ecole else dataclasses.replace(classee, ecole=code)


def _ecoles_hors_facturation(regles: Mapping[Any, RegleEcole]) -> set[str]:
    """Écoles dont toutes les lignes de `Règles_facturation` sont `Hors facturation`."""
    statuts: dict[str, set[str]] = {}
    for regle in regles.values():
        statuts.setdefault(regle.ecole, set()).add(_forme(regle.statut_lot))
    return {ecole for ecole, vus in statuts.items() if vus == {_forme(STATUT_HORS_FACTURATION)}}


def _filtre_ecoles(ecoles: Sequence[str] | None, master: Master, regles: Mapping[Any, RegleEcole],
                   classees: Sequence[SeanceClassee], lignes: Sequence[LigneSessions]) -> set[str] | None:
    """Codes des écoles demandées ; None sans filtre.

    Raises:
        ValueError: liste vide, ou école inconnue de `Règles_facturation`, `Écoles`,
            `Correspondances` et `Sessions` (jamais d'audit vide en silence sur une faute de frappe).
    """
    if ecoles is None:
        return None
    if isinstance(ecoles, str):
        ecoles = [ecoles]
    demandees = [valeur for valeur in ecoles if _code(valeur)]
    if not demandees:
        raise ValueError("liste d'écoles vide : nommez au moins une école (code de Règles_facturation ou de Écoles)")
    connues = {regle.ecole for regle in regles.values()}
    connues |= {_code(ligne.get("Code facturation")) for ligne in master.ecoles}
    connues |= {_code_tolerant(ligne.get("École"), master.ecoles) for ligne in master.correspondances}
    connues |= {classee.ecole for classee in classees}
    connues |= {ligne.ecole for ligne in lignes}
    connues.discard("")
    filtre = set()
    for valeur in demandees:
        code = _code_tolerant(valeur, master.ecoles)
        if code not in connues:
            raise ValueError(
                f"école « {valeur} » inconnue : absente des onglets Règles_facturation, Écoles, Correspondances "
                "et Sessions. Vérifiez le code demandé."
            )
        filtre.add(code)
    return filtre


def _index_brouillons(brouillons: Sequence[Any]) -> dict[str, str]:
    """Clé de séance → identifiant du premier brouillon (non émis) qui la contient."""
    index: dict[str, str] = {}
    for brouillon in brouillons:
        identifiant = _champ(brouillon, "identifiant")
        if not identifiant:
            continue
        for ligne in _champ(brouillon, "lignes") or ():
            cle = _champ(ligne, "cle_seance")
            if cle:
                index.setdefault(str(cle), str(identifiant))
    return index


class _Ensembles:
    """Union-find des séances et des lignes d'une même situation de doublon (clés : id())."""

    def __init__(self) -> None:
        self._parent: dict[int, int] = {}

    def trouver(self, x: int) -> int:
        self._parent.setdefault(x, x)
        while self._parent[x] != x:
            self._parent[x] = self._parent[self._parent[x]]
            x = self._parent[x]
        return x

    def unir(self, a: int, b: int) -> None:
        ra, rb = self.trouver(a), self.trouver(b)
        if ra != rb:
            self._parent[rb] = ra


def _ordre_ecart(ecart: Ecart) -> tuple:
    if ecart.nature == ECART_VOLUME:
        heure: tuple[int, int] = (99, 0)
    elif ecart.seance is not None:
        heure = _hm(ecart.seance.debut)
    elif ecart.lignes and ecart.lignes[0].debut is not None:
        heure = _hm(ecart.lignes[0].debut)
    else:
        heure = (98, 0)
    return (ecart.date, ecart.ecole or "", heure, NATURES_ECART.index(ecart.nature),
            tuple(ligne.rang for ligne in ecart.lignes))


def auditer(*, entite: str, du: date, au: date, seances: Sequence[Seance], master: Master,
            maintenant: datetime, ecoles: Sequence[str] | None = None,
            brouillons: Sequence[Any] = ()) -> RapportAudit:
    """Audit agenda ↔ facturé sur [du, au] (bornes incluses). Fonction pure.

    Args:
        entite: code de l'entité du rapport.
        seances: séances de l'agenda (celles hors de la période sont ignorées).
        master: MASTER lu (`lire_master`) ; seuls `Règles_facturation`, `Correspondances`,
            `Écoles`, `Sessions` et `Factures` (N° abandonnés) servent.
        maintenant: une séance est tenue si `fin <= maintenant`.
        ecoles: codes (ou libellés) des écoles à auditer ; None pour toutes.
        brouillons: brouillons non émis (F-16a), lus par `identifiant` et `lignes[].cle_seance`.

    Raises:
        ValueError: période inversée ; liste d'écoles vide ou école inconnue ; ligne mal formée de
            `Correspondances` ou de `Règles_facturation` (message du module qui la lit).
    """
    du, au = _controler_periode(du, au)
    regles = regles_par_ecole(master.regles_facturation)
    classees_brutes, a_classer = classer(seances, master.correspondances)
    lignes, illisibles = lire_sessions(master)
    classees = [_avec_code(classee, master.ecoles) for classee in classees_brutes]
    hors = _ecoles_hors_facturation(regles)
    bon_de_commande = {regle.ecole for regle in regles.values() if regle.bc_requis}
    filtre = _filtre_ecoles(ecoles, master, regles, classees, lignes)
    notes: list[str] = []
    if filtre is not None:
        for code in sorted(filtre & hors):
            notes.append(f"école {code} « {STATUT_HORS_FACTURATION} » (Règles_facturation) : non auditée")

    def auditee(code: str) -> bool:
        return code not in hors and (filtre is None or code in filtre)

    def dans_periode(jour: date) -> bool:
        return du <= jour <= au

    # Séances classées.
    a_facturer: list[SeanceClassee] = []
    a_venir: list[SeanceClassee] = []
    hors_facturation: list[SeanceClassee] = []
    for classee in classees:
        if not dans_periode(classee.seance.date) or (filtre is not None and classee.ecole not in filtre):
            continue
        if classee.ecole in hors:
            hors_facturation.append(classee)
        elif classee.seance.fin <= maintenant:
            a_facturer.append(classee)
        else:
            a_venir.append(classee)

    # Lignes de Sessions.
    concernees: list[LigneSessions] = []
    non_datees: list[LigneSessions] = []
    for ligne in lignes:
        if ligne.abandonnee or not auditee(ligne.ecole):
            continue
        if ligne.date_seance is None:
            non_datees.append(ligne)
        elif dans_periode(ligne.date_seance):
            concernees.append(ligne)

    rapprochement = apparier(a_facturer + a_venir, concernees)
    tenues = {id(classee) for classee in a_facturer}
    index_brouillons = _index_brouillons(brouillons)

    calculs: dict[int, float | ValueError] = {}

    def facturables(classee: SeanceClassee) -> float | ValueError:
        if id(classee) not in calculs:
            try:
                calculs[id(classee)] = heures_facturables(classee.seance, classee.correspondance)
            except ValueError as exc:
                calculs[id(classee)] = exc
        return calculs[id(classee)]

    def ecart(nature: str, ecole: str | None, jour: date, seance: Seance | None, lignes_ecart: Iterable[LigneSessions],
              heures_agenda: float | None, heures_facturees: float | None, detail: str) -> Ecart:
        brouillon = index_brouillons.get(chemins.cle_seance(seance)) if seance is not None else None
        return Ecart(
            nature=nature, ecole=ecole or None, date=jour, seance=seance,
            lignes=tuple(sorted(lignes_ecart, key=_ordre_ligne)), heures_agenda=heures_agenda,
            heures_facturees=heures_facturees, detail=detail, dans_brouillon=brouillon,
        )

    ecarts: list[Ecart] = []

    # Doublons : plusieurs lignes pour une séance, ou lignes identiques (école, date, début, fin),
    # tous N° confondus. Une situation = un seul écart.
    ensembles = _Ensembles()
    objets: dict[int, SeanceClassee | LigneSessions] = {}
    for appariement in rapprochement.appariements:
        objets[id(appariement.seance)] = appariement.seance
        ensembles.trouver(id(appariement.seance))
        for ligne in appariement.lignes:
            objets[id(ligne)] = ligne
            ensembles.unir(id(appariement.seance), id(ligne))
    identiques: dict[tuple, list[LigneSessions]] = {}
    for ligne in concernees:
        if _a_des_heures(ligne):
            identiques.setdefault((ligne.ecole, ligne.date_seance, _hm(ligne.debut), _hm(ligne.fin)), []).append(ligne)
    for groupe in identiques.values():
        for ligne in groupe:
            objets[id(ligne)] = ligne
            ensembles.unir(id(groupe[0]), id(ligne))
    composantes: dict[int, list[SeanceClassee | LigneSessions]] = {}
    for cle, objet in objets.items():
        composantes.setdefault(ensembles.trouver(cle), []).append(objet)
    en_doublon: set[int] = set()
    for membres in composantes.values():
        lignes_c = sorted((m for m in membres if isinstance(m, LigneSessions)), key=_ordre_ligne)
        if len(lignes_c) < 2:
            continue
        seances_c = sorted((m for m in membres if isinstance(m, SeanceClassee)), key=_ordre_seance)
        en_doublon.update(id(m) for m in membres)
        principale = seances_c[0] if seances_c else None
        heures_agenda = None
        if len(seances_c) == 1 and not isinstance(facturables(principale), ValueError):
            heures_agenda = facturables(principale)
        durees = [ligne.duree_h for ligne in lignes_c]
        heures_facturees = None if None in durees else _somme(durees)
        numeros = ", ".join(f"ligne {ligne.rang} ({ligne.n_facture or 'sans N°'})" for ligne in lignes_c)
        if principale is None:
            objet = f"un même créneau {_plage(lignes_c[0].debut, lignes_c[0].fin)} absent de l'agenda"
        else:
            objet = "la " + " et la ".join(_desc_seance(c.seance) for c in seances_c)
        detail = (f"{len(lignes_c)} lignes de {_ONGLET_SESSIONS} pour {objet} : {numeros} ; une séance ne se "
                  "facture qu'une fois")
        if principale is not None and any(_autres_heures(c, ligne) for c in seances_c for ligne in lignes_c):
            detail = f"{_A_VERIFIER} — {detail}"
        ecarts.append(ecart(DOUBLON, lignes_c[0].ecole, lignes_c[0].date_seance,
                            principale.seance if principale else None, lignes_c, heures_agenda,
                            heures_facturees, detail))

    # Paires d'une seule ligne : heures, et même jour à d'autres heures.
    for appariement in rapprochement.appariements:
        classee = appariement.seance
        if id(classee) in en_doublon or id(classee) not in tenues:
            continue  # une séance à venir ne donne jamais d'écart d'heures
        (ligne,) = appariement.lignes
        seance = classee.seance
        autres = _autres_heures(classee, ligne)
        prefixe = f"{_A_VERIFIER} — {_desc_seance(seance)}, " if autres else ""
        calcul = facturables(classee)
        if isinstance(calcul, ValueError):
            ecarts.append(ecart(HEURES_INCALCULABLES, classee.ecole, seance.date, seance, [ligne], None,
                                ligne.duree_h, f"{prefixe}{_desc_ligne(ligne)} : {calcul}"))
        elif ligne.duree_h is None:
            ecarts.append(ecart(ECART_HEURES, classee.ecole, seance.date, seance, [ligne], calcul, None,
                                f"{prefixe}{_desc_ligne(ligne)} : Durée_h vide, {_h(calcul)} h facturables "
                                "d'après l'agenda"))
        elif _differe(calcul, ligne.duree_h):
            ecarts.append(ecart(ECART_HEURES, classee.ecole, seance.date, seance, [ligne], calcul, ligne.duree_h,
                                f"{prefixe}{_desc_ligne(ligne)} : {_h(calcul)} h facturables d'après l'agenda, "
                                f"{_h(ligne.duree_h)} h facturées"))
        elif autres:
            ecarts.append(ecart(MEME_JOUR_AUTRES_HEURES, classee.ecole, seance.date, seance, [ligne], calcul,
                                ligne.duree_h,
                                f"à vérifier : {_desc_seance(seance)} rapprochée de la {_desc_ligne(ligne)} "
                                f"(niveau « {appariement.niveau} »), {_h(calcul)} h des deux côtés"))

    # Séances et lignes restées seules.
    lignes_seules_jour: dict[tuple[str, date], list[LigneSessions]] = {}
    for ligne in rapprochement.sessions_seules:
        if id(ligne) not in en_doublon:
            lignes_seules_jour.setdefault((_code(ligne.ecole), ligne.date_seance), []).append(ligne)
    seances_seules_jour: dict[tuple[str, date], list[SeanceClassee]] = {}
    for classee in rapprochement.agenda_seul:
        seances_seules_jour.setdefault((_code(classee.ecole), classee.seance.date), []).append(classee)
    a_classer_jour: dict[date, list[SeanceAClasser]] = {}
    for a_ranger in a_classer:
        a_classer_jour.setdefault(a_ranger.seance.date, []).append(a_ranger)

    for classee in rapprochement.agenda_seul:
        if id(classee) not in tenues:
            continue
        seance = classee.seance
        detail = f"{_desc_seance(seance)} ({classee.promotion or '—'} / {classee.module or '—'}) : aucune ligne de {_ONGLET_SESSIONS}"
        raison = rapprochement.raisons.get(chemins.cle_seance(seance))
        if raison:
            detail += f" ({raison})"
        autres = [ligne for ligne in lignes_seules_jour.get((_code(classee.ecole), seance.date), [])
                  if _a_des_heures(ligne)]
        if autres:
            detail += f" — {_A_VERIFIER} : " + " ; ".join(_desc_ligne(ligne) for ligne in autres)
        calcul = facturables(classee)
        ecarts.append(ecart(TENUE_NON_FACTUREE, classee.ecole, seance.date, seance, [],
                            None if isinstance(calcul, ValueError) else calcul, None, detail))

    for ligne in rapprochement.sessions_seules:
        if id(ligne) in en_doublon:
            continue
        detail = f"{_desc_ligne(ligne, cours=True)} : aucune séance de l'agenda"
        seules = seances_seules_jour.get((_code(ligne.ecole), ligne.date_seance), [])
        if seules and _a_des_heures(ligne):
            detail += f" — {_A_VERIFIER} : " + " ; ".join(_desc_seance(c.seance) for c in seules)
        elif seules:
            raisons = sorted({rapprochement.raisons.get(chemins.cle_seance(c.seance), "") for c in seules} - {""})
            if raisons:
                detail += f" ({' ; '.join(raisons)})"
        a_ranger = a_classer_jour.get(ligne.date_seance, [])
        if a_ranger:
            detail += " — créneau à classer ce jour-là : " + " ; ".join(_desc_seance(a.seance) for a in a_ranger)
        ecarts.append(ecart(FACTUREE_ABSENTE, ligne.ecole, ligne.date_seance, None, [ligne], None, ligne.duree_h, detail))

    # Créneaux à classer tenus : jamais « tenue non facturée » (leur école est inconnue).
    for a_ranger in a_classer:
        seance = a_ranger.seance
        if dans_periode(seance.date) and seance.fin <= maintenant:
            ecarts.append(ecart(
                TENUE_NON_RECONNUE, None, seance.date, seance, [], None, None,
                f"{_desc_seance(seance)} : créneau à classer, "
                f"{_RAISONS_A_CLASSER.get(a_ranger.raison, a_ranger.raison)} ; jamais facturé par défaut",
            ))

    # Écart de volume des écoles à bon de commande, par (école, promotion).
    lignes_a_venir = {id(ligne) for a in rapprochement.appariements if id(a.seance) not in tenues for ligne in a.lignes}
    for ecole in sorted(bon_de_commande):
        if not auditee(ecole):
            continue
        groupes: dict[str, dict[str, Any]] = {}

        def groupe(promotion: str) -> dict[str, Any]:
            return groupes.setdefault(_forme(promotion), {
                "promotion": promotion, "agenda": [], "facture": [], "lignes": [], "comparable": True,
            })

        for classee in a_facturer:
            if classee.ecole == ecole:
                calcul = facturables(classee)
                g = groupe(classee.promotion)
                if isinstance(calcul, ValueError):
                    g["comparable"] = False
                else:
                    g["agenda"].append(calcul)
        for ligne in concernees:
            if ligne.ecole == ecole and id(ligne) not in lignes_a_venir:
                g = groupe(ligne.promotion)
                g["lignes"].append(ligne)
                if ligne.duree_h is None:
                    g["comparable"] = False
                else:
                    g["facture"].append(ligne.duree_h)
        for cle in sorted(groupes):
            g = groupes[cle]
            agenda_h, facture_h = _somme(g["agenda"]), _somme(g["facture"])
            if g["comparable"] and _differe(agenda_h, facture_h):
                ecarts.append(ecart(
                    ECART_VOLUME, ecole, au, None, g["lignes"], agenda_h, facture_h,
                    f"promotion {g['promotion'] or '—'}, du {du:%d/%m/%Y} au {au:%d/%m/%Y} : {_h(agenda_h)} h "
                    f"facturables d'après l'agenda, {_h(facture_h)} h facturées (bon de commande requis ; "
                    f"lignes de {_ONGLET_SESSIONS} : {_rangs(g['lignes']) or 'aucune'})",
                ))

    rapprochees = {id(a.seance) for a in rapprochement.appariements}
    ecoles_auditees = (
        filtre - hors if filtre is not None else
        ({regle.ecole for regle in regles.values()} - hors)
        | {classee.ecole for classee in a_facturer + a_venir}
        | {ligne.ecole for ligne in concernees + non_datees if ligne.ecole}
    )
    return RapportAudit(
        entite=entite,
        du=du,
        au=au,
        etabli_le=maintenant,
        ecoles=tuple(sorted(ecoles_auditees)),
        nb_seances_tenues=len(a_facturer),
        nb_rapprochees=sum(id(classee) in rapprochees for classee in a_facturer),
        ecarts=tuple(sorted(ecarts, key=_ordre_ecart)),
        a_venir=tuple(sorted(a_venir, key=_ordre_seance)),
        hors_facturation=tuple(sorted(hors_facturation, key=_ordre_seance)),
        lignes_non_datees=tuple(sorted(non_datees, key=_ordre_ligne)),
        lignes_illisibles=tuple(illisibles),
        notes=tuple(notes),
    )


# ──────────────────────────────────────────────────────────────────────
# Rapport
# ──────────────────────────────────────────────────────────────────────


def rapport_texte(rapport: RapportAudit) -> str:
    """Le rapport lisible : entité, période, écarts (nature, école, date, détail), puis lignes
    illisibles, lignes sans date, séances à venir, écoles hors facturation et notes."""
    etabli = rapport.etabli_le
    lignes = [
        f"Audit agenda ↔ facturé — entité {rapport.entite}",
        f"Période : du {rapport.du:%d/%m/%Y} au {rapport.au:%d/%m/%Y} (établi le {etabli:%d/%m/%Y} à {etabli:%H:%M})",
        f"Écoles auditées : {', '.join(rapport.ecoles) or 'aucune'}",
        f"Séances tenues à facturer : {rapport.nb_seances_tenues} ; rapprochées d'une ligne de "
        f"{_ONGLET_SESSIONS} : {rapport.nb_rapprochees}",
        "",
    ]
    if not rapport.ecarts:
        lignes.append("Aucun écart sur la période.")
    else:
        comptes = [(nature, sum(e.nature == nature for e in rapport.ecarts)) for nature in NATURES_ECART]
        resume = ", ".join(f"{nature} : {n}" for nature, n in comptes if n)
        lignes.append(f"ÉCARTS : {len(rapport.ecarts)} ({resume})")
        for numero, e in enumerate(rapport.ecarts, start=1):
            lignes.append(f"{numero:>3}. {e.date:%d/%m/%Y}  {e.ecole or '—'}  [{e.nature}]")
            lignes.append(f"     {e.detail}")
            if e.dans_brouillon:
                lignes.append(f"     Présente dans le brouillon non émis {e.dans_brouillon} (préparation du mois).")
    if rapport.lignes_illisibles:
        lignes += ["", f"LIGNES DE {_ONGLET_SESSIONS.upper()} ILLISIBLES : {len(rapport.lignes_illisibles)} (à corriger dans le MASTER)"]
        lignes += [f"  - ligne {rang} : {message}" for rang, message in rapport.lignes_illisibles]
    if rapport.lignes_non_datees:
        lignes += ["", f"LIGNES DE {_ONGLET_SESSIONS.upper()} SANS DATE DE SÉANCE : {len(rapport.lignes_non_datees)} "
                       "(non rapprochées ; renseignez « Date_séance »)"]
        lignes += [f"  - {_desc_ligne(ligne)}, école {ligne.ecole or '—'}" for ligne in rapport.lignes_non_datees]
    if rapport.a_venir:
        lignes += ["", f"SÉANCES À VENIR : {len(rapport.a_venir)} (jamais des écarts)"]
        lignes += [f"  - {c.seance.date:%d/%m/%Y}  {c.ecole}  {_desc_seance(c.seance)}" for c in rapport.a_venir]
    if rapport.hors_facturation:
        par_ecole: dict[str, int] = {}
        for classee in rapport.hors_facturation:
            par_ecole[classee.ecole] = par_ecole.get(classee.ecole, 0) + 1
        lignes += ["", f"HORS FACTURATION (non auditées) : "
                       + ", ".join(f"{ecole} ({n} séance(s))" for ecole, n in sorted(par_ecole.items()))]
    if rapport.notes:
        lignes += ["", "NOTES"]
        lignes += [f"  - {note}" for note in rapport.notes]
    return fichiers.masquer_secrets("\n".join(lignes) + "\n")


def _cellule_sure(texte: str) -> str:
    """Texte d'une cellule CSV qu'Excel ne lira jamais comme une formule."""
    return f"'{texte}" if texte[:1] in ("=", "+", "-", "@") else texte


def _csv(rapport: RapportAudit) -> str:
    tampon = io.StringIO()
    ecrivain = csv.writer(tampon, delimiter=";", lineterminator="\r\n")
    ecrivain.writerow(_COLONNES_CSV)
    for e in rapport.ecarts:
        debut = fin = intitule = ""
        if e.seance is not None:
            debut, fin, intitule = f"{e.seance.debut:%H:%M}", f"{e.seance.fin:%H:%M}", e.seance.intitule
        elif len(e.lignes) == 1 and e.nature != ECART_VOLUME:
            debut, fin = _texte_heure(e.lignes[0].debut), _texte_heure(e.lignes[0].fin)
        numeros = ", ".join(dict.fromkeys(ligne.n_facture for ligne in e.lignes if ligne.n_facture))
        ecrivain.writerow([
            e.nature, e.ecole or "", f"{e.date:%d/%m/%Y}", debut, fin, _cellule_sure(intitule), _rangs(e.lignes),
            _cellule_sure(numeros), "" if e.heures_agenda is None else _h(e.heures_agenda),
            "" if e.heures_facturees is None else _h(e.heures_facturees), e.dans_brouillon or "",
            _cellule_sure(e.detail),
        ])
    return fichiers.masquer_secrets(tampon.getvalue())


def _reserver(chemin: Path) -> bool:
    """Crée le fichier vide s'il n'existe pas (`O_EXCL`) : le nom est alors réservé à cet appel."""
    try:
        descripteur = os.open(chemin, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except FileExistsError:
        return False
    os.close(descripteur)
    return True


def ecrire_rapport(rapport: RapportAudit, dossier: Path) -> tuple[Path, Path]:
    """Écrit `audit_<du>_<au>_<AAAAMMJJ-HHMMSS>.txt` (texte du rapport, UTF-8) et `.csv` (un écart
    par ligne, UTF-8 avec BOM, séparateur « ; ») dans le dossier, qu'il crée au besoin.

    Les écritures sont atomiques. Un nom déjà pris reçoit le suffixe `-1`, `-2`… : rien n'est
    jamais écrasé. L'horodatage est celui du rapport (`etabli_le`).

    Returns:
        (chemin du .txt, chemin du .csv).
    """
    dossier = Path(dossier)
    dossier.mkdir(parents=True, exist_ok=True)
    base = f"audit_{rapport.du:%Y-%m-%d}_{rapport.au:%Y-%m-%d}_{rapport.etabli_le:%Y%m%d-%H%M%S}"
    suffixe = 0
    while True:
        nom = base if suffixe == 0 else f"{base}-{suffixe}"
        txt, csv_ = dossier / f"{nom}.txt", dossier / f"{nom}.csv"
        if not csv_.exists() and _reserver(txt):
            break
        suffixe += 1
    try:
        fichiers.ecrire_texte_atomique(txt, rapport_texte(rapport))
        fichiers.ecrire_texte_atomique(csv_, "﻿" + _csv(rapport))
    except BaseException:
        for chemin in (txt, csv_):
            try:
                chemin.unlink(missing_ok=True)
            except OSError:
                pass
        raise
    return txt, csv_


# ──────────────────────────────────────────────────────────────────────
# Lancement
# ──────────────────────────────────────────────────────────────────────


def _mois_de(du: date, au: date) -> list[str]:
    """Mois civils que touche [du, au], puis le mois qui suit (fin de période d'une école à cycle)."""
    mois = [f"{du.year}-{du.month:02d}"]
    dernier = f"{au.year}-{au.month:02d}"
    while mois[-1] != dernier:
        mois.append(chemins.mois_suivant(mois[-1]))
    mois.append(chemins.mois_suivant(dernier))
    return mois


def _brouillons_de(donnees: Any, entite: str) -> list[_BrouillonLu]:
    """Brouillons d'une préparation lue (format du § 4).

    Raises:
        ValueError: préparation d'une autre entité, de format inconnu, ou mal formée.
    """
    if not isinstance(donnees, dict):
        raise ValueError("un objet JSON est attendu")
    if donnees.get("entite") != entite:
        raise ValueError(
            f"entité « {donnees.get('entite')} », alors que le profil est celui de « {entite} » "
            "(les entités ne se mélangent jamais)"
        )
    if donnees.get("format", FORMAT_PREPARATION) != FORMAT_PREPARATION:
        raise ValueError(f"format {donnees.get('format')!r} inconnu (attendu {FORMAT_PREPARATION})")
    liste = donnees.get("brouillons")
    if not isinstance(liste, list):
        raise ValueError("liste « brouillons » absente")
    resultat = []
    for brouillon in liste:
        if (not isinstance(brouillon, dict) or not isinstance(brouillon.get("identifiant"), str)
                or not isinstance(brouillon.get("lignes"), list)):
            raise ValueError("brouillon mal formé (identifiant et lignes attendus)")
        if brouillon.get("entite", entite) != entite:
            raise ValueError(f"brouillon {brouillon['identifiant']} d'une autre entité (« {brouillon.get('entite')} »)")
        lignes = []
        for ligne in brouillon["lignes"]:
            if not isinstance(ligne, dict) or not isinstance(ligne.get("cle_seance"), str):
                raise ValueError(f"brouillon {brouillon['identifiant']} : ligne sans « cle_seance »")
            lignes.append(_LigneLue(cle_seance=ligne["cle_seance"]))
        resultat.append(_BrouillonLu(identifiant=brouillon["identifiant"], lignes=tuple(lignes)))
    return resultat


def _brouillons_non_emis(profil: Profil, du: date, au: date) -> tuple[list[_BrouillonLu], list[str]]:
    """Brouillons non émis des préparations des mois concernés, et ce qui a été ignoré (notes)."""
    brouillons: list[_BrouillonLu] = []
    notes: list[str] = []
    for mois in _mois_de(du, au):
        chemin = chemins.dossier_mois(profil, mois) / NOM_PREPARATION
        if not chemin.is_file():
            continue
        relatif = chemin.relative_to(profil.dossier_entite).as_posix()
        try:
            brouillons.extend(_brouillons_de(fichiers.lire_json(chemin), profil.entite))
        except (ValueError, OSError) as exc:  # FichierAbime est une ValueError
            notes.append(f"préparation de {mois} ignorée ({relatif}) : {exc}")
    if not brouillons:
        return [], notes
    try:
        emises = {piece.identifiant for piece in journal.pieces_emission(profil) if piece.statut == journal.EMISE}
    except (ValueError, OSError) as exc:
        notes.append(f"pièces d'émission illisibles, séances non rapprochées des brouillons : {exc}")
        return [], notes
    return [b for b in brouillons if b.identifiant not in emises], notes


def lancer_audit(profil: Profil, du: date, au: date, *, contenu_ics: bytes | None = None,
                 maintenant: datetime | None = None,
                 ecoles: Sequence[str] | None = None) -> tuple[RapportAudit, Path]:
    """Audit de la période [du, au] pour l'entité du profil, rapport écrit sous `_sorties/audits/`.

    Lit le MASTER et l'agenda (`contenu_ics`, ou téléchargement par l'adresse iCal du profil), et
    les préparations des mois concernés, dont les erreurs sont ignorées et notées. Ne prend aucun
    verrou et ne modifie pas le MASTER.

    Returns:
        (le rapport, le chemin du .txt ; le .csv est à côté, même nom).

    Raises:
        ValueError: racine dans un dépôt git, période inversée, agenda illisible, MASTER invalide
            (`MasterInvalide`), règle ou correspondance mal formée, école inconnue.
        FileNotFoundError: racine ou MASTER introuvable.
        agenda.AgendaIndisponible: agenda impossible à télécharger.
        master.MasterVerrouille: MASTER impossible à ouvrir.
        OSError: écriture du rapport impossible.
    """
    chemins.refuser_racine_dans_un_depot(profil.racine)
    du, au = _controler_periode(du, au)
    if not Path(profil.racine).is_dir():
        raise FileNotFoundError(
            f"dossier de données introuvable : {profil.racine} ; si c'est un lecteur Google Drive, "
            "vérifiez qu'il est lancé"
        )
    if maintenant is None:
        maintenant = datetime.now(ZoneInfo(profil.fuseau))
    master = lire_master(profil.master)
    contenu = contenu_ics if contenu_ics is not None else agenda.telecharger_ics(profil)
    seances = agenda.lire_ics(contenu, debut=du, fin=au, fuseau=profil.fuseau)
    brouillons, notes = _brouillons_non_emis(profil, du, au)
    rapport = auditer(entite=profil.entite, du=du, au=au, seances=seances, master=master,
                      maintenant=maintenant, ecoles=ecoles, brouillons=brouillons)
    if notes:
        rapport = dataclasses.replace(rapport, notes=rapport.notes + tuple(notes))
    txt, _ = ecrire_rapport(rapport, chemins.dossier_audits(profil))
    return rapport, txt
