"""Audit agenda ↔ facturé (F-19, `facturation/audit.py` ; banque de tests : T-AUD-01 à T-AUD-09 et
T-AUD-12 à T-AUD-32 ; T-AUD-10, T-AUD-11 et T-AUD-33 à T-AUD-43 — `lancer_audit` et la commande
`facturation audit` — dans tests/integration/test_audit.py).

Tests d'acceptation écrits avant l'implémentation : chaque test importe `facturation.audit` lui-même,
pour que l'absence du module fasse échouer les tests sans bloquer la collecte des autres. Données
fictives créées à la volée (écoles ABC, ECO, GHI, DEF, HFX, ZZZ ; entité ROSARI) ; MASTER construits
en mémoire (`master.Master`), ou écrits dans `tmp_path` quand le rang Excel compte.

Contrat (contrats des vagues 3 à 5, § 6, avec les § 15 et 16 qui priment) :
- `auditer(*, entite, du, au, seances, master, maintenant, ecoles=None, brouillons=())` est pure ;
- séances concernées : classées pour une école qui n'est pas `Hors facturation` (les écoles
  `Plus tard` sont auditées), filtrées par `ecoles` ; tenue (`fin <= maintenant`) et datée entre
  `du` et `au` → « à facturer » ; à venir → `a_venir`, jamais un écart ; créneau à classer tenu →
  « tenue non reconnue » (école None), jamais « tenue non facturée » ;
- lignes `Sessions` concernées : celles de `lire_sessions` dans la période, hors lignes d'un N°
  abandonné ; sans `Date_séance` → `lignes_non_datees` ; illisible → `lignes_illisibles` (rang) ;
- appariement déterministe par (école, date), dans l'ordre : même début et même fin (« exact »),
  même début (« début »), même intitulé normalisé (« intitulé »), même (promotion, module)
  (« cours »), lignes sans heure un pour un par ordre chronologique si les nombres sont égaux
  (« jour sans heure »), sinon la raison « N séance(s) pour M ligne(s) sans heure le JJ/MM » ;
- écarts : séance seule et tenue → « tenue non facturée » ; ligne seule → « facturée absente de
  l'agenda » ; |heures facturables recalculées − Durée_h| > 0,005 → « écart d'heures » ;
  `ValueError` de `heures_facturables` → « heures incalculables » ; deux lignes appariées à une
  séance, ou deux lignes identiques (école, date, début, fin), tous N° confondus → « doublon » ;
  école à bon de commande : Σ heures facturables ≠ Σ Durée_h par (école, promotion) → « écart de
  volume » ; `dans_brouillon` = identifiant du brouillon non émis qui contient la séance ;
- `RapportAudit.code_retour` : 0 sans écart, 3 avec écarts ; `sans_ecart(ecole)` ;
- `ecrire_rapport` : `audit_<du>_<au>_<AAAAMMJJ-HHMMSS>.txt` et `.csv` (UTF-8 avec BOM, `;`),
  écrits de façon atomique, avec des noms uniques.

Lectures retenues (`💡 [proposé — à valider]`) :
- `heures_agenda` d'un écart d'heures porte les heures facturables recalculées depuis l'agenda
  (7.0 pour une journée de 8 h au coefficient 7/8), `heures_facturees` la `Durée_h` de `Sessions` ;
  pour un écart de volume, les deux sommes ;
- `nb_seances_tenues` compte les séances « à facturer » (tenues, dans la période, d'une école
  auditée) ; `nb_rapprochees`, celles d'entre elles qui ont trouvé leur ligne ;
- une séance à venir déjà journalisée dans `Sessions` (facture produite en avance) est appariée :
  sa ligne n'est pas « facturée absente de l'agenda » ;
- une école classée mais absente de `Règles_facturation` est auditée (elle n'est pas `Hors
  facturation`) ; les lignes `Sessions` d'une école `Hors facturation` ou hors du filtre `ecoles`
  sont ignorées comme ses séances ;
- les lignes sans heure n'entrent qu'au niveau « jour sans heure » quand leur intitulé et leur
  (promotion, module) ne désignent aucune séance : les tests leur donnent des cellules vides, pour
  rester vrais quelle que soit la place des lignes sans heure dans les niveaux 3 et 4 ;
- une séance et une ligne `Sessions` de la même école, le même jour, à d'autres heures, sortent en
  écart dont le texte (nature et détail) dit « à vérifier : même jour, autres heures » et nomme la
  ligne (« ligne N ») : jamais une « tenue non facturée » muette (constat écarté de la revue de P-0b,
  T-AUD-13). Cela vaut pour une paire rapprochée au niveau « début » ou « intitulé », même sans écart
  d'heures, comme pour une séance et une ligne restées seules ; la nature est libre (une huitième
  nature ou le détail d'une nature du contrat). Sans un tel écart, le texte du rapport ne porte pas
  la mention (pas de légende des natures qui la citerait toujours) ;
- `NATURES_ECART` commence par les sept natures du contrat ; toute nature produite y figure ;
- `rapport_texte` d'un rapport sans écart dit « aucun écart » ; il cite l'entité, la période
  (JJ/MM/AAAA), et pour chaque écart sa nature, son école et sa date ;
- `auditer` lève `ValueError` sur une période inversée (message « inversée »), et laisse remonter
  la `ValueError` d'une ligne de `Correspondances` ou de `Règles_facturation` mal formée.
"""

from __future__ import annotations

import csv
import dataclasses
import importlib
import io
import random
import re
from datetime import date, datetime, time
from types import SimpleNamespace
from zoneinfo import ZoneInfo

import pytest

from facturation.agenda import Seance
from facturation.chemins import cle_seance
from facturation.classement import SeanceClassee
from facturation.master import ONGLETS_PREVUS, Master, lire_master
from facturation.sessions import LigneSessions
from tests.conftest import creer_master

pytestmark = pytest.mark.unit

PARIS = ZoneInfo("Europe/Paris")
DU, AU = date(2026, 9, 1), date(2026, 10, 31)
MAINTENANT = datetime(2026, 11, 26, 16, 0, tzinfo=PARIS)

NATURES_CONTRAT = (
    "tenue non facturée", "facturée absente de l'agenda", "écart d'heures", "doublon",
    "tenue non reconnue", "heures incalculables", "écart de volume",
)
TENUE_NON_FACTUREE, ABSENTE, ECART_HEURES, DOUBLON, NON_RECONNUE, INCALCULABLES, VOLUME = NATURES_CONTRAT
MEME_JOUR = "même jour, autres heures"

# Intitulés fictifs de l'agenda, reconnus par une ligne `exact` de Correspondances.
I_ABC = "Cours fictif ABC M1"
I_ECO = "Cours fictif ECO M2"
I_GHI = "Atelier fictif GHI"
I_DEF = "Cours fictif DEF M1"
I_HFX = "Cours fictif HFX"
I_COEF = "Journée fictive ABC"
I_FORFAIT = "Jury fictif ABC"
I_INCALC = "Séminaire fictif ABC"
I_ZZZ = "Cours fictif ZZZ"
I_INCONNU = "Réunion fictive sans correspondance"

ECOLES = [
    {"École": "École Fictive A", "Code facturation": "ABC"},
    {"École": "Institut Fictif B", "Code facturation": "ECO"},
    {"École": "Atelier Fictif G", "Code facturation": "GHI"},
    {"École": "Organisme Fictif D", "Code facturation": "DEF"},
    {"École": "Structure Fictive H", "Code facturation": "HFX"},
]


def _aud():
    """Module testé, importé dans le test : son absence fait échouer le test, pas la collecte."""
    return importlib.import_module("facturation.audit")


# ──────────────────────────────────────────────────────────────────────
# Données fictives
# ──────────────────────────────────────────────────────────────────────


def _slug(texte: str) -> str:
    return re.sub(r"[^0-9A-Za-z]+", "-", texte).strip("-").lower()


def _seance(jour: date, debut: str, fin: str, intitule: str, *, uid: str | None = None,
            description: str = "") -> Seance:
    """Séance fictive de l'agenda, heures locales de Paris."""
    d = datetime.combine(jour, time.fromisoformat(debut), tzinfo=PARIS)
    f = datetime.combine(jour, time.fromisoformat(fin), tzinfo=PARIS)
    return Seance(
        uid=uid or f"{jour:%Y%m%d}-{debut.replace(':', '')}-{_slug(intitule)}@exemple.invalid",
        date=jour, debut=d, fin=f, duree_h=(f.timestamp() - d.timestamp()) / 3600,
        intitule=intitule, description=description,
    )


def _regle(ecole: str, statut: str = "Lot 1", *, bc: bool = False) -> dict:
    """Ligne fictive de `Règles_facturation`, toutes colonnes présentes."""
    ligne = dict.fromkeys(ONGLETS_PREVUS["Règles_facturation"])
    ligne.update({"École": ecole, "Statut lot": statut, "Bon de commande requis": "oui" if bc else "non"})
    return ligne


def _corresp(cle: str, ecole: str, module: str, promotion: str, *, mode: str | None = None,
             valeur: float | None = None, type_cle: str = "exact") -> dict:
    """Ligne fictive de `Correspondances`, toutes colonnes présentes."""
    ligne = dict.fromkeys(ONGLETS_PREVUS["Correspondances"])
    ligne.update({
        "Clé": cle, "Type de clé": type_cle, "École": ecole, "Module": module, "Promotion": promotion,
        "Clé tarif": f"{ecole}|{module}|{promotion}", "Mode heures": mode, "Valeur": valeur,
    })
    return ligne


def _facture(n: str, statut: str, ecole: str = "ABC") -> dict:
    ligne = dict.fromkeys(ONGLETS_PREVUS["Factures"])
    ligne.update({"N° facture": n, "École": ecole, "Statut": statut})
    return ligne


def _ses(ecole: str, jour: date | str | None, debut: str | None, fin: str | None, duree: float | None, *,
         n: str = "R_ABC001", promotion: str = "M1", module: str = "Module A", summary: str = "",
         periode: str = "Septembre 2026") -> dict:
    """Ligne fictive de `Sessions` (heures en texte « HH:MM », comme les nouvelles lignes, § 15)."""
    return {
        "Date_log": "2026-11-02", "École": ecole, "Période_facture": periode, "N°_Facture": n,
        "Date_séance": jour, "Début": debut, "Fin": fin, "Durée_h": duree, "Promotion": promotion,
        "Module / UE": module, "Taux": 45, "Total HT": None, "Summary calendrier verbatim": summary,
    }


REGLES = [
    _regle("ABC"), _regle("ECO"), _regle("GHI", "Plus tard"), _regle("DEF", bc=True),
    _regle("HFX", "Hors facturation"),
]
CORRESPONDANCES = [
    _corresp(I_ABC, "ABC", "Module A", "M1"),
    _corresp(I_ECO, "ECO", "Module B", "M2"),
    _corresp(I_GHI, "GHI", "Atelier G", "L3"),
    _corresp(I_DEF, "DEF", "Module D", "M1"),
    _corresp(I_HFX, "HFX", "Module H", "M1"),
    _corresp(I_COEF, "ABC", "Journée A", "M1", mode="coefficient", valeur=0.875),
    _corresp(I_FORFAIT, "ABC", "Jury A", "M1", mode="forfait", valeur=7.5),
    _corresp(I_INCALC, "ABC", "Séminaire A", "M1", mode="coefficient", valeur=None),
    _corresp(I_ZZZ, "ZZZ", "Module Z", "M1"),
]


def _master(sessions=(), *, correspondances=None, regles=None, ecoles=None, factures=()) -> Master:
    """MASTER fictif en mémoire ; la i-ème ligne de `Sessions` porte le rang Excel i + 2."""
    return Master(
        sessions=[dict(ligne) for ligne in sessions],
        correspondances=list(CORRESPONDANCES if correspondances is None else correspondances),
        regles_facturation=list(REGLES if regles is None else regles),
        ecoles=list(ECOLES if ecoles is None else ecoles),
        factures=list(factures),
    )


def _auditer(seances, master, *, du=DU, au=AU, maintenant=MAINTENANT, **options):
    """`auditer` sur l'entité ROSARI ; contrôle au passage que chaque nature produite est connue."""
    aud = _aud()
    rapport = aud.auditer(entite="ROSARI", du=du, au=au, seances=list(seances), master=master,
                          maintenant=maintenant, **options)
    assert tuple(aud.NATURES_ECART[:len(NATURES_CONTRAT)]) == NATURES_CONTRAT
    for ecart in rapport.ecarts:
        assert ecart.nature in aud.NATURES_ECART, ecart
    return rapport


def _de_nature(rapport, nature: str) -> list:
    return [e for e in rapport.ecarts if e.nature == nature]


def _cles(ecarts) -> list[tuple[str, date]]:
    return sorted((e.ecole or "", e.date) for e in ecarts)


def _texte_ecart(ecart) -> str:
    return f"{ecart.nature} {ecart.detail}"


def _classee(seance: Seance, ecole: str = "ABC", module: str = "Module A", promotion: str = "M1") -> SeanceClassee:
    return SeanceClassee(
        seance=seance, ecole=ecole, module=module, promotion=promotion, cle_tarif=f"{ecole}|{module}|{promotion}",
        correspondance=_corresp(seance.intitule, ecole, module, promotion), par="exact",
    )


def _ls(rang: int, ecole: str, jour: date, debut: str | None, fin: str | None, duree: float | None, *,
        n: str = "R_ABC001", promotion: str = "M1", module: str = "Module A", summary: str = "",
        abandonnee: bool = False) -> LigneSessions:
    return LigneSessions(
        rang=rang, ecole=ecole, n_facture=n, periode="Septembre 2026", date_seance=jour,
        debut=time.fromisoformat(debut) if debut else None, fin=time.fromisoformat(fin) if fin else None,
        duree_h=duree, promotion=promotion, module=module, taux=45.0, total_ht=None, summary=summary,
        abandonnee=abandonnee,
    )


def _paires(rapprochement) -> set[tuple[str, tuple[int, ...], str]]:
    """Appariements sous forme comparable : (clé de séance, rangs des lignes, niveau)."""
    return {
        (cle_seance(a.seance.seance), tuple(sorted(ligne.rang for ligne in a.lignes)), a.niveau)
        for a in rapprochement.appariements
    }


def _brouillon_fictif(identifiant: str, seances, *, ecole: str = "ABC", mois: str = "2026-10"):
    """Brouillon non émis de F-16a (pas encore développé), réduit à ce qu'un audit lit : son
    identifiant et ses lignes (clé de séance, uid, dates)."""
    lignes = tuple(
        SimpleNamespace(
            cle_seance=cle_seance(s), uid=s.uid, date=s.date, debut=s.debut, fin=s.fin, intitule=s.intitule,
            ecole=ecole, promotion="M1", module="Module A", libelle="Module A", cle_tarif=f"{ecole}|Module A|M1",
            classee_par="exact", duree_agenda_h=s.duree_h, heures=s.duree_h, taux=45.0, montant=None,
            ref_contrat=None, tenue=True,
        )
        for s in seances
    )
    return SimpleNamespace(
        identifiant=identifiant, entite="ROSARI", ecole=ecole, mois=mois, lignes=lignes, statut="prêt",
        blocages=(), avertissements=(), empreinte="0" * 64,
    )


# ──────────────────────────────────────────────────────────────────────
# Contrat des types
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    ("classe", "champs"),
    [
        ("Appariement", ("seance", "lignes", "niveau")),
        ("Rapprochement", ("appariements", "agenda_seul", "sessions_seules", "raisons")),
        ("Ecart", ("nature", "ecole", "date", "seance", "lignes", "heures_agenda", "heures_facturees", "detail",
                   "dans_brouillon")),
        ("RapportAudit", ("entite", "du", "au", "etabli_le", "ecoles", "nb_seances_tenues", "nb_rapprochees",
                          "ecarts", "a_venir", "hors_facturation", "lignes_non_datees", "lignes_illisibles")),
    ],
)
def test_t_aud_14_types_figes_du_contrat(classe, champs):
    """T-AUD-14 — `Appariement`, `Rapprochement`, `Ecart` et `RapportAudit` sont des dataclasses figées
    qui portent, dans l'ordre, les champs du contrat ; `NATURES_ECART` commence par ses sept natures."""
    aud = _aud()
    cls = getattr(aud, classe)
    assert dataclasses.is_dataclass(cls) and cls.__dataclass_params__.frozen
    assert tuple(f.name for f in dataclasses.fields(cls))[:len(champs)] == champs
    assert tuple(aud.NATURES_ECART[:7]) == NATURES_CONTRAT
    assert callable(aud.apparier) and callable(aud.auditer)
    assert callable(aud.rapport_texte) and callable(aud.ecrire_rapport) and callable(aud.lancer_audit)


# ──────────────────────────────────────────────────────────────────────
# Appariement
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    ("seance", "ligne", "niveau"),
    [
        (("09:00", "12:30", I_ABC), dict(debut="09:00", fin="12:30", summary=I_ABC), "exact"),
        (("09:00", "12:00", I_ABC), dict(debut="09:00", fin="12:30", summary="Autre", promotion="M2",
                                         module="Module Z"), "début"),
        (("14:00", "17:00", "Cours Fictif  ABC  M1"), dict(debut="09:00", fin="12:00", summary="  cours fictif abc m1 ",
                                                          promotion="M2", module="Module Z"), "intitulé"),
        (("14:00", "17:00", I_ABC), dict(debut="10:00", fin="13:00", summary="Autre libellé"), "cours"),
    ],
    ids=["exact", "debut", "intitule", "cours"],
)
def test_t_aud_15_apparier_par_niveau(seance, ligne, niveau):
    """T-AUD-15 — `apparier` rapproche une séance et une ligne de la même école, le même jour, au premier
    niveau qui convient : même début et même fin, même début, même intitulé (casse et espaces
    normalisés), même (promotion, module)."""
    jour = date(2026, 9, 15)
    classee = _classee(_seance(jour, *seance))
    options = {"promotion": "M1", "module": "Module A", **ligne}
    debut, fin = options.pop("debut"), options.pop("fin")
    lignes = [_ls(2, "ABC", jour, debut, fin, 3.0, **options)]
    rapprochement = _aud().apparier([classee], lignes)
    assert _paires(rapprochement) == {(cle_seance(classee.seance), (2,), niveau)}
    assert list(rapprochement.agenda_seul) == [] and list(rapprochement.sessions_seules) == []


def test_t_aud_16_apparier_niveau_par_niveau_sur_la_journee():
    """T-AUD-16 — Les niveaux s'appliquent dans l'ordre à toute la journée de l'école : une ligne exacte
    pour la séance B n'est jamais prise par la séance A au niveau de l'intitulé."""
    jour = date(2026, 9, 15)
    a = _classee(_seance(jour, "09:00", "12:00", "Cours A"))
    b = _classee(_seance(jour, "14:00", "17:00", "Cours B"))
    lignes = [
        _ls(2, "ABC", jour, "14:00", "17:00", 3.0, summary="Cours A", promotion="", module=""),
        _ls(3, "ABC", jour, "10:00", "13:00", 3.0, summary="Cours A", promotion="", module=""),
    ]
    rapprochement = _aud().apparier([a, b], lignes)
    assert _paires(rapprochement) == {
        (cle_seance(b.seance), (2,), "exact"),
        (cle_seance(a.seance), (3,), "intitulé"),
    }


def test_t_aud_17_apparier_lignes_sans_heure_un_pour_un():
    """T-AUD-17 — Lignes sans heure : un pour un, par ordre chronologique, quand il reste autant de
    séances que de lignes ce jour-là ; une séance déjà appariée exactement n'est pas recomptée."""
    jour = date(2026, 10, 15)
    matin = _classee(_seance(jour, "09:00", "12:00", I_ABC))
    apres_midi = _classee(_seance(jour, "14:00", "17:00", I_ABC))
    sans_heure = [
        _ls(3, "ABC", jour, None, None, 3.0, promotion="", module=""),
        _ls(5, "ABC", jour, None, None, 3.0, promotion="", module=""),
    ]
    rapprochement = _aud().apparier([apres_midi, matin], sans_heure)
    assert _paires(rapprochement) == {
        (cle_seance(matin.seance), (3,), "jour sans heure"),
        (cle_seance(apres_midi.seance), (5,), "jour sans heure"),
    }

    exacte = _ls(2, "ABC", jour, "09:00", "12:00", 3.0)
    rapprochement = _aud().apparier([matin, apres_midi], [exacte, sans_heure[0]])
    assert _paires(rapprochement) == {
        (cle_seance(matin.seance), (2,), "exact"),
        (cle_seance(apres_midi.seance), (3,), "jour sans heure"),
    }


def test_t_aud_18_apparier_sans_heure_nombres_differents():
    """T-AUD-18 — Deux séances pour une ligne sans heure le même jour : aucune n'est appariée, chacune
    porte la raison « 2 séance(s) pour 1 ligne(s) sans heure le 15/10 », la ligne reste seule."""
    jour = date(2026, 10, 15)
    matin = _classee(_seance(jour, "09:00", "12:00", I_ABC))
    apres_midi = _classee(_seance(jour, "14:00", "17:00", I_ABC))
    ligne = _ls(4, "ABC", jour, None, None, 6.0, promotion="", module="")
    rapprochement = _aud().apparier([matin, apres_midi], [ligne])
    assert list(rapprochement.appariements) == []
    assert set(rapprochement.agenda_seul) == {matin, apres_midi}
    assert list(rapprochement.sessions_seules) == [ligne]
    raison = "2 séance(s) pour 1 ligne(s) sans heure le 15/10"
    assert rapprochement.raisons == {cle_seance(matin.seance): raison, cle_seance(apres_midi.seance): raison}


def test_t_aud_19_apparier_jamais_une_autre_ecole_ni_un_autre_jour():
    """T-AUD-19 — Une ligne d'une autre école, ou d'un autre jour, n'est jamais appariée, même à la minute
    près (garde cassée : la ligne de la bonne école et du bon jour est appariée)."""
    jour = date(2026, 9, 15)
    classee = _classee(_seance(jour, "09:00", "12:30", I_ABC))
    autre_ecole = _ls(2, "ECO", jour, "09:00", "12:30", 3.5, summary=I_ABC)
    autre_jour = _ls(3, "ABC", date(2026, 9, 16), "09:00", "12:30", 3.5, summary=I_ABC)
    rapprochement = _aud().apparier([classee], [autre_ecole, autre_jour])
    assert list(rapprochement.appariements) == []
    assert list(rapprochement.agenda_seul) == [classee]
    assert set(rapprochement.sessions_seules) == {autre_ecole, autre_jour}

    bonne = _ls(4, "ABC", jour, "09:00", "12:30", 3.5)
    rapprochement = _aud().apparier([classee], [autre_ecole, autre_jour, bonne])
    assert _paires(rapprochement) == {(cle_seance(classee.seance), (4,), "exact")}


def test_t_aud_20_apparier_deterministe():
    """T-AUD-20 — Le résultat ne dépend pas de l'ordre des entrées, y compris quand deux séances
    pourraient prendre la même ligne, ou une séance deux lignes, au même niveau (le 18/09)."""
    j1, j2, j4 = date(2026, 9, 15), date(2026, 9, 16), date(2026, 9, 18)
    seances = [
        _classee(_seance(j1, "09:00", "12:00", "Cours A")),
        _classee(_seance(j1, "14:00", "17:00", "Cours B")),
        _classee(_seance(j2, "09:00", "12:00", I_ECO), "ECO", "Module B", "M2"),
        _classee(_seance(j2, "14:00", "17:00", I_ABC)),
        _classee(_seance(j4, "09:00", "10:00", "Cours C")),
        _classee(_seance(j4, "14:00", "15:00", "Cours C")),
        _classee(_seance(j4, "16:00", "17:00", "Cours D")),
    ]
    lignes = [
        _ls(2, "ABC", j1, "14:00", "17:00", 3.0, summary="Cours A", promotion="", module=""),
        _ls(3, "ABC", j1, "10:00", "13:00", 3.0, summary="Cours A", promotion="", module=""),
        _ls(4, "ECO", j2, "09:00", "12:00", 3.0, promotion="M2", module="Module B"),
        _ls(5, "ABC", j2, None, None, 3.0, promotion="", module=""),
        _ls(6, "ABC", date(2026, 9, 17), "09:00", "12:00", 3.0),
        _ls(7, "ABC", j4, "11:00", "12:00", 1.0, summary="Cours C", promotion="", module=""),
        _ls(8, "ABC", j4, "10:30", "11:30", 1.0, summary="Cours D", promotion="", module=""),
        _ls(9, "ABC", j4, "18:00", "19:00", 1.0, summary="Cours D", promotion="", module=""),
    ]
    aud = _aud()
    reference = aud.apparier(seances, lignes)
    hasard = random.Random(19)
    for _ in range(5):
        s, lg = seances[:], lignes[:]
        hasard.shuffle(s)
        hasard.shuffle(lg)
        autre = aud.apparier(s, lg)
        assert _paires(autre) == _paires(reference)
        assert set(autre.agenda_seul) == set(reference.agenda_seul)
        assert set(autre.sessions_seules) == set(reference.sessions_seules)
    assert len(reference.appariements) == 6 and len(reference.agenda_seul) == 1
    assert {ligne.rang for ligne in reference.sessions_seules} in ({6, 8}, {6, 9})


# ──────────────────────────────────────────────────────────────────────
# Audit : séances non facturées et rapprochement
# ──────────────────────────────────────────────────────────────────────


def test_t_aud_01_deux_mois_trois_ecoles_seances_non_facturees_exactes():
    """T-AUD-01 — Période fictive de deux mois et trois écoles (dont une `Plus tard`) : le rapport liste
    exactement les séances tenues non facturées attendues, et rien d'autre ; les séances hors de la
    période ne comptent pas."""
    seances = [
        _seance(date(2026, 9, 15), "09:00", "12:30", I_ABC),
        _seance(date(2026, 9, 22), "09:00", "12:30", I_ABC),
        _seance(date(2026, 10, 13), "09:00", "12:30", I_ABC),
        _seance(date(2026, 10, 20), "09:00", "12:30", I_ABC),
        _seance(date(2026, 9, 16), "14:00", "17:00", I_ECO),
        _seance(date(2026, 10, 14), "14:00", "17:00", I_ECO),
        _seance(date(2026, 9, 30), "10:00", "12:00", I_GHI),
        _seance(date(2026, 8, 31), "09:00", "12:30", I_ABC),   # avant la période
        _seance(date(2026, 11, 3), "09:00", "12:30", I_ABC),   # après la période, tenue
    ]
    sessions = [
        _ses("ABC", date(2026, 9, 15), "09:00", "12:30", 3.5),
        _ses("ABC", date(2026, 10, 13), "09:00", "12:30", 3.5, n="R_ABC002", periode="Octobre 2026"),
        _ses("ECO", date(2026, 9, 16), "14:00", "17:00", 3.0, n="R_ECO001", promotion="M2", module="Module B"),
    ]
    rapport = _auditer(seances, _master(sessions))
    attendues = [
        ("ABC", date(2026, 9, 22)), ("ABC", date(2026, 10, 20)), ("ECO", date(2026, 10, 14)),
        ("GHI", date(2026, 9, 30)),
    ]
    non_facturees = _de_nature(rapport, TENUE_NON_FACTUREE)
    assert _cles(non_facturees) == sorted(attendues)
    assert [e for e in rapport.ecarts if e.nature != TENUE_NON_FACTUREE] == []
    for ecart in non_facturees:
        assert ecart.seance is not None and ecart.seance.date == ecart.date
        assert tuple(ecart.lignes) == ()
    assert rapport.nb_seances_tenues == 7
    assert rapport.nb_rapprochees == 3
    assert (rapport.entite, rapport.du, rapport.au, rapport.etabli_le) == ("ROSARI", DU, AU, MAINTENANT)
    assert rapport.code_retour == 3


SEANCES_PILOTE = [
    _seance(date(2026, 9, 15), "09:00", "12:30", I_ABC),
    _seance(date(2026, 9, 22), "09:00", "12:30", I_ABC),
    _seance(date(2026, 10, 13), "09:00", "12:30", I_ABC),
    _seance(date(2026, 10, 20), "09:00", "12:30", I_ABC),
]
SESSIONS_PILOTE = [_ses("ABC", s.date, "09:00", "12:30", 3.5, n=f"R_ABC00{s.date.month - 8}") for s in SEANCES_PILOTE]


def test_t_aud_02_ecole_pilote_entierement_journalisee():
    """T-AUD-02 — École pilote fictive entièrement journalisée : aucun écart, code 0. Retirer une ligne de
    `Sessions` produit exactement un écart, « tenue non facturée », pour sa séance (garde cassée)."""
    rapport = _auditer(SEANCES_PILOTE, _master(SESSIONS_PILOTE))
    assert list(rapport.ecarts) == []
    assert rapport.code_retour == 0 and rapport.sans_ecart("ABC")
    assert rapport.nb_seances_tenues == 4 and rapport.nb_rapprochees == 4

    for retiree, seance in enumerate(SEANCES_PILOTE):
        sessions = SESSIONS_PILOTE[:retiree] + SESSIONS_PILOTE[retiree + 1:]
        rapport = _auditer(SEANCES_PILOTE, _master(sessions))
        assert len(rapport.ecarts) == 1, rapport.ecarts
        (ecart,) = rapport.ecarts
        assert (ecart.nature, ecart.ecole, ecart.date, ecart.seance) == (TENUE_NON_FACTUREE, "ABC", seance.date, seance)
        assert rapport.code_retour == 3 and not rapport.sans_ecart("ABC")
        assert rapport.nb_rapprochees == 3


@pytest.mark.parametrize("second_numero", ["R_ABC002", "R_ABC001"], ids=["deux-numeros", "meme-numero"])
def test_t_aud_03_seance_journalisee_deux_fois_doublon(second_numero):
    """T-AUD-03 — Même séance journalisée sous deux N° (ou deux fois sous le même) : un seul écart
    « doublon », qui porte les deux lignes ; ni « tenue non facturée », ni « facturée absente ». Deux
    lignes identiques sans séance dans l'agenda sont aussi un doublon."""
    jour = date(2026, 9, 15)
    seance = _seance(jour, "09:00", "12:30", I_ABC)
    sessions = [
        _ses("ABC", jour, "09:00", "12:30", 3.5, n="R_ABC001"),
        _ses("ABC", jour, "09:00", "12:30", 3.5, n=second_numero),
    ]
    rapport = _auditer([seance], _master(sessions))
    doublons = _de_nature(rapport, DOUBLON)
    assert len(doublons) == 1, rapport.ecarts
    (doublon,) = doublons
    assert (doublon.ecole, doublon.date, doublon.seance) == ("ABC", jour, seance)
    assert sorted(ligne.rang for ligne in doublon.lignes) == [2, 3]
    assert _de_nature(rapport, TENUE_NON_FACTUREE) == [] and _de_nature(rapport, ABSENTE) == []

    sans_seance = _auditer([], _master(sessions))
    assert any(sorted(ligne.rang for ligne in e.lignes) == [2, 3] for e in _de_nature(sans_seance, DOUBLON))

    # Garde cassée : une seule ligne, aucun doublon.
    assert _de_nature(_auditer([seance], _master(sessions[:1])), DOUBLON) == []


def test_t_aud_04_seance_a_venir_jamais_non_facturee():
    """T-AUD-04 — Une séance à venir (en cours comprise) va dans `a_venir` et n'est jamais « non
    facturée » ; une séance à venir déjà journalisée (facture produite en avance) ne fait pas de sa
    ligne une « facturée absente » ; une séance qui finit à l'instant est tenue."""
    maintenant = datetime(2026, 10, 20, 12, 0, tzinfo=PARIS)
    du, au = date(2026, 10, 1), date(2026, 10, 31)
    journalisee = _seance(date(2026, 10, 13), "09:00", "12:30", I_ABC)
    finie_a_l_instant = _seance(date(2026, 10, 20), "09:00", "12:00", I_ABC)
    en_cours = _seance(date(2026, 10, 20), "11:00", "13:00", I_ECO)
    a_venir = _seance(date(2026, 10, 27), "09:00", "12:30", I_ABC)
    a_venir_deja_facturee = _seance(date(2026, 10, 29), "09:00", "12:30", I_ABC)
    sessions = [
        _ses("ABC", date(2026, 10, 13), "09:00", "12:30", 3.5, periode="Octobre 2026"),
        _ses("ABC", date(2026, 10, 29), "09:00", "12:30", 3.5, periode="Octobre 2026"),
    ]
    seances = [journalisee, finie_a_l_instant, en_cours, a_venir, a_venir_deja_facturee]
    rapport = _auditer(seances, _master(sessions), du=du, au=au, maintenant=maintenant)
    assert [(e.nature, e.seance) for e in rapport.ecarts] == [(TENUE_NON_FACTUREE, finie_a_l_instant)]
    venir = {classee.seance for classee in rapport.a_venir}
    assert {en_cours, a_venir, a_venir_deja_facturee} <= venir
    assert finie_a_l_instant not in venir and journalisee not in venir
    assert rapport.nb_seances_tenues == 2

    # Garde cassée : huit jours plus tard, la séance du 27/10 est tenue et non facturée.
    plus_tard = _auditer(seances, _master(sessions), du=du, au=au,
                         maintenant=datetime(2026, 10, 28, 12, 0, tzinfo=PARIS))
    assert a_venir in {e.seance for e in _de_nature(plus_tard, TENUE_NON_FACTUREE)}


def test_t_aud_05_creneau_non_reconnu_et_tenu():
    """T-AUD-05 — Créneau non reconnu (ou de code inconnu) et tenu : « tenue non reconnue » seulement,
    école None ; jamais « tenue non facturée ». Un créneau à classer à venir ne donne rien."""
    maintenant = datetime(2026, 10, 20, 12, 0, tzinfo=PARIS)
    non_reconnu = _seance(date(2026, 9, 15), "09:00", "11:00", I_INCONNU)
    code_inconnu = _seance(date(2026, 9, 22), "09:00", "12:30", I_ABC, description="Salle fictive #ZZQ-M1-STATS")
    futur = _seance(date(2026, 10, 27), "09:00", "10:00", "Réunion fictive à venir")
    seances = [non_reconnu, code_inconnu, futur]
    rapport = _auditer(seances, _master([]), maintenant=maintenant)
    non_reconnues = _de_nature(rapport, NON_RECONNUE)
    assert sorted(e.seance.date for e in non_reconnues) == [date(2026, 9, 15), date(2026, 9, 22)]
    assert all(e.ecole is None for e in non_reconnues)
    assert [e for e in rapport.ecarts if e.nature != NON_RECONNUE] == []

    # Garde cassée : reconnu par une ligne de Correspondances, le créneau devient « tenue non facturée ».
    classe = [*CORRESPONDANCES, _corresp(I_INCONNU, "ABC", "Module A", "M1")]
    rapport = _auditer(seances, _master([], correspondances=classe), maintenant=maintenant)
    assert [e.seance for e in _de_nature(rapport, TENUE_NON_FACTUREE)] == [non_reconnu]


def test_t_aud_06_ecole_hors_facturation_ignoree():
    """T-AUD-06 — École `Hors facturation` : ses séances et ses lignes `Sessions` ne donnent aucun écart ;
    ses séances sont listées dans `hors_facturation`, et l'école n'est pas auditée."""
    hfx_seule = _seance(date(2026, 9, 15), "09:00", "12:00", I_HFX)
    hfx_journalisee = _seance(date(2026, 9, 22), "09:00", "12:00", I_HFX)
    abc = _seance(date(2026, 9, 15), "14:00", "17:30", I_ABC)
    sessions = [
        _ses("HFX", date(2026, 9, 22), "09:00", "12:00", 3.0, n="R_HFX001", module="Module H"),
        _ses("HFX", date(2026, 9, 29), "09:00", "12:00", 3.0, n="R_HFX001", module="Module H"),
    ]
    seances = [hfx_seule, hfx_journalisee, abc]
    rapport = _auditer(seances, _master(sessions))
    assert _cles(rapport.ecarts) == [("ABC", date(2026, 9, 15))]
    assert {classee.seance for classee in rapport.hors_facturation} >= {hfx_seule, hfx_journalisee}
    assert "HFX" not in rapport.ecoles

    # Garde cassée : la même école au Lot 1 est auditée.
    regles = [r for r in REGLES if r["École"] != "HFX"] + [_regle("HFX")]
    rapport = _auditer(seances, _master(sessions, regles=regles))
    assert ("HFX", date(2026, 9, 15)) in _cles(_de_nature(rapport, TENUE_NON_FACTUREE))
    assert ("HFX", date(2026, 9, 29)) in _cles(_de_nature(rapport, ABSENTE))


def test_t_aud_07_ligne_sessions_mal_formee_signalee_avec_son_rang(tmp_path):
    """T-AUD-07 — Ligne de `Sessions` mal formée (date ou heure illisible) : signalée avec son rang Excel
    (ligne vide comprise dans le compte) et la colonne fautive, sans plantage ; le texte du rapport la
    cite."""
    chemin = creer_master(
        tmp_path / "donnees", regles=REGLES, correspondances=CORRESPONDANCES, ecoles=ECOLES,
        sessions=[
            _ses("ABC", date(2026, 9, 15), "09:00", "12:30", 3.5),
            None,
            _ses("ABC", "fin sept", "09:00", "12:30", 3.5),
            _ses("ABC", date(2026, 9, 29), "midi", "12:30", 3.5),
        ],
    )
    master = lire_master(chemin)
    seances = [_seance(date(2026, 9, d), "09:00", "12:30", I_ABC) for d in (15, 22, 29)]
    rapport = _auditer(seances, master)
    illisibles = dict(rapport.lignes_illisibles)
    assert set(illisibles) == {4, 5}
    assert "Date_séance" in illisibles[4] and "fin sept" in illisibles[4]
    assert "Début" in illisibles[5] and "midi" in illisibles[5]
    texte = _aud().rapport_texte(rapport)
    assert "fin sept" in texte and "midi" in texte


@pytest.mark.parametrize(
    ("intitule", "debut", "fin", "module", "facturables"),
    [(I_COEF, "09:00", "17:00", "Journée A", 7.0), (I_FORFAIT, "09:00", "16:00", "Jury A", 7.5)],
    ids=["coefficient-7-8", "forfait"],
)
def test_t_aud_08_coefficient_et_forfait_sans_ecart(intitule, debut, fin, module, facturables):
    """T-AUD-08 — Coefficient 7/8 (9 h–17 h, `Durée_h` = 7) ou forfait (7 h facturées 7,5 h) : aucun écart.
    Garde cassée : une `Durée_h` égale à la durée de l'agenda donne un « écart d'heures » entre les heures
    facturables recalculées et la `Durée_h`."""
    jour = date(2026, 9, 15)
    seance = _seance(jour, debut, fin, intitule)
    ligne = _ses("ABC", jour, debut, fin, facturables, module=module)
    rapport = _auditer([seance], _master([ligne]))
    assert list(rapport.ecarts) == []

    brute = seance.duree_h
    rapport = _auditer([seance], _master([{**ligne, "Durée_h": brute}]))
    assert len(rapport.ecarts) == 1, rapport.ecarts
    (ecart,) = rapport.ecarts
    assert (ecart.nature, ecart.seance, ecart.ecole, ecart.date) == (ECART_HEURES, seance, "ABC", jour)
    assert (ecart.heures_agenda, ecart.heures_facturees) == (facturables, brute)
    assert [ligne.rang for ligne in ecart.lignes] == [2]


def test_t_aud_09_ecole_historique_en_toutes_lettres():
    """T-AUD-09 — `École` historique écrite en toutes lettres dans `Sessions` : reliée à son code par
    l'onglet `Écoles`, sans écart. Garde cassée : sans la ligne de `Écoles`, la ligne est illisible et la
    séance redevient non facturée."""
    jour = date(2026, 9, 15)
    seance = _seance(jour, "09:00", "12:30", I_ABC)
    ligne = _ses(" école fictive a ", jour, "09:00", "12:30", 3.5, n="R_ABC001")
    rapport = _auditer([seance], _master([ligne]))
    assert list(rapport.ecarts) == [] and list(rapport.lignes_illisibles) == []
    assert rapport.nb_rapprochees == 1

    ecoles = [e for e in ECOLES if e["Code facturation"] != "ABC"]
    rapport = _auditer([seance], _master([ligne], ecoles=ecoles))
    assert [rang for rang, _ in rapport.lignes_illisibles] == [2]
    assert _cles(_de_nature(rapport, TENUE_NON_FACTUREE)) == [("ABC", jour)]


def test_t_aud_32_seance_dans_un_brouillon_non_emis_annotee():
    """T-AUD-32 — (T-AUD-10 au niveau d'`auditer`) Séance tenue non facturée présente dans un brouillon
    non émis : l'écart porte l'identifiant du brouillon (`dans_brouillon`), que le texte du rapport
    cite ; une séance absente des brouillons n'est pas annotée (garde cassée : sans brouillon, aucune
    annotation)."""
    dans = _seance(date(2026, 10, 13), "09:00", "12:30", I_ABC)
    hors = _seance(date(2026, 10, 20), "09:00", "12:30", I_ABC)
    brouillon = _brouillon_fictif("ABC-2026-10", [dans])
    rapport = _auditer([dans, hors], _master([]), brouillons=[brouillon])
    par_seance = {e.seance: e for e in _de_nature(rapport, TENUE_NON_FACTUREE)}
    assert set(par_seance) == {dans, hors}
    assert par_seance[dans].dans_brouillon == "ABC-2026-10"
    assert par_seance[hors].dans_brouillon is None
    assert "ABC-2026-10" in _aud().rapport_texte(rapport)

    rapport = _auditer([dans, hors], _master([]))
    assert [e.dans_brouillon for e in rapport.ecarts] == [None, None]


def test_t_aud_12_lignes_d_un_numero_abandonne_ignorees():
    """T-AUD-12 — Les lignes d'un N° abandonné (statut de `Factures`) sont ignorées : la séance qu'elles
    seules journalisent est non facturée, une ligne seule ne donne pas « facturée absente », et elles ne
    font pas de doublon. Garde cassée : sans le statut, les mêmes lignes comptent."""
    s1 = _seance(date(2026, 9, 15), "09:00", "12:30", I_ABC)
    s2 = _seance(date(2026, 9, 29), "09:00", "12:30", I_ABC)
    sessions = [
        _ses("ABC", date(2026, 9, 15), "09:00", "12:30", 3.5, n="R_ABC009"),
        _ses("ABC", date(2026, 9, 22), "09:00", "12:30", 3.5, n="R_ABC009"),
        _ses("ABC", date(2026, 9, 29), "09:00", "12:30", 3.5, n="R_ABC009"),
        _ses("ABC", date(2026, 9, 29), "09:00", "12:30", 3.5, n="R_ABC010"),
    ]
    abandon = [_facture("R_ABC009", "abandonnée")]
    rapport = _auditer([s1, s2], _master(sessions, factures=abandon))
    assert [(e.nature, e.seance) for e in rapport.ecarts] == [(TENUE_NON_FACTUREE, s1)]

    rapport = _auditer([s1, s2], _master(sessions, factures=[_facture("R_ABC009", "émise")]))
    assert _de_nature(rapport, TENUE_NON_FACTUREE) == []
    assert _cles(_de_nature(rapport, ABSENTE)) == [("ABC", date(2026, 9, 22))]
    assert [e.seance for e in _de_nature(rapport, DOUBLON)] == [s2]


@pytest.mark.parametrize(
    ("seance", "ligne"),
    [
        (("09:00", "12:00", I_ABC), dict(debut="09:00", fin="12:30", duree=3.5, summary=I_ABC)),
        (("14:00", "17:00", I_ABC), dict(debut="09:00", fin="12:30", duree=3.5, summary="Autre cours fictif",
                                         promotion="M2", module="Module Z")),
        (("09:00", "12:00", I_ABC), dict(debut="09:30", fin="12:30", duree=3.0, summary=I_ABC)),
    ],
    ids=["meme-debut", "autre-debut", "memes-heures-facturees"],
)
def test_t_aud_13_meme_jour_autres_heures_a_verifier(seance, ligne):
    """T-AUD-13 — 💡 [proposé — à valider], constat écarté de la revue de P-0b : une séance tenue et une
    ligne `Sessions` de la même école, le même jour, à d'autres heures (09:00–12:00 contre 09:00–12:30)
    sortent en écart « à vérifier : même jour, autres heures » qui nomme la ligne de `Sessions` ; jamais
    une « tenue non facturée » qui ne le dit pas. Cela vaut que la paire soit rapprochée (même début,
    même intitulé) ou non, et même quand la `Durée_h` égale les heures facturables (09:00–12:00 contre
    09:30–12:30, 3 h des deux côtés). Garde cassée : la ligne un autre jour ne le dit plus."""
    jour = date(2026, 9, 15)
    s = _seance(jour, *seance)
    options = dict(ligne)
    lg = _ses("ABC", jour, options.pop("debut"), options.pop("fin"), options.pop("duree"), **options)
    aud = _aud()
    rapport = _auditer([s], _master([lg]))
    touches = [e for e in rapport.ecarts if e.seance == s or 2 in {x.rang for x in e.lignes}]
    assert touches, rapport.ecarts
    for ecart in touches:
        assert MEME_JOUR in _texte_ecart(ecart), ecart
        assert re.search(r"\bligne 2\b", ecart.detail), ecart
    assert MEME_JOUR in aud.rapport_texte(rapport)

    ailleurs = _auditer([s], _master([{**lg, "Date_séance": date(2026, 9, 16)}]))
    assert [e for e in ailleurs.ecarts if MEME_JOUR in _texte_ecart(e)] == []
    assert MEME_JOUR not in aud.rapport_texte(ailleurs)
    assert [e.seance for e in _de_nature(ailleurs, TENUE_NON_FACTUREE)] == [s]


# ──────────────────────────────────────────────────────────────────────
# Audit : heures, volume, période, filtres
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    ("duree", "ecart"),
    [(3.5, False), (3.504, False), (3.496, False), (3.51, True), (3.49, True), (3.0, True)],
)
def test_t_aud_21_ecart_d_heures_au_dela_d_un_demi_centieme(duree, ecart):
    """T-AUD-21 — « écart d'heures » si |heures facturables recalculées − `Durée_h`| > 0,005, sinon rien."""
    jour = date(2026, 9, 15)
    seance = _seance(jour, "09:00", "12:30", I_ABC)
    rapport = _auditer([seance], _master([_ses("ABC", jour, "09:00", "12:30", duree)]))
    if ecart:
        assert [(e.nature, e.heures_agenda, e.heures_facturees) for e in rapport.ecarts] == [(ECART_HEURES, 3.5, duree)]
    else:
        assert list(rapport.ecarts) == []


def test_t_aud_22_heures_incalculables():
    """T-AUD-22 — Une séance appariée dont `heures_facturables` lève `ValueError` (coefficient sans
    `Valeur`) donne « heures incalculables », qui cite la cause, et pas d'« écart d'heures »."""
    jour = date(2026, 9, 15)
    seance = _seance(jour, "09:00", "17:00", I_INCALC)
    rapport = _auditer([seance], _master([_ses("ABC", jour, "09:00", "17:00", 7.0, module="Séminaire A")]))
    assert len(rapport.ecarts) == 1, rapport.ecarts
    (ecart,) = rapport.ecarts
    assert (ecart.nature, ecart.seance, ecart.ecole) == (INCALCULABLES, seance, "ABC")
    assert "Valeur" in ecart.detail
    assert ecart.heures_agenda is None


def _volume(ecole: str, intitule: str, module: str, durees: tuple[float, float]):
    jours = (date(2026, 10, 6), date(2026, 10, 13))
    seances = [_seance(j, "09:00", "12:00", intitule) for j in jours]
    sessions = [
        _ses(ecole, j, "09:00", "12:00", d, n=f"R_{ecole}001", module=module, periode="Octobre 2026")
        for j, d in zip(jours, durees)
    ]
    return _auditer(seances, _master(sessions))


def test_t_aud_23_ecart_de_volume_pour_une_ecole_a_bon_de_commande():
    """T-AUD-23 — École à bon de commande : Σ heures facturables ≠ Σ `Durée_h` par (école, promotion) sur
    la période → « écart de volume » ; volumes égaux → rien ; une école sans bon de commande n'a jamais
    d'écart de volume (garde cassée)."""
    assert list(_volume("DEF", I_DEF, "Module D", (3, 3)).ecarts) == []

    rapport = _volume("DEF", I_DEF, "Module D", (3, 2.5))
    volumes = _de_nature(rapport, VOLUME)
    assert len(volumes) == 1, rapport.ecarts
    (volume,) = volumes
    assert volume.ecole == "DEF" and "M1" in volume.detail
    assert (volume.heures_agenda, volume.heures_facturees) == (6.0, 5.5)
    assert len(_de_nature(rapport, ECART_HEURES)) == 1

    sans_bc = _volume("ABC", I_ABC, "Module A", (3, 2.5))
    assert _de_nature(sans_bc, VOLUME) == []
    assert len(_de_nature(sans_bc, ECART_HEURES)) == 1


def test_t_aud_24_ligne_sans_date_de_seance():
    """T-AUD-24 — Une ligne sans `Date_séance` va dans `lignes_non_datees` et ne donne pas d'écart.
    Garde cassée : datée, la même ligne est « facturée absente de l'agenda »."""
    jour = date(2026, 9, 15)
    seance = _seance(jour, "09:00", "12:30", I_ABC)
    sessions = [_ses("ABC", jour, "09:00", "12:30", 3.5), _ses("ABC", None, "09:00", "12:30", 3.5)]
    rapport = _auditer([seance], _master(sessions))
    assert [ligne.rang for ligne in rapport.lignes_non_datees] == [3]
    assert list(rapport.ecarts) == []

    datee = [sessions[0], {**sessions[1], "Date_séance": date(2026, 9, 22)}]
    rapport = _auditer([seance], _master(datee))
    assert _cles(rapport.ecarts) == [("ABC", date(2026, 9, 22))]
    assert [e.nature for e in rapport.ecarts] == [ABSENTE]
    (absente,) = rapport.ecarts
    assert absente.seance is None and [ligne.rang for ligne in absente.lignes] == [3]
    assert list(rapport.lignes_non_datees) == []


def test_t_aud_25_periode_bornes_incluses_le_reste_ignore():
    """T-AUD-25 — Seules comptent les séances et les lignes datées du `du` au `au`, bornes incluses."""
    seances = [_seance(date(2026, m, j), "09:00", "12:30", I_ABC) for m, j in ((8, 31), (9, 1), (10, 31), (11, 1))]
    sessions = [
        _ses("ABC", date(2026, 8, 30), "09:00", "12:30", 3.5),
        _ses("ABC", date(2026, 11, 2), "09:00", "12:30", 3.5),
    ]
    rapport = _auditer(seances, _master(sessions))
    assert _cles(rapport.ecarts) == [("ABC", date(2026, 9, 1)), ("ABC", date(2026, 10, 31))]
    assert {e.nature for e in rapport.ecarts} == {TENUE_NON_FACTUREE}


def test_t_aud_26_filtre_par_ecoles():
    """T-AUD-26 — `ecoles` limite l'audit à ces écoles, séances et lignes `Sessions` comprises."""
    seances = [_seance(date(2026, 9, 15), "09:00", "12:30", I_ABC), _seance(date(2026, 9, 16), "14:00", "17:00", I_ECO)]
    sessions = [_ses("ECO", date(2026, 9, 23), "14:00", "17:00", 3.0, n="R_ECO001", promotion="M2", module="Module B")]
    rapport = _auditer(seances, _master(sessions), ecoles=["ABC"])
    assert _cles(rapport.ecarts) == [("ABC", date(2026, 9, 15))]
    assert tuple(rapport.ecoles) == ("ABC",)

    rapport = _auditer(seances, _master(sessions))
    assert {("ECO", date(2026, 9, 16)), ("ECO", date(2026, 9, 23))} <= set(_cles(rapport.ecarts))


def test_t_aud_27_code_retour_et_sans_ecart_par_ecole():
    """T-AUD-27 — `code_retour` vaut 3 dès qu'il y a un écart, 0 sinon ; `sans_ecart(ecole)` répond école
    par école."""
    seances = [_seance(date(2026, 9, 15), "09:00", "12:30", I_ABC), _seance(date(2026, 9, 16), "14:00", "17:00", I_ECO)]
    sessions = [_ses("ECO", date(2026, 9, 16), "14:00", "17:00", 3.0, n="R_ECO001", promotion="M2", module="Module B")]
    rapport = _auditer(seances, _master(sessions))
    assert rapport.code_retour == 3
    assert not rapport.sans_ecart("ABC") and rapport.sans_ecart("ECO")

    vide = _auditer([], _master([]))
    assert list(vide.ecarts) == [] and vide.code_retour == 0 and vide.sans_ecart("ABC")


def test_t_aud_31_ecole_sans_regle_auditee():
    """T-AUD-31 — 💡 [proposé — à valider] Une école classée mais absente de `Règles_facturation` n'est pas
    `Hors facturation` : elle est auditée."""
    seance = _seance(date(2026, 9, 15), "09:00", "12:30", I_ZZZ)
    rapport = _auditer([seance], _master([]))
    assert [(e.nature, e.ecole, e.seance) for e in rapport.ecarts] == [(TENUE_NON_FACTUREE, "ZZZ", seance)]


# ──────────────────────────────────────────────────────────────────────
# Erreurs
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    ("cas", "motif"),
    [
        ("periode-inversee", "invers"),
        ("correspondances", "Clé fictive mal formée"),
        ("regles", "Règles_facturation"),
    ],
)
def test_t_aud_30_erreurs_de_l_audit(cas, motif):
    """T-AUD-30 — Période inversée, ligne de `Correspondances` sans clé tarif ou ligne de
    `Règles_facturation` illisible : `ValueError` dont le message nomme la cause."""
    seances = [_seance(date(2026, 9, 15), "09:00", "12:30", I_ABC)]
    options = {}
    master = _master([])
    if cas == "periode-inversee":
        options = {"du": AU, "au": DU}
    elif cas == "correspondances":
        mal_formee = {**_corresp("Clé fictive mal formée", "ABC", "Module A", "M1"), "Clé tarif": None}
        master = _master([], correspondances=[*CORRESPONDANCES, mal_formee])
    else:
        master = _master([], regles=[*REGLES, _regle("ZZQ", "Lot 9")])
    with pytest.raises(ValueError, match=motif):
        _auditer(seances, master, **options)


# ──────────────────────────────────────────────────────────────────────
# Rapport
# ──────────────────────────────────────────────────────────────────────


def _rapport_varie():
    seances = [
        _seance(date(2026, 9, 22), "09:00", "12:30", I_ABC),
        _seance(date(2026, 9, 29), "09:00", "12:30", I_ABC),
        _seance(date(2026, 9, 30), "10:00", "11:00", "Réunion fictive non reconnue"),
    ]
    sessions = [
        _ses("ECO", date(2026, 9, 23), "14:00", "17:00", 3.0, n="R_ECO001", promotion="M2", module="Module B"),
        _ses("ABC", date(2026, 9, 29), "09:00", "12:30", 3.0),
        _ses("ABC", "fin sept", "09:00", "12:30", 3.5),
    ]
    return _auditer(seances, _master(sessions))


def test_t_aud_28_rapport_texte():
    """T-AUD-28 — Le texte du rapport cite l'entité, la période (JJ/MM/AAAA), chaque écart (nature, école,
    date ; intitulé d'un créneau non reconnu) et les lignes illisibles avec leur rang ; un rapport sans
    écart le dit."""
    aud = _aud()
    rapport = _rapport_varie()
    assert {e.nature for e in rapport.ecarts} >= {TENUE_NON_FACTUREE, ABSENTE, ECART_HEURES, NON_RECONNUE}
    texte = aud.rapport_texte(rapport)
    for fragment in (
        "ROSARI", "01/09/2026", "31/10/2026", TENUE_NON_FACTUREE, ABSENTE, ECART_HEURES, NON_RECONNUE,
        "22/09/2026", "23/09/2026", "29/09/2026", "30/09/2026", "ABC", "ECO", "Réunion fictive non reconnue",
        "fin sept", "ligne 4",
    ):
        assert fragment in texte, fragment

    vide = aud.rapport_texte(_auditer([], _master([])))
    assert "aucun écart" in vide.lower()
    assert "01/09/2026" in vide and "31/10/2026" in vide


def _date_du_nom(texte: str) -> date:
    return datetime.strptime(texte.replace("-", ""), "%Y%m%d").date()


def test_t_aud_29_ecrire_rapport_txt_et_csv(tmp_path):
    """T-AUD-29 — `ecrire_rapport` écrit `audit_<du>_<au>_<AAAAMMJJ-HHMMSS>.txt` (le texte du rapport) et
    `.csv` (UTF-8 avec BOM, séparateur « ; », une ligne par écart au moins) dans le dossier, qu'il crée ;
    sans fichier temporaire laissé ; un second appel n'écrase rien."""
    aud = _aud()
    rapport = _rapport_varie()
    dossier = tmp_path / "_sorties" / "audits"
    txt, csv_ = aud.ecrire_rapport(rapport, dossier)
    assert (txt.parent, csv_.parent) == (dossier, dossier)
    assert (txt.suffix, csv_.suffix) == (".txt", ".csv") and txt.stem == csv_.stem
    nom = re.fullmatch(r"audit_(\d{4}-?\d{2}-?\d{2})_(\d{4}-?\d{2}-?\d{2})_(\d{8}-\d{6})(?:-\d+)?", txt.stem)
    assert nom, txt.name
    assert (_date_du_nom(nom[1]), _date_du_nom(nom[2])) == (DU, AU)

    assert aud.rapport_texte(rapport).strip() in txt.read_text(encoding="utf-8-sig")
    octets = csv_.read_bytes()
    assert octets.startswith(b"\xef\xbb\xbf")
    lignes = list(csv.reader(io.StringIO(octets.decode("utf-8-sig")), delimiter=";"))
    assert len(lignes[0]) >= 3, lignes[0]
    for ecart in rapport.ecarts:
        dates = (f"{ecart.date:%d/%m/%Y}", f"{ecart.date:%Y-%m-%d}")
        assert any(
            any(ecart.nature in cellule for cellule in ligne) and any(d in cellule for cellule in ligne for d in dates)
            for ligne in lignes[1:]
        ), ecart
    assert sorted(p.name for p in dossier.iterdir()) == sorted([txt.name, csv_.name])

    txt2, csv2 = aud.ecrire_rapport(rapport, dossier)
    assert len({txt, csv_, txt2, csv2}) == 4
    assert len(list(dossier.iterdir())) == 4
    assert txt.read_text(encoding="utf-8-sig").strip() == txt2.read_text(encoding="utf-8-sig").strip()
