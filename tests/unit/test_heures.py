"""Heures facturables et taux horaire (F-13, ADR009, ADR016 ; banque de tests : T-HRS-01 à T-HRS-51).

Tests d'acceptation écrits avant l'implémentation : chaque test importe `facturation.heures`
lui-même, pour que l'absence du module fasse échouer les tests sans bloquer la collecte des
autres. Les séances (`facturation.agenda.Seance`) et les lignes du MASTER (dictionnaires
{en-tête : valeur}, forme de `facturation.master`) sont construites à la main, avec des
données fictives : écoles ABC, ECO et XYZ, modules et taux inventés.

Contrat (vague 2) :
- `heures_facturables(seance, correspondance) -> float`, selon la colonne `Mode heures` :
  `agenda` (défaut, cellule vide comprise) → `seance.duree_h` ; `coefficient` → durée × `Valeur` ;
  `forfait` → `Valeur` ; arrondi à 2 décimales, sans autre arrondi ; mode inconnu, ou `Valeur`
  absente ou négative en `coefficient` et `forfait` → `ValueError` qui nomme l'intitulé de la
  séance et le mode ;
- `taux_horaire(cle_tarif, tarifs) -> float` : clé `ÉCOLE|Module|Classe`, comparée sans tenir
  compte de la casse ni des espaces en début et en fin, aux colonnes `École`, `Module / Matière`,
  `Classe / Programme` de `Tarifs` ; clé introuvable ou taux vide → `ValueError` (ne jamais
  inventer un taux).

Lectures retenues là où le contrat se tait (à valider) : le mode est reconnu sans tenir compte
de la casse ni des espaces autour ; une correspondance sans colonne `Mode heures` vaut `agenda` ;
une `Valeur` nulle est acceptée (elle n'est pas négative) ; une `Valeur` ou un taux non numérique
est refusé comme une valeur absente ; les espaces sont retirés autour de chaque segment de la clé
tarif et autour de chaque cellule de `Tarifs`. 💡 [proposé — à valider] une `Valeur` saisie alors
que `Mode heures` est vide est ignorée (T-HRS-03) : question ouverte, relevée en revue.

Revue croisée de F-13 (T-HRS-25 à T-HRS-40) : T-HRS-25 à T-HRS-31 décrivent des défauts relevés
et sont écrits rouges, avant le correctif ; T-HRS-32 à T-HRS-40 couvrent des gardes déjà codées
que les relecteurs ont trouvées sans test, chacune vérifiée en la cassant (AGENTS.md §5).
"""

from __future__ import annotations

import copy
import importlib
import math
import re
import unicodedata
from datetime import date, datetime, time
from zoneinfo import ZoneInfo

import pytest

from facturation.agenda import Seance

pytestmark = pytest.mark.unit

PARIS = ZoneInfo("Europe/Paris")
JOUR = date(2026, 11, 12)  # un jeudi fictif

CLE_STATS_M1 = "ABC|Statistiques fictives|M1"

TARIFS = [  # onglet `Tarifs` fictif, colonnes de MASTER_SCHEMA (dont `Notes` et `Réf. contrat`)
    {"École": "ABC", "Module / Matière": "Statistiques fictives", "Classe / Programme": "M1",
     "Taux HT/h": 45, "Notes": "", "Réf. contrat": "CT-FICTIF-01"},
    {"École": "ABC", "Module / Matière": "Statistiques fictives", "Classe / Programme": "M2",
     "Taux HT/h": 50.0, "Notes": "", "Réf. contrat": "CT-FICTIF-01"},
    {"École": "ABC", "Module / Matière": "Économétrie fictive", "Classe / Programme": "M1",
     "Taux HT/h": 52.5, "Notes": None, "Réf. contrat": None},
    {"École": "ABC", "Module / Matière": "Jury fictif", "Classe / Programme": "M2",
     "Taux HT/h": 60, "Notes": "forfait de soutenance", "Réf. contrat": "CT-FICTIF-02"},
    {"École": "ECO", "Module / Matière": "Statistiques fictives", "Classe / Programme": "M1",
     "Taux HT/h": 40, "Notes": "", "Réf. contrat": "CT-FICTIF-03"},
]


def _hrs():
    """Module testé, importé dans le test : son absence fait échouer le test, pas la collecte."""
    return importlib.import_module("facturation.heures")


def _seance(debut: str = "09:00", fin: str = "12:00", *, jour: date = JOUR,
            intitule: str = "Cours fictif de statistiques") -> Seance:
    """Une séance de l'agenda, à Paris ; durée calculée sur les instants, comme `agenda.lire_ics`."""
    h1, m1 = map(int, debut.split(":"))
    h2, m2 = map(int, fin.split(":"))
    t1 = datetime(jour.year, jour.month, jour.day, h1, m1, tzinfo=PARIS)
    t2 = datetime(jour.year, jour.month, jour.day, h2, m2, tzinfo=PARIS)
    return _seance_entre(t1, t2, intitule=intitule)


def _seance_entre(t1: datetime, t2: datetime, *, intitule: str = "Cours fictif de statistiques") -> Seance:
    return Seance(
        uid=f"fictif-{t1:%Y%m%dT%H%M}@example.org",
        date=t1.date(),
        debut=t1,
        fin=t2,
        duree_h=(t2.timestamp() - t1.timestamp()) / 3600,
        intitule=intitule,
        description="",
    )


def _correspondance(mode=None, valeur=None, **autres) -> dict:
    """Une ligne de `Correspondances` (colonnes de MASTER_SCHEMA) ; `None` = cellule vide."""
    ligne = {
        "Clé": "Cours fictif de statistiques",
        "Type de clé": "exact",
        "École": "ABC",
        "Module": "Statistiques fictives",
        "Promotion": "M1",
        "Clé tarif": CLE_STATS_M1,
        "Mode heures": mode,
        "Valeur": valeur,
        "Gardée le": date(2026, 11, 2),
        "Notes": "",
    }
    ligne.update(autres)
    return ligne


def _heures(seance: Seance, correspondance: dict) -> float:
    return _hrs().heures_facturables(seance, correspondance)


def _taux(cle, tarifs=TARIFS) -> float:
    return _hrs().taux_horaire(cle, tarifs)


def _segments(cle: str) -> list[str]:
    return [s.strip().casefold() for s in cle.split("|") if s.strip()]


# ──────────────────────────────────────────────────────────────────────
# heures_facturables — critères d'acceptation de F-13 (ROADMAP)
# ──────────────────────────────────────────────────────────────────────


def test_t_hrs_01_forfait_de_jury():
    """T-HRS-01 — critère F-13 : un forfait de jury, 7 h d'agenda donnent 7,5 h facturées."""
    jury = _seance("09:00", "16:00", intitule="Jury de soutenance fictif")
    assert jury.duree_h == 7.0
    assert _heures(jury, _correspondance("forfait", 7.5)) == 7.5


@pytest.mark.parametrize("coefficient", [0.875, 7 / 8], ids=["0.875", "7/8"])
def test_t_hrs_02_journee_avec_coefficient_sept_huitiemes(coefficient):
    """T-HRS-02 — critère F-13 : une journée 9 h–17 h avec un coefficient 7/8 donne 7 h."""
    journee = _seance("09:00", "17:00", intitule="Journée fictive d'économétrie")
    resultat = _heures(journee, _correspondance("coefficient", coefficient))
    assert resultat == 7.0
    assert isinstance(resultat, float)


@pytest.mark.parametrize("mode", [None, "", "   "], ids=["None", "vide", "espaces"])
def test_t_hrs_03_sans_regle_heures_de_l_agenda(mode):
    """T-HRS-03 — critère F-13 : sans règle (cellule `Mode heures` vide), les heures facturées sont celles de l'agenda."""
    assert _heures(_seance("09:00", "12:30"), _correspondance(mode, None)) == 3.5
    # Une `Valeur` restée dans la ligne ne compte pas : le mode par défaut est `agenda`.
    assert _heures(_seance("09:00", "12:30"), _correspondance(mode, 0.5)) == 3.5


def test_t_hrs_04_correspondance_sans_colonne_mode_heures():
    """T-HRS-04 — une correspondance sans colonne `Mode heures` ni `Valeur` vaut le mode par défaut, `agenda`."""
    ligne = _correspondance()
    del ligne["Mode heures"], ligne["Valeur"]
    assert _heures(_seance("14:00", "18:00"), ligne) == 4.0
    assert _heures(_seance("14:00", "18:00"), {}) == 4.0


# ──────────────────────────────────────────────────────────────────────
# heures_facturables — les trois modes
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize("valeur", [None, 0.875, 7.5], ids=["sans-valeur", "coefficient-oublie", "forfait-oublie"])
def test_t_hrs_05_mode_agenda_explicite_ignore_valeur(valeur):
    """T-HRS-05 — mode `agenda` : la durée de la séance, quelle que soit la `Valeur` restée dans la ligne."""
    assert _heures(_seance("08:30", "12:00"), _correspondance("agenda", valeur)) == 3.5


def test_t_hrs_06_duree_lue_dans_seance_duree_h():
    """T-HRS-06 — la durée vient de `seance.duree_h` : 1 h – 4 h la nuit du passage à l'heure d'hiver font 4 h."""
    nuit = _seance_entre(
        datetime(2026, 10, 25, 1, 0, tzinfo=PARIS),
        datetime(2026, 10, 25, 4, 0, tzinfo=PARIS),
        intitule="Cours fictif de nuit",
    )
    assert nuit.duree_h == 4.0  # 3 h à l'horloge, 4 h réelles
    assert _heures(nuit, _correspondance("agenda")) == 4.0
    assert _heures(nuit, _correspondance("coefficient", 0.5)) == 2.0


@pytest.mark.parametrize(
    ("mode", "valeur", "attendu"),
    [
        ("Agenda", None, 8.0),
        (" AGENDA ", None, 8.0),
        ("Coefficient", 0.875, 7.0),
        (" coefficient  ", 0.875, 7.0),
        ("FORFAIT", 7.5, 7.5),
        ("Forfait ", 7.5, 7.5),
    ],
)
def test_t_hrs_07_mode_sans_casse_ni_espaces(mode, valeur, attendu):
    """T-HRS-07 — le mode saisi dans Excel est reconnu sans tenir compte de la casse ni des espaces autour."""
    assert _heures(_seance("09:00", "17:00"), _correspondance(mode, valeur)) == attendu


# ──────────────────────────────────────────────────────────────────────
# heures_facturables — arrondi
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    ("debut", "fin", "mode", "valeur", "attendu"),
    [
        ("09:00", "10:00", "coefficient", 1 / 3, 0.33),
        ("09:00", "11:00", "coefficient", 1 / 3, 0.67),
        ("09:00", "10:30", "coefficient", 0.875, 1.31),  # 1,3125
        ("09:00", "09:50", "agenda", None, 0.83),  # 50 min
        ("09:00", "12:00", "forfait", 2.3333, 2.33),
    ],
)
def test_t_hrs_08_arrondi_a_deux_decimales(debut, fin, mode, valeur, attendu):
    """T-HRS-08 — le résultat est arrondi à 2 décimales, dans les trois modes."""
    assert _heures(_seance(debut, fin), _correspondance(mode, valeur)) == attendu


@pytest.mark.parametrize(
    ("debut", "fin", "mode", "valeur", "attendu"),
    [
        ("09:00", "09:45", "agenda", None, 0.75),  # ni 1 h, ni 0,5 h
        ("09:00", "10:10", "agenda", None, 1.17),  # pas 1,25 h : aucun pas d'un quart d'heure
        ("09:00", "12:00", "coefficient", 0.9, 2.7),
        ("09:00", "12:00", "forfait", 7.25, 7.25),
        ("09:00", "12:00", "coefficient", 1.25, 3.75),  # coefficient au-dessus de 1
    ],
)
def test_t_hrs_09_aucun_autre_arrondi(debut, fin, mode, valeur, attendu):
    """T-HRS-09 — aucun autre arrondi : ni à l'heure, ni à la demi-heure, ni au quart d'heure."""
    assert _heures(_seance(debut, fin), _correspondance(mode, valeur)) == attendu


@pytest.mark.parametrize(
    ("mode", "valeur", "attendu"),
    [("forfait", 7, 7.0), ("coefficient", 1, 3.0), ("agenda", None, 3.0)],
)
def test_t_hrs_10_resultat_toujours_float(mode, valeur, attendu):
    """T-HRS-10 — le résultat est un float, même quand la `Valeur` lue dans Excel est un entier."""
    resultat = _heures(_seance("09:00", "12:00"), _correspondance(mode, valeur))
    assert resultat == attendu
    assert isinstance(resultat, float)


@pytest.mark.parametrize("mode", ["coefficient", "forfait"])
def test_t_hrs_11_valeur_nulle_acceptee(mode):
    """T-HRS-11 — une `Valeur` nulle n'est pas négative : elle est acceptée et donne 0 h."""
    resultat = _heures(_seance("09:00", "12:00"), _correspondance(mode, 0))
    assert resultat == 0.0
    assert isinstance(resultat, float)


# ──────────────────────────────────────────────────────────────────────
# heures_facturables — erreurs
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize("mode", ["demi-journée", "horaire", "forfaitaire"])
def test_t_hrs_12_mode_inconnu(mode):
    """T-HRS-12 — mode inconnu (même voisin d'un mode connu) → ValueError qui nomme l'intitulé et le mode."""
    seance = _seance("09:00", "17:00", intitule="Atelier fictif du jeudi")
    with pytest.raises(ValueError) as erreur:
        _heures(seance, _correspondance(mode, 0.875))
    message = str(erreur.value)
    assert "Atelier fictif du jeudi" in message
    assert mode.casefold() in message.casefold()


@pytest.mark.parametrize("mode", ["coefficient", "forfait"])
@pytest.mark.parametrize("valeur", [None, "", "   "], ids=["None", "vide", "espaces"])
def test_t_hrs_13_valeur_absente(mode, valeur):
    """T-HRS-13 — `Valeur` absente en `coefficient` ou `forfait` → ValueError qui nomme l'intitulé, le mode et la colonne."""
    seance = _seance("09:00", "16:00", intitule="Jury fictif de fin d'année")
    with pytest.raises(ValueError) as erreur:
        _heures(seance, _correspondance(mode, valeur))
    message = str(erreur.value).casefold()
    assert "jury fictif de fin d'année" in message
    assert mode in message
    assert "valeur" in message


@pytest.mark.parametrize(("mode", "valeur"), [("coefficient", -0.5), ("forfait", -7.5), ("forfait", -1)])
def test_t_hrs_14_valeur_negative(mode, valeur):
    """T-HRS-14 — `Valeur` négative en `coefficient` ou `forfait` → ValueError qui nomme l'intitulé, le mode et la colonne."""
    seance = _seance("09:00", "16:00", intitule="Séminaire fictif")
    with pytest.raises(ValueError) as erreur:
        _heures(seance, _correspondance(mode, valeur))
    message = str(erreur.value).casefold()
    assert "séminaire fictif" in message
    assert mode in message
    assert "valeur" in message


@pytest.mark.parametrize("mode", ["coefficient", "forfait"])
def test_t_hrs_15_valeur_non_numerique(mode):
    """T-HRS-15 — `Valeur` non numérique (texte saisi dans Excel) → ValueError (pas TypeError) qui nomme l'intitulé et le mode."""
    seance = _seance("09:00", "16:00", intitule="Soutenance fictive")
    with pytest.raises(ValueError) as erreur:
        _heures(seance, _correspondance(mode, "sept heures"))
    message = str(erreur.value).casefold()
    assert "soutenance fictive" in message
    assert mode in message


def test_t_hrs_16_calcul_sans_ecriture():
    """T-HRS-16 — le calcul n'écrit rien : la correspondance reste identique après l'appel."""
    ligne = _correspondance("coefficient", 0.875)
    avant = copy.deepcopy(ligne)
    _heures(_seance("09:00", "17:00"), ligne)
    assert ligne == avant


# ──────────────────────────────────────────────────────────────────────
# taux_horaire
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    ("cle", "attendu"),
    [
        (CLE_STATS_M1, 45.0),
        ("ABC|Statistiques fictives|M2", 50.0),
        ("ABC|Économétrie fictive|M1", 52.5),
        ("ECO|Statistiques fictives|M1", 40.0),
    ],
)
def test_t_hrs_17_taux_de_la_ligne_designee(cle, attendu):
    """T-HRS-17 — le taux de la seule ligne dont l'école, le module et la classe correspondent ; un float, même lu entier."""
    taux = _taux(cle)
    assert taux == attendu
    assert isinstance(taux, float)


@pytest.mark.parametrize(
    ("cle", "attendu"),
    [
        ("abc|statistiques FICTIVES|m1", 45.0),
        ("  ABC|Statistiques fictives|M1  ", 45.0),
        (" ABC | Statistiques fictives | M2 ", 50.0),
        ("abc|ÉCONOMÉTRIE FICTIVE|m1", 52.5),
        ("abc|économétrie fictive|M1", 52.5),
        ("xyz|algèbre fictive|l3", 38.0),  # cellules entourées d'espaces dans Tarifs
    ],
)
def test_t_hrs_18_cle_sans_casse_ni_espaces(cle, attendu):
    """T-HRS-18 — clé et cellules comparées sans tenir compte de la casse (accents compris) ni des espaces autour."""
    tarifs = TARIFS + [
        {"École": " XYZ ", "Module / Matière": "  Algèbre fictive ", "Classe / Programme": " L3 ",
         "Taux HT/h": 38, "Notes": "", "Réf. contrat": ""},
    ]
    assert _taux(cle, tarifs) == attendu


@pytest.mark.parametrize(
    ("cle", "tarifs"),
    [
        ("ZZZ|Statistiques fictives|M1", TARIFS),  # école inconnue
        ("ABC|Probabilités fictives|M1", TARIFS),  # module inconnu
        ("ABC|Statistiques fictives|M3", TARIFS),  # classe voisine absente
        ("ABC|Statistiques fictives|M", TARIFS),  # pas de préfixe : M ne désigne ni M1 ni M2
        ("ABC|Statistiques|M1", TARIFS),  # pas de préfixe sur le module
        ("ECO|Économétrie fictive|M1", TARIFS),  # module d'une autre école
        (CLE_STATS_M1, []),  # onglet Tarifs vide
    ],
    ids=["ecole", "module", "classe", "prefixe-classe", "prefixe-module", "autre-ecole", "tarifs-vides"],
)
def test_t_hrs_19_cle_introuvable(cle, tarifs):
    """T-HRS-19 — clé introuvable → ValueError qui nomme la clé ; aucun taux voisin n'est pris à la place."""
    with pytest.raises(ValueError) as erreur:
        _taux(cle, tarifs)
    message = str(erreur.value).casefold()
    for segment in _segments(cle):
        assert segment in message


@pytest.mark.parametrize("taux", [None, "", "   "], ids=["None", "vide", "espaces"])
def test_t_hrs_20_taux_vide(taux):
    """T-HRS-20 — ligne trouvée mais taux vide → ValueError qui nomme la clé et le taux (ne jamais inventer un taux)."""
    tarifs = TARIFS + [
        {"École": "ABC", "Module / Matière": "Projet fictif", "Classe / Programme": "M1",
         "Taux HT/h": taux, "Notes": "taux à confirmer", "Réf. contrat": ""},
    ]
    cle = "ABC|Projet fictif|M1"
    with pytest.raises(ValueError) as erreur:
        _taux(cle, tarifs)
    message = str(erreur.value).casefold()
    for segment in _segments(cle):
        assert segment in message
    assert "taux" in message


def test_t_hrs_21_taux_non_numerique():
    """T-HRS-21 — taux non numérique (texte saisi dans Excel) → ValueError qui nomme la clé, jamais un taux deviné."""
    tarifs = TARIFS + [
        {"École": "ABC", "Module / Matière": "Projet fictif", "Classe / Programme": "M1",
         "Taux HT/h": "à négocier", "Notes": "", "Réf. contrat": ""},
    ]
    with pytest.raises(ValueError) as erreur:
        _taux("ABC|Projet fictif|M1", tarifs)
    message = str(erreur.value).casefold()
    assert "projet fictif" in message


@pytest.mark.parametrize(
    "cle",
    [None, "", "   ", "ABC|Statistiques fictives", "ABC", "ABC|Statistiques fictives|M1|extra", "||"],
    ids=["None", "vide", "espaces", "deux-segments", "un-segment", "quatre-segments", "segments-vides"],
)
def test_t_hrs_22_cle_absente_ou_mal_formee(cle):
    """T-HRS-22 — clé tarif absente (cellule vide) ou qui n'a pas la forme `ÉCOLE|Module|Classe` → ValueError."""
    with pytest.raises(ValueError):
        _taux(cle)


def test_t_hrs_23_lignes_incompletes_et_colonnes_en_plus():
    """T-HRS-23 — des lignes incomplètes de `Tarifs` (cellules vides) ne gênent pas la recherche ; les colonnes en plus sont ignorées."""
    tarifs = [
        {"École": None, "Module / Matière": None, "Classe / Programme": None, "Taux HT/h": None,
         "Notes": "ligne vide", "Réf. contrat": None},
        {"École": "ABC", "Module / Matière": None, "Classe / Programme": None, "Taux HT/h": 99,
         "Notes": "ligne commencée", "Réf. contrat": None},
        {"École": "ABC", "Module / Matière": "Statistiques fictives", "Classe / Programme": None,
         "Taux HT/h": 98, "Notes": "", "Réf. contrat": None},
    ] + TARIFS
    assert _taux(CLE_STATS_M1, tarifs) == 45.0


# ──────────────────────────────────────────────────────────────────────
# Parcours : heures × taux d'une correspondance
# ──────────────────────────────────────────────────────────────────────


def test_t_hrs_24_heures_et_taux_d_une_correspondance():
    """T-HRS-24 — d'une séance et de sa correspondance : heures facturables et taux de la `Clé tarif`, sans rien inventer."""
    jury = _seance("09:00", "16:00", intitule="Jury de soutenance fictif")
    ligne_jury = _correspondance("forfait", 7.5, Module="Jury fictif", Promotion="M2",
                                 **{"Clé tarif": "ABC|Jury fictif|M2"})
    journee = _seance("09:00", "17:00", intitule="Journée fictive de statistiques")
    ligne_journee = _correspondance("coefficient", 0.875)

    heures_jury = _heures(jury, ligne_jury)
    heures_journee = _heures(journee, ligne_journee)
    taux_jury = _taux(ligne_jury["Clé tarif"])
    taux_journee = _taux(ligne_journee["Clé tarif"])

    assert (heures_jury, taux_jury) == (7.5, 60.0)
    assert (heures_journee, taux_journee) == (7.0, 45.0)
    assert heures_jury * taux_jury + heures_journee * taux_journee == 765.0


# ──────────────────────────────────────────────────────────────────────
# Revue croisée de F-13 — défauts relevés (rouges avant le correctif)
# ──────────────────────────────────────────────────────────────────────

ABSENTE = re.compile(r"absente|manquante|manque")


def _tarifs_renommes(ancien: str, nouveau: str | None) -> list[dict]:
    """TARIFS où la colonne `ancien` est renommée `nouveau`, ou retirée si `nouveau` vaut None."""
    lignes = []
    for ligne in TARIFS:
        copie = {}
        for colonne, valeur in ligne.items():
            if colonne != ancien:
                copie[colonne] = valeur
            elif nouveau is not None:
                copie[nouveau] = valeur
        lignes.append(copie)
    return lignes


@pytest.mark.parametrize(
    ("debut", "fin", "mode", "valeur"),
    [
        ("09:00", "09:00", None, None),
        ("09:00", "09:00", "agenda", None),
        ("09:00", "09:00", "coefficient", 0.875),
        ("12:00", "09:00", "agenda", None),
        ("12:00", "09:00", "coefficient", 0.5),
    ],
    ids=["nulle-mode-vide", "nulle-agenda", "nulle-coefficient", "negative-agenda", "negative-coefficient"],
)
def test_t_hrs_25_duree_nulle_ou_negative_refusee(debut, fin, mode, valeur):
    """T-HRS-25 — 💡 [proposé — à valider] une durée d'agenda nulle (événement sans fin ni durée,
    que F-10 lit comme durant zéro) ou négative n'est pas facturée en silence en `agenda` ni en
    `coefficient` : ValueError qui nomme la séance et la durée (« ne jamais inventer un volume
    horaire »). Le `forfait`, qui ne lit pas la durée, n'est pas visé ici."""
    seance = _seance(debut, fin, intitule="Cours fictif sans fin")
    assert seance.duree_h <= 0
    with pytest.raises(ValueError) as erreur:
        _heures(seance, _correspondance(mode, valeur))
    message = str(erreur.value).casefold()
    assert "cours fictif sans fin" in message
    assert "durée" in message


@pytest.mark.parametrize(
    ("colonne", "remplacee_par"),
    [
        ("Classe / Programme", "Classe / Groupe"),  # le nom de la colonne de l'onglet Cours
        ("École", "Ecole"),
        ("Taux HT/h", "Taux HT / h"),
        ("Taux HT/h", None),
    ],
    ids=["classe-groupe", "ecole-sans-accent", "taux-mal-nomme", "taux-absent"],
)
def test_t_hrs_26_colonne_de_tarifs_absente(colonne, remplacee_par):
    """T-HRS-26 — une colonne attendue absente de `Tarifs` (ou mal nommée) → ValueError qui nomme
    la colonne et dit qu'elle manque, au lieu de « clé introuvable » ou de « taux vide » : Adrian
    chercherait une ligne ou un taux qui existent bel et bien."""
    with pytest.raises(ValueError) as erreur:
        _taux(CLE_STATS_M1, _tarifs_renommes(colonne, remplacee_par))
    message = str(erreur.value)
    assert colonne in message
    bas = message.casefold()
    assert ABSENTE.search(bas)
    assert "introuvable" not in bas
    assert "taux vide" not in bas


def test_t_hrs_27_en_tete_de_tarifs_decompose():
    """T-HRS-27 — un en-tête `École` en forme décomposée (NFD, collé depuis le web ou macOS) ne
    donne pas « clé introuvable » : soit l'en-tête est reconnu et le taux rendu, soit l'erreur dit
    que la colonne manque."""
    nfd = unicodedata.normalize("NFD", "École")
    assert nfd != "École"
    try:
        taux = _taux(CLE_STATS_M1, _tarifs_renommes("École", nfd))
    except ValueError as erreur:
        message = str(erreur).casefold()
        assert "introuvable" not in message
        assert ABSENTE.search(message)
    else:
        assert taux == 45.0


def test_t_hrs_28_meme_cle_une_ligne_au_taux_vide():
    """T-HRS-28 — même clé sur deux lignes, l'une au taux 45, l'autre au taux vide → ValueError qui
    dit qu'il y a 2 lignes, pas seulement « taux vide » (Adrian voit 45 dans le MASTER)."""
    tarifs = TARIFS + [{**TARIFS[0], "Taux HT/h": None}]
    with pytest.raises(ValueError) as erreur:
        _taux(CLE_STATS_M1, tarifs)
    message = str(erreur.value).casefold()
    assert "statistiques fictives" in message
    assert re.search(r"\b(2|deux) lignes", message)


def test_t_hrs_29_taux_differents_affiches_distincts():
    """T-HRS-29 — deux taux qui ne diffèrent que par un bruit de calcul (formule Excel) : le message
    montre deux valeurs distinctes, pas « (45, 45) »."""
    tarifs = TARIFS + [{**TARIFS[0], "Taux HT/h": 45.0000001}]
    with pytest.raises(ValueError) as erreur:
        _taux(CLE_STATS_M1, tarifs)
    assert re.search(r"45[.,]0000001", str(erreur.value))


@pytest.mark.parametrize("valeur", ["sept heures et demie", time(7, 30)], ids=["texte", "heure"])
def test_t_hrs_30_message_forfait_non_numerique(valeur):
    """T-HRS-30 — en `forfait`, le message d'une `Valeur` non numérique donne un exemple d'heures
    décimales (7,5), pas le coefficient 0,875."""
    with pytest.raises(ValueError) as erreur:
        _heures(_seance("09:00", "16:00"), _correspondance("forfait", valeur))
    message = str(erreur.value)
    assert "0,875" not in message
    assert re.search(r"7[,.]5", message)


@pytest.mark.parametrize(
    ("mode", "valeur", "cause"),
    [
        ("coefficient", datetime(2026, 9, 7), "date"),  # 7/9 saisi dans Excel devient le 7 septembre
        ("forfait", time(7, 30), "heure"),  # 7:30 saisi dans Excel devient une heure
    ],
    ids=["date", "heure"],
)
def test_t_hrs_31_message_valeur_lue_comme_date_ou_heure(mode, valeur, cause):
    """T-HRS-31 — une `Valeur` qu'Excel a convertie en date ou en heure : le message nomme cette
    cause, au lieu de « pas comme du texte »."""
    with pytest.raises(ValueError) as erreur:
        _heures(_seance("09:00", "17:00"), _correspondance(mode, valeur))
    message = str(erreur.value).casefold()
    assert cause in message
    assert "texte" not in message


# ──────────────────────────────────────────────────────────────────────
# Revue croisée de F-13 — gardes déjà codées, sans test jusqu'ici
# ──────────────────────────────────────────────────────────────────────

LIGNES_INCOMPLETES = [  # comme T-HRS-23 : des lignes de Tarifs commencées puis laissées
    {"École": None, "Module / Matière": None, "Classe / Programme": None, "Taux HT/h": None,
     "Notes": "ligne vide", "Réf. contrat": None},
    {"École": "ABC", "Module / Matière": None, "Classe / Programme": None, "Taux HT/h": 99,
     "Notes": "ligne commencée", "Réf. contrat": None},
    {"École": "ABC", "Module / Matière": "Statistiques fictives", "Classe / Programme": None,
     "Taux HT/h": 98, "Notes": "", "Réf. contrat": None},
]


def _ligne_projet(taux) -> dict:
    return {"École": "ABC", "Module / Matière": "Projet fictif", "Classe / Programme": "M1",
            "Taux HT/h": taux, "Notes": "", "Réf. contrat": ""}


@pytest.mark.parametrize("taux", [0, 0.0, -45], ids=["zero", "zero-float", "negatif"])
def test_t_hrs_32_taux_nul_ou_negatif(taux):
    """T-HRS-32 — taux nul ou négatif → ValueError qui nomme la clé et le taux : jamais une facture à 0 €."""
    with pytest.raises(ValueError) as erreur:
        _taux("ABC|Projet fictif|M1", TARIFS + [_ligne_projet(taux)])
    message = str(erreur.value).casefold()
    assert "projet fictif" in message
    assert "taux" in message


def test_t_hrs_33_meme_cle_taux_differents_ou_egaux():
    """T-HRS-33 — même clé sur deux lignes : des taux différents lèvent une ValueError qui les cite
    (aucun n'est choisi) ; le même taux, saisi autrement (casse, 45 et 45.0), est accepté."""
    differents = TARIFS + [{**TARIFS[0], "Taux HT/h": 50}]
    with pytest.raises(ValueError) as erreur:
        _taux(CLE_STATS_M1, differents)
    message = str(erreur.value)
    assert "45" in message
    assert "50" in message

    egaux = TARIFS + [{**TARIFS[0], "École": "abc", "Module / Matière": " STATISTIQUES FICTIVES ",
                       "Classe / Programme": "m1", "Taux HT/h": 45.0}]
    assert _taux(CLE_STATS_M1, egaux) == 45.0


@pytest.mark.parametrize(
    ("debut", "fin", "mode", "valeur", "attendu"),
    [
        ("09:00", "12:00", "forfait", 1.125, 1.13),  # round() donne 1,12
        ("09:00", "10:30", "coefficient", 0.75, 1.13),  # 1,5 h × 0,75 = 1,125
        ("09:00", "10:06", "coefficient", 0.95, 1.05),  # 1,1 h × 0,95 = 1,045 ; le calcul en float donne 1,04
        ("09:00", "12:00", "forfait", 2.675, 2.68),  # round() donne 2,67
    ],
)
def test_t_hrs_34_arrondi_moitie_vers_le_haut(debut, fin, mode, valeur, attendu):
    """T-HRS-34 — l'arrondi à 2 décimales prend la moitié vers le haut, sur le nombre tel qu'il s'écrit."""
    assert _heures(_seance(debut, fin), _correspondance(mode, valeur)) == attendu


@pytest.mark.parametrize("mode", ["coefficient", "forfait"])
@pytest.mark.parametrize(
    "valeur",
    [True, False, math.nan, math.inf, "0.875", "7,5"],
    ids=["VRAI", "FAUX", "nan", "inf", "texte-point", "texte-virgule"],
)
def test_t_hrs_35_valeur_booleenne_non_finie_ou_texte(mode, valeur):
    """T-HRS-35 — une `Valeur` VRAI/FAUX, NaN, infinie ou numérique mais saisie en texte → ValueError
    qui nomme l'intitulé et le mode (VRAI ne vaut ni un coefficient de 1, ni un forfait d'1 h)."""
    seance = _seance("09:00", "17:00", intitule="Atelier fictif")
    with pytest.raises(ValueError) as erreur:
        _heures(seance, _correspondance(mode, valeur))
    message = str(erreur.value).casefold()
    assert "atelier fictif" in message
    assert mode in message


@pytest.mark.parametrize(
    "taux",
    [True, math.nan, math.inf, "quarante-cinq", "45 €/jour"],
    ids=["VRAI", "nan", "inf", "texte", "texte-par-jour"],
)
def test_t_hrs_36_taux_booleen_non_fini_ou_texte(taux):
    """T-HRS-36 — un taux VRAI, NaN, infini ou saisi en texte hors des formes reconnues → ValueError
    qui nomme la clé (VRAI ne vaut pas 1 €/h).

    Révisé par P-0b : le constat du MASTER réel (contrats des vagues 3 à 5, § 15) montre des taux
    saisis en texte (« 60 €/h ») ; la décision du § 15 fait lire « 45 » et « 45,5 » comme des
    nombres. Ces deux cas, refusés à l'origine, sont désormais couverts par T-HRS-41 ; ils sont
    remplacés ici par des textes qui restent refusés."""
    with pytest.raises(ValueError) as erreur:
        _taux("ABC|Projet fictif|M1", TARIFS + [_ligne_projet(taux)])
    assert "projet fictif" in str(erreur.value).casefold()


def test_t_hrs_37_cle_ou_cellule_en_forme_decomposee():
    """T-HRS-37 — `É` composé (NFC) ou décomposé (NFD) se valent, dans la cellule comme dans la clé."""
    nfd = unicodedata.normalize("NFD", "Économétrie fictive")
    assert nfd != "Économétrie fictive"
    cellule_nfd = [
        {**ligne, "Module / Matière": nfd} if ligne["Module / Matière"] == "Économétrie fictive" else ligne
        for ligne in TARIFS
    ]
    assert _taux("ABC|Économétrie fictive|M1", cellule_nfd) == 52.5
    assert _taux(f"ABC|{nfd}|M1") == 52.5


@pytest.mark.parametrize("classe", [3, 3.0], ids=["entier", "float"])
def test_t_hrs_38_classe_saisie_comme_nombre(classe):
    """T-HRS-38 — une classe saisie comme nombre (3, lu 3 ou 3.0) est trouvée par la clé `…|3`."""
    tarifs = TARIFS + [{"École": "ABC", "Module / Matière": "Statistiques fictives",
                        "Classe / Programme": classe, "Taux HT/h": 41, "Notes": "", "Réf. contrat": ""}]
    assert _taux("ABC|Statistiques fictives|3", tarifs) == 41.0


def test_t_hrs_39_seance_non_classee():
    """T-HRS-39 — aucune correspondance (`None`, séance à classer) → TypeError qui nomme la séance :
    un créneau à classer n'est jamais facturé par défaut."""
    seance = _seance(intitule="Cours fictif non classé")
    with pytest.raises(TypeError) as erreur:
        _heures(seance, None)
    assert "Cours fictif non classé" in str(erreur.value)


@pytest.mark.parametrize(
    "cle",
    ["||", "ABC|Statistiques fictives|", "|Statistiques fictives|M1", "ABC||M1",
     "ABC|Statistiques fictives", "ABC|Statistiques fictives|M1|extra"],
    ids=["segments-vides", "classe-vide", "ecole-vide", "module-vide", "deux-segments", "quatre-segments"],
)
def test_t_hrs_40_cle_mal_formee_jamais_un_taux_voisin(cle):
    """T-HRS-40 — une clé mal formée est refusée comme telle, même quand une ligne incomplète de
    `Tarifs` lui ressemble : `ABC|Statistiques fictives|` ne prend pas le taux 98 de la ligne sans classe."""
    with pytest.raises(ValueError) as erreur:
        _taux(cle, LIGNES_INCOMPLETES + TARIFS)
    assert "mal formée" in str(erreur.value).casefold()


# ══════════════════════════════════════════════════════════════════════
# P-0b — Socle du lot 1 (banque de tests : T-HRS-41 à T-HRS-49)
# ══════════════════════════════════════════════════════════════════════
#
# Le contrat donnait T-HRS-25 comme point de départ ; T-HRS-25 à T-HRS-40 étant déjà pris (revue
# de F-13), la suite commence à T-HRS-41. Le test « T-HRS-25 » du contrat (`ligne_tarif` rend
# `Réf. contrat` ; deux références pour un même taux lèvent `ValueError`) est T-HRS-43 et T-HRS-44.
#
# Contrat (contrats des vagues 3 à 5, § 1.2 et 3.6), et décisions du § 15 qui priment :
# - `ligne_tarif(cle_tarif, tarifs) -> dict` : même recherche et mêmes `ValueError` que
#   `taux_horaire`, qui s'appuie désormais dessus ; même taux, `Réf. contrat` différentes →
#   `ValueError` qui cite les références ;
# - `montant(h, t) = Decimal(repr(h)) * Decimal(repr(t))`, valeur exacte ;
#   `arrondi_centime` : `ROUND_HALF_UP`, au centime ;
# - § 15 : `Taux HT/h` est souvent du texte ; sont lus « 60 », « 60 € », « 60 €/h », « 60,5 €/h »,
#   « 60.50 », espaces (insécables compris) tolérés ; tout autre texte, ou une cellule vide →
#   `ValueError` qui nomme la clé tarif et la cellule ;
# - § 15 : la colonne `École` de `Tarifs` porte parfois le libellé de l'école (onglet `Écoles`) ;
#   un libellé inconnu donne une erreur qui nomme l'onglet, la ligne et la valeur.
#
# Lectures retenues (`💡 [proposé — à valider]`) :
# - « espaces tolérés » vaut aussi autour de « € » et de « / » (« 60 € / h ») ;
# - un taux en texte nul (« 0 €/h ») reste refusé comme un taux nul ;
# - la correspondance libellé → code de `Tarifs` passe par un argument nommé facultatif
#   `ecoles=` (lignes de l'onglet `Écoles`) de `ligne_tarif` et de `taux_horaire` ; sans lui, rien
#   ne change. Une ligne au libellé inconnu qui porte le module et la classe de la clé lève une
#   `ValueError` qui cite l'onglet `Tarifs`, la ligne (son module) et la valeur, et renvoie à
#   l'onglet `Écoles` : le taux n'est ni deviné ni déclaré « introuvable ».

from decimal import Decimal  # noqa: E402

ECOLES = [
    {"École": "École Fictive A", "Code facturation": "ABC"},
    {"École": "Institut Fictif B", "Code facturation": "ECO"},
]


@pytest.mark.parametrize(
    ("saisie", "attendu"),
    [
        ("60", 60.0), ("60 €", 60.0), ("60 €/h", 60.0), ("60€/h", 60.0), ("60,5 €/h", 60.5), ("60.50", 60.5),
        ("60 €/h", 60.0), (" 60,5 €/h ", 60.5), (" 60 € / h ", 60.0), ("45,75", 45.75),
    ],
    ids=["entier", "euro", "euro-heure", "sans-espace", "virgule", "point", "insecable", "insecables-fines",
         "espaces-partout", "deux-decimales"],
)
def test_t_hrs_41_taux_saisi_en_texte(saisie, attendu):
    """T-HRS-41 — décision du § 15 : un taux saisi en texte (« 60 €/h », forme du MASTER réel) se lit
    comme le nombre qu'il écrit, virgule ou point décimal, espaces insécables compris."""
    taux = _taux("ABC|Projet fictif|M1", TARIFS + [_ligne_projet(saisie)])
    assert taux == attendu
    assert isinstance(taux, float)


@pytest.mark.parametrize(
    "saisie",
    ["€/h", "à définir", "60 €/jour", "soixante €/h", "60-65 €/h", "60 €/h ou 65 €/h", "0 €/h", None, ""],
    ids=["sans-chiffre", "a-definir", "par-jour", "en-lettres", "fourchette", "deux-taux", "nul", "None", "vide"],
)
def test_t_hrs_42_taux_en_texte_illisible_refuse(saisie):
    """T-HRS-42 — tout autre texte (sans chiffre, par jour, fourchette, deux taux…), un taux nul ou une
    cellule vide → `ValueError` qui nomme la clé tarif et la cellule (colonne « Taux HT/h », contenu
    cité) : ne jamais inventer un taux."""
    with pytest.raises(ValueError) as erreur:
        _taux("ABC|Projet fictif|M1", TARIFS + [_ligne_projet(saisie)])
    message = str(erreur.value)
    assert "projet fictif" in message.casefold()
    assert "Taux HT/h" in message
    if saisie:
        assert saisie in message


def test_t_hrs_43_ligne_tarif_rend_la_ligne_et_sa_reference():
    """T-HRS-43 — `ligne_tarif` rend la ligne de `Tarifs` désignée par la clé, `Réf. contrat` comprise,
    avec la même recherche que `taux_horaire` (casse et espaces ignorés) et les mêmes refus."""
    hrs = _hrs()
    ligne = hrs.ligne_tarif(" abc | STATISTIQUES FICTIVES | m1 ", TARIFS)
    assert (ligne["Réf. contrat"], ligne["Taux HT/h"], ligne["Classe / Programme"]) == ("CT-FICTIF-01", 45, "M1")
    assert hrs.ligne_tarif("ABC|Économétrie fictive|M1", TARIFS)["Réf. contrat"] in (None, "")
    for cle in ("ABC|Statistiques fictives|M3", "XYZ|Statistiques fictives|M1", "ABC|Statistiques fictives"):
        with pytest.raises(ValueError) as erreur:
            hrs.ligne_tarif(cle, TARIFS)
        with pytest.raises(ValueError) as erreur_taux:
            hrs.taux_horaire(cle, TARIFS)
        assert str(erreur.value) == str(erreur_taux.value)


def test_t_hrs_44_meme_taux_references_differentes():
    """T-HRS-44 — même clé sur deux lignes au même taux mais aux `Réf. contrat` différentes :
    `ValueError` qui cite les deux références (`ligne_tarif`, et `taux_horaire` qui s'appuie
    dessus) ; même taux et même référence : la ligne est rendue."""
    hrs = _hrs()
    doublon = {**TARIFS[0], "Réf. contrat": "CT-FICTIF-09"}
    for fonction in (hrs.ligne_tarif, hrs.taux_horaire):
        with pytest.raises(ValueError) as erreur:
            fonction(CLE_STATS_M1, TARIFS + [doublon])
        message = str(erreur.value)
        assert "CT-FICTIF-01" in message and "CT-FICTIF-09" in message
    meme_reference = {**TARIFS[0], "Notes": "doublon fictif"}
    assert hrs.ligne_tarif(CLE_STATS_M1, TARIFS + [meme_reference])["Réf. contrat"] == "CT-FICTIF-01"
    assert hrs.taux_horaire(CLE_STATS_M1, TARIFS + [meme_reference]) == 45.0


@pytest.mark.parametrize(
    ("heures", "taux", "attendu"),
    [(3.5, 45.0, "157.5"), (1.13, 52.5, "59.325"), (0.1, 0.2, "0.02"), (7.5, 60, "450.0"), (2.25, 41.33, "92.9925")],
)
def test_t_hrs_45_montant_exact(heures, taux, attendu):
    """T-HRS-45 — `montant(h, t)` = `Decimal(repr(h)) * Decimal(repr(t))` : la valeur exacte, sans
    arrondi ni bruit binaire (0,1 × 0,2 = 0,02 et non 0,020000000000000004)."""
    resultat = _hrs().montant(heures, taux)
    assert isinstance(resultat, Decimal)
    assert resultat == Decimal(attendu)
    assert resultat == Decimal(repr(heures)) * Decimal(repr(taux))


@pytest.mark.parametrize(
    ("valeur", "attendu"),
    [("59.325", "59.33"), ("2.675", "2.68"), ("92.9925", "92.99"), ("10", "10.00"), ("157.5", "157.50"),
     ("0.005", "0.01"), ("1890.004", "1890.00")],
)
def test_t_hrs_46_arrondi_centime(valeur, attendu):
    """T-HRS-46 — `arrondi_centime` : au centime, la moitié vers le haut (59,325 → 59,33, là où
    l'arrondi bancaire donnerait 59,32), deux décimales écrites."""
    resultat = _hrs().arrondi_centime(Decimal(valeur))
    assert isinstance(resultat, Decimal)
    assert str(resultat) == attendu


def test_t_hrs_47_total_arrondi_une_seule_fois():
    """T-HRS-47 — un total vaut la somme exacte des montants, arrondie une seule fois (§ 1.2) : trois
    lignes de 1,005 € donnent 3,02 €, pas 3,03 € (somme des lignes arrondies)."""
    hrs = _hrs()
    montants = [hrs.montant(0.5, 2.01) for _ in range(3)]
    assert montants == [Decimal("1.005")] * 3
    assert hrs.arrondi_centime(sum(montants, Decimal(0))) == Decimal("3.02")
    assert sum((hrs.arrondi_centime(m) for m in montants), Decimal(0)) == Decimal("3.03")


TARIFS_LIBELLES = [
    {"École": "École Fictive A", "Module / Matière": "Projet fictif", "Classe / Programme": "M1",
     "Taux HT/h": "55 €/h", "Notes": "", "Réf. contrat": "CT-FICTIF-05"},
    {"École": "Institut Fictif B", "Module / Matière": "Projet fictif", "Classe / Programme": "M1",
     "Taux HT/h": 48, "Notes": "", "Réf. contrat": ""},
]


def test_t_hrs_48_ecole_designee_par_son_libelle():
    """T-HRS-48 — § 15 : la colonne `École` de `Tarifs` peut porter le libellé de l'école ; avec les
    lignes de l'onglet `Écoles` (`ecoles=`), la clé `ABC|…` trouve la ligne « École Fictive A »,
    sans la confondre avec celle d'une autre école ; sans `ecoles=`, rien n'est deviné."""
    hrs = _hrs()
    tarifs = TARIFS + TARIFS_LIBELLES
    assert hrs.taux_horaire("ABC|Projet fictif|M1", tarifs, ecoles=ECOLES) == 55.0
    assert hrs.ligne_tarif("ABC|Projet fictif|M1", tarifs, ecoles=ECOLES)["Réf. contrat"] == "CT-FICTIF-05"
    assert hrs.taux_horaire("ECO|Projet fictif|M1", tarifs, ecoles=ECOLES) == 48.0
    assert hrs.taux_horaire(CLE_STATS_M1, tarifs, ecoles=ECOLES) == 45.0
    with pytest.raises(ValueError, match="introuvable"):
        hrs.taux_horaire("ABC|Projet fictif|M1", tarifs)


def test_t_hrs_49_libelle_inconnu_refuse():
    """T-HRS-49 — § 15 : une ligne de `Tarifs` dont l'école est un libellé inconnu de `Écoles`, et qui
    porte le module et la classe de la clé, lève `ValueError` qui cite l'onglet `Tarifs`, la ligne
    (son module), la valeur et l'onglet `Écoles` : le taux n'est ni deviné ni « introuvable »."""
    inconnue = {"École": "École Fictive Inconnue", "Module / Matière": "Projet fictif",
                "Classe / Programme": "M1", "Taux HT/h": 50, "Notes": "", "Réf. contrat": ""}
    with pytest.raises(ValueError) as erreur:
        _hrs().taux_horaire("ABC|Projet fictif|M1", TARIFS + [inconnue], ecoles=ECOLES)
    message = str(erreur.value)
    for attendu in ("Tarifs", "Projet fictif", "École Fictive Inconnue", "Écoles"):
        assert attendu in message, attendu


# ──────────────────────────────────────────────────────────────────────
# Revue de P-0b (T-HRS-50, T-HRS-51)
# ──────────────────────────────────────────────────────────────────────
#
# Constats des deux relecteurs, écrits rouges avant le correctif :
# - T-HRS-50 : avec `ecoles=`, la colonne `École` de `Tarifs` était traduite en code, mais pas le
#   segment école de la clé : une clé écrite avec le libellé (« École Fictive A|… ») était déclarée
#   « introuvable » alors qu'une ligne la porte mot pour mot, et elle se trouve sans `ecoles=`. Le
#   § 15 (« la correspondance d'une école accepte le code ou le libellé ») vaut pour les deux côtés ;
# - T-HRS-51 : `arrondi_centime` arrondissait un `float` sur sa valeur binaire (2.675 → 2,67 au lieu
#   de 2,68), sans erreur. Le § 1.2 veut des montants en `Decimal`. Lecture retenue (`💡 [proposé —
#   à valider]`) : un `float` lève `TypeError`, comme `montant` refuse un type inattendu ; un entier
#   (la somme d'une liste vide vaut 0) reste accepté.


def test_t_hrs_50_cle_qui_nomme_l_ecole_par_son_libelle():
    """T-HRS-50 — avec `ecoles=`, le segment école de la clé peut aussi porter le libellé de l'école : il
    trouve la ligne de `Tarifs` écrite avec le libellé comme celle écrite avec le code, sans la
    confondre avec celle d'une autre école."""
    hrs = _hrs()
    tarifs = TARIFS + TARIFS_LIBELLES
    assert hrs.taux_horaire("École Fictive A|Projet fictif|M1", tarifs) == 55.0  # sans `ecoles=` : mot pour mot
    assert hrs.taux_horaire("École Fictive A|Projet fictif|M1", tarifs, ecoles=ECOLES) == 55.0
    assert hrs.taux_horaire("Institut Fictif B|Projet fictif|M1", tarifs, ecoles=ECOLES) == 48.0
    ligne = hrs.ligne_tarif(" école fictive a |Statistiques fictives|M1", tarifs, ecoles=ECOLES)
    assert (ligne["École"], ligne["Réf. contrat"]) == ("ABC", "CT-FICTIF-01")


@pytest.mark.parametrize("valeur", [2.675, 59.325, 0.1, 157.5])
def test_t_hrs_51_arrondi_centime_refuse_un_float(valeur):
    """T-HRS-51 — `arrondi_centime` lève `TypeError` pour un `float` (💡) : un montant arrondi sur la
    valeur binaire d'un float peut être faux d'un centime sans alerte ; le même nombre en `Decimal`
    (écrit comme il s'affiche) est arrondi, et un entier reste accepté."""
    hrs = _hrs()
    with pytest.raises(TypeError):
        hrs.arrondi_centime(valeur)
    attendu = Decimal(repr(valeur)).quantize(Decimal("0.01"), rounding="ROUND_HALF_UP")
    assert hrs.arrondi_centime(Decimal(repr(valeur))) == attendu
    assert hrs.arrondi_centime(0) == Decimal("0.00")
