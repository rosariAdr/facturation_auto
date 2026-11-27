"""Règles de facturation par école (F-14, ADR015, ADR026, ADR027, ADR030 ; banque de tests : T-REG-01 à T-REG-49).

Tests d'acceptation écrits avant l'implémentation : chaque test importe `facturation.regles`
lui-même, pour que l'absence du module fasse échouer les tests sans bloquer la collecte des
autres. Données fictives uniquement : codes école `ABC`, `DEF`…, contrats `CTR-FICTIF-…`,
bons de commande `BC-FICTIF-…`. Les lignes de `Règles_facturation` sont construites à la main,
sous la forme que rend `facturation.master` (un dictionnaire {en-tête : valeur} par ligne,
toutes les colonnes de MASTER_SCHEMA présentes, cellule vide = None) ; les séances sont des
`facturation.agenda.Seance` construites à la main.

Lectures retenues là où le contrat de la vague 2 ne tranche pas (`💡 [proposé — à valider]`) :
- un texte est lu sans les espaces de début et de fin ; une cellule faite d'espaces est vide ;
- le code école sert de clé, en majuscules (`" abc "` → `ABC`, comme les codes de F-15) ;
  deux lignes dont les codes ne diffèrent que par la casse ou les espaces sont un doublon ;
  une ligne remplie sans code école est refusée (`ValueError` qui nomme la colonne « École ») ;
- `Rythme`, `Colonne Jour` et `Regroupement` sont lus sans tenir compte de la casse ni des
  espaces et rangés en minuscules ; une valeur hors liste lève `ValueError` qui la nomme
  (au chargement pour `Colonne Jour` et `Regroupement` ; au chargement ou au calcul de la
  période pour `Rythme`) ;
- oui/non : `oui`/`non` dans toute casse, ou un booléen Excel ; vide → False ; toute autre
  valeur → `ValueError` qui nomme la colonne, l'école et la valeur ;
- `Début de cycle` est un jour du mois, lu comme entier (15, 15.0, "15") ; vide → None ;
  un cycle sans début, ou avec un début hors de 2 à 28, lève `ValueError` (au chargement ou
  au calcul de la période) ; les bornes 1 et 29 à 31 sont gardées par T-REG-34 ;
- `Date limite d'envoi` reste un texte libre (« 15 à 12 h ») ;
- les champs sans défaut dans le contrat (contrat, canal, pièces jointes, notes, statut,
  délai, date limite) valent "" ou None quand la cellule est vide ;
- `ecoles_du_lot` reçoit le dictionnaire rendu par `regles_par_ecole` et compare le statut
  sans tenir compte de la casse ni des espaces ;
- un mois hors de 1 à 12 lève `ValueError` dont le message dit « mois ».

Constats de la revue de F-14 (T-REG-30 à T-REG-44) : les tests marqués `💡 [proposé — à valider]`
fixent le minimum attendu sans choisir la forme de la correction. T-REG-31 est réservé aux écoles
à plusieurs contrats : le contrat de la vague (« une école en double → ValueError ») contredit
MASTER_SCHEMA, ADR030 §3 et le critère F-14 (« une facture par contrat ») ; Adrian tranche.

Révision de P-0b (T-REG-31, T-REG-45 à T-REG-48) : la décision Q-B1 (contrats des vagues 3 à 5,
§ 16) admet une ligne par contrat ; « la clé devient (école, n° de contrat) ». `regles_par_ecole`
rend donc un dictionnaire indexé par le couple (code école, n° de contrat), le n° valant "" quand
la cellule est vide. Les tests qui indexaient le résultat par le seul code école (T-REG-02, 03,
09, 24, 29f, 32, 40b, 42) sont réécrits avec cette clé, sans autre changement : ils avaient raison
sous le contrat de la vague 2, que le § 16 remplace. Une école en double sans contrat (T-REG-10)
reste refusée.
"""

from __future__ import annotations

import calendar
import dataclasses
import importlib
import re
from datetime import date, datetime, time, timedelta
from zoneinfo import ZoneInfo

import pytest

from facturation.agenda import Seance

pytestmark = pytest.mark.unit

PARIS = ZoneInfo("Europe/Paris")

# Colonnes de l'onglet `Règles_facturation` (MASTER_SCHEMA), dans l'ordre.
COLONNES = (
    "École", "Statut lot", "Rythme", "Début de cycle", "Date limite d'envoi", "Délai de paiement",
    "N° de contrat", "Une facture par contrat", "Bon de commande requis", "Colonne Jour",
    "Regroupement", "Canal d'envoi", "Pièces jointes", "Notes",
)
# Champs de `RegleEcole`, dans l'ordre du contrat (un par colonne ci-dessus).
CHAMPS = (
    "ecole", "statut_lot", "rythme", "debut_cycle", "date_limite", "delai_paiement", "n_contrat",
    "une_facture_par_contrat", "bc_requis", "colonne_jour", "regroupement", "canal",
    "pieces_jointes", "notes",
)
VIDES = (None, "", "   ")


def _module():
    """Module testé, importé dans le test : son absence fait échouer le test, pas la collecte."""
    return importlib.import_module("facturation.regles")


def _ligne(ecole="ABC", **valeurs) -> dict:
    """Ligne fictive de `Règles_facturation` : toutes les colonnes, vides sauf celles données.

    Les arguments nomment les colonnes sans accents ni espaces : `rythme="cycle"`,
    `debut_cycle=15`, `bc="oui"`… (voir `_ALIAS`).
    """
    ligne = dict.fromkeys(COLONNES)
    ligne["École"] = ecole
    for alias, valeur in valeurs.items():
        ligne[_ALIAS[alias]] = valeur
    return ligne


_ALIAS = {
    "statut": "Statut lot",
    "rythme": "Rythme",
    "debut_cycle": "Début de cycle",
    "date_limite": "Date limite d'envoi",
    "delai": "Délai de paiement",
    "contrat": "N° de contrat",
    "par_contrat": "Une facture par contrat",
    "bc": "Bon de commande requis",
    "jour": "Colonne Jour",
    "regroupement": "Regroupement",
    "canal": "Canal d'envoi",
    "pj": "Pièces jointes",
    "notes": "Notes",
}


def _regle(**valeurs):
    """La règle de l'école `ABC` (ou de l'école donnée) lue depuis une seule ligne."""
    ligne = _ligne(**valeurs)
    regles = _module().regles_par_ecole([ligne])
    (regle,) = regles.values()
    return regle


def _seance(jour: date, debut_h: int = 9, fin_h: int = 12) -> Seance:
    """Séance fictive de l'agenda, construite à la main."""
    debut = datetime(jour.year, jour.month, jour.day, debut_h, tzinfo=PARIS)
    fin = datetime(jour.year, jour.month, jour.day, fin_h, tzinfo=PARIS)
    return Seance(
        uid=f"fictif-{jour.isoformat()}@example.org",
        date=jour,
        debut=debut,
        fin=fin,
        duree_h=float(fin_h - debut_h),
        intitule="Cours fictif",
        description="#ABC-M1-STATS",
    )


def _dans(periode: tuple[date, date], seances: list[Seance]) -> list[date]:
    debut, fin = periode
    return [s.date for s in seances if debut <= s.date <= fin]


def _dernier_jour(annee: int, mois: int) -> date:
    return date(annee, mois, calendar.monthrange(annee, mois)[1])


def _mois_suivant(annee: int, mois: int) -> tuple[int, int]:
    return (annee + 1, 1) if mois == 12 else (annee, mois + 1)


# ──────────────────────────────────────────────────────────────────────
# RegleEcole et regles_par_ecole
# ──────────────────────────────────────────────────────────────────────


def test_t_reg_01_regle_ecole_dataclass_figee_aux_champs_du_contrat():
    """T-REG-01 — `RegleEcole` est une dataclass figée, aux champs du contrat, dans l'ordre."""
    m = _module()
    assert dataclasses.is_dataclass(m.RegleEcole)
    assert tuple(f.name for f in dataclasses.fields(m.RegleEcole)) == CHAMPS
    regle = _regle(statut="Lot 1")
    assert isinstance(regle, m.RegleEcole)
    with pytest.raises(dataclasses.FrozenInstanceError):
        regle.rythme = "cycle"


def test_t_reg_02_ligne_complete_lue_champ_par_champ():
    """T-REG-02 — une ligne entièrement remplie : chaque colonne arrive dans son champ, sans espaces autour."""
    ligne = _ligne(
        " ABC ",
        statut=" Lot 1 ",
        rythme="cycle",
        debut_cycle=15,
        date_limite=" 15 à 12 h ",
        delai=" 45 jours fin de mois ",
        contrat=" CTR-FICTIF-001 ",
        par_contrat="oui",
        bc="non",
        jour="masquée",
        regroupement="jour",
        canal=" Portail fournisseur ",
        pj=" PDF + récapitulatif ",
        notes=" Mention fictive demandée ",
    )
    regles = _module().regles_par_ecole([ligne])
    assert isinstance(regles, dict)
    assert list(regles) == [("ABC", "CTR-FICTIF-001")]
    regle = regles[("ABC", "CTR-FICTIF-001")]
    assert (regle.ecole, regle.statut_lot, regle.rythme, regle.debut_cycle) == ("ABC", "Lot 1", "cycle", 15)
    assert (regle.date_limite, regle.delai_paiement, regle.n_contrat) == (
        "15 à 12 h", "45 jours fin de mois", "CTR-FICTIF-001",
    )
    assert regle.une_facture_par_contrat is True
    assert regle.bc_requis is False
    assert (regle.colonne_jour, regle.regroupement) == ("masquée", "jour")
    assert (regle.canal, regle.pieces_jointes, regle.notes) == (
        "Portail fournisseur", "PDF + récapitulatif", "Mention fictive demandée",
    )


def test_t_reg_03_plusieurs_ecoles_et_liste_vide():
    """T-REG-03 — une entrée par école, indexée par son code ; aucune ligne → dictionnaire vide."""
    m = _module()
    assert m.regles_par_ecole([]) == {}
    regles = m.regles_par_ecole([_ligne("ABC", statut="Lot 1"), _ligne("DEF", statut="Plus tard")])
    assert sorted(regles) == [("ABC", ""), ("DEF", "")]
    assert regles[("DEF", "")].ecole == "DEF"
    assert regles[("DEF", "")].statut_lot == "Plus tard"


@pytest.mark.parametrize("vide", VIDES, ids=["None", "chaine-vide", "espaces"])
def test_t_reg_04_valeurs_vides_prennent_les_defauts(vide):
    """T-REG-04 — cellules vides : rythme `mensuel`, colonne Jour `affichée`, regroupement `séance`,
    oui/non à False ; les champs sans défaut restent vides."""
    regle = _regle(
        statut=vide, rythme=vide, debut_cycle=vide, date_limite=vide, delai=vide, contrat=vide,
        par_contrat=vide, bc=vide, jour=vide, regroupement=vide, canal=vide, pj=vide, notes=vide,
    )
    assert regle.rythme == "mensuel"
    assert regle.colonne_jour == "affichée"
    assert regle.regroupement == "séance"
    assert regle.une_facture_par_contrat is False
    assert regle.bc_requis is False
    assert regle.debut_cycle in (None, "")
    for champ in ("statut_lot", "date_limite", "delai_paiement", "n_contrat", "canal", "pieces_jointes", "notes"):
        assert getattr(regle, champ) in ("", None), champ


@pytest.mark.parametrize(
    "valeur, attendu",
    [("oui", True), ("Oui", True), (" OUI ", True), ("non", False), ("Non ", False), ("NON", False),
     (True, True), (False, False)],
    ids=["oui", "Oui", "OUI-espaces", "non", "Non-espace", "NON", "booleen-vrai", "booleen-faux"],
)
def test_t_reg_05_oui_non_lus_en_booleens(valeur, attendu):
    """T-REG-05 — `Une facture par contrat` et `Bon de commande requis` : oui/non dans toute casse,
    ou booléen Excel, deviennent des booléens."""
    regle = _regle(par_contrat=valeur, bc=valeur)
    assert regle.une_facture_par_contrat is attendu
    assert regle.bc_requis is attendu


@pytest.mark.parametrize("colonne", ["par_contrat", "bc"])
@pytest.mark.parametrize("valeur", ["peut-être", "oui et non", "1 sur 2"])
def test_t_reg_06_oui_non_invalide_refuse(colonne, valeur):
    """T-REG-06 — une autre valeur qu'oui/non est refusée : `ValueError` qui nomme la colonne,
    l'école et la valeur (jamais un booléen deviné)."""
    with pytest.raises(ValueError) as erreur:
        _regle(**{colonne: valeur})
    message = str(erreur.value)
    assert _ALIAS[colonne] in message
    assert "ABC" in message
    assert valeur in message


@pytest.mark.parametrize(
    "colonne, saisie, champ, attendu",
    [
        ("rythme", "MENSUEL", "rythme", "mensuel"),
        ("rythme", " Cycle ", "rythme", "cycle"),
        ("rythme", "Fin de module", "rythme", "fin de module"),
        ("jour", "Masquée", "colonne_jour", "masquée"),
        ("jour", " AFFICHÉE ", "colonne_jour", "affichée"),
        ("regroupement", " Jour ", "regroupement", "jour"),
        ("regroupement", "SÉANCE", "regroupement", "séance"),
    ],
)
def test_t_reg_07_valeurs_de_liste_sans_casse_ni_espaces(colonne, saisie, champ, attendu):
    """T-REG-07 — `Rythme`, `Colonne Jour`, `Regroupement` : lus sans tenir compte de la casse ni des
    espaces, rangés en minuscules."""
    valeurs = {colonne: saisie}
    if attendu == "cycle":
        valeurs["debut_cycle"] = 15
    assert getattr(_regle(**valeurs), champ) == attendu


@pytest.mark.parametrize(
    "colonne, valeur",
    [("rythme", "trimestriel"), ("rythme", "mensuelle"), ("jour", "cachée"), ("regroupement", "semaine")],
)
def test_t_reg_08_valeur_hors_liste_refusee(colonne, valeur):
    """T-REG-08 — une valeur hors liste (`Rythme`, `Colonne Jour`, `Regroupement`) lève `ValueError`
    qui la nomme, avec l'école ; pour `Rythme`, au chargement ou au plus tard au calcul de la période."""
    m = _module()
    with pytest.raises(ValueError) as erreur:
        regle = _regle(**{colonne: valeur})
        m.periode_a_facturer(regle, 2026, 11)
    message = str(erreur.value)
    assert valeur in message
    assert "ABC" in message


def test_t_reg_09_code_ecole_sans_espaces_et_en_majuscules():
    """T-REG-09 — le code école sert de clé, sans espaces autour et en majuscules."""
    regles = _module().regles_par_ecole([_ligne(" abc ", statut="Lot 1"), _ligne("Ec2")])
    assert sorted(regles) == [("ABC", ""), ("EC2", "")]
    assert regles[("ABC", "")].ecole == "ABC"
    assert regles[("EC2", "")].ecole == "EC2"


@pytest.mark.parametrize("second", ["ABC", " ABC ", "abc"], ids=["identique", "espaces", "minuscules"])
def test_t_reg_10_ecole_en_double_refusee(second):
    """T-REG-10 — une école en double (y compris à la casse ou aux espaces près) lève `ValueError`
    qui nomme l'école."""
    lignes = [_ligne("ABC", statut="Lot 1"), _ligne(second, statut="Plus tard")]
    with pytest.raises(ValueError, match=r"(?i)double|plusieurs fois|deux fois|déjà") as erreur:
        _module().regles_par_ecole(lignes)
    assert "ABC" in str(erreur.value)


@pytest.mark.parametrize("vide", VIDES, ids=["None", "chaine-vide", "espaces"])
def test_t_reg_11_ligne_sans_code_ecole_refusee(vide):
    """T-REG-11 — une ligne remplie sans code école est refusée (`ValueError` qui nomme la colonne
    « École »), plutôt que perdue en silence."""
    with pytest.raises(ValueError, match="École"):
        _module().regles_par_ecole([_ligne("ABC"), _ligne(vide, statut="Lot 1", rythme="mensuel")])


@pytest.mark.parametrize("saisie", [15, 15.0, "15", " 15 "], ids=["entier", "decimal", "texte", "texte-espaces"])
def test_t_reg_12_debut_de_cycle_lu_comme_jour_entier(saisie):
    """T-REG-12 — `Début de cycle` est un jour du mois, lu comme entier quelle que soit la forme de la cellule."""
    regle = _regle(rythme="cycle", debut_cycle=saisie)
    assert regle.debut_cycle == 15
    assert isinstance(regle.debut_cycle, int)
    assert _module().periode_a_facturer(regle, 2026, 11) == (date(2026, 10, 15), date(2026, 11, 14))


# ──────────────────────────────────────────────────────────────────────
# periode_a_facturer
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    "annee, mois, attendu",
    [
        (2026, 11, (date(2026, 11, 1), date(2026, 11, 30))),
        (2026, 12, (date(2026, 12, 1), date(2026, 12, 31))),
        (2027, 2, (date(2027, 2, 1), date(2027, 2, 28))),
        (2028, 2, (date(2028, 2, 1), date(2028, 2, 29))),
        (2027, 4, (date(2027, 4, 1), date(2027, 4, 30))),
    ],
    ids=["novembre", "decembre", "fevrier", "fevrier-bissextile", "avril"],
)
def test_t_reg_13_mensuel_du_premier_au_dernier_jour(annee, mois, attendu):
    """T-REG-13 — rythme `mensuel` : du 1er au dernier jour du mois, `Début de cycle` ignoré."""
    m = _module()
    assert m.periode_a_facturer(_regle(rythme="mensuel"), annee, mois) == attendu
    assert m.periode_a_facturer(_regle(rythme="mensuel", debut_cycle=15), annee, mois) == attendu


@pytest.mark.parametrize(
    "debut_cycle, annee, mois, attendu",
    [
        (15, 2026, 11, (date(2026, 10, 15), date(2026, 11, 14))),
        (15, 2027, 1, (date(2026, 12, 15), date(2027, 1, 14))),
        (10, 2026, 12, (date(2026, 11, 10), date(2026, 12, 9))),
        (2, 2027, 3, (date(2027, 2, 2), date(2027, 3, 1))),
        (28, 2028, 3, (date(2028, 2, 28), date(2028, 3, 27))),
    ],
    ids=["15-novembre", "15-janvier-changement-d-annee", "10-decembre", "2-mars", "28-mars-bissextile"],
)
def test_t_reg_14_cycle_du_jour_d_du_mois_precedent_a_la_veille(debut_cycle, annee, mois, attendu):
    """T-REG-14 — rythme `cycle`, début d : du d du mois précédent au d-1 du mois (cycle qui se termine dans le mois)."""
    regle = _regle(rythme="cycle", debut_cycle=debut_cycle)
    assert _module().periode_a_facturer(regle, annee, mois) == attendu


@pytest.mark.parametrize(
    "valeurs",
    [{"rythme": "mensuel"}, {"rythme": None}, {"rythme": "cycle", "debut_cycle": 2},
     {"rythme": "cycle", "debut_cycle": 15}, {"rythme": "cycle", "debut_cycle": 28}],
    ids=["mensuel", "rythme-vide", "cycle-2", "cycle-15", "cycle-28"],
)
def test_t_reg_15_periodes_successives_sans_trou_ni_chevauchement(valeurs):
    """T-REG-15 — sur treize mois, chaque période commence le lendemain de la précédente :
    aucune séance oubliée, aucune facturée deux fois."""
    m = _module()
    regle = _regle(**valeurs)
    annee, mois = 2027, 1
    periode = m.periode_a_facturer(regle, annee, mois)
    for _ in range(12):
        debut, fin = periode
        assert debut <= fin
        annee, mois = _mois_suivant(annee, mois)
        periode = m.periode_a_facturer(regle, annee, mois)
        assert periode[0] == fin + timedelta(days=1), (annee, mois)


@pytest.mark.parametrize("debut_cycle", [None, "", 0, 32, "quinze"], ids=["None", "vide", "zero", "32", "texte"])
def test_t_reg_16_cycle_sans_debut_valide_refuse(debut_cycle):
    """T-REG-16 — un cycle sans début, ou au début hors de 2 à 28, lève `ValueError` qui nomme le
    début de cycle et l'école (au chargement ou au calcul de la période) : jamais de cycle deviné."""
    m = _module()
    with pytest.raises(ValueError) as erreur:
        regle = _regle(rythme="cycle", debut_cycle=debut_cycle)
        m.periode_a_facturer(regle, 2026, 11)
    message = str(erreur.value)
    assert re.search(r"(?i)début de cycle", message)
    assert "ABC" in message


def test_t_reg_17_fin_de_module_pas_encore_prise_en_charge():
    """T-REG-17 — rythme `fin de module` : la règle se charge, mais le calcul de la période lève
    `NotImplementedError` explicite (critère F-14 : facturée seulement à la fin, à venir)."""
    regle = _regle(rythme="fin de module")
    assert regle.rythme == "fin de module"
    with pytest.raises(NotImplementedError, match=r"(?i)fin de module"):
        _module().periode_a_facturer(regle, 2026, 11)


@pytest.mark.parametrize("mois", [0, 13])
@pytest.mark.parametrize("valeurs", [{"rythme": "mensuel"}, {"rythme": "cycle", "debut_cycle": 15}], ids=["mensuel", "cycle"])
def test_t_reg_18_mois_invalide_refuse(valeurs, mois):
    """T-REG-18 — un mois hors de 1 à 12 lève `ValueError` dont le message parle du mois."""
    with pytest.raises(ValueError, match=r"(?i)mois"):
        _module().periode_a_facturer(_regle(**valeurs), 2026, mois)


# ──────────────────────────────────────────────────────────────────────
# echeance_selon_regle
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize("delai", [None, "", "   "], ids=["None", "vide", "espaces"])
@pytest.mark.parametrize(
    "emission, attendu",
    [(date(2026, 12, 1), date(2026, 12, 31)), (date(2026, 12, 15), date(2027, 1, 14)),
     (date(2027, 2, 1), date(2027, 3, 3))],
)
def test_t_reg_19_sans_delai_emission_plus_30_jours(delai, emission, attendu):
    """T-REG-19 — délai vide : date d'émission + 30 jours (ADR030)."""
    assert _module().echeance_selon_regle(emission, delai) == attendu


@pytest.mark.parametrize(
    "delai, emission, attendu",
    [
        ("30 jours", date(2026, 12, 1), date(2026, 12, 31)),
        ("45 jours", date(2026, 12, 1), date(2027, 1, 15)),
        ("60 jours", date(2026, 12, 1), date(2027, 1, 30)),
        ("30 jours", date(2026, 11, 20), date(2026, 12, 20)),
        (" 45 Jours ", date(2026, 12, 1), date(2027, 1, 15)),
        ("30 JOURS", date(2027, 2, 1), date(2027, 3, 3)),
    ],
)
def test_t_reg_20_n_jours(delai, emission, attendu):
    """T-REG-20 — « N jours » : émission + N jours, sans tenir compte de la casse ni des espaces."""
    assert _module().echeance_selon_regle(emission, delai) == attendu


@pytest.mark.parametrize(
    "delai, emission, attendu",
    [
        ("45 jours fin de mois", date(2026, 12, 1), date(2027, 1, 31)),
        ("45 jours fin de mois", date(2026, 11, 20), date(2027, 1, 31)),
        ("45 jours fin de mois", date(2026, 11, 1), date(2026, 12, 31)),
        ("45 jours fin de mois", date(2027, 1, 1), date(2027, 2, 28)),
        ("45 jours fin de mois", date(2028, 1, 15), date(2028, 2, 29)),
        ("30 jours fin de mois", date(2026, 12, 1), date(2026, 12, 31)),
        (" 45 JOURS FIN DE MOIS ", date(2026, 12, 1), date(2027, 1, 31)),
    ],
)
def test_t_reg_21_n_jours_fin_de_mois(delai, emission, attendu):
    """T-REG-21 — « N jours fin de mois » : dernier jour du mois de (émission + N jours)."""
    assert _module().echeance_selon_regle(emission, delai) == attendu


@pytest.mark.parametrize(
    "delai, emission, attendu",
    [
        ("fin de mois", date(2026, 12, 1), date(2026, 12, 31)),
        ("fin de mois", date(2027, 2, 1), date(2027, 2, 28)),
        ("fin de mois", date(2028, 2, 1), date(2028, 2, 29)),
        ("fin de mois", date(2027, 1, 31), date(2027, 1, 31)),
        (" Fin de mois ", date(2026, 11, 1), date(2026, 11, 30)),
    ],
)
def test_t_reg_22_fin_de_mois(delai, emission, attendu):
    """T-REG-22 — « fin de mois » : dernier jour du mois d'émission."""
    assert _module().echeance_selon_regle(emission, delai) == attendu


@pytest.mark.parametrize(
    "delai",
    ["à réception", "trente jours", "jours", "30 jours fin de semaine", "fin de trimestre", "payable fin de mois"],
)
def test_t_reg_23_delai_illisible_refuse(delai):
    """T-REG-23 — tout autre texte lève `ValueError` qui le cite : jamais d'échéance devinée
    (« payable fin de mois », écarté par ADR030, n'est pas lu comme « fin de mois »)."""
    with pytest.raises(ValueError, match=re.escape(delai)):
        _module().echeance_selon_regle(date(2026, 12, 1), delai)


def test_t_reg_24_echeance_depuis_la_regle_lue():
    """T-REG-24 — l'échéance se calcule depuis le délai de la règle lue : vide → + 30 jours,
    « 45 jours fin de mois » → fin du mois de (émission + 45 jours), comme T-REG-21."""
    m = _module()
    regles = m.regles_par_ecole([_ligne("ABC"), _ligne("DEF", delai="45 jours fin de mois")])
    emission = date(2026, 12, 1)
    assert m.echeance_selon_regle(emission, regles[("ABC", "")].delai_paiement) == date(2026, 12, 31)
    assert m.echeance_selon_regle(emission, regles[("DEF", "")].delai_paiement) == date(2027, 1, 31)


# ──────────────────────────────────────────────────────────────────────
# ecoles_du_lot
# ──────────────────────────────────────────────────────────────────────


def test_t_reg_25_ecoles_du_lot_triees():
    """T-REG-25 — codes des écoles dont le statut vaut le lot demandé (« Lot 1 » par défaut), triés ;
    statut comparé sans tenir compte de la casse ni des espaces ; statut vide : dans aucun lot."""
    m = _module()
    regles = m.regles_par_ecole([
        _ligne("MNO", statut=None),
        _ligne("GHI", statut="Plus tard"),
        _ligne("DEF", statut=" lot 1 "),
        _ligne("JKL", statut="Hors facturation"),
        _ligne("ABC", statut="Lot 1"),
        _ligne("AAA", statut="LOT 1"),
    ])
    assert m.ecoles_du_lot(regles) == ["AAA", "ABC", "DEF"]
    assert m.ecoles_du_lot(regles, lot="Lot 1") == ["AAA", "ABC", "DEF"]
    assert m.ecoles_du_lot(regles, lot="Plus tard") == ["GHI"]
    assert m.ecoles_du_lot(regles, lot="Hors facturation") == ["JKL"]
    assert m.ecoles_du_lot(regles, lot="Lot 2") == []
    assert m.ecoles_du_lot({}) == []


# ──────────────────────────────────────────────────────────────────────
# blocages
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize("ref", [None, "", "   "], ids=["None", "vide", "espaces"])
def test_t_reg_26_bon_de_commande_requis_et_absent_bloque(ref):
    """T-REG-26 — bon de commande requis et absent : un blocage qui le dit."""
    resultat = _module().blocages(_regle(bc="oui"), ref_bon_commande=ref)
    assert isinstance(resultat, list)
    assert len(resultat) == 1
    assert "bon de commande" in resultat[0].lower()


@pytest.mark.parametrize(
    "bc, ref",
    [("oui", "BC-FICTIF-0001"), ("non", None), ("non", ""), (None, None), ("non", "BC-FICTIF-0002")],
    ids=["requis-present", "non-requis-absent", "non-requis-vide", "colonne-vide", "non-requis-present"],
)
def test_t_reg_27_rien_ne_bloque(bc, ref):
    """T-REG-27 — bon de commande présent, ou non requis : liste vide."""
    assert _module().blocages(_regle(bc=bc), ref_bon_commande=ref) == []


def test_t_reg_28_reference_du_bon_de_commande_nommee():
    """T-REG-28 — `ref_bon_commande` se passe par son nom (argument nommé seulement)."""
    regle = _regle(bc="oui")
    with pytest.raises(TypeError):
        _module().blocages(regle, "BC-FICTIF-0001")


# ──────────────────────────────────────────────────────────────────────
# Critères d'acceptation de F-14 (references/ROADMAP.md)
# ──────────────────────────────────────────────────────────────────────


def test_t_reg_29a_sans_rythme_facturee_chaque_mois():
    """T-REG-29a — une école sans rythme renseigné est facturée chaque mois : chaque séance tombe
    dans la période de son mois, et d'aucun autre."""
    m = _module()
    regle = _regle(rythme=None)
    seances = [_seance(date(2026, 10, 31)), _seance(date(2026, 11, 1)), _seance(date(2026, 11, 30)),
               _seance(date(2026, 12, 1))]
    assert _dans(m.periode_a_facturer(regle, 2026, 10), seances) == [date(2026, 10, 31)]
    assert _dans(m.periode_a_facturer(regle, 2026, 11), seances) == [date(2026, 11, 1), date(2026, 11, 30)]
    assert _dans(m.periode_a_facturer(regle, 2026, 12), seances) == [date(2026, 12, 1)]
    for mois in range(1, 13):
        assert m.periode_a_facturer(regle, 2027, mois) == (date(2027, mois, 1), _dernier_jour(2027, mois))


def test_t_reg_29b_bons_de_commande_colonne_jour_masquee():
    """T-REG-29b — une école à bons de commande : la règle porte « bon de commande requis » et la
    colonne Jour masquée (la facture par ligne de bon de commande est faite par F-16)."""
    regle = _regle(bc="oui", jour="masquée")
    assert regle.bc_requis is True
    assert regle.colonne_jour == "masquée"
    assert _module().blocages(regle, ref_bon_commande="BC-FICTIF-0001") == []


def test_t_reg_29c_sans_bon_de_commande_preparee_puis_bloquee():
    """T-REG-29c — une école sans bon de commande reçu est préparée (sa période se calcule)
    puis bloquée (un blocage la retient)."""
    m = _module()
    regle = _regle(statut="Lot 1", bc="oui")
    assert m.ecoles_du_lot({"ABC": regle}) == ["ABC"]
    assert m.periode_a_facturer(regle, 2026, 11) == (date(2026, 11, 1), date(2026, 11, 30))
    assert m.blocages(regle, ref_bon_commande=None) != []


def test_t_reg_29d_fin_de_module_pas_facturee_chaque_mois():
    """T-REG-29d — une école en fin de module n'est pas facturée au mois : aucun mois ne donne de
    période (calcul à la fin du module, pas encore pris en charge)."""
    m = _module()
    regle = _regle(rythme="fin de module")
    for mois in (10, 11, 12):
        with pytest.raises(NotImplementedError):
            m.periode_a_facturer(regle, 2026, mois)


def test_t_reg_29e_cycle_propre_prepare_avant_sa_date_limite():
    """T-REG-29e — une école à cycle propre : son cycle de novembre se termine avant sa date limite
    (« 15 à 12 h ») ; la séance du 15/11 relève du cycle suivant."""
    m = _module()
    regle = _regle(rythme="cycle", debut_cycle=15, date_limite="15 à 12 h")
    assert regle.date_limite == "15 à 12 h"
    seances = [_seance(date(2026, 10, 14)), _seance(date(2026, 10, 15)), _seance(date(2026, 11, 14)),
               _seance(date(2026, 11, 15))]
    novembre = m.periode_a_facturer(regle, 2026, 11)
    assert novembre[1] < date(2026, 11, 15)
    assert _dans(novembre, seances) == [date(2026, 10, 15), date(2026, 11, 14)]
    assert _dans(m.periode_a_facturer(regle, 2026, 10), seances) == [date(2026, 10, 14)]
    assert _dans(m.periode_a_facturer(regle, 2026, 12), seances) == [date(2026, 11, 15)]


def test_t_reg_29f_une_facture_par_contrat_quand_la_regle_le_demande():
    """T-REG-29f — une facture par contrat quand la règle le demande : le drapeau et la référence
    du contrat sont portés par la règle ; sans demande, le drapeau est faux."""
    regles = _module().regles_par_ecole([
        _ligne("ABC", par_contrat="oui", contrat="CTR-FICTIF-001"),
        _ligne("DEF", par_contrat="non"),
    ])
    assert regles[("ABC", "CTR-FICTIF-001")].une_facture_par_contrat is True
    assert regles[("ABC", "CTR-FICTIF-001")].n_contrat == "CTR-FICTIF-001"
    assert regles[("DEF", "")].une_facture_par_contrat is False


# ──────────────────────────────────────────────────────────────────────
# Constats de la revue de F-14
# ──────────────────────────────────────────────────────────────────────


def _issue(**valeurs) -> str:
    """« acceptée » si la ligne se lit, « refusée » si sa lecture lève `ValueError`."""
    try:
        _regle(**valeurs)
    except ValueError:
        return "refusée"
    return "acceptée"


@pytest.mark.parametrize("statut", ["Lot1", "Lot 1 (pilote)", "Plus-tard", "Lot 1.", "Lot I"])
def test_t_reg_30_statut_lot_hors_liste_refuse(statut):
    """T-REG-30 — un `Statut lot` hors de la liste de MASTER_SCHEMA (`Lot 1`, `Plus tard`,
    `Hors facturation`) est refusé dès la lecture : `ValueError` qui nomme la colonne, l'école et
    la valeur. Sinon l'école sort du lot sans un mot, et ses séances ne sont jamais facturées.
    La casse et les espaces restent tolérés (T-REG-25). Le sort d'un statut vide n'est pas fixé
    ici : `💡 [proposé — à valider]`."""
    with pytest.raises(ValueError) as erreur:
        _regle(statut=statut)
    message = str(erreur.value)
    assert "Statut lot" in message
    assert "ABC" in message
    assert statut in message


@pytest.mark.parametrize("statut", ["Hors facturation", "Plus tard"])
def test_t_reg_32_erreur_d_une_ecole_hors_lot_ne_retient_pas_le_lot(statut):
    """T-REG-32 — `💡 [proposé — à valider]` une valeur illisible sur la ligne d'une école
    « Hors facturation » ou « Plus tard » ne retient pas les écoles du lot 1 (esprit d'ADR031 :
    bloquée sans retenir les autres). La règle de l'école du lot reste utilisable. La forme du
    signalement de l'école hors lot n'est pas fixée ici. Si l'école hors lot figure dans le
    résultat, sa valeur illisible ne devient jamais une échéance devinée."""
    m = _module()
    regles = m.regles_par_ecole([
        _ligne("XYZ", statut=statut, delai="à négocier"),
        _ligne("ABC", statut="Lot 1"),
    ])
    assert m.ecoles_du_lot(regles) == ["ABC"]
    assert m.periode_a_facturer(regles[("ABC", "")], 2026, 11) == (date(2026, 11, 1), date(2026, 11, 30))
    if ("XYZ", "") in regles:
        with pytest.raises(ValueError):
            m.echeance_selon_regle(date(2026, 12, 1), regles[("XYZ", "")].delai_paiement)


def test_t_reg_33_delai_illisible_refuse_au_chargement():
    """T-REG-33 — un délai de paiement illisible est refusé dès la lecture de la règle, par un
    message qui nomme l'école, la colonne et la valeur (T-REG-23 ne teste que le calcul de
    l'échéance, dont le message ne nomme pas l'école)."""
    with pytest.raises(ValueError) as erreur:
        _regle(delai="à réception")
    message = str(erreur.value)
    assert "ABC" in message
    assert "Délai de paiement" in message
    assert "à réception" in message


@pytest.mark.parametrize("debut_cycle", [1, 29, 30, 31])
def test_t_reg_34_debut_de_cycle_aux_bornes_refuse(debut_cycle):
    """T-REG-34 — début de cycle 1 ou 29 à 31 : refusé au chargement comme au calcul de la période
    (règle construite à la main), par un message qui nomme le début de cycle et l'école. Au 1er,
    la facture de novembre couvrirait octobre ; de 29 à 31, le jour manque certains mois. Au 1er,
    le message propose le rythme « mensuel »."""
    m = _module()
    with pytest.raises(ValueError) as au_chargement:
        _regle(rythme="cycle", debut_cycle=debut_cycle)
    with pytest.raises(ValueError) as au_calcul:
        m.periode_a_facturer(m.RegleEcole("ABC", rythme="cycle", debut_cycle=debut_cycle), 2027, 3)
    for erreur in (au_chargement, au_calcul):
        message = str(erreur.value)
        assert re.search(r"(?i)début de cycle", message)
        assert "ABC" in message
        if debut_cycle == 1:
            assert "mensuel" in message


@pytest.mark.parametrize(
    "delai", ["3650000 jours", "99999999 jours", "99999999999 jours", "3650000 jours fin de mois"],
)
def test_t_reg_35_delai_demesure_refuse(delai):
    """T-REG-35 — un délai qui mène hors du calendrier lève `ValueError` qui le cite, au calcul
    de l'échéance (pas une `OverflowError` muette) et dès la lecture de la règle, avec le nom de
    l'école. Le sort de « 0 jours » n'est pas fixé ici : `💡 [proposé — à valider]`."""
    m = _module()
    with pytest.raises(ValueError, match=re.escape(delai)):
        m.echeance_selon_regle(date(2026, 12, 1), delai)
    with pytest.raises(ValueError) as erreur:
        _regle(delai=delai)
    message = str(erreur.value)
    assert "ABC" in message
    assert "Délai de paiement" in message


@pytest.mark.parametrize("vide", VIDES, ids=["None", "chaine-vide", "espaces"])
def test_t_reg_36_une_facture_par_contrat_sans_numero_signalee(vide):
    """T-REG-36 — `💡 [proposé — à valider]` « Une facture par contrat » à `oui` sans « N° de
    contrat » : la règle se contredit (la facture ne porterait aucune référence de contrat,
    ADR030 §3). Elle est refusée à la lecture (le message nomme la colonne et l'école), ou un
    blocage le dit ; jamais acceptée sans remarque."""
    m = _module()
    try:
        regle = _regle(par_contrat="oui", contrat=vide)
    except ValueError as erreur:
        message = str(erreur)
        assert "N° de contrat" in message
        assert "ABC" in message
        return
    raisons = m.blocages(regle, ref_bon_commande=None)
    assert any("contrat" in raison.lower() for raison in raisons), raisons


def test_t_reg_37_ligne_sans_ecole_retrouvable():
    """T-REG-37 — une ligne remplie sans code école est refusée par un message qui permet de la
    retrouver dans le MASTER : il cite son contenu ou l'école qui la précède. Le rang compté dans
    la liste (« règle n° 2 ») n'est pas le numéro de ligne Excel : l'en-tête est décalé et
    `lire_master` a retiré les lignes vides."""
    with pytest.raises(ValueError, match="École") as erreur:
        _module().regles_par_ecole([_ligne("ABC"), _ligne(None, notes="Note fictive Z9")])
    message = str(erreur.value)
    assert "Note fictive Z9" in message or "ABC" in message, message


def test_t_reg_38_debut_de_cycle_decimal_refuse():
    """T-REG-38 — un début de cycle décimal (15.5) est refusé, jamais tronqué en 15 : `ValueError`
    qui nomme la colonne et la valeur."""
    with pytest.raises(ValueError) as erreur:
        _regle(rythme="cycle", debut_cycle=15.5)
    message = str(erreur.value)
    assert "Début de cycle" in message
    assert "15.5" in message


@pytest.mark.parametrize("nombre, texte", [(40, "quinze"), (-3, "-3")], ids=["40-et-quinze", "moins-3"])
def test_t_reg_39_debut_de_cycle_hors_cycle_lu_de_facon_coherente(nombre, texte):
    """T-REG-39 — `💡 [proposé — à valider]` hors du rythme `cycle`, « Début de cycle » est soit
    contrôlé, soit ignoré, mais de la même façon pour un nombre et pour un texte : aujourd'hui 40
    et -3 passent, alors que « quinze » et « -3 » saisis en texte bloquent la lecture."""
    assert _issue(rythme="mensuel", debut_cycle=nombre) == _issue(rythme="mensuel", debut_cycle=texte)


def test_t_reg_40a_lot_vide_ne_designe_aucune_ecole():
    """T-REG-40a — un lot vide ne désigne aucune école, pas même celles au statut vide."""
    m = _module()
    assert m.ecoles_du_lot(m.regles_par_ecole([_ligne("MNO", statut=None)]), lot="") == []


def test_t_reg_40b_ligne_entierement_vide_ignoree():
    """T-REG-40b — une ligne entièrement vide est ignorée, comme le promet la docstring de
    `regles_par_ecole` (fonction publique : l'appelant ne passe pas forcément par `lire_master`)."""
    assert list(_module().regles_par_ecole([_ligne("ABC"), dict.fromkeys(COLONNES)])) == [("ABC", "")]


@pytest.mark.parametrize("saisie", [12345.0, 12345], ids=["decimal", "entier"])
def test_t_reg_41_numero_de_contrat_saisi_en_nombre(saisie):
    """T-REG-41 — un n° de contrat saisi comme un nombre s'affiche sans « .0 »."""
    assert _regle(contrat=saisie).n_contrat == "12345"


def test_t_reg_42_colonne_du_schema_manquante_refusee():
    """T-REG-42 — `💡 [proposé — à valider]` une colonne de MASTER_SCHEMA absente de la ligne
    (ici mal orthographiée) est refusée par un message qui la nomme, au lieu de prendre sa valeur
    par défaut : « Bon de commande requise » à la place de « Bon de commande requis » donnerait
    « non requis », et la facture partirait sans bon de commande. Une colonne en plus reste
    permise, comme dans `lire_master`."""
    m = _module()
    ligne = _ligne("ABC", statut="Lot 1")
    del ligne["Bon de commande requis"]
    ligne["Bon de commande requise"] = "oui"
    with pytest.raises(ValueError, match="Bon de commande requis"):
        m.regles_par_ecole([ligne])
    assert list(m.regles_par_ecole([{**_ligne("ABC"), "Commentaire": "Note fictive"}])) == [("ABC", "")]


@pytest.mark.parametrize("saisie", [datetime(2026, 11, 15), time(12, 0)], ids=["date", "heure"])
def test_t_reg_43_date_limite_convertie_par_excel(saisie):
    """T-REG-43 — une « Date limite d'envoi » qu'Excel a convertie en date ou en heure n'est pas
    rangée sous sa forme Python (« 2026-11-15 00:00:00 », « 12:00:00 ») : refusée par un message
    qui nomme la colonne et l'école, ou rendue sous une forme lisible (`💡 [proposé — à valider]`)."""
    try:
        regle = _regle(date_limite=saisie)
    except ValueError as erreur:
        message = str(erreur)
        assert "Date limite d'envoi" in message
        assert "ABC" in message
        return
    assert str(saisie) not in regle.date_limite
    assert not re.search(r"\d{1,2}:\d{2}:\d{2}", regle.date_limite), regle.date_limite


def test_t_reg_44_erreurs_de_plusieurs_ecoles_dans_un_seul_message():
    """T-REG-44 — `💡 [proposé — à valider]` deux écoles du lot aux lignes illisibles : une seule
    `ValueError` les cite toutes les deux, pour tout corriger en une fois au lieu de découvrir
    les erreurs une à une, relance après relance."""
    with pytest.raises(ValueError) as erreur:
        _module().regles_par_ecole([
            _ligne("ABC", statut="Lot 1", rythme="trimestriel"),
            _ligne("DEF", statut="Lot 1", bc="peut-être"),
        ])
    message = str(erreur.value)
    for attendu in ("ABC", "trimestriel", "DEF", "peut-être"):
        assert attendu in message


# ──────────────────────────────────────────────────────────────────────
# P-0b — une ligne par contrat (décision Q-B1, contrats des vagues 3 à 5, § 16)
# ──────────────────────────────────────────────────────────────────────
#
# Lectures retenues (`💡 [proposé — à valider]`) : la clé est (code école, n° de contrat sans
# espaces autour), le n° valant "" quand la cellule est vide ; quand une école a plusieurs lignes,
# chacune doit porter un n° de contrat (« à condition qu'elles portent des N° de contrat
# distincts ») : une ligne sans n° à côté d'une ligne avec n° est refusée, comme deux lignes sans
# n° ; deux n° qui ne diffèrent que par les espaces autour sont le même contrat.


def test_t_reg_31_une_ligne_par_contrat():
    """T-REG-31 — une école peut avoir une ligne par contrat, à n° de contrat distincts : chaque règle
    est rangée sous (école, n° de contrat), avec son propre délai et son propre drapeau ; l'école
    n'apparaît qu'une fois dans son lot, et la période se calcule pour chaque règle."""
    m = _module()
    regles = m.regles_par_ecole([
        _ligne("ABC", statut="Lot 1", contrat="CTR-FICTIF-001", par_contrat="oui"),
        _ligne(" abc ", statut="Lot 1", contrat=" CTR-FICTIF-002 ", delai="45 jours fin de mois", par_contrat="oui"),
        _ligne("DEF", statut="Lot 1"),
    ])
    assert sorted(regles) == [("ABC", "CTR-FICTIF-001"), ("ABC", "CTR-FICTIF-002"), ("DEF", "")]
    premier, second = regles[("ABC", "CTR-FICTIF-001")], regles[("ABC", "CTR-FICTIF-002")]
    assert (premier.ecole, premier.n_contrat, premier.delai_paiement) == ("ABC", "CTR-FICTIF-001", "")
    assert (second.ecole, second.n_contrat, second.delai_paiement) == ("ABC", "CTR-FICTIF-002", "45 jours fin de mois")
    assert premier.une_facture_par_contrat and second.une_facture_par_contrat
    assert m.ecoles_du_lot(regles) == ["ABC", "DEF"]
    for regle in (premier, second):
        assert m.periode_a_facturer(regle, 2026, 12) == (date(2026, 12, 1), date(2026, 12, 31))
        assert m.blocages(regle, ref_bon_commande=None) == []


@pytest.mark.parametrize("second", ["CTR-FICTIF-001", " CTR-FICTIF-001 "], ids=["identique", "espaces"])
def test_t_reg_45_meme_contrat_deux_fois_refuse(second):
    """T-REG-45 — une école avec le même n° de contrat sur deux lignes reste refusée : `ValueError` qui
    nomme l'école et le contrat."""
    with pytest.raises(ValueError) as erreur:
        _module().regles_par_ecole([
            _ligne("ABC", statut="Lot 1", contrat="CTR-FICTIF-001"),
            _ligne("ABC", statut="Lot 1", contrat=second, delai="30 jours"),
        ])
    message = str(erreur.value)
    assert "ABC" in message and "CTR-FICTIF-001" in message


@pytest.mark.parametrize(
    "contrats", [(None, "CTR-FICTIF-001"), ("CTR-FICTIF-001", "   ")], ids=["sans-puis-avec", "avec-puis-sans"],
)
def test_t_reg_46_ligne_sans_contrat_a_cote_d_un_contrat_refusee(contrats):
    """T-REG-46 — quand une école a plusieurs lignes, chacune porte son n° de contrat : une ligne sans
    n° à côté d'une ligne avec n° est refusée (`ValueError` qui nomme l'école et la colonne « N° de
    contrat ») ; la règle à appliquer n'est jamais devinée."""
    premier, second = contrats
    with pytest.raises(ValueError) as erreur:
        _module().regles_par_ecole([
            _ligne("ABC", statut="Lot 1", contrat=premier),
            _ligne("ABC", statut="Lot 1", contrat=second),
        ])
    message = str(erreur.value)
    assert "ABC" in message and "N° de contrat" in message


def test_t_reg_47_deux_lignes_sans_contrat_toujours_refusees():
    """T-REG-47 — deux lignes sans n° de contrat pour une même école restent une `ValueError` (§ 16),
    dont le message nomme l'école et la colonne « N° de contrat » : c'est elle qui permet d'avoir une
    ligne par contrat."""
    with pytest.raises(ValueError) as erreur:
        _module().regles_par_ecole([_ligne("ABC", statut="Lot 1"), _ligne("ABC", statut="Plus tard")])
    message = str(erreur.value)
    assert "ABC" in message and "N° de contrat" in message


def test_t_reg_48_contrats_distincts_sans_message_de_doublon():
    """T-REG-48 — trois contrats d'une même école, plus une autre école : aucune erreur, quatre règles,
    une par (école, contrat) (la garde de T-REG-45 ne refuse pas des contrats distincts)."""
    regles = _module().regles_par_ecole([
        _ligne("ABC", contrat="CTR-FICTIF-001"), _ligne("ABC", contrat="CTR-FICTIF-002"),
        _ligne("ABC", contrat=12345.0), _ligne("ECO", contrat="CTR-FICTIF-001"),
    ])
    assert sorted(regles) == [("ABC", "12345"), ("ABC", "CTR-FICTIF-001"), ("ABC", "CTR-FICTIF-002"),
                              ("ECO", "CTR-FICTIF-001")]
    assert all(regle.ecole == cle[0] and regle.n_contrat == cle[1] for cle, regle in regles.items())


# Revue de P-0b (T-REG-49), écrit rouge avant le correctif : `regles_par_ecole` comparait les n° de
# contrat à la casse près (« CTR-FICTIF-001 » et « ctr-fictif-001 » donnaient deux règles, donc deux
# contrats et deux brouillons possibles), alors que le code école se compare sans casse et que
# `heures.ligne_tarif` compare les `Réf. contrat` sans casse (NFC, casefold). Le § 16 refuse une
# école « avec le même contrat deux fois » : une faute de casse ne fait pas un second contrat.


@pytest.mark.parametrize("second", ["ctr-fictif-001", "Ctr-Fictif-001"], ids=["minuscules", "casse-melangee"])
def test_t_reg_49_meme_contrat_a_la_casse_pres_refuse(second):
    """T-REG-49 — une école avec le même n° de contrat à la casse près sur deux lignes est refusée
    comme un doublon : `ValueError` qui nomme l'école et le contrat."""
    with pytest.raises(ValueError) as erreur:
        _module().regles_par_ecole([
            _ligne("ABC", statut="Lot 1", contrat="CTR-FICTIF-001"),
            _ligne("ABC", statut="Lot 1", contrat=second, delai="30 jours"),
        ])
    message = str(erreur.value)
    assert "ABC" in message and re.search(r"(?i)ctr-fictif-001", message), message
