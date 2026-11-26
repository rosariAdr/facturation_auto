"""Règles de facturation par école (F-14, ADR015, ADR026, ADR027, ADR030 ; banque de tests : T-REG-01 à T-REG-29).

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
  un cycle sans début, ou avec un début hors de 1 à 28, lève `ValueError` (au chargement ou
  au calcul de la période) ; les débuts 1 et 29 à 31 ne sont pas testés (voir le rapport) ;
- `Date limite d'envoi` reste un texte libre (« 15 à 12 h ») ;
- les champs sans défaut dans le contrat (contrat, canal, pièces jointes, notes, statut,
  délai, date limite) valent "" ou None quand la cellule est vide ;
- `ecoles_du_lot` reçoit le dictionnaire rendu par `regles_par_ecole` et compare le statut
  sans tenir compte de la casse ni des espaces ;
- un mois hors de 1 à 12 lève `ValueError` dont le message dit « mois ».
"""

from __future__ import annotations

import calendar
import dataclasses
import importlib
import re
from datetime import date, datetime, timedelta
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
    return regles[ligne["École"].strip().upper()]


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
    assert list(regles) == ["ABC"]
    regle = regles["ABC"]
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
    assert sorted(regles) == ["ABC", "DEF"]
    assert regles["DEF"].ecole == "DEF"
    assert regles["DEF"].statut_lot == "Plus tard"


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
    assert sorted(regles) == ["ABC", "EC2"]
    assert regles["ABC"].ecole == "ABC"
    assert regles["EC2"].ecole == "EC2"


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
    """T-REG-16 — un cycle sans début, ou au début hors de 1 à 28, lève `ValueError` qui nomme le
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
    « 45 jours fin de mois » → fin du mois d'émission + 45 jours."""
    m = _module()
    regles = m.regles_par_ecole([_ligne("ABC"), _ligne("DEF", delai="45 jours fin de mois")])
    emission = date(2026, 12, 1)
    assert m.echeance_selon_regle(emission, regles["ABC"].delai_paiement) == date(2026, 12, 31)
    assert m.echeance_selon_regle(emission, regles["DEF"].delai_paiement) == date(2027, 1, 31)


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
    assert regles["ABC"].une_facture_par_contrat is True
    assert regles["ABC"].n_contrat == "CTR-FICTIF-001"
    assert regles["DEF"].une_facture_par_contrat is False
