"""Numérotation des factures (F-15, ADR002, ADR029 ; banque de tests : T-NUM-01 à T-NUM-11).

Tests d'acceptation écrits avant l'implémentation : chaque test importe
`facturation.numerotation` lui-même, pour que son absence fasse échouer les tests sans
bloquer la collecte des autres. Données fictives : écoles ABC, ECO, AB et ABCD, préfixes
d'entité R_ et K_.

Format d'un numéro (ADR002, ADR029) : `<préfixe d'entité><CODE ÉCOLE><NNN>`, par exemple
`R_ABC008` avec le préfixe `R_`. Variante historique reconnue : `R_ABC_007`.
"""

from __future__ import annotations

import hashlib
import importlib
from datetime import date

import pytest
from openpyxl import load_workbook

from tests.conftest import ENTETES, creer_master

pytestmark = pytest.mark.unit


def _num():
    """Module testé, importé dans le test : son absence fait échouer le test, pas la collecte."""
    return importlib.import_module("facturation.numerotation")


def _prochain(numeros: list, code_ecole: str = "ABC", prefixe: str = "R_") -> str:
    return _num().prochain_numero(numeros, prefixe=prefixe, code_ecole=code_ecole)


# ──────────────────────────────────────────────────────────────────────
# prochain_numero
# ──────────────────────────────────────────────────────────────────────


def test_t_num_01_prochain_numero():
    """T-NUM-01 — le plus grand numéro de la série plus un, sur 3 chiffres, quels que soient l'ordre et les doublons."""
    assert _prochain(["R_ABC003", "R_ABC007", "R_ABC007", "R_ABC001"]) == "R_ABC008"
    assert _prochain(["R_ABC009"]) == "R_ABC010"
    assert _prochain(["R_ABC099"]) == "R_ABC100"


def test_t_num_02_serie_vide():
    """T-NUM-02 — une série vide commence à 001, même si d'autres écoles ou préfixes ont des numéros."""
    assert _prochain([]) == "R_ABC001"
    assert _prochain(["R_ECO004", "K_ABC002"]) == "R_ABC001"


def test_t_num_03_autres_ecoles_et_prefixes_ignores():
    """T-NUM-03 — seuls comptent le préfixe et le code demandés : AB, ABC et ABCD sont trois séries distinctes."""
    numeros = ["R_ABC002", "R_ECO009", "K_ABC015", "R_ABCD020", "R_AB030"]
    assert _prochain(numeros, "ABC") == "R_ABC003"
    assert _prochain(numeros, "AB") == "R_AB031"
    assert _prochain(numeros, "ABCD") == "R_ABCD021"
    assert _prochain(numeros, "ABC", prefixe="K_") == "K_ABC016"


def test_t_num_04_variante_historique_avec_tiret_bas():
    """T-NUM-04 — `R_ABC_007` compte pour trouver le plus grand ; le numéro suivant est au format courant."""
    assert _prochain(["R_ABC005", "R_ABC_007"]) == "R_ABC008"
    assert _prochain(["R_ABC_012"]) == "R_ABC013"


def test_t_num_05_depassement_de_999():
    """T-NUM-05 — 999 est le dernier numéro possible d'une série : au-delà, erreur."""
    assert _prochain(["R_ABC998"]) == "R_ABC999"
    with pytest.raises(ValueError):
        _prochain(["R_ABC999"])


# ──────────────────────────────────────────────────────────────────────
# numeros_manquants
# ──────────────────────────────────────────────────────────────────────


def test_t_num_06_numeros_manquants():
    """T-NUM-06 — les trous de la série, au format courant et dans l'ordre (ADR029 : à noter, jamais réutilisés)."""
    num = _num()
    numeros = ["R_ABC001", "R_ABC002", "R_ABC004", "R_ABC_005", "R_ABC008", "R_ABC008", "R_ECO003", "K_ABC006"]
    assert num.numeros_manquants(numeros, prefixe="R_", code_ecole="ABC") == ["R_ABC003", "R_ABC006", "R_ABC007"]
    assert num.numeros_manquants(["R_ABC001", "R_ABC002"], prefixe="R_", code_ecole="ABC") == []
    assert num.numeros_manquants([], prefixe="R_", code_ecole="ABC") == []


# ──────────────────────────────────────────────────────────────────────
# valider_code_ecole
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize("code", ["ABC", "ECO", "AB", "ABCD", "EC2"])
def test_t_num_07_code_existant_accepte(code):
    """T-NUM-07 — un code existant : majuscules et chiffres, quelle que soit sa longueur (ADR029)."""
    _num().valider_code_ecole(code)


@pytest.mark.parametrize("code", ["abc", "AB-C", "AB C", "", "AB_C"])
def test_t_num_08_code_mal_forme_refuse(code):
    """T-NUM-08 — minuscules, tiret, espace, tiret bas ou code vide : refusés."""
    with pytest.raises(ValueError):
        _num().valider_code_ecole(code)


@pytest.mark.parametrize(("code", "accepte"), [("ABC", True), ("AB", False), ("ABCD", False), ("AB1", False), ("A1C", False)])
def test_t_num_09_code_d_une_nouvelle_ecole(code, accepte):
    """T-NUM-09 — `nouveau=True` : exactement 3 lettres (ADR029)."""
    num = _num()
    if accepte:
        num.valider_code_ecole(code, nouveau=True)
    else:
        with pytest.raises(ValueError):
            num.valider_code_ecole(code, nouveau=True)


# ──────────────────────────────────────────────────────────────────────
# lire_numeros
# ──────────────────────────────────────────────────────────────────────

ENTETES_FACTURES = [
    "N° facture", "École", "Contrat", "Période", "Date d'émission", "Échéance", "Total HT",
    "Date d'envoi", "Date de paiement", "Statut", "Note",
]


def _session(numero) -> list:
    """Ligne fictive de `Sessions`, colonnes dans l'ordre de MASTER_SCHEMA."""
    return [date(2026, 10, 1), "ABC", "Sept 2026", numero, date(2026, 9, 15), "09:00", "12:30", 3.5,
            "M1 S1", "Module A", 60, 210, "#ABC-M1 Cours fictif"]


def _empreinte(chemin) -> str:
    return hashlib.sha256(chemin.read_bytes()).hexdigest()


def test_t_num_10_lire_numeros_sessions_et_factures(tmp_path):
    """T-NUM-10 — `lire_numeros` réunit les numéros de `Sessions` (N°_Facture) et de `Factures`
    (N° facture), ignore les cellules vides et ne modifie pas le MASTER."""
    master = creer_master(tmp_path / "donnees")
    wb = load_workbook(master)
    for numero in ("R_ABC001", "R_ABC001", "R_ABC002", None, "R_ECO001"):
        wb["Sessions"].append(_session(numero))
    factures = wb.create_sheet("Factures")
    factures.append(ENTETES_FACTURES)
    factures.append(["R_ABC003", "ABC", "CT-EXEMPLE-0001", "Oct 2026", date(2026, 11, 1)])
    factures.append(["R_ECO002", "ECO", None, "Oct 2026", date(2026, 11, 1)])
    wb.save(master)
    avant = _empreinte(master)

    num = _num()
    numeros = num.lire_numeros(master)

    assert _empreinte(master) == avant
    assert all(isinstance(n, str) for n in numeros), numeros
    assert set(numeros) == {"R_ABC001", "R_ABC002", "R_ECO001", "R_ABC003", "R_ECO002"}
    assert num.prochain_numero(numeros, prefixe="R_", code_ecole="ABC") == "R_ABC004"


def test_t_num_11_lire_numeros_sans_onglet_factures_colonne_par_son_nom(tmp_path):
    """T-NUM-11 — sans onglet `Factures`, seuls les numéros de `Sessions` ; la colonne N°_Facture est
    trouvée par son en-tête, même déplacée."""
    master = creer_master(tmp_path / "donnees")
    wb = load_workbook(master)
    del wb["Sessions"]
    sessions = wb.create_sheet("Sessions")
    sessions.append(list(reversed(ENTETES["Sessions"])))
    for numero in ("R_ABC004", "R_ABC005"):
        sessions.append(list(reversed(_session(numero))))
    wb.save(master)

    assert set(_num().lire_numeros(master)) == {"R_ABC004", "R_ABC005"}
