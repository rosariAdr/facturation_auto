"""Numérotation des factures (F-15, ADR002, ADR029 ; banque de tests : T-NUM-01 à T-NUM-16).

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
import re
import shutil
import zipfile
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


# ──────────────────────────────────────────────────────────────────────
# Revue croisée (T-NUM-12 à T-NUM-16)
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize("code", ["123", "1AB", "0"])
def test_t_num_12_code_ecole_commence_par_une_lettre(code):
    """T-NUM-12 — un code école commence par une lettre (décision de l'orchestrateur, revue de F-15) :
    `R_123001` se lirait comme un préfixe suivi de chiffres. Refusé aussi par `prochain_numero`."""
    num = _num()
    with pytest.raises(ValueError):
        num.valider_code_ecole(code)
    with pytest.raises(ValueError):
        num.prochain_numero([], prefixe="R_", code_ecole=code)


@pytest.mark.parametrize("prefixe", ["R", "r_", "", "ROSA_", "R-", "R__", "1_", "R_ "])
def test_t_num_13_prefixe_invalide_refuse(prefixe):
    """T-NUM-13 — un préfixe d'entité hors de la règle du profil (1 à 3 majuscules suivies de « _ ») est
    refusé par `prochain_numero` comme par `numeros_manquants` : les entités ne se mélangent jamais."""
    num = _num()
    with pytest.raises(ValueError):
        num.prochain_numero(["R_ABC001"], prefixe=prefixe, code_ecole="ABC")
    with pytest.raises(ValueError):
        num.numeros_manquants(["R_ABC001"], prefixe=prefixe, code_ecole="ABC")


def test_t_num_14_onglet_factures_sans_sa_colonne(tmp_path):
    """T-NUM-14 — un onglet `Factures` sans sa colonne `N° facture` : `ValueError` qui nomme l'onglet et la
    colonne, plutôt qu'une liste incomplète qui ferait proposer un numéro déjà attribué."""
    master = creer_master(tmp_path / "donnees")
    wb = load_workbook(master)
    wb["Sessions"].append(_session("R_ABC001"))
    factures = wb.create_sheet("Factures")
    factures.append([e for e in ENTETES_FACTURES if e != "N° facture"])
    factures.append(["ABC", "CT-EXEMPLE-0001", "Oct 2026"])
    wb.save(master)

    with pytest.raises(ValueError) as exc:
        _num().lire_numeros(master)
    assert "Factures" in str(exc.value) and "N° facture" in str(exc.value), str(exc.value)


def _fausser_dimension(chemin, remplacement: bytes):
    """Déclare une dimension fausse dans chaque feuille, comme certains outils d'export : `remplacement` est
    un motif `re.sub` appliqué à `<dimension ref="A1:<col><ligne>"/>`."""
    temporaire = chemin.with_suffix(".tmp")
    with zipfile.ZipFile(chemin) as source, zipfile.ZipFile(temporaire, "w", zipfile.ZIP_DEFLATED) as copie:
        for info in source.infolist():
            donnees = source.read(info.filename)
            if info.filename.startswith("xl/worksheets/sheet"):
                donnees, n = re.subn(rb'<dimension ref="A1:([A-Z]+)(\d+)"\s*/>', remplacement, donnees)
                assert n == 1, info.filename  # la garde ne passe pas à vide
            copie.writestr(info, donnees)
    shutil.move(temporaire, chemin)


@pytest.mark.parametrize(
    "remplacement",
    [rb'<dimension ref="A1:A1"/>', rb'<dimension ref="A1:\g<1>1"/>'],
    ids=["A1-A1", "en-tete-seul"],
)
def test_t_num_15_dimension_declaree_fausse(tmp_path, remplacement):
    """T-NUM-15 — un classeur qui déclare `A1:A1` (ou sa seule ligne d'en-tête) alors que des numéros existent :
    `lire_numeros` les trouve quand même. En lecture seule, openpyxl se fie à la dimension déclarée si on ne la
    réinitialise pas : la colonne paraît absente, ou les numéros manquent en silence (numéro en double proposé)."""
    master = creer_master(tmp_path / "donnees")
    wb = load_workbook(master)
    for numero in ("R_ABC001", "R_ABC002"):
        wb["Sessions"].append(_session(numero))
    factures = wb.create_sheet("Factures")
    factures.append(ENTETES_FACTURES)
    factures.append(["R_ABC003", "ABC"])
    wb.save(master)
    _fausser_dimension(master, remplacement)

    numeros = _num().lire_numeros(master)
    assert numeros == ["R_ABC001", "R_ABC002", "R_ABC003"]
    assert _num().prochain_numero(numeros, prefixe="R_", code_ecole="ABC") == "R_ABC004"


def test_t_num_16_lire_numeros_trie_et_sans_doublon(tmp_path):
    """T-NUM-16 — `lire_numeros` renvoie une liste triée, sans doublon, même quand un numéro revient dans
    `Sessions` et dans `Factures`, et quel que soit l'ordre des lignes."""
    master = creer_master(tmp_path / "donnees")
    wb = load_workbook(master)
    for numero in ("R_ECO002", "R_ABC002", "K_ABC001", "R_ABC002", "R_ABC001"):
        wb["Sessions"].append(_session(numero))
    factures = wb.create_sheet("Factures")
    factures.append(ENTETES_FACTURES)
    for numero in ("R_ABC003", "R_ABC001", "R_ECO002"):
        factures.append([numero, "ABC"])
    wb.save(master)

    numeros = _num().lire_numeros(master)
    assert isinstance(numeros, list)
    assert numeros == ["K_ABC001", "R_ABC001", "R_ABC002", "R_ABC003", "R_ECO002"]
