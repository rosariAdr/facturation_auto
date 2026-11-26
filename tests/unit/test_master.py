"""Lecture sûre du MASTER (F-11 ; banque de tests : T-MST-01 à T-MST-07, T-MST-17 à T-MST-19, T-MST-21).

Tests d'acceptation écrits avant l'implémentation : chaque test importe `facturation.master`
lui-même, pour que son absence fasse échouer les tests sans bloquer la collecte des autres.
Classeurs fictifs créés dans `tmp_path` (`tests.conftest.creer_master`).

Forme de `Master` supposée par ces tests (`💡 [proposé — à valider]`, le contrat ne la fixe
pas) : un attribut par onglet lu — `coordonnees`, `sessions`, `regles_facturation`,
`correspondances`, `factures` —, chacun une liste de dictionnaires {en-tête: valeur}, une
entrée par ligne non vide.
"""

from __future__ import annotations

import importlib
import re
import shutil
import unicodedata
import zipfile
from datetime import date
from pathlib import Path

import pytest
from openpyxl import load_workbook

from tests.conftest import ENTETES, creer_master

pytestmark = pytest.mark.unit

COLONNES_PREVUES = {  # MASTER_SCHEMA, onglets prévus
    "Règles_facturation": [
        "École", "Statut lot", "Rythme", "Début de cycle", "Date limite d'envoi", "Délai de paiement",
        "N° de contrat", "Une facture par contrat", "Bon de commande requis", "Colonne Jour",
        "Regroupement", "Canal d'envoi", "Pièces jointes", "Notes",
    ],
    "Correspondances": [
        "Clé", "Type de clé", "École", "Module", "Promotion", "Clé tarif", "Mode heures", "Valeur",
        "Gardée le", "Notes",
    ],
    "Factures": [
        "N° facture", "École", "Contrat", "Période", "Date d'émission", "Échéance", "Total HT",
        "Date d'envoi", "Date de paiement", "Statut", "Note",
    ],
}
ATTRIBUTS_PREVUS = {"Règles_facturation": "regles_facturation", "Correspondances": "correspondances", "Factures": "factures"}


def _module():
    """Module testé, importé dans le test : son absence fait échouer le test, pas la collecte."""
    return importlib.import_module("facturation.master")


def _session(numero: str) -> list:
    """Ligne fictive de `Sessions`, colonnes dans l'ordre de MASTER_SCHEMA."""
    return [date(2026, 10, 1), "ABC", "Sept 2026", numero, date(2026, 9, 15), "09:00", "12:30", 3.5,
            "M1 S1", "Module A", 60, 210, "#ABC-M1 Cours fictif"]


def _retirer_colonne(chemin: Path, onglet: str, entete: str):
    wb = load_workbook(chemin)
    ws = wb[onglet]
    colonne = [c.value for c in ws[1]].index(entete) + 1
    ws.delete_cols(colonne)
    wb.save(chemin)


def _ajouter_onglet(chemin: Path, nom: str, lignes: list[list]):
    wb = load_workbook(chemin)
    ws = wb.create_sheet(nom)
    ws.append(COLONNES_PREVUES[nom])
    for ligne in lignes:
        ws.append(ligne)
    wb.save(chemin)


@pytest.fixture
def master(tmp_path: Path) -> Path:
    return creer_master(tmp_path / "donnees")


def test_t_mst_01_lecture_d_un_master_complet(master):
    """T-MST-01 — un MASTER fictif complet se lit : coordonnées et séances, par le nom des colonnes."""
    wb = load_workbook(master)
    wb["Sessions"].append(_session("R_ABC001"))
    wb.save(master)

    lu = _module().lire_master(master)

    emetteurs = [ligne for ligne in lu.coordonnees if ligne["Type"] == "Émetteur"]
    assert [e["Nom / Raison sociale"] for e in emetteurs] == ["FORMATEUR FICTIF ROSARI EI"]
    assert [(s["N°_Facture"], s["Durée_h"]) for s in lu.sessions] == [("R_ABC001", 3.5)]


def test_t_mst_02_onglets_prevus_absents_listes_vides(master):
    """T-MST-02 — `Règles_facturation`, `Correspondances` et `Factures` absents : listes vides, pas d'erreur."""
    lu = _module().lire_master(master)
    assert (lu.regles_facturation, lu.correspondances, lu.factures) == ([], [], [])


@pytest.mark.parametrize("onglet", ["Coordonnées", "Sessions"])
def test_t_mst_03_onglet_obligatoire_absent(master, onglet):
    """T-MST-03 — un onglet obligatoire absent : `MasterInvalide` (une `ValueError`) qui le nomme."""
    module = _module()
    wb = load_workbook(master)
    del wb[onglet]
    wb.save(master)
    with pytest.raises(module.MasterInvalide) as exc:
        module.lire_master(master)
    assert isinstance(exc.value, ValueError)
    assert onglet in str(exc.value)


@pytest.mark.parametrize(
    ("onglet", "colonne"),
    [("Coordonnées", "Type"), ("Coordonnées", "Adresse L1"), ("Sessions", "N°_Facture"), ("Sessions", "Durée_h")],
)
def test_t_mst_04_en_tete_obligatoire_absent(master, onglet, colonne):
    """T-MST-04 — une colonne obligatoire absente : l'erreur nomme l'onglet et la colonne."""
    module = _module()
    _retirer_colonne(master, onglet, colonne)
    with pytest.raises(module.MasterInvalide) as exc:
        module.lire_master(master)
    assert onglet in str(exc.value) and colonne in str(exc.value), str(exc.value)


def test_t_mst_05_colonnes_lues_par_leur_nom(master):
    """T-MST-05 — colonnes de `Sessions` dans un autre ordre, plus une colonne en trop : lecture par le nom."""
    wb = load_workbook(master)
    del wb["Sessions"]
    sessions = wb.create_sheet("Sessions")
    sessions.append(["Commentaire libre", *reversed(ENTETES["Sessions"])])
    sessions.append(["note fictive", *reversed(_session("R_ABC002"))])
    wb.save(master)

    (seance,) = _module().lire_master(master).sessions
    assert (seance["N°_Facture"], seance["Durée_h"], seance["École"]) == ("R_ABC002", 3.5, "ABC")


def test_t_mst_06_onglets_prevus_presents_lus(master):
    """T-MST-06 — les onglets prévus présents sont lus, une entrée par ligne non vide."""
    _ajouter_onglet(master, "Factures", [
        ["R_ABC001", "ABC", "CT-EXEMPLE-0001", "Oct 2026", date(2026, 11, 1), date(2026, 12, 1), 210, None, None, "émise"],
        [None, None, None],  # ligne vide : ignorée
        ["R_ABC002", "ABC", "CT-EXEMPLE-0001", "Nov 2026", date(2026, 12, 1), date(2026, 12, 31), 420, None, None, "envoyée"],
    ])
    _ajouter_onglet(master, "Règles_facturation", [["ABC", "Lot 1", "mensuel"]])
    _ajouter_onglet(master, "Correspondances", [["#ABC-M1", "code", "ABC", "Module A", "M1 S1"]])

    lu = _module().lire_master(master)

    assert [(f["N° facture"], f["Statut"]) for f in lu.factures] == [("R_ABC001", "émise"), ("R_ABC002", "envoyée")]
    assert [(r["École"], r["Statut lot"]) for r in lu.regles_facturation] == [("ABC", "Lot 1")]
    assert [(c["Clé"], c["Type de clé"], c["Module"]) for c in lu.correspondances] == [("#ABC-M1", "code", "Module A")]


@pytest.mark.parametrize(
    ("onglet", "colonne"),
    [("Factures", "N° facture"), ("Règles_facturation", "Délai de paiement"), ("Correspondances", "Clé")],
)
def test_t_mst_07_onglet_prevu_present_sans_une_colonne(master, onglet, colonne):
    """T-MST-07 (`💡 [proposé — à valider]`) — un onglet prévu présent mais sans une colonne de
    MASTER_SCHEMA : l'erreur nomme l'onglet et la colonne."""
    module = _module()
    _ajouter_onglet(master, onglet, [])
    _retirer_colonne(master, onglet, colonne)
    with pytest.raises(module.MasterInvalide) as exc:
        module.lire_master(master)
    assert onglet in str(exc.value) and colonne in str(exc.value), str(exc.value)


@pytest.mark.parametrize(
    ("onglet", "doublon"),
    [
        ("Sessions", "Durée_h"),
        ("Coordonnées", "Email"),
        ("Sessions", unicodedata.normalize("NFD", "École")),  # même en-tête, autre forme Unicode
    ],
    ids=["sessions", "coordonnees", "sessions-nfd"],
)
def test_t_mst_17_colonne_en_double(master, onglet, doublon):
    """T-MST-17 — une colonne en double dans un onglet : `MasterInvalide` qui nomme l'onglet et la colonne,
    au lieu de lire l'une des deux au hasard."""
    module = _module()
    wb = load_workbook(master)
    ws = wb[onglet]
    ws.cell(row=1, column=ws.max_column + 1, value=doublon)
    wb.save(master)
    with pytest.raises(module.MasterInvalide) as exc:
        module.lire_master(master)
    nom = unicodedata.normalize("NFC", doublon)
    assert onglet in str(exc.value) and nom in unicodedata.normalize("NFC", str(exc.value)), str(exc.value)


def test_t_mst_18_poste_courant_non_vide(monkeypatch):
    """T-MST-18 — `poste_courant()` renvoie un nom non vide, sur une ligne, même sans COMPUTERNAME ni nom réseau."""
    module = _module()
    poste = module.poste_courant()
    assert isinstance(poste, str) and poste.strip() and "\n" not in poste, repr(poste)

    monkeypatch.delenv("COMPUTERNAME", raising=False)
    monkeypatch.setattr(module.platform, "node", lambda: "")
    secours = module.poste_courant()
    assert isinstance(secours, str) and secours.strip() and "\n" not in secours, repr(secours)


def _fausser_dimension(chemin: Path, remplacement: bytes):
    """Réécrit la dimension déclarée (`<dimension ref=…/>`) de chaque feuille, comme le font certains
    outils d'export : `remplacement` est un motif `re.sub` appliqué à la référence `A1:<col><ligne>`."""
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
def test_t_mst_19_dimension_declaree_fausse(master, remplacement):
    """T-MST-19 — une dimension déclarée fausse (A1:A1, ou la seule ligne d'en-tête) ne fait ni refuser un
    MASTER valide ni perdre de lignes en silence : en lecture seule, openpyxl s'y fie si on ne la réinitialise pas."""
    wb = load_workbook(master)
    wb["Sessions"].append(_session("R_ABC001"))
    wb["Sessions"].append(_session("R_ABC002"))
    wb.save(master)
    attendu = _module().lire_master(master)
    assert len(attendu.coordonnees) > 1 and len(attendu.sessions) == 2

    _fausser_dimension(master, remplacement)
    lu = _module().lire_master(master)
    assert lu.coordonnees == attendu.coordonnees
    assert [s["N°_Facture"] for s in lu.sessions] == ["R_ABC001", "R_ABC002"]


def test_t_mst_21a_ouverture_refusee_par_le_systeme(master, monkeypatch):
    """T-MST-21 — `PermissionError` à l'ouverture (classeur verrouillé par Excel, droits) : `MasterVerrouille`
    « fermez le classeur », pas une trace brute."""
    module = _module()
    assert hasattr(module, "load_workbook"), "le module n'ouvre plus le classeur par `load_workbook`"

    def refus(chemin, *args, **kwargs):
        raise PermissionError(13, "Permission denied", str(chemin))

    monkeypatch.setattr(module, "load_workbook", refus)
    with pytest.raises(Exception) as exc:
        module.lire_master(master)
    assert isinstance(exc.value, module.MasterVerrouille), repr(exc.value)
    assert "fermez le classeur" in str(exc.value).lower(), str(exc.value)


def _abimer(chemin: Path, partie: str, contenu: bytes | None):
    """Remplace une partie de l'archive xlsx par `contenu`, ou la retire si `contenu` est None."""
    temporaire = chemin.with_suffix(".tmp")
    with zipfile.ZipFile(chemin) as source, zipfile.ZipFile(temporaire, "w", zipfile.ZIP_DEFLATED) as copie:
        assert partie in source.namelist(), partie  # la garde ne passe pas à vide
        for info in source.infolist():
            if info.filename != partie:
                copie.writestr(info, source.read(info.filename))
            elif contenu is not None:
                copie.writestr(info, contenu)
    shutil.move(temporaire, chemin)


@pytest.mark.parametrize(
    ("partie", "contenu"),
    [
        ("xl/workbook.xml", None),
        ("[Content_Types].xml", None),
        ("xl/workbook.xml", b"<workbook pas du xml"),
        ("xl/worksheets/sheet1.xml", b"<worksheet pas du xml"),
    ],
    ids=["sans-workbook", "sans-content-types", "workbook-corrompu", "feuille-corrompue"],
)
def test_t_mst_21b_xlsx_abime(master, partie, contenu):
    """T-MST-21 — un xlsx abîmé (partie absente, XML corrompu, y compris dans une feuille lue plus tard) :
    `MasterInvalide` qui nomme le fichier, pas un `KeyError` ni un `ParseError` bruts."""
    module = _module()
    _abimer(master, partie, contenu)
    with pytest.raises(Exception) as exc:
        module.lire_master(master)
    assert isinstance(exc.value, module.MasterInvalide), repr(exc.value)
    assert master.name in str(exc.value), str(exc.value)
