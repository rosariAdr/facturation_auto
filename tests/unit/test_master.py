"""Lecture sûre du MASTER (F-11 ; banque de tests : T-MST-01 à T-MST-07, T-MST-17 à T-MST-19, T-MST-21 à T-MST-40).

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


# ══════════════════════════════════════════════════════════════════════
# P-0b — Socle du lot 1 (banque de tests : T-MST-22 à T-MST-38)
# ══════════════════════════════════════════════════════════════════════
#
# Contrat (contrats des vagues 3 à 5, § 1.4, 1.5 et 3.5), et constats des § 15 et 17 qui priment :
# - `Master` gagne `tarifs`, `ecoles`, `cours`, `rangs` (ligne Excel de chaque élément de chaque
#   liste) et `suivi_invalide` (onglet → message) ;
# - onglets de suivi lus de façon tolérante : absent → [] ; colonne exigée manquante ou en double
#   → [] et `suivi_invalide[onglet]` ; `lire_master` ne lève pas pour eux ;
# - `verrou_ecriture` : verrou seul, sans contrôle « classeur ouvert » ni sauvegarde ;
# - `verifier_classeur_non_ouvert` : `ClasseurVerrouille`, message propre au classeur d'école ;
# - `index_colonnes`, `derniere_ligne_remplie` ;
# - `fidelite_aller_retour` : sur une copie, une ligne par perte (parties de l'archive, formules
#   dont la valeur en cache se perd) ; § 17 : image, graphique ou dessin avec ancre perdus sont
#   signalés ; les parties vides des exports Google (dessin sans ancre, `xl/metadata`) ne le sont
#   pas ; Pillow devient une dépendance, pour que les images survivent.
#
# Lectures retenues (`💡 [proposé — à valider]`) :
# - `rangs` est indexé par le nom de l'attribut de la liste (`"sessions"`, `"tarifs"`…), comme
#   « chaque liste » du contrat ; `suivi_invalide` l'est par le nom de l'onglet (« Tarifs ») ;
# - un onglet de suivi absent n'a pas d'entrée dans `suivi_invalide` (il n'est pas invalide :
#   `profil verifier` le signale déjà comme absent) ;
# - `index_colonnes` rend des numéros de colonne openpyxl (à partir de 1) et refuse aussi une
#   colonne attendue en double ;
# - `derniere_ligne_remplie` ignore les cellules sans valeur (mise en forme seule) ;
# - les formules signalées sont celles des onglets que l'outil lit (MASTER_SCHEMA) : la formule
#   du pied de `Devis`, seul onglet à formule du MASTER réel (§ 15), ne l'est pas ;
# - la perte d'une image est simulée en retirant Pillow à openpyxl (`PILImage = None`), celle
#   d'un graphique en lui faisant ignorer les dessins : c'est ce qu'il fait d'un graphique qu'il
#   ne sait pas relire. Une forme ancrée (`xdr:sp`) est perdue par openpyxl sans simulation.

import hashlib  # noqa: E402
import importlib.util  # noqa: E402
import os  # noqa: E402
import stat  # noqa: E402
import struct  # noqa: E402
import tomllib  # noqa: E402
import warnings  # noqa: E402
import zlib  # noqa: E402

from openpyxl import Workbook  # noqa: E402
from openpyxl.chart import BarChart, Reference  # noqa: E402
from openpyxl.styles import Font  # noqa: E402

DEPOT = Path(__file__).resolve().parents[2]


def _fichiers():
    return importlib.import_module("facturation.fichiers")


def _sha(chemin: Path) -> str:
    return hashlib.sha256(chemin.read_bytes()).hexdigest()


def _parties(chemin: Path) -> list[str]:
    with zipfile.ZipFile(chemin) as archive:
        return sorted(archive.namelist())


def _png_1x1() -> bytes:
    """Une image PNG de 1 × 1 pixel, construite sans Pillow."""
    def bloc(nature: bytes, donnees: bytes) -> bytes:
        return (struct.pack(">I", len(donnees)) + nature + donnees
                + struct.pack(">I", zlib.crc32(nature + donnees) & 0xFFFFFFFF))
    entete = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + bloc(b"IHDR", entete) + bloc(b"IDAT", zlib.compress(b"\x00\xff\x00\x00"))
            + bloc(b"IEND", b""))


_NS_DESSIN = ('xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
              'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
              'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
_DE = "<xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
_A = "<xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>"
DESSIN_IMAGE = (
    f'<xdr:oneCellAnchor>{_DE}<xdr:ext cx="952500" cy="952500"/>'
    '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Logo fictif"/><xdr:cNvPicPr/></xdr:nvPicPr>'
    '<xdr:blipFill><a:blip r:embed="rIdImageFictive"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
    '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'
    '<xdr:clientData/></xdr:oneCellAnchor>'
)
DESSIN_FORME = (
    f'<xdr:twoCellAnchor>{_DE}{_A}<xdr:sp macro="" textlink="">'
    '<xdr:nvSpPr><xdr:cNvPr id="2" name="Forme fictive"/><xdr:cNvSpPr/></xdr:nvSpPr>'
    '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>'
    '<xdr:clientData/></xdr:twoCellAnchor>'
)


def _ajouter_dessin(chemin: Path, contenu: str, *, image: bool = False, metadata: bool = False):
    """Ajoute à la première feuille un dessin (`contenu` : ancres DrawingML, vide pour un dessin sans
    ancre), avec son image si `image`, et une partie `xl/metadata.xml` si `metadata` (exports Google)."""
    temporaire = chemin.with_suffix(".tmp")
    with zipfile.ZipFile(chemin) as source, zipfile.ZipFile(temporaire, "w", zipfile.ZIP_DEFLATED) as copie:
        noms = source.namelist()
        assert "xl/worksheets/sheet1.xml" in noms and "xl/worksheets/_rels/sheet1.xml.rels" not in noms
        for info in source.infolist():
            donnees = source.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                assert donnees.count(b"</worksheet>") == 1  # la garde ne passe pas à vide
                donnees = donnees.replace(b"</worksheet>", (
                    b'<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
                    b'r:id="rIdDessinFictif"/></worksheet>'))
            elif info.filename == "[Content_Types].xml":
                ajout = (b'<Override PartName="/xl/drawings/drawing1.xml" '
                         b'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>')
                if image:
                    ajout += b'<Default Extension="png" ContentType="image/png"/>'
                if metadata:
                    ajout += (b'<Override PartName="/xl/metadata.xml" ContentType="application/'
                              b'vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>')
                donnees = donnees.replace(b"</Types>", ajout + b"</Types>")
            copie.writestr(info, donnees)
        copie.writestr("xl/worksheets/_rels/sheet1.xml.rels", (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rIdDessinFictif" Target="../drawings/drawing1.xml" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"/></Relationships>'))
        copie.writestr("xl/drawings/drawing1.xml", (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr {_NS_DESSIN}>{contenu}</xdr:wsDr>'))
        if image:
            copie.writestr("xl/drawings/_rels/drawing1.xml.rels", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rIdImageFictive" Target="../media/image1.png" '
                'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"/></Relationships>'))
            copie.writestr("xl/media/image1.png", _png_1x1())
        if metadata:
            copie.writestr("xl/metadata.xml", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>'))
    shutil.move(temporaire, chemin)


def _aller_retour_openpyxl(chemin: Path, dossier: Path) -> list[str]:
    """Parties de l'archive après un enregistrement par openpyxl (pour vérifier qu'une perte a bien lieu)."""
    copie = dossier / f"aller_retour_{chemin.parent.parent.name}_{chemin.name}"
    load_workbook(chemin).save(copie)
    return _parties(copie)


def _ajouter_graphique(chemin: Path):
    wb = load_workbook(chemin)
    ws = wb.create_sheet("Graphique fictif")
    for i in range(1, 5):
        ws.append([f"Mois fictif {i}", i * 10])
    graphique = BarChart()
    graphique.add_data(Reference(ws, min_col=2, min_row=1, max_row=4))
    ws.add_chart(graphique, "D2")
    wb.save(chemin)


def _formules_dans(chemin: Path, onglet: str, n: int = 2):
    """Ajoute `n` lignes dont une cellule est une formule (`Total HT` dans `Sessions`, `Montant HT`
    dans `Devis`)."""
    wb = load_workbook(chemin)
    ws = wb[onglet]
    for i in range(n):
        ligne = ws.max_row + 1
        if onglet == "Sessions":
            ws.append([*_session(f"R_ABC10{i}")[:11], f"=H{ligne}*K{ligne}", "#ABC-M1 Cours fictif"])
        else:
            ws.append(["ABC", "Module fictif", f"D-FICTIF-{i}", date(2026, 9, 1), f"=SUM(E1:E{ligne - 1})", "émis"])
    wb.save(chemin)


# ──────────────────────────────────────────────────────────────────────
# Onglets de suivi et rangs
# ──────────────────────────────────────────────────────────────────────


def test_t_mst_22_tarifs_se_lit_sans_ref_contrat(tmp_path):
    """T-MST-22 — `Tarifs` se lit sans la colonne `Réf. contrat` (facultative, comme `Notes`), et avec
    elle quand elle existe ; aucune entrée dans `suivi_invalide`."""
    module = _module()
    sans = creer_master(tmp_path / "sans")
    wb = load_workbook(sans)
    assert "Réf. contrat" not in [c.value for c in wb["Tarifs"][1]]  # la garde ne passe pas à vide
    wb["Tarifs"].append(["ABC", "Statistiques fictives", "M1", 45, "note fictive"])
    wb.save(sans)
    lu = module.lire_master(sans)
    assert [(t["École"], t["Taux HT/h"]) for t in lu.tarifs] == [("ABC", 45)]
    assert lu.suivi_invalide == {}

    avec = creer_master(tmp_path / "avec", tarifs=[["ABC", "Statistiques fictives", "M1", "60 €/h", "", "CT-FICTIF-01"]])
    lu = module.lire_master(avec)
    assert [(t["Taux HT/h"], t["Réf. contrat"]) for t in lu.tarifs] == [("60 €/h", "CT-FICTIF-01")]
    assert lu.suivi_invalide == {}


@pytest.mark.parametrize(
    ("onglet", "attribut", "colonne"),
    [("Tarifs", "tarifs", "Taux HT/h"), ("Tarifs", "tarifs", "École"), ("Écoles", "ecoles", "Code facturation"),
     ("Cours", "cours", "Statut")],
)
def test_t_mst_23_colonne_de_suivi_manquante_ne_leve_rien(tmp_path, onglet, attribut, colonne):
    """T-MST-23 — une colonne exigée manquante dans un onglet de suivi ne lève rien : l'onglet donne []
    et `suivi_invalide[onglet]` nomme la colonne ; les autres onglets restent lus."""
    module = _module()
    chemin = creer_master(
        tmp_path / "donnees", sessions=[_session("R_ABC001")],
        tarifs=[["ABC", "Statistiques fictives", "M1", 45, "", ""]],
        ecoles=[["École Fictive A", "ABC"]],
        cours=[["ABC", "Module fictif", "M1", 30, 10, 6, 4, "R_ABC001", None, "En cours"]],
    )
    _retirer_colonne(chemin, onglet, colonne)
    lu = module.lire_master(chemin)
    assert getattr(lu, attribut) == []
    assert set(lu.suivi_invalide) == {onglet}
    assert colonne in lu.suivi_invalide[onglet]
    assert [s["N°_Facture"] for s in lu.sessions] == ["R_ABC001"]


def test_t_mst_27_colonne_de_suivi_en_double(tmp_path):
    """T-MST-27 — une colonne exigée en double dans un onglet de suivi : [] et `suivi_invalide` nomme
    l'onglet et la colonne, sans lever ; un onglet de suivi valide reste lu."""
    module = _module()
    chemin = creer_master(tmp_path / "donnees", ecoles=[["École Fictive A", "ABC"]],
                          tarifs=[["ABC", "Statistiques fictives", "M1", 45, "", ""]])
    wb = load_workbook(chemin)
    ws = wb["Écoles"]
    ws.cell(row=1, column=ws.max_column + 1, value="Code facturation")
    wb.save(chemin)
    lu = module.lire_master(chemin)
    assert lu.ecoles == []
    assert "Code facturation" in lu.suivi_invalide["Écoles"]
    assert "Tarifs" not in lu.suivi_invalide and len(lu.tarifs) == 1


def test_t_mst_28_onglets_de_suivi_lus_ou_absents(tmp_path):
    """T-MST-28 — `Tarifs`, `Écoles` et `Cours` sont lus par le nom de leurs colonnes ; un onglet de
    suivi absent donne [] sans erreur ni entrée dans `suivi_invalide`."""
    module = _module()
    chemin = creer_master(
        tmp_path / "donnees", ecoles=[["École Fictive A", "ABC", "SAS FICTIVE A"], None, ["Institut Fictif B", "ECO"]],
        cours=[["ABC", "Module fictif", "M1", 30, 10, 6, 4, "R_ABC001, R_ABC002", None, "Facturé (octobre)"]],
    )
    lu = module.lire_master(chemin)
    assert [(e["École"], e["Code facturation"]) for e in lu.ecoles] == [
        ("École Fictive A", "ABC"), ("Institut Fictif B", "ECO"),
    ]
    assert [(c["École"], c["N° Facture liée"], c["Statut"]) for c in lu.cours] == [
        ("ABC", "R_ABC001, R_ABC002", "Facturé (octobre)"),
    ]
    assert lu.tarifs == [] and lu.suivi_invalide == {}

    wb = load_workbook(chemin)
    for onglet in ("Tarifs", "Écoles", "Cours"):
        del wb[onglet]
    wb.save(chemin)
    lu = module.lire_master(chemin)
    assert (lu.tarifs, lu.ecoles, lu.cours) == ([], [], [])
    assert lu.suivi_invalide == {}


def test_t_mst_38_onglet_de_suivi_abime_leve_toujours(tmp_path):
    """T-MST-38 — la tolérance des onglets de suivi ne couvre pas un fichier abîmé : une feuille `Tarifs`
    dont le XML se révèle corrompu pendant la lecture des lignes (balise de fin mal fermée) lève
    toujours `MasterInvalide`, qui nomme le fichier, au lieu d'une simple entrée de `suivi_invalide`
    (§ 3.5)."""
    module = _module()
    chemin = creer_master(tmp_path / "donnees", tarifs=[["ABC", "Statistiques fictives", "M1", 45, "", ""]])
    partie = f"xl/worksheets/sheet{load_workbook(chemin).sheetnames.index('Tarifs') + 1}.xml"
    with zipfile.ZipFile(chemin) as archive:
        original = archive.read(partie)
    assert original.count(b"</sheetData>") == 1  # la garde ne passe pas à vide
    _abimer(chemin, partie, original.replace(b"</sheetData>", b"</sheetDat>"))
    with pytest.raises(module.MasterInvalide) as exc:
        module.lire_master(chemin)
    assert chemin.name in str(exc.value), str(exc.value)


def test_t_mst_26_rangs_donnent_les_lignes_excel(tmp_path):
    """T-MST-26 — `rangs[liste]` donne la ligne Excel de chaque élément, lignes vides comprises dans le
    compte (elles ne figurent pas dans les listes) ; une entrée par attribut de `Master`."""
    module = _module()
    chemin = creer_master(
        tmp_path / "donnees",
        sessions=[_session("R_ABC001"), None, None, _session("R_ABC002"), _session("R_ABC003")],
        factures=[None, ["R_ABC001", "ABC", None, "Sept 2026"]],
        tarifs=[["ABC", "Statistiques fictives", "M1", 45, "", ""], None, ["ECO", "Module fictif", "M2", 50, "", ""]],
    )
    lu = module.lire_master(chemin)
    assert [s["N°_Facture"] for s in lu.sessions] == ["R_ABC001", "R_ABC002", "R_ABC003"]
    assert lu.rangs["sessions"] == [2, 5, 6]
    assert lu.rangs["factures"] == [3]
    assert lu.rangs["tarifs"] == [2, 4]
    for attribut in ("coordonnees", "sessions", "factures", "tarifs", "ecoles", "cours"):
        assert len(lu.rangs.get(attribut, [])) == len(getattr(lu, attribut)), attribut
    assert lu.coordonnees[0]["Type"] == "Émetteur" and lu.rangs["coordonnees"][0] == 2


# ──────────────────────────────────────────────────────────────────────
# Verrou et classeurs d'école
# ──────────────────────────────────────────────────────────────────────


def test_t_mst_24_verrou_ecriture_bloque_un_autre_poste(tmp_path):
    """T-MST-24 — `verrou_ecriture` ne refuse pas un MASTER ouvert dans Excel (il ne garde que les
    écritures sous `_sorties/`) et ne fait aucune sauvegarde ; il pose le verrou de l'entité, le
    retire à la sortie, bloque un autre poste (`MasterVerrouille` qui le nomme) et n'est pas
    réentrant."""
    module = _module()
    master = creer_master(tmp_path / "donnees")
    ouvert = master.parent / f"~${master.name}"
    ouvert.write_bytes(b"")  # Excel a le MASTER ouvert
    verrou = master.with_suffix(".verrou")
    avant = master.read_bytes()

    with module.verrou_ecriture(master, poste="POSTE-A"):
        assert verrou.exists()
        with pytest.raises(module.MasterVerrouille):  # pas réentrant
            with module.verrou_ecriture(master, poste="POSTE-A"):
                pass
    assert not verrou.exists()
    assert not (master.parent / "sauvegardes").exists()
    assert master.read_bytes() == avant

    ouvert.unlink()
    with module.ecriture_master(master, poste="POSTE-B"):
        with pytest.raises(module.MasterVerrouille, match="POSTE-B"):
            with module.verrou_ecriture(master, poste="POSTE-A"):
                pass
    assert not verrou.exists()


@pytest.mark.parametrize(
    ("proprietaire", "logiciel"),
    [("~$ROSARI_ABC.xlsx", "Excel"), (".~lock.ROSARI_ABC.xlsx#", "LibreOffice")],
)
def test_t_mst_29_classeur_d_ecole_ouvert(tmp_path, proprietaire, logiciel):
    """T-MST-29 — `verifier_classeur_non_ouvert` lève `ClasseurVerrouille` pour un classeur d'école
    ouvert dans Excel ou LibreOffice, avec un message propre au classeur (son nom, le logiciel),
    pas « Le MASTER… » ; fermé, il ne lève rien."""
    module = _module()
    classeur = tmp_path / "ABC" / "ROSARI_ABC.xlsx"
    classeur.parent.mkdir()
    Workbook().save(classeur)
    assert module.verifier_classeur_non_ouvert(classeur) is None
    (classeur.parent / proprietaire).write_bytes(b"")
    with pytest.raises(_fichiers().ClasseurVerrouille) as erreur:
        module.verifier_classeur_non_ouvert(classeur)
    message = str(erreur.value)
    assert "ROSARI_ABC.xlsx" in message and logiciel in message
    assert "MASTER" not in message


def test_t_mst_30_classeur_d_ecole_absent_ou_en_lecture_seule(tmp_path):
    """T-MST-30 — un classeur d'école absent ne lève rien (il sera créé) ; en lecture seule, il lève
    `ClasseurVerrouille`, qui le nomme."""
    module = _module()
    classeur = tmp_path / "ABC" / "ROSARI_ABC.xlsx"
    assert module.verifier_classeur_non_ouvert(classeur) is None
    classeur.parent.mkdir()
    Workbook().save(classeur)
    os.chmod(classeur, stat.S_IREAD)
    try:
        with pytest.raises(_fichiers().ClasseurVerrouille, match="ROSARI_ABC.xlsx"):
            module.verifier_classeur_non_ouvert(classeur)
    finally:
        os.chmod(classeur, stat.S_IREAD | stat.S_IWRITE)


def test_t_mst_31_index_colonnes():
    """T-MST-31 — `index_colonnes` : en-tête (NFC, sans espaces autour) → numéro de colonne ; une
    colonne attendue absente ou en double lève `MasterInvalide`, qui nomme l'onglet et la colonne."""
    module = _module()
    wb = Workbook()
    ws = wb.active
    ws.title = "Factures"
    ws.append(["Note libre", unicodedata.normalize("NFD", " Échéance "), "N° facture", "Statut"])
    index = module.index_colonnes(ws, ["N° facture", "Statut", "Échéance"], onglet="Factures")
    assert {cle: index[cle] for cle in ("N° facture", "Statut", "Échéance")} == {
        "N° facture": 3, "Statut": 4, "Échéance": 2,
    }
    with pytest.raises(module.MasterInvalide) as erreur:
        module.index_colonnes(ws, ["N° facture", "Date d'envoi"], onglet="Factures")
    assert "Factures" in str(erreur.value) and "Date d'envoi" in str(erreur.value)
    ws.cell(row=1, column=5, value="Statut")
    with pytest.raises(module.MasterInvalide, match="Statut"):
        module.index_colonnes(ws, ["N° facture", "Statut"], onglet="Factures")


def test_t_mst_32_derniere_ligne_remplie():
    """T-MST-32 — `derniere_ligne_remplie` : dernière ligne qui porte une valeur ; une cellule mise en
    forme mais vide, plus bas, ne compte pas ; en-tête seul → 1."""
    module = _module()
    wb = Workbook()
    ws = wb.active
    ws.append(["N° facture", "Statut"])
    assert module.derniere_ligne_remplie(ws) == 1
    ws.append(["R_ABC001", "émise"])
    ws.append([])
    ws.append([None, "envoyée"])
    ws.cell(row=12, column=1).font = Font(bold=True)  # mise en forme seule
    assert ws.max_row == 12  # la garde ne passe pas à vide
    assert module.derniere_ligne_remplie(ws) == 4


# ──────────────────────────────────────────────────────────────────────
# Fidélité de l'aller-retour openpyxl
# ──────────────────────────────────────────────────────────────────────


def test_t_mst_25_fidelite_graphique_perdu_et_formules(tmp_path, monkeypatch):
    """T-MST-25 — `fidelite_aller_retour` signale un graphique perdu à l'enregistrement et les
    formules d'un onglet lu par l'outil (leur valeur en cache se perd : « elles se liront vides »),
    sans toucher l'original ; elle ne signale rien sur un classeur simple ni sur un graphique
    conservé."""
    module = _module()
    simple = creer_master(tmp_path / "simple")
    assert module.fidelite_aller_retour(simple) == []

    graphique = creer_master(tmp_path / "graphique")
    _ajouter_graphique(graphique)
    assert any(p.startswith("xl/charts/") for p in _parties(graphique))
    assert module.fidelite_aller_retour(graphique) == []  # graphique relu et conservé
    import openpyxl.reader.excel as lecteur
    monkeypatch.setattr(lecteur, "find_images", lambda archive, chemin: ([], []))
    assert not any(p.startswith("xl/charts/") for p in _aller_retour_openpyxl(graphique, tmp_path))
    pertes = module.fidelite_aller_retour(graphique)
    assert any(re.search(r"(?i)graphique|chart", ligne) for ligne in pertes), pertes
    monkeypatch.undo()

    formules = creer_master(tmp_path / "formules")
    _formules_dans(formules, "Sessions", n=2)
    empreinte, contenu = _sha(formules), sorted(os.listdir(formules.parent))
    pertes = module.fidelite_aller_retour(formules)
    lignes = [ligne for ligne in pertes if "se liront vides" in ligne]
    assert len(lignes) == 1, pertes
    assert "Sessions" in lignes[0] and re.search(r"\b2\b", lignes[0]), lignes[0]
    assert _sha(formules) == empreinte and sorted(os.listdir(formules.parent)) == contenu


def test_t_mst_33_fidelite_formule_de_devis_non_signalee(tmp_path):
    """T-MST-33 — la formule du pied de `Devis` (seule formule du MASTER réel, § 15) n'est pas
    signalée : l'outil ne lit pas cet onglet."""
    module = _module()
    chemin = creer_master(tmp_path / "donnees")
    _formules_dans(chemin, "Devis", n=1)
    assert any(isinstance(c.value, str) and c.value.startswith("=")
               for ligne in load_workbook(chemin)["Devis"].iter_rows() for c in ligne)
    assert module.fidelite_aller_retour(chemin) == []


def test_t_mst_34_fidelite_image_perdue_sans_pillow(tmp_path, monkeypatch):
    """T-MST-34 — sans Pillow, openpyxl abandonne l'image d'un classeur sans prévenir (§ 17) :
    `fidelite_aller_retour` le signale, sans toucher l'original ni laisser de fichier à côté."""
    module = _module()
    chemin = creer_master(tmp_path / "donnees")
    _ajouter_dessin(chemin, DESSIN_IMAGE, image=True)
    import openpyxl.reader.drawings as dessins
    monkeypatch.setattr(dessins, "PILImage", None)
    assert "xl/media/image1.png" not in _aller_retour_openpyxl(chemin, tmp_path)
    empreinte, contenu = _sha(chemin), sorted(os.listdir(chemin.parent))
    pertes = module.fidelite_aller_retour(chemin)
    assert any(re.search(r"(?i)image|media", ligne) for ligne in pertes), pertes
    assert _sha(chemin) == empreinte and sorted(os.listdir(chemin.parent)) == contenu


def test_t_mst_35_fidelite_dessin_ancre_perdu_mais_pas_un_dessin_vide(tmp_path):
    """T-MST-35 — une forme ancrée, qu'openpyxl ne sait pas garder, est signalée (« dessin avec
    ancre », § 17) ; un dessin sans ancre et une partie `xl/metadata` (exports Google, sans effet
    visible) ne le sont pas, bien qu'openpyxl les perde aussi."""
    module = _module()
    forme = creer_master(tmp_path / "forme")
    _ajouter_dessin(forme, DESSIN_FORME)
    assert "xl/drawings/drawing1.xml" not in _aller_retour_openpyxl(forme, tmp_path)
    pertes = module.fidelite_aller_retour(forme)
    assert any(re.search(r"(?i)dessin|drawing", ligne) for ligne in pertes), pertes

    vide = creer_master(tmp_path / "vide")
    _ajouter_dessin(vide, "", metadata=True)
    restantes = _aller_retour_openpyxl(vide, tmp_path)
    assert "xl/drawings/drawing1.xml" not in restantes and "xl/metadata.xml" not in restantes
    assert module.fidelite_aller_retour(vide) == []


def test_t_mst_36_pillow_dependance_declaree():
    """T-MST-36 — Pillow figure dans les dépendances du paquet (`[project] dependencies`), pour que les
    images des classeurs survivent à l'enregistrement (§ 17)."""
    projet = tomllib.loads((DEPOT / "pyproject.toml").read_text(encoding="utf-8"))["project"]
    noms = [re.split(r"[<>=!~;\[ ]", dependance, maxsplit=1)[0].strip().casefold()
            for dependance in projet["dependencies"]]
    assert "pillow" in noms, projet["dependencies"]


def test_t_mst_37_image_conservee_avec_pillow(tmp_path):
    """T-MST-37 — Pillow installé, l'image d'un classeur survit à l'enregistrement par openpyxl :
    `fidelite_aller_retour` ne signale rien."""
    if importlib.util.find_spec("PIL") is None:
        pytest.fail("Pillow, dépendance déclarée, n'est pas installé : python -m pip install -e \".[dev]\"")
    module = _module()
    chemin = creer_master(tmp_path / "donnees")
    _ajouter_dessin(chemin, DESSIN_IMAGE, image=True)
    assert "xl/media/image1.png" in _aller_retour_openpyxl(chemin, tmp_path)
    assert module.fidelite_aller_retour(chemin) == []


# ──────────────────────────────────────────────────────────────────────
# Revue de P-0b (T-MST-39, T-MST-40)
# ──────────────────────────────────────────────────────────────────────
#
# Constat des relecteurs : une bonne partie de `fidelite_aller_retour` n'était couverte par aucun
# test. Le comptage des validations et des règles de mise en forme conditionnelle par feuille
# pouvait être mis à zéro, et la comparaison des classeurs (fusions, noms définis…) rendue inactive,
# sans qu'aucun test ne rougisse ; F-17 s'appuiera sur cette fonction pour refuser d'écrire dans un
# classeur d'école (§ 17). Ces tests couvrent des gardes déjà codées : ils sont verts, et chacune a
# été vérifiée en la cassant (AGENTS.md §5).
# - T-MST-39 : une validation de données x14 (liste vers un autre onglet) et une barre de données
#   x14, qu'openpyxl abandonne sans les relire (extension `extLst` de la feuille), sont signalées ;
# - T-MST-40 : une fusion ou un nom défini perdus à l'enregistrement sont signalés. openpyxl les garde
#   (§ 17) : la perte est simulée en lui faisant sauter leur écriture.

_NS_X14 = 'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"'
_NS_XM = 'xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"'
EXTENSION_VALIDATION = (
    f'<extLst><ext uri="{{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}}" {_NS_X14}>'
    f'<x14:dataValidations count="1" {_NS_XM}>'
    '<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1">'
    "<x14:formula1><xm:f>'Sessions'!$B$2:$B$5</xm:f></x14:formula1><xm:sqref>B12:B20</xm:sqref>"
    '</x14:dataValidation></x14:dataValidations></ext></extLst>'
)
EXTENSION_BARRE = (
    f'<extLst><ext uri="{{78C0D931-6437-407d-A8EE-F0AAD7539E65}}" {_NS_X14}>'
    f'<x14:conditionalFormattings><x14:conditionalFormatting {_NS_XM}>'
    '<x14:cfRule type="dataBar" id="{00000000-0000-0000-0000-0000000F1C70}">'
    '<x14:dataBar minLength="0" maxLength="100"><x14:cfvo type="autoMin"/><x14:cfvo type="autoMax"/>'
    '</x14:dataBar></x14:cfRule><xm:sqref>D12:D20</xm:sqref>'
    '</x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>'
)


def _ajouter_extension(chemin: Path, extension: str):
    """Ajoute une extension `extLst` à la fin de la première feuille (`Coordonnées`)."""
    temporaire = chemin.with_suffix(".tmp")
    with zipfile.ZipFile(chemin) as source, zipfile.ZipFile(temporaire, "w", zipfile.ZIP_DEFLATED) as copie:
        for info in source.infolist():
            donnees = source.read(info.filename)
            if info.filename == "xl/worksheets/sheet1.xml":
                assert donnees.count(b"</worksheet>") == 1 and b"extLst" not in donnees  # garde non vide
                donnees = donnees.replace(b"</worksheet>", extension.encode("utf-8") + b"</worksheet>")
            copie.writestr(info, donnees)
    shutil.move(temporaire, chemin)


@pytest.mark.parametrize(
    ("extension", "libelle"),
    [(EXTENSION_VALIDATION, "validation(s) de données perdue(s)"),
     (EXTENSION_BARRE, "règle(s) de mise en forme conditionnelle perdue(s)")],
    ids=["validation-x14", "barre-de-donnees-x14"],
)
def test_t_mst_39_fidelite_extensions_x14_perdues(tmp_path, extension, libelle):
    """T-MST-39 — une validation de données ou une règle de mise en forme conditionnelle x14,
    qu'openpyxl abandonne à l'enregistrement, est signalée pour son onglet (« 1 avant, 0 après »),
    sans toucher l'original."""
    module = _module()
    chemin = creer_master(tmp_path / "donnees")
    _ajouter_extension(chemin, extension)
    relu = tmp_path / "relu.xlsx"
    with warnings.catch_warnings():  # openpyxl prévient qu'il abandonne l'extension : c'est voulu ici
        warnings.simplefilter("ignore")
        classeur = load_workbook(chemin)
        assert classeur.sheetnames[0] == "Coordonnées"
        classeur.save(relu)
    with zipfile.ZipFile(relu) as archive:
        assert b"x14:" not in archive.read("xl/worksheets/sheet1.xml")  # la perte a bien lieu
    empreinte = _sha(chemin)
    pertes = module.fidelite_aller_retour(chemin)
    lignes = [ligne for ligne in pertes if libelle in ligne]
    assert len(lignes) == 1, pertes
    assert "Coordonnées" in lignes[0] and "1 avant, 0 après" in lignes[0], lignes[0]
    assert _sha(chemin) == empreinte


def test_t_mst_40_fidelite_fusion_et_nom_defini_perdus(tmp_path, monkeypatch):
    """T-MST-40 — une plage fusionnée ou un nom défini perdus à l'enregistrement sont signalés (onglet,
    ou nom du nom défini) ; conservés, ils ne le sont pas."""
    from openpyxl.workbook.defined_name import DefinedName
    import openpyxl.workbook._writer as ecrivain_classeur
    import openpyxl.worksheet._writer as ecrivain_feuille

    module = _module()
    chemin = creer_master(tmp_path / "donnees")
    wb = load_workbook(chemin)
    wb["Coordonnées"].merge_cells("J1:K2")
    wb.defined_names["Zone_fictive"] = DefinedName("Zone_fictive", attr_text="'Coordonnées'!$A$1:$B$3")
    wb.save(chemin)
    assert module.fidelite_aller_retour(chemin) == []  # openpyxl garde fusion et nom défini

    monkeypatch.setattr(ecrivain_feuille.WorksheetWriter, "write_merged_cells", lambda self: None)
    pertes = module.fidelite_aller_retour(chemin)
    assert any("fusionnées" in ligne and "Coordonnées" in ligne for ligne in pertes), pertes
    monkeypatch.undo()

    monkeypatch.setattr(ecrivain_classeur.WorkbookWriter, "write_names", lambda self: None)
    pertes = module.fidelite_aller_retour(chemin)
    assert any("nom(s) défini(s)" in ligne and "Zone_fictive" in ligne for ligne in pertes), pertes
    monkeypatch.undo()
