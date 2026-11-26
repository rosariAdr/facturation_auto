"""Clonage du gabarit, sur des gabarits miniatures construits par les tests (T-CLO-01 à 11).

Chaque gabarit miniature porte un exemple de tout ce que le clonage doit reprendre :
préfixe texte, cellule déverrouillée, colonnes groupées et masquées, mise en page,
zone d'impression, volets figés, protection de feuille, etc.
"""

from __future__ import annotations

import os
from copy import copy
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, NamedStyle, PatternFill, Protection, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table
from openpyxl.worksheet.worksheet import Worksheet

from facturation import pcf_format as fmt
from tests.empreinte import ecarts, empreinte_cellules, empreinte_feuille, largeurs_effectives

pytestmark = pytest.mark.unit

TITRE = Font(bold=True, size=14)


@pytest.fixture(autouse=True)
def _plans_vides():
    """Chaque test part d'un cache de gabarits vide et le laisse vide."""
    fmt._PLANS_GABARIT.clear()
    yield
    fmt._PLANS_GABARIT.clear()


def _gabarit_mini(chemin: Path, preparer=None, titre: str = "Titre") -> Path:
    """Gabarit miniature : un exemple de chaque propriété que le clonage doit reprendre."""
    wb = Workbook()
    ws = wb.active
    ws.title = "MODELE"
    ws["A1"] = titre
    ws["A1"].font = TITRE
    ws.merge_cells("A1:C1")
    ws["A2"] = "0042"
    ws["A2"].quotePrefix = True
    ws["B2"] = 12.5
    ws["B2"].number_format = '0.00 "€"'
    ws["B2"].protection = Protection(locked=False)
    ws["C2"] = "=B2*2"
    ws["C2"].fill = PatternFill("solid", fgColor="FFEEDD")
    ws["C2"].border = Border(left=Side(style="thin"), bottom=Side(style="double"))
    for coord in ("A3", "B3", "C3"):  # même style que A1 : réutilisé par la voie rapide
        ws[coord].font = TITRE
    ws["D4"].alignment = Alignment(horizontal="center", wrap_text=True)
    ws.merge_cells("D4:E5")
    ws.column_dimensions["B"].width = 21.5
    ws.column_dimensions.group("F", "H", hidden=True, outline_level=1)
    ws.row_dimensions[3].height = 30
    ws.row_dimensions[6].hidden = True
    ws.row_dimensions[7].outlineLevel = 1
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_properties.tabColor = "1072BA"
    ws.page_margins.left = 0.3
    ws.print_options.horizontalCentered = True
    ws.oddFooter.center.text = "Page &P"
    ws.protection.sheet = True
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    ws.sheet_view.tabSelected = True  # onglet actif de son fichier, comme le vrai gabarit
    ws.freeze_panes = "A2"
    ws.print_area = "A1:H10"
    ws.print_title_rows = "1:1"
    if preparer:
        preparer(wb, ws)
    wb.save(chemin)
    return chemin


def _classeur_ecole() -> Workbook:
    """Classeur de destination avec un onglet déjà actif, comme un classeur d'école."""
    wb = Workbook()
    wb.active.title = "R_TEST001"
    return wb


def _cloner(gabarit: Path, wb: Workbook | None = None, nom: str = "CLONE") -> Worksheet:
    return fmt._xworkbook_clone_modele(gabarit, wb or _classeur_ecole(), nom)


def _source(gabarit: Path) -> Worksheet:
    return load_workbook(gabarit)["MODELE"]


def _assert_fidele(source: Worksheet, clone: Worksheet) -> None:
    attendu, obtenu = empreinte_cellules(source), empreinte_cellules(clone)
    assert attendu == obtenu, f"cellules différentes : {ecarts(attendu, obtenu)}"
    attendu, obtenu = empreinte_feuille(source), empreinte_feuille(clone)
    assert attendu == obtenu, f"propriétés de feuille différentes : {ecarts(attendu, obtenu)}"


def test_t_clo_01_contrat_du_style_interne_openpyxl():
    """T-CLO-01 — la structure interne sur laquelle repose la voie rapide existe toujours.

    À relancer avant de relever la borne d'openpyxl dans pyproject.toml.
    """
    assert fmt.StyleArray is not None
    vide = fmt.StyleArray()
    for champ in ("fontId", "fillId", "borderId", "numFmtId", "protectionId",
                  "alignmentId", "quotePrefix", "pivotButton", "xfId"):
        assert getattr(vide, champ) == 0, champ
    ws = Workbook().active
    a, b = ws["A1"], ws["B1"]
    a.font = Font(bold=True)
    a.protection = Protection(locked=False)
    a.number_format = "0.00"
    a.quotePrefix = True
    assert isinstance(a._style, fmt.StyleArray)
    b._style = copy(a._style)
    assert b._style == a._style and b._style is not a._style
    assert b.font.b and b.protection.locked is False and b.number_format == "0.00" and b.quotePrefix
    b.font = Font(italic=True)
    assert a.font.b and not a.font.i, "la copie du tableau d'index ne doit pas être partagée"


def test_t_clo_02_clone_fidele_en_memoire(tmp_path):
    """T-CLO-02 — valeurs, styles complets et propriétés de feuille sont tous repris."""
    gabarit = _gabarit_mini(tmp_path / "gabarit.xlsx")
    clone = _cloner(gabarit)
    _assert_fidele(_source(gabarit), clone)
    # Les cas que la copie d'origine perdait, nommés pour un message d'échec lisible :
    assert clone["A2"].quotePrefix
    assert clone["B2"].protection.locked is False
    assert all(largeurs_effectives(clone)[c][1:] == (True, 1) for c in "FGH")
    assert clone.page_setup.orientation == "landscape"
    assert clone.sheet_properties.pageSetUpPr.fitToPage
    assert clone.freeze_panes == "A2"
    assert clone.protection.sheet
    assert clone.print_title_rows == "$1:$1"


def test_t_clo_03_clone_fidele_apres_enregistrement(tmp_path):
    """T-CLO-03 — le clone enregistré puis relu reste identique au gabarit."""
    gabarit = _gabarit_mini(tmp_path / "gabarit.xlsx")
    wb = _classeur_ecole()
    _cloner(gabarit, wb)
    wb.save(tmp_path / "ecole.xlsx")
    _assert_fidele(_source(gabarit), load_workbook(tmp_path / "ecole.xlsx")["CLONE"])


def test_t_clo_04_voie_publique_identique_a_la_voie_rapide(tmp_path, monkeypatch):
    """T-CLO-04 — sans la structure interne d'openpyxl, la copie reste exacte."""
    gabarit = _gabarit_mini(tmp_path / "gabarit.xlsx")
    rapide = empreinte_cellules(_cloner(gabarit))
    monkeypatch.setattr(fmt, "StyleArray", None)
    lente = empreinte_cellules(_cloner(gabarit))
    assert rapide == lente, ecarts(rapide, lente)


def test_t_clo_05_un_style_modifie_ne_touche_pas_les_voisines(tmp_path):
    """T-CLO-05 — A1, A3, B3 et C3 partagent un style : en changer un laisse les autres."""
    clone = _cloner(_gabarit_mini(tmp_path / "gabarit.xlsx"))
    clone["B3"].font = Font(italic=True, color="FF0000")
    for coord in ("A1", "A3", "C3"):
        assert clone[coord].font.b and not clone[coord].font.i, coord


def test_t_clo_06_la_selection_d_onglet_n_est_pas_reprise(tmp_path):
    """T-CLO-06 — le gabarit est l'onglet actif de son fichier ; son clone ne doit pas être
    sélectionné en plus de l'onglet actif de l'école, sinon Excel groupe les deux onglets."""
    gabarit = _gabarit_mini(tmp_path / "gabarit.xlsx")
    assert _source(gabarit).sheet_view.tabSelected
    wb = _classeur_ecole()
    clone = _cloner(gabarit, wb)
    assert not clone.sheet_view.tabSelected
    assert wb.active.title == "R_TEST001"


def test_t_clo_07_fusion_en_echec_arrete_la_generation(tmp_path, monkeypatch):
    """T-CLO-07 — une fusion qui lève une erreur n'est plus ignorée : elle est nommée."""
    gabarit = _gabarit_mini(tmp_path / "gabarit.xlsx")

    def refuser(self, range_string=None, **_):
        raise ValueError("plage refusée")

    monkeypatch.setattr(Worksheet, "merge_cells", refuser)
    with pytest.raises(ValueError, match=r"Fusion A1:C1 du gabarit .*plage refusée"):
        _cloner(gabarit)


def test_t_clo_08_fusion_ignoree_sans_erreur_est_detectee(tmp_path, monkeypatch):
    """T-CLO-08 — openpyxl ignore sans erreur une plage incluse dans une autre : le
    contrôle après copie la signale quand même."""
    gabarit = _gabarit_mini(tmp_path / "gabarit.xlsx")
    monkeypatch.setattr(Worksheet, "merge_cells", lambda self, *a, **k: None)
    with pytest.raises(ValueError, match=r"absentes .*A1:C1, D4:E5"):
        _cloner(gabarit)


def test_t_clo_09_gabarit_lu_une_fois_puis_relu_s_il_change(tmp_path, monkeypatch):
    """T-CLO-09 — deux factures d'affilée ne relisent pas le gabarit ; un gabarit modifié
    est relu et sa nouvelle version est utilisée."""
    gabarit = _gabarit_mini(tmp_path / "gabarit.xlsx")
    lectures = []
    vrai_load = fmt.load_workbook
    monkeypatch.setattr(fmt, "load_workbook", lambda *a, **k: lectures.append(a) or vrai_load(*a, **k))
    wb = _classeur_ecole()
    _cloner(gabarit, wb, "F1")
    _cloner(gabarit, wb, "F2")
    assert len(lectures) == 1
    assert wb["F2"]["A1"].value == "Titre"

    _gabarit_mini(gabarit, titre="Titre v2")
    etat = gabarit.stat()
    os.utime(gabarit, ns=(etat.st_atime_ns, etat.st_mtime_ns + 2_000_000_000))
    assert _cloner(gabarit, wb, "F3")["A1"].value == "Titre v2"
    assert len(lectures) == 2
    assert len(fmt._PLANS_GABARIT) == 1, "l'ancienne version du gabarit doit quitter le cache"


def _mfc(wb, ws):
    ws.conditional_formatting.add("B2", CellIsRule(operator="lessThan", formula=["0"], font=Font(color="FF0000")))


def _validation(wb, ws):
    dv = DataValidation(type="list", formula1='"oui,non"')
    ws.add_data_validation(dv)
    dv.add("B4")


def _commentaire(wb, ws):
    ws["B4"].comment = Comment("note", "test")


def _style_nomme(wb, ws):
    wb.add_named_style(NamedStyle(name="Titre perso", font=Font(bold=True)))
    ws["B4"].style = "Titre perso"


def _tableau(wb, ws):
    ws["J1"], ws["K1"], ws["J2"], ws["K2"] = "Col1", "Col2", 1, 2
    ws.add_table(Table(displayName="T1", ref="J1:K2"))


@pytest.mark.parametrize(
    ("preparer", "message"),
    [
        (_mfc, "mise en forme conditionnelle"),
        (_validation, "validation des données"),
        (_tableau, "tableau structuré"),
        (_commentaire, "B4 du gabarit : commentaire"),
        (_style_nomme, "style nommé « Titre perso »"),
    ],
    ids=["mfc", "validation", "tableau", "commentaire", "style-nomme"],
)
def test_t_clo_10_le_non_recopiable_est_refuse(tmp_path, preparer, message):
    """T-CLO-10 — ce que le clonage ne sait pas recopier arrête la génération, au lieu
    de disparaître en silence de la facture."""
    gabarit = _gabarit_mini(tmp_path / "gabarit.xlsx", preparer=preparer)
    with pytest.raises(ValueError, match=message):
        _cloner(gabarit)


def test_t_clo_11_gabarit_sans_onglet_modele(tmp_path):
    """T-CLO-11 — un gabarit sans onglet MODELE donne une erreur qui liste ses onglets."""
    wb = Workbook()
    wb.active.title = "Feuil1"
    wb.save(tmp_path / "sans_modele.xlsx")
    with pytest.raises(ValueError, match="pas d'onglet MODELE .*Feuil1"):
        _cloner(tmp_path / "sans_modele.xlsx")
