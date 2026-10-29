"""Non-régression du format des factures (banque de tests : T-FMT-01 à T-FMT-07).

T-FMT-01 à 06 reprennent les tests de la v1 (`scripts/test_pcf_format.py`), sur un
MASTER fictif. T-FMT-07 trace les défauts relevés le 26/10/2026 : marqué xfail
strict tant que F-18 (fiabilisation du gabarit) n'est pas livré ; le jour où il
passe, pytest échoue pour rappeler de retirer le marqueur.
"""

from __future__ import annotations

import re
import subprocess
from datetime import date

import pytest
from openpyxl import Workbook

from facturation import pcf_format as fmt
from facturation.verification import trouver_libreoffice

pytestmark = pytest.mark.regression


def _classeur() -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    return wb


def test_t_fmt_01_clone_du_gabarit(racine_fictive):
    """T-FMT-01 — le gabarit versionné est vide de coordonnées, sa structure est intacte."""
    ws = fmt.clone_template(_classeur(), "TEST_CLONE", entite="ROSARI")
    pm = fmt.post_mortem(ws, entite="ROSARI")
    assert pm["total_formule"] == "=SUM(Q20:R79)"
    assert pm["emetteur_nom"] is None
    assert pm["iban"] is None
    assert pm["destinataire_nom"] is None


def test_t_fmt_02_coordonnees_injectees(racine_fictive):
    """T-FMT-02 — émetteur, SIRET, IBAN et BIC injectés depuis le MASTER."""
    ws = fmt.clone_template(_classeur(), "TEST_COORDS", entite="ROSARI")
    fmt.apply_emetteur(ws, entite="ROSARI")
    fmt.apply_metadata(ws, n_facture="R_TEST002", date_emission=date(2026, 4, 30), entite="ROSARI")
    fmt.apply_rib(ws, entite="ROSARI")
    pm = fmt.post_mortem(ws, entite="ROSARI")
    assert pm["emetteur_nom"] == "FORMATEUR FICTIF ROSARI"
    assert pm["siret"] == "00000000000000"
    assert pm["iban"].startswith("FR76")
    assert pm["bic"] == "TESTFRPP"
    assert ws["C16"].value == "N° SIRET :"
    assert ws["C17"].value == "N° de Facture :"


def test_t_fmt_03_destinataire_complet(racine_fictive):
    """T-FMT-03 — la rue ne disparaît pas dans la fusion M7:P8."""
    ws = fmt.clone_template(_classeur(), "TEST_DEST", entite="ROSARI")
    dest = {
        "nom": "ECOLE TEST", "adresse_l1": "1 rue de la Paix",
        "adresse_l2": "75002 Paris", "adresse_l3": "France", "email": "compta@test.fr",
    }
    fmt.apply_destinataire(ws, dest=dest, entite="ROSARI")
    lues = [ws.cell(row=r, column=13).value for r in range(7, 11)]
    assert lues == ["ECOLE TEST", "1 rue de la Paix", "75002 Paris, France", "compta@test.fr"]


def test_t_fmt_04_destinataire_vide(racine_fictive):
    """T-FMT-04 — dest=None laisse l'encadré, sans aucune valeur."""
    ws = fmt.clone_template(_classeur(), "TEST_DEST_VIDE", entite="ROSARI")
    fmt.apply_destinataire(ws, dest=None, entite="ROSARI")
    assert all(ws.cell(row=r, column=13).value is None for r in range(7, 11))


def test_t_fmt_05_une_seance(racine_fictive):
    """T-FMT-05 — une séance : ligne 20 remplie, 21 à 79 masquées, formule correcte."""
    ws = fmt.clone_template(_classeur(), "TEST_1_SESSION", entite="ROSARI")
    fmt.apply_metadata(ws, n_facture="R_TEST005", date_emission=date(2026, 4, 30), entite="ROSARI")
    groups = [("BTS SIO A1", "Maths BTS SIO", [(date(2026, 3, 12), 3.5, 60)])]
    fmt.place_sessions(ws, ecole_name="ECOLE B", groups=groups, entite="ROSARI")
    pm = fmt.post_mortem(ws, entite="ROSARI")
    assert pm["n_facture"] == "R_TEST005"
    assert pm["ecole"] == "ECOLE B"
    assert pm["premier_libelle"] == "Maths BTS SIO"
    assert "30/04/2026" in str(pm["titre"])
    assert sum(1 for r in range(21, 80) if ws.row_dimensions[r].hidden) == 59
    assert ws.cell(row=20, column=13).value == 3.5
    assert ws.cell(row=20, column=15).value == 60
    assert ws.cell(row=20, column=17).value == "=O20*M20"


def test_t_fmt_06_bon_de_commande_et_jour_masque(racine_fictive):
    """T-FMT-06 — ligne Bon de commande écrite, colonnes Jour masquées à la demande."""
    ws = fmt.clone_template(_classeur(), "TEST_BC", entite="ROSARI")
    fmt.apply_metadata(
        ws, n_facture="R_TEST006", date_emission=date(2026, 7, 31),
        entite="ROSARI", ref_bon_commande="PO-TEST-0001",
    )
    fmt.place_sessions(
        ws, ecole_name="EPITA", groups=[("AIS S1", "UE de test", [(date(2026, 7, 31), 18, 90)])],
        entite="ROSARI", masquer_jour=True,
    )
    fmt.apply_print_setup(ws, entite="ROSARI")
    assert ws["C14"].value == "Bon de commande :"
    assert ws["D14"].value == "PO-TEST-0001"
    assert ws.column_dimensions["K"].hidden and ws.column_dimensions["L"].hidden
    assert ws.page_setup.orientation == "portrait"
    assert ws.print_area


def test_gabarit_standard_pour_toute_entite(tmp_path):
    """ADR015 — une entité sans gabarit propre utilise le gabarit standard."""
    assert fmt.gabarit_pour("ROSARI") == fmt.GABARIT_STANDARD
    assert fmt.gabarit_pour("KAUR") == fmt.GABARIT_STANDARD
    with pytest.raises(ValueError):
        fmt.gabarit_pour("rosari")
    with pytest.raises(ValueError):
        fmt.gabarit_pour("STANDARD")


def _facture_17_seances(racine):
    """Facture fictive structurée comme R_DOR005 : 17 séances sur 4 promotions."""
    wb = _classeur()
    ws = fmt.clone_template(wb, "R_DOR999 Sept 2026", entite="ROSARI")
    fmt.apply_emetteur(ws, entite="ROSARI")
    fmt.apply_destinataire(ws, dest=fmt.get_destinataire("DOR"), entite="ROSARI")
    fmt.apply_metadata(ws, n_facture="R_DOR999", date_emission=date(2026, 10, 1), entite="ROSARI")
    j = lambda n: date(2026, 9, n)  # noqa: E731
    groups = [
        ("BTS SIO A1", "Mathématiques pour l'informatique", [(j(n), 3.5, 60) for n in (9, 11, 15, 21, 28)]),
        ("B3", "Maths pour la data et l'IA : remise à niveau", [(j(n), 3.5, 60) for n in (16, 16, 17, 17)]),
        ("B3", "Conception de base de données avec MySQL et PostgreSQL", [(j(18), 3.5, 60)] * 2),
        ("M1", "Statistiques et régression", [(j(n), 3.5, 65) for n in (23, 23, 24, 24, 25, 25)]),
    ]
    fmt.place_sessions(ws, ecole_name="ECOLE B", groups=groups, entite="ROSARI")
    fmt.apply_rib(ws, entite="ROSARI")
    fmt.apply_print_setup(ws, entite="ROSARI")
    return wb, ws


@pytest.mark.xfail(raises=AssertionError, reason="T-FMT-07 / F-18 : lignes 27 à 79 du gabarit sans format de nombre")
def test_t_fmt_07a_formats_homogenes_sur_toutes_les_lignes(racine_fictive):
    """T-FMT-07a — heures, taux et totaux gardent le même format au-delà de la 7e séance."""
    _, ws = _facture_17_seances(racine_fictive)
    cols = {"heures": 13, "taux": 15, "total": 17}
    reference = {nom: ws.cell(row=20, column=c).number_format for nom, c in cols.items()}
    for r in range(20, 37):
        for nom, c in cols.items():
            assert ws.cell(row=r, column=c).number_format == reference[nom], (r, nom)


@pytest.mark.skipif(trouver_libreoffice() is None, reason="LibreOffice absent")
@pytest.mark.xfail(raises=AssertionError, reason="T-FMT-07 / F-18 : les mentions légales passent en page 2")
def test_t_fmt_07b_rendu_pdf_sur_une_page(racine_fictive, tmp_path):
    """T-FMT-07b — la facture de 17 séances tient sur une page, mentions légales comprises."""
    wb, _ = _facture_17_seances(racine_fictive)
    xlsx = tmp_path / "R_DOR999.xlsx"
    wb.save(xlsx)
    subprocess.run(
        [str(trouver_libreoffice()), "--headless", "--convert-to", "pdf", "--outdir", str(tmp_path), str(xlsx)],
        check=True, capture_output=True, timeout=180,
    )
    pdf = (tmp_path / "R_DOR999.pdf").read_bytes()
    pages = len(re.findall(rb"/Type\s*/Page(?!s)", pdf))
    assert pages == 1, f"{pages} pages"
