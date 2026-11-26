"""Recette : une facture complète, du gabarit au fichier relu sur le disque (T-REC-01 à 03).

Ces tests regardent le résultat comme le ferait une relecture humaine du classeur :
rien de ce que le gabarit apporte ne doit se perdre en route, hors des zones que la
mise en forme réécrit volontairement : bandeau, destinataire, bon de commande, séances
et mentions légales (F-18, F-24).
Données fictives uniquement.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from facturation import pcf_format as fmt
from tests.conftest import creer_master
from tests.empreinte import ecarts, empreinte_cellules, empreinte_feuille, largeurs_effectives

pytestmark = pytest.mark.recette

DEPOT = Path(__file__).resolve().parents[2]
SPEC_EXEMPLE = DEPOT / "examples" / "spec_exemple.json"
L = fmt.LAYOUT_STANDARD


def _gabarit():
    return load_workbook(fmt.GABARIT_STANDARD)["MODELE"]


def _zone_reecrite(plage) -> bool:
    """Vrai si la fusion touche une zone que la mise en forme refait volontairement."""
    d_min, d_max = L["destinataire_rows"]
    col_dest = L["destinataire_col"]
    seances = L["cours_first_row"] <= plage.max_row and plage.min_row <= L["cours_last_row"]
    destinataire = (d_min <= plage.max_row and plage.min_row <= d_max
                    and plage.max_col >= col_dest and plage.min_col <= col_dest + 3)
    bon_de_commande = plage.min_row <= L["bon_commande_row"] <= plage.max_row
    # F-18 refusionne le bandeau dans la zone d'impression (B2:S3 devient B2:R3) ; F-24 réécrit
    # les mentions légales sous le total et en refait les fusions (TVA sur C:K).
    bandeau = plage.min_row <= L["bandeau_row"] + 1 and plage.max_row >= L["bandeau_row"]
    mentions = L["net_a_payer_row"] <= plage.max_row and plage.min_row <= L["derniere_ligne_utile"]
    return seances or destinataire or bon_de_commande or bandeau or mentions


def _colonnes_retaillees(facture) -> set:
    """Colonnes que F-18 retaille exprès : marges et libellés (apply_metadata), et dernier
    libellé du tableau, qui reprend la largeur des colonnes Jour masquées (place_sessions)."""
    lettres = {get_column_letter(c) for c in range(1, L["bon_commande_label_col"] + 1)}
    if facture.column_dimensions[get_column_letter(L["cols_table"]["date"])].hidden:
        lettres.add(get_column_letter(L["cols_table"]["libelle"] + 3))
    return lettres


def _verifier_structure_preservee(gabarit, facture) -> None:
    """Ce que le gabarit apporte hors des zones réécrites se retrouve dans la facture."""
    fusions = {str(p) for p in facture.merged_cells.ranges}
    perdues = sorted(str(p) for p in gabarit.merged_cells.ranges if not _zone_reecrite(p) and str(p) not in fusions)
    assert not perdues, f"{facture.title} : fusions du gabarit perdues : {perdues[:10]}"

    # Largeurs : hors des colonnes que F-18 retaille (effet vérifié par test_gabarit.py : libellés
    # entiers, facture sur une page), la facture peut masquer des colonnes (Jour), pas en changer
    # la largeur.
    retaillees = _colonnes_retaillees(facture)
    attendu = {k: v[0] for k, v in largeurs_effectives(gabarit).items() if k not in retaillees}
    obtenu = {k: v[0] for k, v in largeurs_effectives(facture).items() if k in attendu}
    assert obtenu == attendu, f"{facture.title} : {ecarts(attendu, obtenu)}"

    # Sous les mentions légales, la feuille est celle du gabarit, cellule pour cellule.
    sous = L["derniere_ligne_utile"]
    attendu = {k: v for k, v in empreinte_cellules(gabarit).items() if k[0] > sous}
    obtenu = {k: v for k, v in empreinte_cellules(facture).items() if k[0] > sous}
    assert attendu, "le gabarit devrait porter des cellules stylées sous les mentions"
    assert obtenu == attendu, f"{facture.title} : {ecarts(attendu, obtenu)}"


def _facture(wb: Workbook, onglet: str, n_facture: str, dest, groups, **meta) -> None:
    """Enchaîne les étapes de nouvelle_facture.py, sans passer par le script."""
    ws = fmt.clone_template(wb, onglet, entite="ROSARI")
    fmt.apply_emetteur(ws, entite="ROSARI")
    fmt.apply_destinataire(ws, dest=dest, entite="ROSARI")
    fmt.apply_metadata(ws, n_facture=n_facture, date_emission=date(2026, 12, 1), entite="ROSARI", **meta)
    fmt.place_sessions(ws, ecole_name="ECOLE B", groups=groups, entite="ROSARI")
    fmt.apply_mentions(ws, entite="ROSARI", date_emission=date(2026, 12, 1))
    fmt.apply_rib(ws, entite="ROSARI")
    fmt.apply_print_setup(ws, entite="ROSARI")


def test_t_rec_01_clone_enregistre_identique_au_gabarit(tmp_path):
    """T-REC-01 — le gabarit cloné, enregistré puis relu, est identique au gabarit :
    les index de style posés par la voie rapide restent justes une fois sur le disque."""
    wb = Workbook()
    wb.active.title = "R_TEST001"
    fmt.clone_template(wb, "CLONE", entite="ROSARI")
    wb.save(tmp_path / "ecole.xlsx")
    gabarit, relu = _gabarit(), load_workbook(tmp_path / "ecole.xlsx")["CLONE"]
    attendu, obtenu = empreinte_cellules(gabarit), empreinte_cellules(relu)
    assert obtenu == attendu, ecarts(attendu, obtenu)
    attendu, obtenu = empreinte_feuille(gabarit), empreinte_feuille(relu)
    assert obtenu == attendu, ecarts(attendu, obtenu)


def test_t_rec_02_deux_factures_dans_le_meme_classeur(racine_fictive, tmp_path):
    """T-REC-02 — deux factures d'affilée dans le classeur d'une école (la seconde clone le
    gabarit depuis le cache) : chacune est complète, et elles ne déteignent pas l'une sur l'autre."""
    wb = Workbook()
    wb.remove(wb.active)
    j = lambda n: date(2026, 11, n)  # noqa: E731
    _facture(
        wb, "R_TEST101 Nov 2026", "R_TEST101", fmt.get_destinataire("ECO", entite="ROSARI"),
        [("B3", "Statistiques", [(j(n), 3.5, 60) for n in (3, 5, 12)]),
         ("M1", "Régression", [(j(n), 7, 65) for n in (18, 19)])],
        ref_bon_commande="PO-TEST-101",
    )
    _facture(
        wb, "R_TEST102 Nov 2026", "R_TEST102", None,
        [("BTS SIO A1", "Maths", [(j(24), 2, 52)])],
    )
    wb.save(tmp_path / "ecole.xlsx")

    relu = load_workbook(tmp_path / "ecole.xlsx")
    gabarit = _gabarit()
    f1, f2 = relu["R_TEST101 Nov 2026"], relu["R_TEST102 Nov 2026"]
    pm1, pm2 = fmt.post_mortem(f1, entite="ROSARI"), fmt.post_mortem(f2, entite="ROSARI")
    assert (pm1["n_facture"], pm2["n_facture"]) == ("R_TEST101", "R_TEST102")
    assert pm1["destinataire_nom"] == "ÉCOLE FICTIVE B" and pm2["destinataire_nom"] is None
    assert f1["D14"].value == "PO-TEST-101" and f2["D14"].value is None
    for facture, pm in ((f1, pm1), (f2, pm2)):
        assert pm["emetteur_nom"] == "FORMATEUR FICTIF ROSARI EI"
        assert pm["iban"].startswith("FR76")
        assert pm["total_formule"] == "=SUM(Q20:R79)"
        assert facture.page_setup.orientation == "portrait" and facture.print_area
        _verifier_structure_preservee(gabarit, facture)
    selectionnes = [ws.title for ws in relu.worksheets if ws.sheet_view.tabSelected]
    assert len(selectionnes) <= 1, f"onglets groupés à l'ouverture dans Excel : {selectionnes}"


def test_t_rec_03_facture_produite_par_le_script(tmp_path):
    """T-REC-03 — la facture écrite par scripts/nouvelle_facture.py, telle qu'on l'ouvre
    ensuite dans Excel : complète, et fidèle au gabarit hors des zones réécrites."""
    racine = tmp_path / "donnees"
    creer_master(racine)
    env = {**os.environ, "PYTHONIOENCODING": "utf-8"}
    env.pop("PCF_DATA_ROOT", None)
    res = subprocess.run(
        [sys.executable, "scripts/nouvelle_facture.py", str(SPEC_EXEMPLE), "--data-root", str(racine)],
        capture_output=True, text=True, encoding="utf-8", env=env, cwd=DEPOT,
    )
    assert res.returncode == 0, res.stderr
    spec = json.loads(SPEC_EXEMPLE.read_text(encoding="utf-8"))
    facture = load_workbook(racine / spec["fichier_sortie"])[spec["onglet"]]
    assert facture["D17"].value == spec["n_facture"]
    assert facture["M7"].value == "ÉCOLE FICTIVE A"
    assert facture.column_dimensions["K"].hidden, "colonne Jour masquée (masquer_jour)"
    _verifier_structure_preservee(_gabarit(), facture)
