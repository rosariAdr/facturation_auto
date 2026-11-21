"""Contrôles purs de `scripts/nouvelle_facture.py`, sur des classeurs fictifs en mémoire.

Revue system design du 18/11/2026 : FA-02, FA-03, FA-05, FA-09, FA-12.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest
from openpyxl import Workbook

pytestmark = pytest.mark.unit

DEPOT = Path(__file__).resolve().parents[2]


def _charger_script():
    """Importe `scripts/nouvelle_facture.py`, qui n'est pas un paquet, par son chemin."""
    spec = importlib.util.spec_from_file_location("nouvelle_facture", DEPOT / "scripts" / "nouvelle_facture.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


nf = _charger_script()


def _classeur(*onglets: tuple[str, str | None]) -> Workbook:
    """Classeur en mémoire : un onglet par (titre, numéro écrit en D17)."""
    wb = Workbook()
    wb.remove(wb.active)
    for titre, numero in onglets:
        wb.create_sheet(titre)["D17"] = numero
    return wb


# ──────────────────────────────────────────────────────────────────────
# FA-02 — conflits avec le classeur cible
# ──────────────────────────────────────────────────────────────────────


def test_classeur_absent_aucun_conflit():
    assert nf.conflits_ecriture(None, "A_ECOLEA001 Sept 2026", "A_ECOLEA001", entite="ALPHA") == []


def test_onglet_identique_ou_a_la_casse_pres():
    wb = _classeur(("A_ECOLEA001 Sept 2026", "A_ECOLEA001"))
    assert nf.conflits_ecriture(wb, "A_ECOLEA001 Sept 2026", "A_ECOLEA009", entite="ALPHA")
    assert nf.conflits_ecriture(wb, "a_ecolea001 SEPT 2026", "A_ECOLEA009", entite="ALPHA")


def test_numero_dans_le_titre_ou_dans_la_cellule():
    par_titre = _classeur(("PAID A_ECOLEA001 Sept 2026", None))
    par_cellule = _classeur(("Facture réglée", "A_ECOLEA001"))
    for wb in (par_titre, par_cellule):
        conflits = nf.conflits_ecriture(wb, "A_ECOLEA001 Oct 2026", "A_ECOLEA001", entite="ALPHA")
        assert len(conflits) == 1 and "A_ECOLEA001" in conflits[0]


def test_controle_ne_cree_aucune_cellule_dans_les_autres_onglets():
    """Le contrôle lit les onglets sans les agrandir : le save ne doit pas les réécrire autrement."""
    wb = _classeur()
    notes = wb.create_sheet("Notes")
    notes["A1"] = "note fictive"
    assert nf.conflits_ecriture(wb, "A_ECOLEA001 Sept 2026", "A_ECOLEA001", entite="ALPHA") == []
    assert (notes.max_row, notes.max_column) == (1, 1)


def test_numero_voisin_n_est_pas_un_conflit():
    """A_ECOLEA001 ne doit pas être confondu avec A_ECOLEA0010 ni avec A_ECOLEA00."""
    wb = _classeur(("A_ECOLEA0010 Sept 2026", "A_ECOLEA0010"), ("A_ECOLEA00 Août 2026", "A_ECOLEA00"))
    assert nf.conflits_ecriture(wb, "A_ECOLEA001 Oct 2026", "A_ECOLEA001", entite="ALPHA") == []
