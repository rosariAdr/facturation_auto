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
# FA-09 — destinataire incomplet
# ──────────────────────────────────────────────────────────────────────

from facturation import pcf_format as fmt  # noqa: E402

COMPLET = {"nom": "SERVICE COMPTABLE ECOLE A", "adresse_l1": "1 avenue Fictive", "adresse_l2": "00000 VILLE-A",
           "adresse_l3": "", "email": ""}


def test_destinataire_complet_sans_email_accepte():
    """L'email et le pays (L3) restent facultatifs, comme dans apply_destinataire."""
    assert fmt.champs_destinataire_manquants(COMPLET) == []


@pytest.mark.parametrize(
    ("champ", "libelle"),
    [("nom", "Nom / Raison sociale"), ("adresse_l1", "Adresse L1"), ("adresse_l2", "Adresse L2")],
)
@pytest.mark.parametrize("vide", ["", "   ", None])
def test_champ_obligatoire_manquant(champ, libelle, vide):
    assert fmt.champs_destinataire_manquants({**COMPLET, champ: vide}) == [libelle]


def test_destinataire_sans_adresse():
    sans_adresse = {**COMPLET, "adresse_l1": "", "adresse_l2": "", "email": "compta@ecole-b.example"}
    assert fmt.champs_destinataire_manquants(sans_adresse) == ["Adresse L1", "Adresse L2"]
