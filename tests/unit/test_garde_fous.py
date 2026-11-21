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
# FA-05 — aperçu complet
# ──────────────────────────────────────────────────────────────────────

from datetime import date  # noqa: E402

EMETTEUR = {"nom": "EMETTEUR FICTIF ALPHA", "adresse_l1": "1 rue Fictive", "adresse_l2": "00000 VILLE-FICTIVE"}
DEST = {"nom": "SERVICE COMPTABLE ECOLE A", "adresse_l1": "1 avenue Fictive", "adresse_l2": "00000 VILLE-A"}
SPEC = {"entite": "ALPHA", "n_facture": "A_ECOLEA001", "date_emission": "2026-10-01",
        "fichier_sortie": "ALPHA/ECOLEA/ALPHA_ECOLEA.xlsx", "onglet": "A_ECOLEA001 Sept 2026"}


def test_apercu_libelles_complets_et_colonnes_alignees():
    court = "Module court"
    long = "[2026][CODE-PRODUIT-FICTIF-0002][PROMO 2][Libellé recopié à l'identique du bon de commande]"
    groupes = [("PROMO 1", court, [(date(2026, 9, 15), 3.0, 50.0)]),
               ("PROMO 2", long, [(date(2026, 9, 16), 2.0, 50.0), (date(2026, 9, 17), 2.0, 50.0)])]
    texte = nf.apercu(SPEC, groupes, DEST, emetteur=EMETTEUR)

    lignes = [ligne for ligne in texte.splitlines() if ligne.startswith("  | ")]
    assert len(lignes) == 3
    assert court in lignes[0] and all(long in ligne for ligne in lignes[1:])
    assert len({ligne.index("h |") for ligne in lignes}) == 1, "colonnes décalées"
    assert "Émetteur      : EMETTEUR FICTIF ALPHA — 1 rue Fictive, 00000 VILLE-FICTIVE" in texte


def test_apercu_promotion_numerique_acceptee():
    """Une promotion saisie comme nombre dans la spec JSON s'affiche, comme avant le correctif."""
    groupes = [(2026, "Module fictif", [(date(2026, 9, 15), 3.0, 50.0)])]
    assert "| 2026 " in nf.apercu(SPEC, groupes, DEST, emetteur=EMETTEUR)


def test_apercu_libelle_sur_plusieurs_lignes_reste_sur_une_ligne():
    groupes = [("PROMO 1", "Partie 1\nPartie 2", [(date(2026, 9, 15), 3.0, 50.0)])]
    texte = nf.apercu(SPEC, groupes, DEST, emetteur=EMETTEUR)
    assert "Partie 1 Partie 2" in texte
