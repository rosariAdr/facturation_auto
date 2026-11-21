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
# FA-03 — entité explicite (ADR001), préfixe et dossier de sortie
# ──────────────────────────────────────────────────────────────────────

import inspect  # noqa: E402

from facturation import pcf_format as fmt  # noqa: E402

API_PUBLIQUE = [
    fmt.get_destinataire, fmt.clone_template, fmt.apply_emetteur, fmt.apply_destinataire,
    fmt.apply_metadata, fmt.place_sessions, fmt.apply_print_setup, fmt.apply_rib, fmt.post_mortem,
]


@pytest.mark.parametrize("fonction", API_PUBLIQUE, ids=lambda f: f.__name__)
def test_entite_obligatoire_et_nommee(fonction):
    """ADR001 : aucune entité par défaut ; `entite=` se passe toujours par son nom."""
    parametre = inspect.signature(fonction).parameters["entite"]
    assert parametre.kind is inspect.Parameter.KEYWORD_ONLY
    assert parametre.default is inspect.Parameter.empty


def test_oublier_l_entite_leve_une_erreur():
    with pytest.raises(TypeError, match="entite"):
        fmt.clone_template(Workbook(), "SANS_ENTITE")


RACINE = Path("donnees")


def test_spec_coherente_acceptee():
    spec = {"entite": "ALPHA", "n_facture": "A_ECOLEA001", "fichier_sortie": "ALPHA/ECOLEA/ALPHA_ECOLEA.xlsx"}
    assert nf.controler_spec(spec, prefixe="A_", racine=RACINE) == RACINE / "ALPHA" / "ECOLEA" / "ALPHA_ECOLEA.xlsx"


@pytest.mark.parametrize("entite", ["alpha", "", "STANDARD"])
def test_entite_invalide_refusee_proprement(entite):
    """Un code d'entité invalide donne un refus nommé, pas une trace d'exception."""
    spec = {"entite": entite, "n_facture": "A_ECOLEA001", "fichier_sortie": "ALPHA/ECOLEA/ALPHA_ECOLEA.xlsx"}
    with pytest.raises(nf.EcritureRefusee, match="entite"):
        nf.controler_spec(spec, prefixe=None, racine=RACINE)


def test_prefixe_par_defaut_quand_aucun_n_est_donne():
    spec = {"entite": "ALPHA", "n_facture": "B_ECOLEA001", "fichier_sortie": "ALPHA/ECOLEA/ALPHA_ECOLEA.xlsx"}
    with pytest.raises(nf.EcritureRefusee, match="préfixe A_"):
        nf.controler_spec(spec, prefixe=None, racine=RACINE)


@pytest.mark.parametrize("n_facture", ["B_ECOLEA001", "a_ECOLEA001", "ECOLEA001", "XA_ECOLEA001"])
def test_numero_hors_prefixe_refuse(n_facture):
    spec = {"entite": "ALPHA", "n_facture": n_facture, "fichier_sortie": "ALPHA/ECOLEA/ALPHA_ECOLEA.xlsx"}
    with pytest.raises(nf.EcritureRefusee, match="préfixe A_"):
        nf.controler_spec(spec, prefixe="A_", racine=RACINE)


@pytest.mark.parametrize("fichier_sortie", [
    "/tmp/ALPHA/x.xlsx", "C:/ALPHA/x.xlsx", "C:ALPHA/x.xlsx", r"\\serveur\partage\ALPHA\x.xlsx",
    "ALPHA/../BETA/x.xlsx", r"ALPHA\..\BETA\x.xlsx", "ALPHA/ECOLEA/../../x.xlsx",
    "BETA/ECOLEA/x.xlsx", "x.xlsx", "ALPHA", "alpha/ECOLEA/x.xlsx", "",
])
def test_fichier_de_sortie_hors_du_dossier_de_l_entite_refuse(fichier_sortie):
    spec = {"entite": "ALPHA", "n_facture": "A_ECOLEA001", "fichier_sortie": fichier_sortie}
    with pytest.raises(nf.EcritureRefusee, match="fichier_sortie"):
        nf.controler_spec(spec, prefixe="A_", racine=RACINE)
