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
# FA-10 — l'adresse iCal secrète n'est jamais recopiée
# ──────────────────────────────────────────────────────────────────────

from facturation import config, verification  # noqa: E402

JETON = "private-0123456789abcdef"
URL_FICTIVE = f"https://calendar.example.org/calendar/ical/personne%40example.org/{JETON}/basic.ics"


def _brut(fichier_url_secrete) -> dict:
    return {
        "utilisateur": {"nom": "Personne A", "entite": "ALPHA"},
        "agenda": {"fichier_url_secrete": fichier_url_secrete},
        "donnees": {"racine": "/donnees"},
    }


@pytest.mark.parametrize("valeur", [URL_FICTIVE, f"webcal://calendar.example.org/{JETON}.ics"])
def test_adresse_collee_a_la_place_du_fichier_refusee_sans_recopie(tmp_path, valeur):
    with pytest.raises(config.ProfilInvalide) as exc:
        config.valider(_brut(valeur), tmp_path / "profil.toml")
    assert any(e.startswith("[agenda] fichier_url_secrete") for e in exc.value.erreurs)
    assert JETON not in str(exc.value)
    assert "calendar.example.org" not in str(exc.value)


def test_nom_de_fichier_accepte(tmp_path):
    profil = config.valider(_brut("agenda.secret"), tmp_path / "profil.toml")
    assert profil.agenda_fichier_url_secrete == tmp_path / "agenda.secret"


def _profil(fichier: Path) -> config.Profil:
    return config.Profil(
        chemin=Path("profil.toml"), nom="Personne A", entite="ALPHA", prefixe_facture="A_",
        agenda_source="ical", agenda_fichier_url_secrete=fichier, fuseau="Europe/Paris",
        racine=Path("donnees"), lancement_jour=25, lancement_heure="16:00",
        date_emission="premier_du_mois_suivant", pdf_moteur="libreoffice",
    )


def test_verification_n_affiche_que_le_nom_du_fichier(tmp_path):
    """Fichier introuvable : seul son nom s'affiche, pas le dossier qui le contient."""
    dossier = tmp_path / "dossier-personnel-0042"
    (constat,) = verification._verifier_agenda(_profil(dossier / "agenda.secret"))
    assert "agenda.secret" in constat.message
    assert "dossier-personnel-0042" not in constat.message


def test_verification_ne_recopie_pas_une_adresse_passee_en_chemin():
    """Défense en profondeur : un Profil construit sans `valider` ne fait pas fuir le jeton."""
    (constat,) = verification._verifier_agenda(_profil(Path(URL_FICTIVE)))
    assert JETON not in constat.message
