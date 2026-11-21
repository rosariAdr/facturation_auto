"""Garde-fous de l'émission d'une facture par `scripts/nouvelle_facture.py`.

Revue system design du 18/11/2026 : FA-02, FA-03, FA-05, FA-09, FA-12, FA-13.
Données fictives uniquement, créées à la volée : entités ALPHA et BETA, écoles
ECOLEA et ECOLEB, coordonnées et montants inventés.

Le script est lancé en sous-processus avec `PYTHONPATH` pointé sur ce dépôt, pour
tester le code de cette copie et non celui d'une installation éditable voisine
(FA-13), et avec un `APPDATA` temporaire, pour ne jamais lire le profil réel.
"""

from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

pytestmark = pytest.mark.integration

DEPOT = Path(__file__).resolve().parents[2]
SCRIPT = DEPOT / "scripts" / "nouvelle_facture.py"

ENTETE_COORDONNEES = [
    "Type", "Code", "Nom / Raison sociale", "Adresse L1", "Adresse L2", "Adresse L3", "Email", "Téléphone",
]
DEST_ECOLEA = ["Destinataire", "ECOLEA", "SERVICE COMPTABLE ECOLE A", "1 avenue Fictive", "00000 VILLE-A", "", "compta@ecole-a.example"]


def creer_master_fictif(racine: Path, entite: str = "ALPHA", destinataires: list[list] | None = None) -> Path:
    """Écrit `racine/<entite>/<entite>_MASTER.xlsx` (onglet Coordonnées fictif) et renvoie son chemin."""
    dossier = racine / entite
    dossier.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Coordonnées"
    ws.append(ENTETE_COORDONNEES)
    ws.append(["Émetteur", entite, f"EMETTEUR FICTIF {entite}", "1 rue Fictive", "00000 VILLE-FICTIVE",
               "", f"{entite.lower()}@example.org", "+33000000000"])
    ws.append(["Info", entite, "SIRET", "00000000000000"])
    ws.append(["Info", entite, "IBAN", "FR76 0000 0000 0000 0000 0000 000"])
    ws.append(["Info", entite, "BIC", "TESTFRPP"])
    for ligne in destinataires if destinataires is not None else [DEST_ECOLEA]:
        ws.append(ligne)
    chemin = dossier / f"{entite}_MASTER.xlsx"
    wb.save(chemin)
    return chemin


def spec_fictive(**surcharges) -> dict:
    """Spec d'une facture fictive de l'entité ALPHA pour l'école ECOLEA."""
    spec = {
        "entite": "ALPHA",
        "ecole": "ECOLEA",
        "ecole_label": "ECOLE A",
        "fichier_sortie": "ALPHA/ECOLEA/ALPHA_ECOLEA.xlsx",
        "onglet": "A_ECOLEA001 Sept 2026",
        "n_facture": "A_ECOLEA001",
        "date_emission": "2026-10-01",
        "lignes": [
            {"promotion": "PROMO 1", "libelle": "Module fictif",
             "seances": [{"date": "2026-09-15", "heures": 3, "taux": 50}]},
        ],
    }
    spec.update(surcharges)
    return spec


def lancer(tmp_path: Path, spec: dict, *options: str) -> subprocess.CompletedProcess:
    """Écrit la spec dans `tmp_path` et lance le script dessus, isolé du profil réel."""
    chemin_spec = tmp_path / f"facture_{len(list(tmp_path.glob('facture_*.json')))}.json"
    chemin_spec.write_text(json.dumps(spec), encoding="utf-8")
    env = {
        **os.environ,
        "PYTHONIOENCODING": "utf-8",
        "PYTHONPATH": str(DEPOT),
        "APPDATA": str(tmp_path / "appdata"),
        "XDG_CONFIG_HOME": str(tmp_path / "config"),
    }
    for variable in ("PCF_DATA_ROOT", "FACTURATION_PROFIL"):
        env.pop(variable, None)
    return subprocess.run(
        [sys.executable, str(SCRIPT), str(chemin_spec), *options],
        capture_output=True, text=True, encoding="utf-8", env=env, cwd=tmp_path,
    )


def empreinte(chemin: Path) -> str:
    return hashlib.sha256(chemin.read_bytes()).hexdigest()


@pytest.fixture
def racine(tmp_path: Path) -> Path:
    racine = tmp_path / "donnees"
    creer_master_fictif(racine)
    return racine


# ──────────────────────────────────────────────────────────────────────
# FA-09 — un destinataire sans adresse bloque la facture (PROC00 étape 4.3)
# ──────────────────────────────────────────────────────────────────────

from facturation import config, verification  # noqa: E402

DEST_SANS_ADRESSE = ["Destinataire", "ECOLEB", "ECOLE B", "", "", "", "compta@ecole-b.example"]


@pytest.fixture
def racine_ecoleb_sans_adresse(tmp_path: Path) -> Path:
    racine = tmp_path / "donnees"
    master = creer_master_fictif(racine, destinataires=[DEST_ECOLEA, DEST_SANS_ADRESSE])
    wb = load_workbook(master)
    wb.create_sheet("Sessions")  # onglet obligatoire pour `profil verifier`
    wb.save(master)
    return racine


def _spec_ecoleb() -> dict:
    return spec_fictive(ecole="ECOLEB", ecole_label="ECOLE B", fichier_sortie="ALPHA/ECOLEB/ALPHA_ECOLEB.xlsx",
                        onglet="A_ECOLEB001 Sept 2026", n_facture="A_ECOLEB001")


@pytest.mark.parametrize("options", [(), ("--dry-run",)], ids=["ecriture", "dry-run"])
def test_destinataire_sans_adresse_bloque(tmp_path, racine_ecoleb_sans_adresse, options):
    res = lancer(tmp_path, _spec_ecoleb(), "--data-root", str(racine_ecoleb_sans_adresse), *options)

    assert res.returncode != 0
    assert "ECOLEB" in res.stderr and "Adresse L1" in res.stderr and "Adresse L2" in res.stderr
    assert "APERÇU" not in res.stdout
    assert not (racine_ecoleb_sans_adresse / "ALPHA" / "ECOLEB").exists()


def test_profil_verifier_signale_le_destinataire_incomplet(tmp_path, racine_ecoleb_sans_adresse):
    profil = tmp_path / "profil.toml"
    profil.write_text(config.rediger("Personne A", "ALPHA", str(racine_ecoleb_sans_adresse)), encoding="utf-8")
    constats = verification.verifier(config.charger(profil))

    incomplets = [c for c in constats if c.sujet == "Coordonnées" and "incomplets" in c.message]
    assert len(incomplets) == 1
    assert incomplets[0].niveau == verification.ATTENTION
    assert "ECOLEB (Adresse L1, Adresse L2)" in incomplets[0].message
    assert "ECOLEA" not in incomplets[0].message
