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
# FA-10 — l'adresse iCal secrète n'est jamais affichée par `profil verifier`
# ──────────────────────────────────────────────────────────────────────

JETON = "private-0123456789abcdef"


def test_profil_verifier_refuse_une_adresse_sans_l_afficher(tmp_path):
    """Une adresse iCal collée dans le profil est refusée ; ni stdout ni stderr ne la recopient."""
    url = f"https://calendar.example.org/calendar/ical/personne%40example.org/{JETON}/basic.ics"
    profil = tmp_path / "profil.toml"
    profil.write_text(
        '[utilisateur]\nnom = "Personne A"\nentite = "ALPHA"\n\n'
        f'[agenda]\nfichier_url_secrete = "{url}"\n\n'
        f'[donnees]\nracine = "{(tmp_path / "donnees").as_posix()}"\n',
        encoding="utf-8",
    )
    env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONPATH": str(DEPOT),
           "APPDATA": str(tmp_path / "appdata"), "XDG_CONFIG_HOME": str(tmp_path / "config")}
    for variable in ("PCF_DATA_ROOT", "FACTURATION_PROFIL"):
        env.pop(variable, None)
    res = subprocess.run(
        [sys.executable, "-m", "facturation", "profil", "verifier", "--profil", str(profil)],
        capture_output=True, text=True, encoding="utf-8", env=env, cwd=tmp_path,
    )

    assert res.returncode != 0
    assert "fichier_url_secrete" in res.stdout + res.stderr
    assert JETON not in res.stdout + res.stderr
    assert "calendar.example.org" not in res.stdout + res.stderr
