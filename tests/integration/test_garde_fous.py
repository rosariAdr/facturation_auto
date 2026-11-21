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
# FA-03 — entité, préfixe du numéro et dossier de sortie concordent
# ──────────────────────────────────────────────────────────────────────


@pytest.fixture
def deux_entites(tmp_path: Path) -> Path:
    """Racine fictive à deux entités : ALPHA et BETA, chacune son MASTER et son émetteur."""
    racine = tmp_path / "donnees"
    creer_master_fictif(racine, "ALPHA")
    creer_master_fictif(racine, "BETA")
    return racine


def spec_beta(**surcharges) -> dict:
    """Spec cohérente de l'entité BETA ; les surcharges introduisent l'incohérence testée."""
    beta = {"entite": "BETA", "n_facture": "B_ECOLEA001", "onglet": "B_ECOLEA001 Sept 2026",
            "fichier_sortie": "BETA/ECOLEA/BETA_ECOLEA.xlsx"}
    return spec_fictive(**{**beta, **surcharges})


def _rien_ecrit(racine: Path) -> bool:
    return not any(p.suffix == ".xlsx" and "MASTER" not in p.name for p in racine.rglob("*"))


def test_facture_de_la_seconde_entite_lue_dans_son_master(tmp_path, deux_entites):
    """Une facture BETA prend l'émetteur, le SIRET et le dossier de BETA, sans toucher à ALPHA."""
    res = lancer(tmp_path, spec_beta(), "--data-root", str(deux_entites))

    assert res.returncode == 0, res.stdout + res.stderr
    ws = load_workbook(deux_entites / "BETA" / "ECOLEA" / "BETA_ECOLEA.xlsx")["B_ECOLEA001 Sept 2026"]
    assert ws["C7"].value == "EMETTEUR FICTIF BETA"
    assert not (deux_entites / "ALPHA" / "ECOLEA").exists()


def test_refus_numero_au_prefixe_d_une_autre_entite(tmp_path, deux_entites):
    """Un numéro « A_… » sur une facture BETA est refusé avant toute écriture."""
    res = lancer(tmp_path, spec_beta(n_facture="A_ECOLEA001"), "--data-root", str(deux_entites))

    assert res.returncode != 0
    assert "B_" in res.stderr and "A_ECOLEA001" in res.stderr
    assert _rien_ecrit(deux_entites)


@pytest.mark.parametrize(
    "fichier_sortie",
    ["ALPHA/ECOLEA/ALPHA_ECOLEA.xlsx", "BETA/../ALPHA/ECOLEA/ALPHA_ECOLEA.xlsx", "BETA_ECOLEA.xlsx", "ABSOLU"],
    ids=["dossier-autre-entite", "remontee", "hors-dossier-entite", "absolu"],
)
def test_refus_fichier_de_sortie_hors_du_dossier_de_l_entite(tmp_path, deux_entites, fichier_sortie):
    """Le fichier de sortie doit rester sous `<racine>/BETA/`, en chemin relatif, sans « .. »."""
    if fichier_sortie == "ABSOLU":
        fichier_sortie = str(tmp_path / "ailleurs" / "BETA_ECOLEA.xlsx")
    res = lancer(tmp_path, spec_beta(fichier_sortie=fichier_sortie), "--data-root", str(deux_entites))

    assert res.returncode != 0
    assert "fichier_sortie" in res.stderr
    assert _rien_ecrit(deux_entites)
    assert not (tmp_path / "ailleurs").exists()


def test_prefixe_lu_dans_le_profil(tmp_path, deux_entites):
    """Avec un profil ALPHA, c'est son `prefixe_facture` qui fait foi, pas l'initiale."""
    profil = tmp_path / "profil.toml"
    profil.write_text(
        '[utilisateur]\nnom = "Personne A"\nentite = "ALPHA"\nprefixe_facture = "AL_"\n\n'
        f'[donnees]\nracine = "{deux_entites.as_posix()}"\n',
        encoding="utf-8",
    )
    refuse = lancer(tmp_path, spec_fictive(), "--profil", str(profil), "--dry-run")
    accepte = lancer(tmp_path, spec_fictive(n_facture="AL_ECOLEA001"), "--profil", str(profil), "--dry-run")

    assert refuse.returncode != 0 and "AL_" in refuse.stderr
    assert accepte.returncode == 0, accepte.stdout + accepte.stderr
