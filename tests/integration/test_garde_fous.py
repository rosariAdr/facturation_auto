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
# FA-12 — post-mortem vérifié : cellules relues comparées aux valeurs attendues
# ──────────────────────────────────────────────────────────────────────

import importlib.util  # noqa: E402
from datetime import date  # noqa: E402

from facturation import pcf_format as fmt  # noqa: E402


@pytest.fixture
def fmt_isole(monkeypatch):
    """Rend à `pcf_format` sa racine et son cache à la fin du test (état global du module)."""
    monkeypatch.setattr(fmt, "_DATA_ROOT", fmt._DATA_ROOT)
    monkeypatch.setattr(fmt, "_MASTER_DATA_CACHE", {})
    return fmt


def _charger_script():
    spec = importlib.util.spec_from_file_location("nouvelle_facture", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


GROUPES = [("PROMO 1", "Module fictif", [(date(2026, 9, 15), 3.0, 50.0), (date(2026, 9, 16), 2.0, 50.0)])]


def _facture_en_memoire(racine: Path):
    """Facture fictive construite par l'API, comme le fait le script."""
    fmt.set_data_root(racine)
    wb = Workbook()
    wb.remove(wb.active)
    ws = fmt.clone_template(wb, "A_ECOLEA001 Sept 2026", entite="ALPHA")
    dest = fmt.get_destinataire("ECOLEA", entite="ALPHA")
    fmt.apply_emetteur(ws, entite="ALPHA")
    fmt.apply_destinataire(ws, dest=dest, entite="ALPHA")
    fmt.apply_metadata(ws, n_facture="A_ECOLEA001", date_emission=date(2026, 10, 1), entite="ALPHA")
    fmt.place_sessions(ws, ecole_name="ECOLE A", groups=GROUPES, entite="ALPHA")
    fmt.apply_rib(ws, entite="ALPHA")
    return ws, dest


def _verifier(ws, dest):
    return fmt.verifier_facture(
        ws, entite="ALPHA", n_facture="A_ECOLEA001", date_emission=date(2026, 10, 1),
        dest=dest, ecole_name="ECOLE A", groups=GROUPES,
    )


def test_facture_conforme_aucun_ecart(racine, fmt_isole):
    ws, dest = _facture_en_memoire(racine)
    assert _verifier(ws, dest) == []


@pytest.mark.parametrize(
    ("cellule", "valeur", "cle"),
    [("D17", "A_ECOLEA009", "n_facture"), ("D88", "FR76 9999", "iban"), ("M21", 7.0, "heures"),
     ("O20", 55.0, "taux"), ("B2", "FACTURE\nDate d'émission: 01/01/2020", "titre")],
)
def test_cellule_alteree_signalee(racine, fmt_isole, cellule, valeur, cle):
    ws, dest = _facture_en_memoire(racine)
    ws[cellule] = valeur
    ecarts = _verifier(ws, dest)
    assert len(ecarts) == 1 and ecarts[0].startswith(f"ERREUR sur cellule {cle}"), ecarts


def test_post_mortem_conforme_annonce_ok(tmp_path, racine):
    res = lancer(tmp_path, spec_fictive(), "--data-root", str(racine))

    assert res.returncode == 0, res.stdout + res.stderr
    assert "OK — facture A_ECOLEA001 conforme" in res.stdout
    assert "ERREUR sur cellule" not in res.stdout + res.stderr


def test_post_mortem_en_ecart_code_de_sortie_non_nul(tmp_path, racine, fmt_isole, monkeypatch, capsys):
    """Une cellule relue qui diffère de l'attendu : message « ERREUR sur cellule » et code 2."""
    nf = _charger_script()
    vrai_apply_rib = nf.fmt.apply_rib

    def apply_rib_fautif(ws, **kwargs):
        vrai_apply_rib(ws, **kwargs)
        ws["D88"] = "FR76 IBAN FAUX"

    monkeypatch.setattr(nf.fmt, "apply_rib", apply_rib_fautif)
    chemin_spec = tmp_path / "facture.json"
    chemin_spec.write_text(json.dumps(spec_fictive()), encoding="utf-8")
    monkeypatch.setattr(sys, "argv", ["nouvelle_facture.py", str(chemin_spec), "--data-root", str(racine)])

    with pytest.raises(SystemExit) as fin:
        nf.main()

    assert fin.value.code == 2
    sortie = capsys.readouterr()
    assert "ERREUR sur cellule iban" in sortie.err
    assert "OK — facture" not in sortie.out
