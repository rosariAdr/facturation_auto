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
# FA-02 — une facture déjà émise ne peut pas être écrasée
# ──────────────────────────────────────────────────────────────────────


def classeur_existant(racine: Path, titre: str = "A_ECOLEA001 Sept 2026", numero: str = "A_ECOLEA001") -> Path:
    """Classeur d'école qui contient déjà une facture fictive (titre, numéro en D17, une ligne)."""
    classeur = racine / spec_fictive()["fichier_sortie"]
    classeur.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = titre
    ws["D17"] = numero
    ws["G20"] = "Module déjà facturé"
    wb.save(classeur)
    return classeur


def test_refus_si_l_onglet_existe(tmp_path, racine):
    """Relancer une spec dont l'onglet existe ne remplace pas la facture émise : refus, classeur intact."""
    classeur = classeur_existant(racine)
    avant = empreinte(classeur)

    res = lancer(tmp_path, spec_fictive(), "--data-root", str(racine))

    assert res.returncode != 0
    assert "existe déjà" in res.stderr
    assert "remplacé" not in res.stdout
    assert empreinte(classeur) == avant


def test_refus_si_l_onglet_existe_a_la_casse_pres(tmp_path, racine):
    """Un onglet qui ne diffère que par la casse est refusé, sans copie suffixée « 1 »."""
    classeur = classeur_existant(racine)
    avant = empreinte(classeur)

    spec = spec_fictive(onglet="a_ecolea001 sept 2026", n_facture="A_ECOLEA002")
    res = lancer(tmp_path, spec, "--data-root", str(racine))

    assert res.returncode != 0
    assert "existe déjà" in res.stderr
    assert empreinte(classeur) == avant
    assert load_workbook(classeur).sheetnames == ["A_ECOLEA001 Sept 2026"]


@pytest.mark.parametrize(
    ("titre", "numero"),
    [("PAID A_ECOLEA001 Sept 2026", "A_ECOLEA001"), ("Facture réglée", "A_ECOLEA001"), ("PAID A_ECOLEA001", None)],
    ids=["titre-et-cellule", "cellule-seule", "titre-seul"],
)
def test_refus_si_le_numero_figure_deja_dans_un_onglet(tmp_path, racine, titre, numero):
    """Le numéro est cherché dans le titre et dans la cellule N° de chaque onglet (factures « PAID … »)."""
    classeur = classeur_existant(racine, titre=titre, numero=numero)
    avant = empreinte(classeur)

    res = lancer(tmp_path, spec_fictive(), "--data-root", str(racine))

    assert res.returncode != 0
    assert "A_ECOLEA001" in res.stderr and titre in res.stderr
    assert empreinte(classeur) == avant
    assert load_workbook(classeur).sheetnames == [titre]


def test_nouvelle_facture_ajoutee_sans_toucher_aux_autres(tmp_path, racine):
    """Un nouvel onglet avec un nouveau numéro s'ajoute ; l'onglet existant garde son contenu."""
    classeur = classeur_existant(racine)

    spec = spec_fictive(onglet="A_ECOLEA002 Oct 2026", n_facture="A_ECOLEA002", date_emission="2026-11-01")
    res = lancer(tmp_path, spec, "--data-root", str(racine))

    assert res.returncode == 0, res.stdout + res.stderr
    wb = load_workbook(classeur)
    assert wb.sheetnames == ["A_ECOLEA001 Sept 2026", "A_ECOLEA002 Oct 2026"]
    ancien = wb["A_ECOLEA001 Sept 2026"]
    assert (ancien["D17"].value, ancien["G20"].value) == ("A_ECOLEA001", "Module déjà facturé")
    assert wb["A_ECOLEA002 Oct 2026"]["D17"].value == "A_ECOLEA002"


def test_apercu_annonce_un_nouvel_onglet(tmp_path, racine):
    """Le dry-run ouvre le classeur cible et annonce « nouvel onglet »."""
    res = lancer(tmp_path, spec_fictive(), "--data-root", str(racine), "--dry-run")

    assert res.returncode == 0, res.stdout + res.stderr
    assert "nouvel onglet" in res.stdout
    assert not (racine / "ALPHA" / "ECOLEA").exists()


def test_apercu_distingue_un_doublon_de_numero(tmp_path, racine):
    """Onglet nouveau mais numéro déjà émis (« PAID … ») : l'aperçu dit doublon, pas remplacement."""
    classeur_existant(racine, titre="PAID A_ECOLEA001 Sept 2026")

    res = lancer(tmp_path, spec_fictive(), "--data-root", str(racine), "--dry-run")

    assert res.returncode != 0
    assert "DOUBLON DE NUMÉRO" in res.stdout
    assert "REMPLACEMENT" not in res.stdout


def test_apercu_annonce_un_remplacement_refuse(tmp_path, racine):
    """Le dry-run annonce le remplacement, le refuse dès l'aperçu et n'écrit rien."""
    classeur = classeur_existant(racine)
    avant = empreinte(classeur)

    res = lancer(tmp_path, spec_fictive(), "--data-root", str(racine), "--dry-run")

    assert res.returncode != 0
    assert "REMPLACEMENT" in res.stdout
    assert "nouvel onglet" not in res.stdout
    assert empreinte(classeur) == avant
