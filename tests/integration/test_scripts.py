"""Scripts et commande `facturation` sur données fictives.

Banque de tests : T-CLI-01 à T-CLI-08, T-COH-01, T-VER-01 à T-VER-04, T-CFG-02, T-CFG-03.
T-CLI-03 à 08 et T-VER-03, 04 : mentions légales et échéance (F-24, ADR030).
"""

from __future__ import annotations

import json
import os
import re
import subprocess
import sys
from pathlib import Path

import pytest
from openpyxl import load_workbook

from facturation import config
from tests.conftest import NDA_FICTIF, creer_master

pytestmark = pytest.mark.integration

DEPOT = Path(__file__).resolve().parents[2]
SPEC_EXEMPLE = DEPOT / "examples" / "spec_exemple.json"


def _lancer(*args: str, env_extra: dict | None = None) -> subprocess.CompletedProcess:
    # PYTHONPATH : le sous-processus importe le paquet de ce dépôt, même si
    # l'interpréteur porte l'installation éditable d'une autre copie (FA-13).
    env = {**os.environ, "PYTHONIOENCODING": "utf-8", **(env_extra or {}), "PYTHONPATH": str(DEPOT)}
    env.pop("PCF_DATA_ROOT", None)
    return subprocess.run(
        [sys.executable, *args], capture_output=True, text=True, encoding="utf-8", env=env, cwd=DEPOT,
    )


def test_t_cli_01_apercu_dry_run(tmp_path):
    """T-CLI-01 — l'aperçu affiche les totaux et n'écrit aucun fichier."""
    racine = tmp_path / "donnees"
    creer_master(racine)
    res = _lancer("scripts/nouvelle_facture.py", str(SPEC_EXEMPLE), "--data-root", str(racine), "--dry-run")
    assert res.returncode == 0, res.stderr
    assert "TOTAL HT     : 1890.00 €" in res.stdout
    assert "[dry-run] Aucun fichier écrit." in res.stdout
    assert not (racine / "ROSARI" / "ABC").exists()


def test_t_cli_02_ecriture_et_post_mortem(tmp_path):
    """T-CLI-02 — l'écriture produit l'onglet attendu et un post-mortem complet."""
    racine = tmp_path / "donnees"
    creer_master(racine)
    res = _lancer("scripts/nouvelle_facture.py", str(SPEC_EXEMPLE), "--data-root", str(racine))
    assert res.returncode == 0, res.stderr
    spec = json.loads(SPEC_EXEMPLE.read_text(encoding="utf-8"))
    wb = load_workbook(racine / spec["fichier_sortie"])
    ws = wb[spec["onglet"]]
    assert ws["D17"].value == spec["n_facture"]
    assert ws["D14"].value == spec["ref_bon_commande"]
    assert ws["M7"].value == "ÉCOLE FICTIVE A"
    post_mortem = {
        ligne.split(":")[0].strip(): ligne for ligne in res.stdout.splitlines() if ligne.startswith("  ")
    }
    for cle in ("destinataire_nom", "siret", "iban"):
        assert cle in post_mortem and "None" not in post_mortem[cle], post_mortem.get(cle)


def test_t_coh_01_ecart_volontaire_detecte(tmp_path):
    """T-COH-01 — un écart dispensées ≠ facturées + à facturer est signalé."""
    racine = tmp_path / "donnees"
    creer_master(racine, cours=[
        ["ECO", "Module juste", "A1", 0, 10, 6, 4, "R_ECO001", None, "En cours"],
        ["ECO", "Module faux", "A2", 0, 10, 6, 2, "R_ECO001", None, "En cours"],
    ])
    res = _lancer("scripts/check_coherence.py", "--entite", "ROSARI", "--data-root", str(racine))
    assert res.returncode == 0, res.stderr
    assert "Écarts" in res.stdout
    assert "Module faux" in res.stdout.split("Écarts")[1]
    assert "Module juste" not in res.stdout.split("Écarts")[1]


def test_coherence_signale_heures_facturees_non_dispensees(tmp_path):
    """Toute ligne de Cours facturée au-delà des heures dispensées est signalée, même à 0 h dispensée (FA-08)."""
    racine = tmp_path / "donnees"
    creer_master(racine, cours=[
        # Facturé sans aucune heure dispensée.
        ["ECO", "Module non dispensé", "A1", 0, 0, 3, 0, "R_ECO001", None, "En cours"],
        # Même cas, mais « à facturer » négatif équilibre l'égalité dispensées = facturées + à facturer.
        ["ECO", "Module équilibré", "A2", 0, 0, 3, -3, "R_ECO001", None, "En cours"],
        # Heures dispensées, mais moins que les heures facturées.
        ["ECO", "Module dépassé", "A3", 0, 10, 12, -2, "R_ECO002", None, "En cours"],
        ["ECO", "Module juste", "A4", 0, 10, 6, 4, "R_ECO001", None, "En cours"],
        ["ECO", "Module à venir", "A5", 20, None, None, None, None, None, "À venir"],
    ])
    res = _lancer("scripts/check_coherence.py", "--entite", "ROSARI", "--data-root", str(racine))
    assert res.returncode == 0, res.stderr
    assert "Aucun écart" not in res.stdout
    titre = "Heures facturées au-delà des heures dispensées"
    assert titre in res.stdout, res.stdout
    signalees = res.stdout.split(titre)[1].split("⚠")[0]
    for module in ("Module non dispensé", "Module équilibré", "Module dépassé"):
        assert module in signalees, signalees
    for module in ("Module juste", "Module à venir"):
        assert module not in signalees, signalees


def test_t_ver_01_verifier_installation(tmp_path):
    """T-VER-01 — `facturation profil verifier` : 0 erreur sur une installation fictive complète."""
    racine = tmp_path / "donnees"
    creer_master(racine)
    profil = tmp_path / "profil.toml"
    profil.write_text(config.rediger("Personne A", "ROSARI", str(racine)), encoding="utf-8")
    res = _lancer("-m", "facturation", "profil", "verifier", "--profil", str(profil))
    assert res.returncode == 0, res.stdout + res.stderr
    assert "0 erreur(s)" in res.stdout
    assert "destinataires sans coordonnées (factures bloquées) : VIDE" in res.stdout


def test_t_ver_02_dossier_de_donnees_absent(tmp_path):
    """T-VER-02 — dossier de données introuvable : erreur bloquante, code de retour 1."""
    profil = tmp_path / "profil.toml"
    profil.write_text(config.rediger("Personne A", "ROSARI", str(tmp_path / "absent")), encoding="utf-8")
    res = _lancer("-m", "facturation", "profil", "verifier", "--profil", str(profil))
    assert res.returncode == 1
    assert "[ERREUR]" in res.stdout and "dossier introuvable" in res.stdout


def test_profil_init_refuse_d_ecraser(tmp_path):
    cible = tmp_path / "profil.toml"
    args = ["-m", "facturation", "profil", "init", "--nom", "Personne A", "--entite", "ROSARI",
            "--racine", str(tmp_path), "--emplacement", str(cible)]
    assert _lancer(*args).returncode == 0
    second = _lancer(*args)
    assert second.returncode == 1 and "--force" in second.stderr


ADRESSE_FICTIVE = "https://calendar.google.com/calendar/ical/demo%40example.org/private-0000/basic.ics"


def test_t_cfg_02_enregistrer_adresse_ical(tmp_path):
    """T-CFG-02 — `profil agenda` range l'adresse hors du profil et la déclare ; verifier la voit."""
    racine = tmp_path / "donnees"
    creer_master(racine)
    profil = tmp_path / "profil.toml"
    profil.write_text(config.rediger("Personne A", "ROSARI", str(racine)), encoding="utf-8")
    res = _lancer("-m", "facturation", "profil", "agenda", "--profil", str(profil), "--adresse", ADRESSE_FICTIVE)
    assert res.returncode == 0, res.stdout + res.stderr
    assert (tmp_path / "agenda.secret").read_text(encoding="utf-8").strip() == ADRESSE_FICTIVE
    assert ADRESSE_FICTIVE not in profil.read_text(encoding="utf-8")
    assert config.charger(profil).agenda_fichier_url_secrete == tmp_path / "agenda.secret"
    verif = _lancer("-m", "facturation", "profil", "verifier", "--profil", str(profil))
    assert "[OK]        Agenda" in verif.stdout
    assert ADRESSE_FICTIVE not in verif.stdout


def test_t_cfg_03_adresse_ical_invalide_refusee(tmp_path):
    """T-CFG-03 — une adresse qui n'est pas un lien .ics est refusée, rien n'est écrit."""
    profil = tmp_path / "profil.toml"
    profil.write_text(config.rediger("Personne A", "ROSARI", str(tmp_path)), encoding="utf-8")
    avant = profil.read_text(encoding="utf-8")
    res = _lancer("-m", "facturation", "profil", "agenda", "--profil", str(profil),
                  "--adresse", "https://calendar.google.com/calendar/u/0/r")
    assert res.returncode == 1 and "adresse iCal invalide" in res.stderr
    assert not (tmp_path / "agenda.secret").exists()
    assert profil.read_text(encoding="utf-8") == avant


@pytest.fixture
def autre_copie(tmp_path: Path) -> Path:
    """Dossier contenant un autre paquet `facturation`, qui s'arrête dès qu'on l'importe.

    Il joue l'installation éditable d'une autre copie du dépôt (FA-13) : placé sur
    le chemin d'import, il passe avant le paquet installé.
    """
    dossier = tmp_path / "autre_copie"
    paquet = dossier / "facturation"
    paquet.mkdir(parents=True)
    (paquet / "__init__.py").write_text(
        'raise SystemExit("AUTRE COPIE importée : " + __file__)\n', encoding="utf-8"
    )
    return dossier


def test_check_coherence_importe_le_paquet_de_son_depot(autre_copie):
    """Lancé par son chemin, check_coherence importe le paquet de son dépôt, pas celui d'une autre copie."""
    env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONPATH": str(autre_copie)}
    res = subprocess.run(
        [sys.executable, "scripts/check_coherence.py", "--help"],
        capture_output=True, text=True, encoding="utf-8", env=env, cwd=DEPOT,
    )
    assert "AUTRE COPIE" not in res.stdout + res.stderr
    assert res.returncode == 0, res.stderr
    assert "--data-root" in res.stdout


def test_lancer_fait_importer_le_paquet_de_ce_depot(tmp_path, autre_copie, monkeypatch):
    """`_lancer` fait importer aux sous-processus le paquet de ce dépôt, même si une autre copie est sur le chemin."""
    monkeypatch.setenv("PYTHONPATH", str(autre_copie))
    sonde = tmp_path / "sonde.py"
    sonde.write_text("import facturation\nprint(facturation.__file__)\n", encoding="utf-8")
    res = _lancer(str(sonde))
    assert res.returncode == 0, res.stdout + res.stderr
    assert Path(res.stdout.strip()).resolve().parent == DEPOT / "facturation"


# ──────────────────────────────────────────────────────────────────────
# F-24 — mentions légales et échéance (ADR030)
# ──────────────────────────────────────────────────────────────────────

CONTRAT = "CT-EXEMPLE-0001"
# Un message de refus nomme la mention attendue : « EI » ou « Entrepreneur individuel ».
MENTION_EI = r"\bEI\b|[Ee]ntrepreneur individuel"


def _spec(tmp_path: Path, **surcharges) -> Path:
    """Spec d'exemple, complétée des surcharges, écrite dans `tmp_path`."""
    spec = json.loads(SPEC_EXEMPLE.read_text(encoding="utf-8"))
    spec.update(surcharges)
    chemin = tmp_path / "spec.json"
    chemin.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
    return chemin


def _valeur_apercu(sortie: str, etiquette: str) -> str | None:
    """Valeur d'une ligne « Étiquette : valeur » de l'aperçu, ou None si la ligne manque."""
    trouvee = re.search(rf"^{re.escape(etiquette)}\s*:\s*(.*)$", sortie, re.MULTILINE)
    return trouvee.group(1).strip() if trouvee else None


def _erreurs(sortie: str) -> list[str]:
    """Lignes « [ERREUR] … » de `profil verifier`."""
    return [ligne for ligne in sortie.splitlines() if ligne.startswith("[ERREUR]")]


def test_t_cli_03_apercu_contrat_echeance_nda(tmp_path):
    """T-CLI-03 — l'aperçu montre le n° de contrat, l'échéance donnée et la déclaration d'activité."""
    racine = tmp_path / "donnees"
    creer_master(racine)
    spec = _spec(tmp_path, ref_contrat=CONTRAT, echeance="2026-11-15")
    res = _lancer("scripts/nouvelle_facture.py", str(spec), "--data-root", str(racine), "--dry-run")
    assert res.returncode == 0, res.stdout + res.stderr
    assert _valeur_apercu(res.stdout, "Contrat") == CONTRAT, res.stdout
    assert _valeur_apercu(res.stdout, "Échéance") == "15/11/2026", res.stdout
    assert _valeur_apercu(res.stdout, "NDA") == NDA_FICTIF, res.stdout


def test_t_cli_04_apercu_echeance_par_defaut(tmp_path):
    """T-CLI-04 — sans échéance dans la spec, l'aperçu annonce la date d'émission plus 30 jours."""
    racine = tmp_path / "donnees"
    creer_master(racine)
    res = _lancer("scripts/nouvelle_facture.py", str(SPEC_EXEMPLE), "--data-root", str(racine), "--dry-run")
    assert res.returncode == 0, res.stdout + res.stderr
    assert _valeur_apercu(res.stdout, "Échéance") == "30/10/2026", res.stdout  # émise le 30/09/2026


def test_t_cli_05_ecriture_contrat_echeance_nda(tmp_path):
    """T-CLI-05 — la facture écrite porte le contrat (ligne 15), l'échéance et la déclaration d'activité."""
    racine = tmp_path / "donnees"
    creer_master(racine)
    spec = _spec(tmp_path, ref_contrat=CONTRAT, echeance="2026-11-15")
    res = _lancer("scripts/nouvelle_facture.py", str(spec), "--data-root", str(racine))
    assert res.returncode == 0, res.stdout + res.stderr
    assert "OK — facture" in res.stdout
    donnees = json.loads(spec.read_text(encoding="utf-8"))
    ws = load_workbook(racine / donnees["fichier_sortie"])[donnees["onglet"]]
    assert (ws["C15"].value, ws["D15"].value) == ("N° de contrat :", CONTRAT)
    assert ws["C91"].value == "Échéance : 15/11/2026"
    assert ws["C86"].value == NDA_FICTIF


def test_t_cli_06_refus_sans_nda(tmp_path):
    """T-CLI-06 — sans `Info` / `NDA` dans le MASTER : écriture refusée, code 1, rien d'écrit."""
    racine = tmp_path / "donnees"
    creer_master(racine, nda=None)
    res = _lancer("scripts/nouvelle_facture.py", str(SPEC_EXEMPLE), "--data-root", str(racine))
    assert res.returncode == 1, res.stdout + res.stderr
    assert "Écriture refusée" in res.stderr and "NDA" in res.stderr, res.stderr
    assert not (racine / "ROSARI" / "ABC").exists()


def test_t_cli_07_refus_emetteur_sans_ei(tmp_path):
    """T-CLI-07 — émetteur sans « EI » ni « Entrepreneur individuel » : écriture refusée, code 1, rien d'écrit."""
    racine = tmp_path / "donnees"
    creer_master(racine, nom_emetteur="FORMATEUR FICTIF ROSARI")
    res = _lancer("scripts/nouvelle_facture.py", str(SPEC_EXEMPLE), "--data-root", str(racine))
    assert res.returncode == 1, res.stdout + res.stderr
    assert "Écriture refusée" in res.stderr and re.search(MENTION_EI, res.stderr), res.stderr
    assert not (racine / "ROSARI" / "ABC").exists()


@pytest.mark.parametrize(
    ("master", "cause"),
    [({"nda": None}, r"NDA"), ({"nom_emetteur": "FORMATEUR FICTIF ROSARI"}, MENTION_EI)],
    ids=["sans-nda", "sans-ei"],
)
def test_t_cli_08_refus_des_l_apercu(tmp_path, master, cause):
    """T-CLI-08 — NDA manquant ou émetteur sans « EI » : refus dès le `--dry-run`, code 1, cause nommée, rien d'écrit."""
    racine = tmp_path / "donnees"
    creer_master(racine, **master)
    res = _lancer("scripts/nouvelle_facture.py", str(SPEC_EXEMPLE), "--data-root", str(racine), "--dry-run")
    assert res.returncode == 1, res.stdout + res.stderr
    assert "Écriture refusée" in res.stderr and re.search(cause, res.stderr), res.stderr
    assert not (racine / "ROSARI" / "ABC").exists()


def test_t_ver_03_verifier_signale_le_nda_manquant(tmp_path):
    """T-VER-03 — `profil verifier` : ERREUR si `Info` / `NDA` manque dans le MASTER."""
    racine = tmp_path / "donnees"
    creer_master(racine, nda=None)
    profil = tmp_path / "profil.toml"
    profil.write_text(config.rediger("Personne A", "ROSARI", str(racine)), encoding="utf-8")
    res = _lancer("-m", "facturation", "profil", "verifier", "--profil", str(profil))
    assert res.returncode == 1, res.stdout + res.stderr
    assert any("NDA" in ligne for ligne in _erreurs(res.stdout)), res.stdout


def test_t_ver_04_verifier_signale_l_emetteur_sans_ei(tmp_path):
    """T-VER-04 — `profil verifier` : ERREUR si le nom de l'émetteur ne porte pas « EI »."""
    racine = tmp_path / "donnees"
    creer_master(racine, nom_emetteur="FORMATEUR FICTIF ROSARI")
    profil = tmp_path / "profil.toml"
    profil.write_text(config.rediger("Personne A", "ROSARI", str(racine)), encoding="utf-8")
    res = _lancer("-m", "facturation", "profil", "verifier", "--profil", str(profil))
    assert res.returncode == 1, res.stdout + res.stderr
    assert any(re.search(MENTION_EI, ligne) for ligne in _erreurs(res.stdout)), res.stdout
