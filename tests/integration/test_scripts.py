"""Scripts et commande `facturation` sur données fictives.

Banque de tests : T-CLI-01 à T-CLI-12, T-COH-01, T-VER-01 à T-VER-10, T-CFG-02, T-CFG-03.
T-CLI-03 à 09 et T-VER-03, 04 : mentions légales et échéance (F-24, ADR030).
T-CLI-10 : facture de plus de 60 séances (F-18).
T-FAC-02 : aide du script, devenu l'enveloppe de `facturation.facture` (P-0a).
T-CLI-12, T-VER-07 à 10 : socle du lot 1 (P-0b) et sa revue.
T-FAC-18 à T-FAC-25 : contrôles de la facture laissés hors de P-0a, côté script (fix-facture).
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


@pytest.mark.parametrize(
    "echeance", ["2026-09-29", "30/10/2026"], ids=["anterieure-a-l-emission", "mal-formee"],
)
@pytest.mark.parametrize("options", [(), ("--dry-run",)], ids=["ecriture", "dry-run"])
def test_t_cli_09_refus_echeance_invalide(tmp_path, echeance, options):
    """T-CLI-09 — échéance de la spec antérieure à l'émission (30/09/2026) ou hors format AAAA-MM-JJ :
    refus propre, code 1, sans trace d'exception, rien d'écrit."""
    racine = tmp_path / "donnees"
    creer_master(racine)
    spec = _spec(tmp_path, echeance=echeance)
    res = _lancer("scripts/nouvelle_facture.py", str(spec), "--data-root", str(racine), *options)
    assert res.returncode == 1, res.stdout + res.stderr
    assert "Écriture refusée" in res.stderr and "echeance" in res.stderr, res.stderr
    assert "Traceback" not in res.stderr, res.stderr
    assert not (racine / "ROSARI" / "ABC").exists()


def test_t_cli_10_facture_de_62_seances(tmp_path):
    """T-CLI-10 — une facture de 62 séances (lignes insérées) s'écrit avec le code 0 et un contrôle
    après écriture sans écart (F-18)."""
    racine = tmp_path / "donnees"
    creer_master(racine)
    seances = [{"date": f"2026-09-{1 + i % 28:02d}", "heures": 1, "taux": 50} for i in range(62)]
    spec = _spec(tmp_path, masquer_jour=False, lignes=[{"promotion": "M1 S1", "libelle": "Module A", "seances": seances}])
    res = _lancer("scripts/nouvelle_facture.py", str(spec), "--data-root", str(racine))
    assert res.returncode == 0, res.stdout + res.stderr
    assert "TOTAL HT     : 3100.00 €" in res.stdout
    assert "OK — facture" in res.stdout
    assert "ERREUR sur cellule" not in res.stdout + res.stderr


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


# ──────────────────────────────────────────────────────────────────────
# F-10 — lecture de l'agenda (T-CLI-11)
# ──────────────────────────────────────────────────────────────────────

AGENDA_FICTIF = DEPOT / "tests" / "fixtures" / "agenda_fictif.ics"


def test_t_cli_11_agenda_lire_un_fichier_local(tmp_path):
    """T-CLI-11 — `facturation agenda lire --mois 2026-11 --fichier …` liste les séances de novembre du
    .ics local (date, début–fin, durée, intitulé), sans profil ni accès réseau."""
    sans_profil = {
        "APPDATA": str(tmp_path / "appdata"), "XDG_CONFIG_HOME": str(tmp_path / "config"),
        "FACTURATION_PROFIL": str(tmp_path / "aucun_profil.toml"),
    }
    res = _lancer("-m", "facturation", "agenda", "lire", "--mois", "2026-11", "--fichier", str(AGENDA_FICTIF),
                  env_extra=sans_profil)
    assert res.returncode == 0, res.stdout + res.stderr
    lignes = res.stdout.splitlines()

    def ligne_du(jour: str, debut: str) -> str:
        """La ligne de la séance : sa date et son heure de début (un en-tête peut citer la date seule)."""
        trouvees = [ligne for ligne in lignes if jour in ligne and debut in ligne]
        assert len(trouvees) == 1, f"{jour} {debut} : {trouvees}\n{res.stdout}"
        return trouvees[0]

    for jour, debut, fin, intitule in (
        ("05/11/2026", "14:00", "17:30", "#ABC-M1 Cours fictif - Module A (déplacé)"),
        ("10/11/2026", "09:00", "12:30", "#ABC-M1 Cours fictif - Module A"),
        ("13/11/2026", "00:30", "02:30", "#ECO-M2 Cours fictif noté en UTC"),
        ("17/11/2026", "09:00", "12:30", "#ABC-M1 Cours fictif - Module A"),
        ("24/11/2026", "09:00", "12:30", "#ABC-M1 Cours fictif - Module A"),
        ("28/11/2026", "10:00", "13:30", "#ABC-M1 Cours fictif - Module A"),
        ("30/11/2026", "20:00", "22:00", "#ABC-M1 Cours fictif du soir"),
    ):
        ligne = ligne_du(jour, debut)
        assert fin in ligne and intitule in ligne, ligne
    assert re.search(r"3[.,]5|3 ?h ?30", ligne_du("05/11/2026", "14:00")), "durée de 3 h 30 absente"
    for absent in ("03/11/2026", "11/11/2026", "Jour férié fictif", "27/10/2026", "15/12/2026", "annulé"):
        assert absent not in res.stdout, absent


def test_t_ver_05_verifier_accepte_un_agenda_secret_avec_bom(tmp_path):
    """T-VER-05 — `profil verifier` : un `agenda.secret` retouché au Bloc-notes ou sous PowerShell 5
    (BOM UTF-8, fin de ligne Windows) n'est pas signalé comme invalide, et l'adresse ne s'affiche pas."""
    jeton = "private-0123456789abcdef"
    adresse = f"https://calendar.example.org/calendar/ical/personne%40example.org/{jeton}/basic.ics"
    racine = tmp_path / "donnees"
    creer_master(racine)
    profil = tmp_path / "profil.toml"
    profil.write_text(config.rediger("Personne A", "ROSARI", str(racine)), encoding="utf-8")
    secret = config.enregistrer_adresse_ical(config.charger(profil), adresse)
    secret.write_bytes(b"\xef\xbb\xbf" + adresse.encode("ascii") + b"\r\n")
    res = _lancer("-m", "facturation", "profil", "verifier", "--profil", str(profil))
    for fragment in (jeton, "calendar.example.org"):
        assert fragment not in res.stdout + res.stderr, f"adresse affichée : {fragment}"
    agenda = [ligne for ligne in res.stdout.splitlines() if "Agenda —" in ligne]
    assert len(agenda) == 1 and agenda[0].startswith("[OK]"), res.stdout


@pytest.mark.parametrize(
    "contenu",
    [
        "https://calendar.example.org/calendar/ical/private-0123456789abcdef/basic.ics\r\n".encode("utf-16"),
        b"\x00\xff\xfe\x81 calendar.example.org private-0123456789abcdef \x9d\x00",
        b"https://calendar.example.org/calendar/ical/private-0123456789abcdef /basic.ics\n",
    ],
    ids=["utf-16", "binaire", "espace"],
)
def test_t_ver_06_verifier_signale_un_agenda_secret_inutilisable(tmp_path, contenu):
    """T-VER-06 — `profil verifier` : un `agenda.secret` illisible (UTF-16, binaire) ou dont l'adresse contient
    une espace donne une ERREUR claire qui nomme le fichier ; la commande ne plante pas et n'affiche pas l'adresse."""
    racine = tmp_path / "donnees"
    creer_master(racine)
    profil = tmp_path / "profil.toml"
    profil.write_text(config.rediger("Personne A", "ROSARI", str(racine)), encoding="utf-8")
    secret = config.enregistrer_adresse_ical(
        config.charger(profil), "https://calendar.example.org/calendar/ical/private-0123456789abcdef/basic.ics",
    )
    secret.write_bytes(contenu)
    res = _lancer("-m", "facturation", "profil", "verifier", "--profil", str(profil))
    assert "Traceback" not in res.stderr, res.stderr
    for fragment in ("private-0123456789abcdef", "calendar.example.org"):
        assert fragment not in res.stdout + res.stderr, f"adresse affichée : {fragment}"
    agenda = [ligne for ligne in res.stdout.splitlines() if "Agenda —" in ligne]
    assert len(agenda) == 1 and agenda[0].startswith("[ERREUR]"), res.stdout + res.stderr
    assert "agenda.secret" in agenda[0], agenda[0]
    assert res.returncode == 1, res.stdout + res.stderr


# ──────────────────────────────────────────────────────────────────────
# P-0a — facture dans le paquet (T-FAC-02 ; les autres T-FAC sont dans tests/unit/test_facture.py)
# ──────────────────────────────────────────────────────────────────────


def test_t_fac_02_aide_du_script_enveloppe():
    """T-FAC-02 — `python scripts/nouvelle_facture.py --help` affiche la première ligne de la documentation du
    script, inchangée par le passage dans le paquet (« Génère une facture depuis une spec JSON »), et ses options."""
    import ast

    script = DEPOT / "scripts" / "nouvelle_facture.py"
    documentation = ast.get_docstring(ast.parse(script.read_text(encoding="utf-8")))
    assert documentation, "le script n'a plus de documentation : --help n'a plus de description"
    premiere_ligne = documentation.splitlines()[0]
    # COLUMNS large : argparse ne replie pas la description, qui se lit d'un seul tenant.
    res = _lancer("scripts/nouvelle_facture.py", "--help", env_extra={"COLUMNS": "200"})
    assert res.returncode == 0, res.stdout + res.stderr
    assert premiere_ligne in res.stdout, res.stdout
    assert "Génère une facture depuis une spec JSON" in res.stdout, res.stdout
    for option in ("--data-root", "--profil", "--dry-run"):
        assert option in res.stdout, option


# ──────────────────────────────────────────────────────────────────────
# P-0b — refonte de la répartition (T-CLI-12) et vérification (T-VER-07 à T-VER-09)
# ──────────────────────────────────────────────────────────────────────
#
# T-CLI-12 fige les sorties des commandes existantes, relevées sur `dev` (4c5bdbb) avant la refonte
# de `cli.py` (répartition sur `args.func`, contrats des vagues 3 à 5, § 3.7) : elles ne doivent
# pas changer. T-VER-07 à T-VER-09 : `profil verifier` signale en ATTENTION chaque onglet de suivi
# invalide (`suivi_invalide`) et chaque perte de `fidelite_aller_retour` ; une installation saine
# n'en signale aucun.

SANS_PROFIL = {"FACTURATION_PROFIL": "aucun_profil_fictif.toml"}

SORTIE_AGENDA_NOVEMBRE = """\
Séances de novembre 2026 (fuseau Europe/Paris) :
  jeu. 05/11/2026  14:00–17:30    3,5 h  #ABC-M1 Cours fictif - Module A (déplacé)
  mar. 10/11/2026  09:00–12:30    3,5 h  #ABC-M1 Cours fictif - Module A
  ven. 13/11/2026  00:30–02:30      2 h  #ECO-M2 Cours fictif noté en UTC
  mar. 17/11/2026  09:00–12:30    3,5 h  #ABC-M1 Cours fictif - Module A
  mar. 24/11/2026  09:00–12:30    3,5 h  #ABC-M1 Cours fictif - Module A
  sam. 28/11/2026  10:00–13:30    3,5 h  #ABC-M1 Cours fictif - Module A
  lun. 30/11/2026  20:00–22:00      2 h  #ABC-M1 Cours fictif du soir

7 séance(s), 21,5 h au total.
"""


def test_t_cli_12_commandes_existantes_memes_sorties(tmp_path):
    """T-CLI-12 — après la refonte de la répartition, `agenda lire`, `profil chemin` et les erreurs
    d'usage rendent exactement les sorties et les codes relevés avant elle."""
    res = _lancer("-m", "facturation", "agenda", "lire", "--mois", "2026-11", "--fichier", str(AGENDA_FICTIF),
                  env_extra=SANS_PROFIL)
    assert (res.returncode, res.stdout, res.stderr) == (0, SORTIE_AGENDA_NOVEMBRE, "")

    profil = tmp_path / "config" / "profil.toml"
    res = _lancer("-m", "facturation", "profil", "chemin", env_extra={"FACTURATION_PROFIL": str(profil)})
    assert (res.returncode, res.stdout, res.stderr) == (0, f"{profil}\n", "")

    res = _lancer("-m", "facturation", "agenda", "lire", "--mois", "2026-13", "--fichier", str(AGENDA_FICTIF),
                  env_extra=SANS_PROFIL)
    assert res.returncode == 2 and res.stdout == ""
    assert "argument --mois: mois invalide : '2026-13', attendu AAAA-MM (ex. 2026-11)" in res.stderr

    res = _lancer("-m", "facturation", "agenda", "lire", "--mois", "2026-11", "--fichier", "absent.ics",
                  env_extra=SANS_PROFIL)
    assert (res.returncode, res.stdout) == (1, "")
    assert res.stderr == "[ERREUR]    Agenda — impossible de lire absent.ics : fichier introuvable\n"

    res = _lancer("-m", "facturation", env_extra=SANS_PROFIL)
    assert res.returncode == 2 and res.stdout == ""
    assert res.stderr.startswith("usage: facturation")

    res = _lancer("-m", "facturation", "profil", "verifier", "--profil", str(tmp_path / "absent.toml"))
    assert res.returncode == 1
    assert res.stdout.startswith("[ERREUR]    Profil — Profil invalide (")


def _installation(tmp_path: Path) -> tuple[Path, Path]:
    """Installation fictive complète : (racine, profil)."""
    racine = tmp_path / "donnees"
    master = creer_master(racine, tarifs=[["ABC", "Module fictif", "M1", "45 €/h", "", ""]],
                          ecoles=[["École Fictive A", "ABC"]])
    profil = tmp_path / "profil.toml"
    profil.write_text(config.rediger("Personne A", "ROSARI", str(racine)), encoding="utf-8")
    return master, profil


def _attentions(sortie: str) -> list[str]:
    return [ligne for ligne in sortie.splitlines() if ligne.startswith("[ATTENTION]")]


def _retirer(master: Path, onglet: str, colonne: str):
    wb = load_workbook(master)
    ws = wb[onglet]
    ws.delete_cols([c.value for c in ws[1]].index(colonne) + 1)
    wb.save(master)


def test_t_ver_07_verifier_signale_chaque_onglet_de_suivi_invalide(tmp_path):
    """T-VER-07 — `profil verifier` : une ATTENTION par onglet de suivi invalide, qui nomme l'onglet et
    la colonne manquante ; ce n'est pas une erreur (code 0)."""
    master, profil = _installation(tmp_path)
    _retirer(master, "Tarifs", "Taux HT/h")
    _retirer(master, "Écoles", "Code facturation")
    res = _lancer("-m", "facturation", "profil", "verifier", "--profil", str(profil))
    assert res.returncode == 0, res.stdout + res.stderr
    attentions = _attentions(res.stdout)
    assert any("Tarifs" in a and "Taux HT/h" in a for a in attentions), res.stdout
    assert any("Écoles" in a and "Code facturation" in a for a in attentions), res.stdout


def test_t_ver_08_verifier_signale_les_pertes_de_l_aller_retour(tmp_path):
    """T-VER-08 — `profil verifier` lance `fidelite_aller_retour` sur le MASTER : des formules dans
    `Sessions` donnent une ATTENTION (« elles se liront vides… »), pas une erreur ; le MASTER n'est
    pas modifié."""
    master, profil = _installation(tmp_path)
    wb = load_workbook(master)
    ws = wb["Sessions"]
    ws.append(["2026-10-01", "ABC", "Septembre 2026", "R_ABC001", None, "09:00", "12:30", 3.5, "M1", "Module fictif",
               45, "=H2*K2", "#ABC-M1 Cours fictif"])
    wb.save(master)
    avant = master.read_bytes()
    res = _lancer("-m", "facturation", "profil", "verifier", "--profil", str(profil))
    assert res.returncode == 0, res.stdout + res.stderr
    assert any("se liront vides" in a and "Sessions" in a for a in _attentions(res.stdout)), res.stdout
    assert master.read_bytes() == avant


def test_t_ver_09_installation_saine_sans_attention_de_suivi_ni_de_fidelite(tmp_path):
    """T-VER-09 — sur une installation fictive saine, `profil verifier` ne signale ni onglet de suivi
    invalide ni perte d'aller-retour (les gardes de T-VER-07 et T-VER-08 ne parlent pas à vide)."""
    _, profil = _installation(tmp_path)
    res = _lancer("-m", "facturation", "profil", "verifier", "--profil", str(profil))
    assert res.returncode == 0, res.stdout + res.stderr
    for attention in _attentions(res.stdout):
        assert "se liront vides" not in attention, attention
        for mot in ("Tarifs", "Écoles", "Cours", "Taux HT/h", "Code facturation", "perdu", "perte"):
            assert mot not in attention, attention


# Revue de P-0b (T-VER-10), écrit rouge avant le correctif : un MASTER que `lire_master` refuse
# (colonne obligatoire absente de `Sessions`…) ne donnait qu'une ATTENTION, et `profil verifier`
# rendait 0, alors que toute préparation échouera. La définition du module (une ERREUR « empêcherait
# de préparer les factures ») appelle une ERREUR et le code de retour 1. Les onglets de suivi
# illisibles restent des ATTENTION (T-VER-07).


def test_t_ver_10_master_refuse_par_l_outil_est_une_erreur(tmp_path):
    """T-VER-10 — `profil verifier` sur un MASTER fictif sans colonne « Total HT » dans `Sessions` :
    une ERREUR qui nomme l'onglet et la colonne, et le code de retour 1."""
    master, profil = _installation(tmp_path)
    _retirer(master, "Sessions", "Total HT")
    res = _lancer("-m", "facturation", "profil", "verifier", "--profil", str(profil))
    erreurs = [ligne for ligne in res.stdout.splitlines() if ligne.startswith("[ERREUR]")]
    assert any("Sessions" in e and "Total HT" in e for e in erreurs), res.stdout
    assert res.returncode == 1, res.stdout + res.stderr


# ──────────────────────────────────────────────────────────────────────
# fix-facture — contrôles laissés hors de P-0a, côté script (T-FAC-18 à T-FAC-25)
# ──────────────────────────────────────────────────────────────────────
#
# Un refus de la spec suit le style des refus existants, avant toute écriture : code 1, un message
# « [ERREUR] » qui nomme la cause, sans trace Python. Les détails (type d'exception, rangs, valeurs) sont dans
# tests/unit/test_facture.py.

CLASSEUR_ABC = Path("ROSARI", "ABC", "ROSARI_ABC.xlsx")
# 32 caractères : un de plus que la limite d'Excel.
TITRE_32 = "R_ABC001 Août 2026 École fictive"


def _lignes_fac() -> list[dict]:
    """Deux lignes, la seconde à trois séances : une séance fautive y est la « séance 3 » de la « ligne 2 »."""
    return [
        {"promotion": "M1 S1", "libelle": "[2026][M1 S1][Module A][3H]",
         "seances": [{"date": "2026-09-15", "heures": 3, "taux": 90}]},
        {"promotion": "M2 S1", "libelle": "[2026][M2 S1][Module B][6H]",
         "seances": [{"date": f"2026-09-{jour}", "heures": 2, "taux": 60} for jour in (16, 17, 18)]},
    ]


def _spec_fac(**surcharges) -> dict:
    """Spec d'exemple à deux lignes (`_lignes_fac`), complétée des surcharges."""
    spec = json.loads(SPEC_EXEMPLE.read_text(encoding="utf-8"))
    spec["lignes"] = _lignes_fac()
    spec.update(surcharges)
    return spec


def _instantane(dossier: Path) -> dict[str, tuple[bytes, int]]:
    """Contenu et date de modification de chaque fichier sous `dossier`, pour constater que rien n'a été écrit : la
    date attrape aussi une réécriture à l'identique."""
    return {
        p.relative_to(dossier).as_posix(): (p.read_bytes(), p.stat().st_mtime_ns)
        for p in sorted(dossier.rglob("*")) if p.is_file()
    }


def _classeur_existant(racine: Path) -> Path:
    """Classeur de l'école ABC qui porte déjà une facture (« R_ABC000 Août 2026 », N° R_ABC000 en D17), daté du
    01/01/2020 : une réécriture à l'identique se voit aussi."""
    from openpyxl import Workbook

    classeur = racine / CLASSEUR_ABC
    classeur.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    wb.active.title = "R_ABC000 Août 2026"
    wb.active["D17"] = "R_ABC000"
    wb.save(classeur)
    os.utime(classeur, ns=(1_577_836_800 * 10**9, 1_577_836_800 * 10**9))
    return classeur


def _message_erreur(res: subprocess.CompletedProcess) -> str:
    """Le texte qui suit le premier « [ERREUR] » (sortie d'erreur, sinon sortie standard), ou "" s'il n'y en a pas :
    l'aperçu qui le précède, et qui cite déjà le classeur, n'en fait pas partie."""
    for flux in (res.stderr, res.stdout):
        if "[ERREUR]" in flux:
            return flux[flux.index("[ERREUR]"):]
    return ""


def _refus_du_script(tmp_path: Path, spec: dict, options: tuple = ()) -> str:
    """Lance le script sur `spec`, avec un classeur d'école qui porte déjà une facture, et attend un refus propre :
    code 1, sans trace Python, un message « [ERREUR] », et rien d'écrit sous la racine (aucun onglet ajouté au
    classeur, aucun fichier créé). Rend le message."""
    racine = tmp_path / "donnees"
    creer_master(racine)
    _classeur_existant(racine)
    avant = _instantane(racine)
    chemin = tmp_path / "spec.json"
    chemin.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")

    res = _lancer("scripts/nouvelle_facture.py", str(chemin), "--data-root", str(racine), *options)

    assert res.returncode == 1, res.stdout + res.stderr
    assert "Traceback" not in res.stderr, res.stderr
    message = _message_erreur(res)
    assert message, f"aucun message « [ERREUR] » : {res.stdout + res.stderr}"
    assert _instantane(racine) == avant, "le refus a laissé un onglet ou un fichier écrit sous la racine"
    return message


def test_t_fac_18_script_apercu_arrondi_au_centime(tmp_path):
    """T-FAC-18 — `--dry-run` : 1,5 h × 33,33 € s'affiche 50.00 € sur la ligne et au TOTAL HT, au format inchangé
    (49.99 € avant le correctif, arrondi du float binaire 49,99499…)."""
    racine = tmp_path / "donnees"
    creer_master(racine)
    spec = _spec(tmp_path, lignes=[{"promotion": "M1 S1", "libelle": "Module A",
                                    "seances": [{"date": "2026-09-15", "heures": 1.5, "taux": 33.33}]}])
    res = _lancer("scripts/nouvelle_facture.py", str(spec), "--data-root", str(racine), "--dry-run")
    assert res.returncode == 0, res.stdout + res.stderr
    seances = [ligne for ligne in res.stdout.splitlines() if ligne.startswith("  | ")]
    assert len(seances) == 1 and seances[0].endswith("|     50.00 €"), res.stdout
    assert "TOTAL HT     : 50.00 €" in res.stdout.splitlines(), res.stdout


@pytest.mark.parametrize("options", [(), ("--dry-run",)], ids=["ecriture", "dry-run"])
@pytest.mark.parametrize("seances", ["vide", "absente"])
def test_t_fac_19_script_refuse_une_ligne_sans_seance(tmp_path, seances, options):
    """T-FAC-19 — première ligne sans séance (`seances` vide ou absente) : refus avant toute écriture, dès le
    `--dry-run`, code 1, message « [ERREUR] » qui cite le rang et le libellé de la ligne, sans trace Python, rien
    d'écrit. Avant le correctif : écrite puis déclarée non conforme (code 2) par un faux écart, ou KeyError brute."""
    spec = _spec_fac()
    if seances == "vide":
        spec["lignes"][0]["seances"] = []
    else:
        del spec["lignes"][0]["seances"]

    message = _refus_du_script(tmp_path, spec, options)

    assert spec["lignes"][0]["libelle"] in message, message
    assert re.search(r"\bligne (n° ?)?1\b", message, re.IGNORECASE), message


@pytest.mark.parametrize(
    ("champ", "valeur", "citee"), [("heures", "1,5", "1,5"), ("taux", float("nan"), "nan")],
    ids=["heures-chaine", "taux-nan"],
)
def test_t_fac_20_script_refuse_une_seance_mal_formee(tmp_path, champ, valeur, citee):
    """T-FAC-20 — séance 3 de la ligne 2 aux heures « 1,5 » (chaîne) ou au taux NaN : refus avant toute écriture,
    code 1, message « [ERREUR] » qui nomme le champ, les rangs et la valeur, sans trace Python. Avant le correctif :
    ValueError brute de `float()`, ou facture écrite puis déclarée non conforme (code 2)."""
    spec = _spec_fac()
    spec["lignes"][1]["seances"][2][champ] = valeur

    message = _refus_du_script(tmp_path, spec)

    assert champ in message and citee in message.lower(), message
    assert re.search(r"\bligne (n° ?)?2\b", message, re.IGNORECASE), message
    assert re.search(r"\bs[ée]ance (n° ?)?3\b", message, re.IGNORECASE), message


def test_t_fac_21_script_refuse_un_taux_negatif(tmp_path):
    """T-FAC-21 — taux négatif : refus avant toute écriture, code 1, message « [ERREUR] » qui nomme le taux et la
    valeur. Avant le correctif, la facture s'écrivait (code 0)."""
    spec = _spec_fac()
    spec["lignes"][1]["seances"][2]["taux"] = -90

    message = _refus_du_script(tmp_path, spec)

    assert "taux" in message and "-90" in message, message


def test_t_fac_22_script_refuse_masquer_jour_non_booleen(tmp_path):
    """T-FAC-22 — `masquer_jour` à la chaîne « false » : refus avant toute écriture, code 1, message « [ERREUR] » qui
    nomme la clé et la valeur. Avant le correctif, la chaîne valait vrai : facture écrite, colonne Jour masquée."""
    message = _refus_du_script(tmp_path, _spec_fac(masquer_jour="false"))

    assert "masquer_jour" in message and "false" in message, message


def test_t_fac_23_script_accepte_une_spec_avec_bom(tmp_path):
    """T-FAC-23 — une spec en UTF-8 avec BOM (Bloc-notes, PowerShell 5) donne exactement le même aperçu que sans BOM,
    puis s'écrit : code 0, facture conforme. Avant le correctif : JSONDecodeError « Unexpected UTF-8 BOM »."""
    racine = tmp_path / "donnees"
    creer_master(racine)
    sans_bom = _spec(tmp_path)
    avec_bom = tmp_path / "spec_bom.json"
    avec_bom.write_bytes(b"\xef\xbb\xbf" + sans_bom.read_bytes())

    attendu = _lancer("scripts/nouvelle_facture.py", str(sans_bom), "--data-root", str(racine), "--dry-run")
    apercu = _lancer("scripts/nouvelle_facture.py", str(avec_bom), "--data-root", str(racine), "--dry-run")
    assert attendu.returncode == 0, attendu.stdout + attendu.stderr
    assert apercu.returncode == 0, apercu.stdout + apercu.stderr
    assert apercu.stdout == attendu.stdout

    ecriture = _lancer("scripts/nouvelle_facture.py", str(avec_bom), "--data-root", str(racine))
    assert ecriture.returncode == 0, ecriture.stdout + ecriture.stderr
    assert "OK — facture R_ABC001 conforme" in ecriture.stdout, ecriture.stdout


# Lance le script avec le classeur cible verrouillé, comme ouvert dans Excel ou LibreOffice : l'ouvrir pour
# l'écrire, ou le remplacer par un fichier temporaire, lève PermissionError ; le lire reste possible.
LANCEUR_CLASSEUR_OUVERT = '''\
import os
import runpy
import sys
from pathlib import Path

import openpyxl

verrouille = Path(sys.argv[1]).resolve()
script = sys.argv[2]
sys.argv = [script, *sys.argv[3:]]


def _vise(chemin):
    return isinstance(chemin, (str, os.PathLike)) and Path(chemin).resolve() == verrouille


def _refuser(chemin):
    raise PermissionError(13, "Permission denied", str(chemin))


_vrai_save = openpyxl.Workbook.save


def _save(self, filename):
    if _vise(filename):
        _refuser(filename)
    return _vrai_save(self, filename)


openpyxl.Workbook.save = _save


def _deplacement_verrouille(vraie):
    def deplacer(src, dst, *args, **kwargs):
        if _vise(dst):
            _refuser(dst)
        return vraie(src, dst, *args, **kwargs)
    return deplacer


os.replace = _deplacement_verrouille(os.replace)
os.rename = _deplacement_verrouille(os.rename)

runpy.run_path(script, run_name="__main__")
'''


def test_t_fac_24_classeur_ouvert_a_l_enregistrement(tmp_path):
    """T-FAC-24 — classeur de l'école ouvert dans Excel ou LibreOffice (PermissionError à l'enregistrement) : message
    « [ERREUR] » qui nomme le classeur et dit de le fermer, sans trace Python, code 1, et aucun fichier partiel
    laissé (le classeur garde ses octets et sa date, aucun fichier temporaire ne reste). Avant le correctif : trace
    Python brute de la PermissionError."""
    racine = tmp_path / "donnees"
    creer_master(racine)
    classeur = _classeur_existant(racine)
    avant = _instantane(racine)
    lanceur = tmp_path / "lanceur_classeur_ouvert.py"
    lanceur.write_text(LANCEUR_CLASSEUR_OUVERT, encoding="utf-8")

    res = _lancer(str(lanceur), str(classeur), "scripts/nouvelle_facture.py", str(SPEC_EXEMPLE),
                  "--data-root", str(racine))

    assert res.returncode == 1, res.stdout + res.stderr
    assert "Traceback" not in res.stderr, res.stderr
    message = _message_erreur(res)
    assert classeur.name in message and re.search(r"ferme", message, re.IGNORECASE), res.stdout + res.stderr
    assert _instantane(racine) == avant, "un fichier partiel ou temporaire est resté sous la racine"


@pytest.mark.parametrize("options", [(), ("--dry-run",)], ids=["ecriture", "dry-run"])
def test_t_fac_25_script_refuse_un_titre_d_onglet_trop_long(tmp_path, options):
    """T-FAC-25 — titre d'onglet de 32 caractères : refus avant toute écriture, dès le `--dry-run`, code 1, message
    « [ERREUR] » qui donne le titre et sa longueur. Avant le correctif : un avertissement d'openpyxl, et l'onglet
    écrit (code 0)."""
    assert len(TITRE_32) == 32
    message = _refus_du_script(tmp_path, _spec_fac(onglet=TITRE_32), options)

    assert TITRE_32 in message and re.search(r"(?<![\w.])32(?!\w)", message), message
