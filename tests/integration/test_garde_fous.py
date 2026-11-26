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
    # Émetteur entrepreneur individuel (« EI ») et n° de déclaration d'activité : obligatoires (ADR030).
    ws.append(["Émetteur", entite, f"EMETTEUR FICTIF {entite} EI", "1 rue Fictive", "00000 VILLE-FICTIVE",
               "", f"{entite.lower()}@example.org", "+33000000000"])
    ws.append(["Info", entite, "SIRET", "00000000000000"])
    ws.append(["Info", entite, "IBAN", "FR76 0000 0000 0000 0000 0000 000"])
    ws.append(["Info", entite, "BIC", "TESTFRPP"])
    ws.append(["Info", entite, "NDA", f"Déclaration d'activité fictive de {entite}, n° 00000000000"])
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
    assert ws["C7"].value == "EMETTEUR FICTIF BETA EI"
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


# ──────────────────────────────────────────────────────────────────────
# FA-05 — l'aperçu montre tout ce qu'il faut valider
# ──────────────────────────────────────────────────────────────────────

LIBELLE_LONG = "[2026][CODE-PRODUIT-FICTIF-0001][PROMO 1][Module fictif au libellé volontairement long]"


def test_apercu_montre_l_emetteur_et_le_libelle_complet(tmp_path, racine):
    """Le dry-run affiche l'émetteur lu dans le MASTER et le libellé entier, sans troncature."""
    assert len(LIBELLE_LONG) > 80
    spec = spec_fictive()
    spec["lignes"][0]["libelle"] = LIBELLE_LONG
    res = lancer(tmp_path, spec, "--data-root", str(racine), "--dry-run")

    assert res.returncode == 0, res.stdout + res.stderr
    assert "Émetteur      : EMETTEUR FICTIF ALPHA EI — 1 rue Fictive, 00000 VILLE-FICTIVE" in res.stdout
    assert LIBELLE_LONG in res.stdout


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
    # verifier_facture contrôle toujours l'échéance, le NDA et la ligne de contrat (F-24).
    fmt.apply_mentions(ws, entite="ALPHA", date_emission=date(2026, 10, 1))
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


# ──────────────────────────────────────────────────────────────────────
# FA-13 — le script importe le paquet de son dépôt, pas celui d'une autre copie
# ──────────────────────────────────────────────────────────────────────


@pytest.fixture
def autre_copie(tmp_path: Path) -> Path:
    """Dossier contenant un autre paquet `facturation`, qui s'arrête dès qu'on l'importe.

    Il joue l'installation éditable d'une autre copie du dépôt : placé sur le chemin
    d'import, il passe avant le paquet installé.
    """
    dossier = tmp_path / "autre_copie"
    paquet = dossier / "facturation"
    paquet.mkdir(parents=True)
    (paquet / "__init__.py").write_text('raise SystemExit("AUTRE COPIE importée : " + __file__)\n', encoding="utf-8")
    return dossier


def test_nouvelle_facture_importe_le_paquet_de_son_depot(tmp_path, autre_copie):
    """Lancé par son chemin, sans `lancer()` qui force PYTHONPATH, le script ignore l'autre copie."""
    env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONPATH": str(autre_copie)}
    res = subprocess.run(
        [sys.executable, str(SCRIPT), "--help"],
        capture_output=True, text=True, encoding="utf-8", env=env, cwd=tmp_path,
    )
    assert "AUTRE COPIE" not in res.stdout + res.stderr
    assert res.returncode == 0, res.stderr
    assert "--data-root" in res.stdout
