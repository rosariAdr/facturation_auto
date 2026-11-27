"""Chemins, mois et noms du lot 1 (P-0b, `facturation/chemins.py` ; banque de tests : T-FIC-20 à T-FIC-34,
puis T-FIC-36, 38 et 42 pour la revue).

Tests d'acceptation écrits avant l'implémentation : chaque test importe `facturation.chemins`
lui-même, pour que l'absence du module fasse échouer les tests sans bloquer la collecte des
autres. Le contrat ne donne pas de préfixe propre à `chemins.py` : ses tests prennent la suite
des numéros de T-FIC (primitives de fichiers du même ticket).

Contrat (contrats des vagues 3 à 5, § 1.1, 1.7, 3.2 et 15) :
- abréviations des mois des onglets : Janv, Févr, Mars, Avr, Mai, Juin, Juil, Août, Sept, Oct,
  Nov, Déc (décision du § 15) ; `libelle_periode("2026-12") == "Déc 2026"` ;
- `onglet_facture` refuse un nom de plus de 31 caractères ou qui contient `[]:*?/\\` ;
- `cle_seance` = `f"{uid}@{debut:%Y-%m-%dT%H:%M%z}"` ;
- `refuser_racine_dans_un_depot` refuse une racine dont un parent contient `.git` ;
- arborescence des sorties sous `<racine>/<ENTITE>/` (§ 1.7) ; journal de la tâche à côté du profil.

Lectures retenues (`💡 [proposé — à valider]`) : `.git` peut être un dossier (clone) ou un
fichier (copie de travail `git worktree`) ; les fonctions de chemin ne créent aucun dossier.
"""

from __future__ import annotations

import importlib
import os
import sys
from datetime import date, datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pytest

from facturation import config
from facturation.agenda import Seance

pytestmark = pytest.mark.unit

PARIS = ZoneInfo("Europe/Paris")


def _chm():
    """Module testé, importé dans le test : son absence fait échouer le test, pas la collecte."""
    return importlib.import_module("facturation.chemins")


def _profil(tmp_path: Path, entite: str = "ROSARI") -> config.Profil:
    """Profil fictif écrit dans `tmp_path`, racine `tmp_path/donnees`."""
    chemin = tmp_path / f"profil_{entite}.toml"
    chemin.write_text(config.rediger("Personne A", entite, str(tmp_path / "donnees")), encoding="utf-8")
    return config.charger(chemin)


# ──────────────────────────────────────────────────────────────────────
# Mois et libellés
# ──────────────────────────────────────────────────────────────────────


def test_t_fic_20_noms_et_abreviations_des_mois():
    """T-FIC-20 — abréviations des onglets (décision du § 15) et noms des mois en français."""
    chm = _chm()
    assert tuple(chm.ABREVIATIONS_MOIS) == (
        "Janv", "Févr", "Mars", "Avr", "Mai", "Juin", "Juil", "Août", "Sept", "Oct", "Nov", "Déc",
    )
    assert tuple(chm.NOMS_MOIS) == (
        "janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre",
        "octobre", "novembre", "décembre",
    )


@pytest.mark.parametrize(
    "mois, attendu",
    [("2026-12", (date(2026, 12, 1), date(2026, 12, 31))), ("2027-02", (date(2027, 2, 1), date(2027, 2, 28))),
     ("2028-02", (date(2028, 2, 1), date(2028, 2, 29))), ("2026-04", (date(2026, 4, 1), date(2026, 4, 30)))],
)
def test_t_fic_21_mois_valide(mois, attendu):
    """T-FIC-21 — `mois_valide("AAAA-MM")` rend le premier et le dernier jour du mois."""
    assert _chm().mois_valide(mois) == attendu


@pytest.mark.parametrize("mois", ["2026-13", "2026-00", "2026-1", "26-12", "2026/12", "décembre", ""])
def test_t_fic_22_mois_invalide_refuse(mois):
    """T-FIC-22 — un mois qui n'est pas `AAAA-MM` (mois de 01 à 12) lève `ValueError` qui le cite."""
    with pytest.raises(ValueError) as erreur:
        _chm().mois_valide(mois)
    if mois:
        assert mois in str(erreur.value)


@pytest.mark.parametrize(
    "mois, precedent, suivant",
    [("2026-12", "2026-11", "2027-01"), ("2027-01", "2026-12", "2027-02"), ("2026-06", "2026-05", "2026-07")],
)
def test_t_fic_23_mois_precedent_et_suivant(mois, precedent, suivant):
    """T-FIC-23 — `mois_precedent` et `mois_suivant`, changement d'année compris ; un mois invalide
    est refusé."""
    chm = _chm()
    assert chm.mois_precedent(mois) == precedent
    assert chm.mois_suivant(mois) == suivant
    with pytest.raises(ValueError):
        chm.mois_precedent("2026-13")
    with pytest.raises(ValueError):
        chm.mois_suivant("2026-13")


@pytest.mark.parametrize(
    "mois, attendu",
    [("2026-12", "Déc 2026"), ("2026-09", "Sept 2026"), ("2027-02", "Févr 2027"), ("2027-07", "Juil 2027"),
     ("2027-08", "Août 2027"), ("2027-05", "Mai 2027")],
)
def test_t_fic_24_libelle_periode(mois, attendu):
    """T-FIC-24 — `libelle_periode` : abréviation du mois et année, comme les onglets récents."""
    assert _chm().libelle_periode(mois) == attendu


def test_t_fic_25_onglet_facture():
    """T-FIC-25 — `onglet_facture` : « N° Période » ; 31 caractères au plus (limite d'Excel)."""
    chm = _chm()
    assert chm.onglet_facture("R_ABC001", "Déc 2026") == "R_ABC001 Déc 2026"
    trente_et_un = chm.onglet_facture("R_ABC001", "x" * 22)
    assert len(trente_et_un) == 31
    with pytest.raises(ValueError):
        chm.onglet_facture("R_ABC001", "x" * 23)


@pytest.mark.parametrize("interdit", list("[]:*?/\\"))
def test_t_fic_26_onglet_facture_caractere_interdit(interdit):
    """T-FIC-26 — un caractère interdit par Excel dans un nom d'onglet lève `ValueError`."""
    with pytest.raises(ValueError):
        _chm().onglet_facture("R_ABC001", f"Déc{interdit}2026")


def test_t_fic_27_cle_seance():
    """T-FIC-27 — `cle_seance` : uid, « @ », puis le début à la minute avec son décalage."""
    debut = datetime(2026, 12, 1, 9, 0, tzinfo=PARIS)
    seance = Seance(uid="fictif-001@example.org", date=debut.date(), debut=debut,
                    fin=datetime(2026, 12, 1, 12, 30, tzinfo=PARIS), duree_h=3.5,
                    intitule="Cours fictif", description="")
    assert _chm().cle_seance(seance) == "fictif-001@example.org@2026-12-01T09:00+0100"


# ──────────────────────────────────────────────────────────────────────
# Racine hors dépôt
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize("nature", ["dossier", "fichier"], ids=["clone", "worktree"])
def test_t_fic_28_racine_dans_un_depot_refusee(tmp_path, nature):
    """T-FIC-28 — une racine de données placée dans un dépôt git (un parent contient `.git`, dossier
    d'un clone ou fichier d'une copie de travail) lève `ValueError` qui parle du dépôt : aucune
    donnée réelle dans le dépôt."""
    depot = tmp_path / "depot"
    depot.mkdir()
    if nature == "dossier":
        (depot / ".git").mkdir()
    else:
        (depot / ".git").write_text("gitdir: ../ailleurs/.git/worktrees/copie\n", encoding="utf-8")
    racine = depot / "donnees" / "sous-dossier"
    racine.mkdir(parents=True)
    with pytest.raises(ValueError, match=r"(?i)dépôt"):
        _chm().refuser_racine_dans_un_depot(racine)


def test_t_fic_29_racine_hors_depot_acceptee(tmp_path):
    """T-FIC-29 — une racine hors de tout dépôt est acceptée (la garde de T-FIC-28 ne refuse pas tout)."""
    racine = tmp_path / "donnees"
    racine.mkdir()
    assert not any((p / ".git").exists() for p in (racine, *racine.parents))
    assert _chm().refuser_racine_dans_un_depot(racine) is None


# ──────────────────────────────────────────────────────────────────────
# Arborescence des sorties
# ──────────────────────────────────────────────────────────────────────


def test_t_fic_30_arborescence_des_sorties(tmp_path):
    """T-FIC-30 — chemins des sorties sous `<racine>/<ENTITE>/_sorties/` (§ 1.7)."""
    chm = _chm()
    profil = _profil(tmp_path)
    entite = tmp_path / "donnees" / "ROSARI"
    sorties = entite / "_sorties"
    assert chm.dossier_mois(profil, "2026-12") == sorties / "2026-12"
    assert chm.journal_mois(profil, "2026-12") == sorties / "2026-12" / "journal.jsonl"
    assert chm.journal_ecritures(profil) == sorties / "journal.jsonl"
    assert chm.dossier_emissions(profil) == sorties / "emissions"
    assert chm.piece_emission(profil, "R_ABC001") == sorties / "emissions" / "R_ABC001.json"
    assert chm.piece_finale(profil, "R_ABC001") == sorties / "emissions" / "R_ABC001.final.json"
    assert chm.pdf_en_attente(profil, "R_ABC001") == sorties / "emissions" / "R_ABC001.pdf"
    assert chm.dossier_abandonnees(profil) == sorties / "emissions" / "abandonnees"
    assert chm.dossier_audits(profil) == sorties / "audits"
    assert chm.fichier_signalements(profil) == sorties / "signalements.jsonl"
    assert not (tmp_path / "donnees").exists(), "un chemin calculé ne crée aucun dossier"


def test_t_fic_31_classeur_et_pdf_d_une_ecole(tmp_path):
    """T-FIC-31 — classeur d'école `<racine>/<E>/<ECOLE>/<E>_<ECOLE>.xlsx`, PDF `<onglet>.pdf` à côté."""
    chm = _chm()
    profil = _profil(tmp_path)
    dossier_ecole = tmp_path / "donnees" / "ROSARI" / "ABC"
    assert chm.classeur_ecole(profil, "ABC") == dossier_ecole / "ROSARI_ABC.xlsx"
    assert chm.pdf_facture(profil, "ABC", "R_ABC001 Déc 2026") == dossier_ecole / "R_ABC001 Déc 2026.pdf"


def test_t_fic_32_entites_jamais_melangees(tmp_path):
    """T-FIC-32 — deux entités, deux arborescences : rien de l'une sous l'autre (ADR001)."""
    chm = _chm()
    rosari, kaur = _profil(tmp_path, "ROSARI"), _profil(tmp_path, "KAUR")
    assert chm.journal_mois(kaur, "2026-12") == tmp_path / "donnees" / "KAUR" / "_sorties" / "2026-12" / "journal.jsonl"
    assert chm.classeur_ecole(kaur, "ABC") == tmp_path / "donnees" / "KAUR" / "ABC" / "KAUR_ABC.xlsx"
    assert chm.journal_ecritures(rosari) != chm.journal_ecritures(kaur)


def test_t_fic_33_mois_invalide_dans_un_chemin(tmp_path):
    """T-FIC-33 — un mois invalide ne fabrique pas de dossier fantaisiste : `ValueError`."""
    chm = _chm()
    profil = _profil(tmp_path)
    for fonction in (chm.dossier_mois, chm.journal_mois):
        with pytest.raises(ValueError):
            fonction(profil, "2026-13")


def test_t_fic_34_journal_de_la_tache_a_cote_du_profil(tmp_path, monkeypatch):
    """T-FIC-34 — le journal de la tâche vit à côté du profil (hors racine), ou à côté de
    l'emplacement par défaut du profil quand aucun chemin n'est donné."""
    chm = _chm()
    profil = tmp_path / "config" / "profil.toml"
    assert chm.journal_tache(profil) == tmp_path / "config" / "journal_tache.jsonl"
    defaut = tmp_path / "appdata" / "facturation" / "profil.toml"
    monkeypatch.setenv("FACTURATION_PROFIL", str(defaut))
    assert config.emplacement_par_defaut() == defaut
    assert chm.journal_tache(None) == defaut.parent / "journal_tache.jsonl"


# ──────────────────────────────────────────────────────────────────────
# Revue de P-0b (T-FIC-36, T-FIC-38, T-FIC-42)
# ──────────────────────────────────────────────────────────────────────
#
# Constats des deux relecteurs. Écrits rouges avant le correctif : T-FIC-38 et T-FIC-42. T-FIC-36
# couvre une garde déjà codée que les relecteurs ont trouvée sans test : il est vert, et a été
# vérifié en cassant la garde (AGENTS.md §5).
# - T-FIC-36 : un n° de facture, un code école ou un nom d'onglet qui sortirait de son dossier
#   (« .. », séparateur, lecteur « C: ») est refusé par chaque fonction de chemin ;
# - T-FIC-38 : `refuser_racine_dans_un_depot` comparait le chemin écrit (`absolute`), pas le chemin
#   réel (`resolve`) : une racine `<dépôt>/../donnees`, hors du dépôt, était refusée à tort, et une
#   jonction ou un lien qui mène dans un dépôt passait ;
# - T-FIC-42 : `\d` accepte les chiffres non ASCII (« ２０２６-12 ») ; le mois brut devenait un
#   dossier `_sorties/２０２６-12` distinct de `_sorties/2026-12`. Lecture retenue (`💡 [proposé —
#   à valider]`) : un tel mois est refusé.


@pytest.mark.parametrize(
    "nom", ["..", ".", "../x", "R_ABC001/..", "ABC\\x", "C:x", "", "   "],
    ids=["parent", "courant", "parent-puis-nom", "nom-puis-parent", "barre-inverse", "lecteur", "vide", "espaces"],
)
def test_t_fic_36_nom_de_fichier_sans_traversee(tmp_path, nom):
    """T-FIC-36 — un n° de facture, un code école ou un nom d'onglet qui sortirait de son dossier
    lève `ValueError`, dans chacune des fonctions de chemin qui le reçoivent."""
    chm = _chm()
    profil = _profil(tmp_path)
    appels = {
        "piece_emission": lambda: chm.piece_emission(profil, nom),
        "piece_finale": lambda: chm.piece_finale(profil, nom),
        "pdf_en_attente": lambda: chm.pdf_en_attente(profil, nom),
        "classeur_ecole": lambda: chm.classeur_ecole(profil, nom),
        "pdf_facture (école)": lambda: chm.pdf_facture(profil, nom, "R_ABC001 Déc 2026"),
        "pdf_facture (onglet)": lambda: chm.pdf_facture(profil, "ABC", nom),
    }
    for fonction, appel in appels.items():
        with pytest.raises(ValueError):
            appel()
            pytest.fail(f"{fonction} accepte « {nom} »")


def test_t_fic_38a_racine_hors_depot_ecrite_avec_un_parent(tmp_path):
    """T-FIC-38 — une racine écrite `<dépôt>/../donnees` est hors du dépôt : elle est acceptée (la
    racine est comparée après résolution du chemin)."""
    depot = tmp_path / "depot"
    (depot / ".git").mkdir(parents=True)
    (tmp_path / "donnees").mkdir()
    assert _chm().refuser_racine_dans_un_depot(depot / ".." / "donnees") is None


def _lien_vers_dossier(lien: Path, cible: Path) -> None:
    """Lien symbolique vers un dossier ; sous Windows sans droit de lien, une jonction (mklink /J)."""
    try:
        os.symlink(cible, lien, target_is_directory=True)
        return
    except (OSError, NotImplementedError):
        if sys.platform != "win32":
            raise
    try:
        import _winapi
        _winapi.CreateJunction(str(cible), str(lien))
    except (ImportError, AttributeError, OSError) as exc:
        pytest.skip(f"ni lien symbolique ni jonction possible ici : {exc}")


def test_t_fic_38b_racine_qui_mene_dans_un_depot_par_un_lien(tmp_path):
    """T-FIC-38 — une racine qui mène dans un dépôt par une jonction ou un lien symbolique est refusée
    comme si elle y était écrite."""
    depot = tmp_path / "depot"
    (depot / ".git").mkdir(parents=True)
    (depot / "donnees").mkdir()
    lien = tmp_path / "raccourci"
    _lien_vers_dossier(lien, depot / "donnees")
    try:
        assert not any((p / ".git").exists() for p in (lien, *lien.parents))  # la garde ne passe pas à vide
        with pytest.raises(ValueError, match=r"(?i)dépôt"):
            _chm().refuser_racine_dans_un_depot(lien)
    finally:
        if sys.platform == "win32":
            os.rmdir(lien)  # retire le lien ou la jonction, pas le dossier visé
        else:
            lien.unlink()


@pytest.mark.parametrize("mois", ["２０２６-12", "٢٠٢٦-12", "2026-1２"], ids=["pleine-chasse", "arabe", "mois-pleine-chasse"])
def test_t_fic_42_mois_en_chiffres_non_ascii_refuse(tmp_path, mois):
    """T-FIC-42 — un mois écrit avec des chiffres non ASCII est refusé (`ValueError` qui le cite), par
    `mois_valide` comme par `dossier_mois` : jamais de second dossier du même mois (💡)."""
    chm = _chm()
    with pytest.raises(ValueError) as erreur:
        chm.mois_valide(mois)
    assert mois in str(erreur.value)
    with pytest.raises(ValueError):
        chm.dossier_mois(_profil(tmp_path), mois)
