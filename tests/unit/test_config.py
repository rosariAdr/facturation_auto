"""Profil utilisateur (banque de tests : T-CFG-01 et suivants)."""

from __future__ import annotations

from pathlib import Path

import pytest

from facturation import config

pytestmark = pytest.mark.unit


def _ecrire(dossier: Path, texte: str) -> Path:
    dossier.mkdir(parents=True, exist_ok=True)
    chemin = dossier / "profil.toml"
    chemin.write_text(texte, encoding="utf-8")
    return chemin


MINIMAL = """
[utilisateur]
nom = "Personne A"
entite = "{entite}"

[donnees]
racine = "{racine}"
"""


def test_t_cfg_01_deux_profils_sorties_separees(tmp_path):
    """T-CFG-01 — deux profils, deux entités : MASTER et sorties ne se croisent pas."""
    racine = (tmp_path / "donnees").as_posix()
    a = config.charger(_ecrire(tmp_path / "a", MINIMAL.format(entite="ALPHA", racine=racine)))
    b = config.charger(_ecrire(tmp_path / "b", MINIMAL.format(entite="BETA", racine=racine)))
    assert a.master != b.master
    assert a.dossier_sorties != b.dossier_sorties
    assert a.master.name == "ALPHA_MASTER.xlsx"
    assert b.dossier_sorties == Path(racine) / "BETA" / "_sorties"


def test_valeurs_par_defaut_du_cadrage(tmp_path):
    """ADR024 — lancement le 25, émission au 1er du mois suivant, iCal, LibreOffice."""
    p = config.charger(_ecrire(tmp_path, MINIMAL.format(entite="ROSARI", racine="/donnees")))
    assert p.lancement_jour == 25
    assert p.lancement_heure == "16:00"
    assert p.date_emission == "premier_du_mois_suivant"
    assert p.agenda_source == "ical"
    assert p.fuseau == "Europe/Paris"
    assert p.pdf_moteur == "libreoffice"
    assert p.prefixe_facture == "R_"


def test_profil_incomplet_nomme_chaque_champ(tmp_path):
    """Un profil incomplet liste tous les champs manquants, section par section."""
    chemin = _ecrire(tmp_path, "[utilisateur]\nnom = \"Personne A\"\n")
    with pytest.raises(config.ProfilInvalide) as exc:
        config.charger(chemin)
    assert "[utilisateur] entite : obligatoire" in exc.value.erreurs
    assert "[donnees] racine : obligatoire" in exc.value.erreurs


def test_cles_inconnues_et_valeurs_invalides(tmp_path):
    """Une faute de frappe ou une valeur hors bornes est refusée avec un message explicite."""
    texte = MINIMAL.format(entite="ROSARI", racine="/donnees") + """
[lancement]
jour = 31
heur = "20:00"

[pdf]
moteur = "word"
"""
    with pytest.raises(config.ProfilInvalide) as exc:
        config.charger(_ecrire(tmp_path, texte))
    erreurs = "\n".join(exc.value.erreurs)
    assert "[lancement] heur : clé inconnue" in erreurs
    assert "[lancement] jour" in erreurs
    assert "[pdf] moteur" in erreurs


def test_entite_invalide(tmp_path):
    with pytest.raises(config.ProfilInvalide) as exc:
        config.charger(_ecrire(tmp_path, MINIMAL.format(entite="rosari", racine="/d")))
    assert any(e.startswith("[utilisateur] entite") for e in exc.value.erreurs)


def test_fichier_absent(tmp_path):
    with pytest.raises(config.ProfilInvalide) as exc:
        config.charger(tmp_path / "absent.toml")
    assert "facturation profil init" in str(exc.value)


def test_rediger_puis_charger_chemin_avec_apostrophe(tmp_path):
    """Le profil rédigé se relit tel quel, même avec un chemin Windows contenant une apostrophe."""
    racine = r"G:\My Drive\Documents d'auto-entrepreneur\Factures"
    chemin = tmp_path / "profil.toml"
    chemin.write_text(config.rediger("Personne A", "ROSARI", racine), encoding="utf-8")
    p = config.charger(chemin)
    assert p.racine.as_posix().endswith("Documents d'auto-entrepreneur/Factures")
    assert p.entite == "ROSARI"


def test_url_secrete_relative_au_profil(tmp_path):
    texte = MINIMAL.format(entite="ROSARI", racine="/d") + '\n[agenda]\nfichier_url_secrete = "agenda.secret"\n'
    p = config.charger(_ecrire(tmp_path, texte))
    assert p.agenda_fichier_url_secrete == tmp_path / "agenda.secret"


def test_emplacement_par_defaut_variable_environnement(tmp_path, monkeypatch):
    monkeypatch.setenv("FACTURATION_PROFIL", str(tmp_path / "ailleurs.toml"))
    assert config.emplacement_par_defaut() == tmp_path / "ailleurs.toml"
