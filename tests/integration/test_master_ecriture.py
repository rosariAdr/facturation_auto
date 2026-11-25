"""Écriture sûre du MASTER (F-11, ADR013, ADR032 ; banque de tests : T-MST-08 à T-MST-11).

Tests d'acceptation écrits avant l'implémentation : chaque test importe `facturation.master`
lui-même. Classeurs fictifs créés dans `tmp_path`.

`ecriture_master(chemin, *, poste)` est un gestionnaire de contexte. Les tests écrivent eux-mêmes
dans le classeur à l'intérieur du bloc, avec openpyxl : ils ne supposent rien de ce que le
gestionnaire renvoie. Le verrou d'un autre poste est créé par un bloc `ecriture_master` ouvert
sous un autre nom de poste, pour ne pas supposer le format du fichier `.verrou`.
"""

from __future__ import annotations

import importlib
import re
from datetime import date, datetime, timedelta
from pathlib import Path

import pytest
from openpyxl import load_workbook

from tests.conftest import creer_master

pytestmark = pytest.mark.integration

ERREURS_DE_PROGRAMMATION = (AttributeError, TypeError, NameError, ImportError, KeyError, IndexError, AssertionError)


def _ecriture_master():
    """Gestionnaire testé, importé dans le test : son absence fait échouer le test, pas la collecte."""
    return importlib.import_module("facturation.master").ecriture_master


def _ecrire_une_seance(master: Path):
    wb = load_workbook(master)
    wb["Sessions"].append([date(2026, 10, 1), "ABC", "Sept 2026", "R_ABC001", date(2026, 9, 15), "09:00",
                           "12:30", 3.5, "M1 S1", "Module A", 60, 210, "#ABC-M1 Cours fictif"])
    wb.save(master)


def _sauvegardes(master: Path) -> list[Path]:
    return sorted((master.parent / "sauvegardes").glob("*.xlsx"))


@pytest.fixture
def master(tmp_path: Path) -> Path:
    return creer_master(tmp_path / "donnees")


def test_t_mst_08_sauvegarde_avant_toute_ecriture(master):
    """T-MST-08 — dès l'entrée dans le bloc, une copie intacte du MASTER est rangée dans
    `sauvegardes/<nom>_AAAAMMJJ-HHMMSS.xlsx` ; le verrou est retiré à la sortie."""
    ecriture_master = _ecriture_master()
    avant = master.read_bytes()

    with ecriture_master(master, poste="POSTE-A"):
        (copie,) = _sauvegardes(master)
        assert copie.read_bytes() == avant
        _ecrire_une_seance(master)

    assert copie.read_bytes() == avant
    assert master.read_bytes() != avant
    nom = re.fullmatch(r"ROSARI_MASTER_(\d{8}-\d{6})\.xlsx", copie.name)
    assert nom, copie.name
    assert abs(datetime.strptime(nom.group(1), "%Y%m%d-%H%M%S") - datetime.now()) < timedelta(minutes=5)
    assert not master.with_suffix(".verrou").exists()


def test_t_mst_09_refus_si_excel_a_le_classeur_ouvert(master):
    """T-MST-09 — un fichier `~$<nom>` à côté du MASTER : refus « fermez le classeur », avant toute
    sauvegarde, sans verrou laissé ni modification."""
    ecriture_master = _ecriture_master()
    (master.parent / f"~${master.name}").write_bytes(b"")
    avant = master.read_bytes()
    entre = []

    with pytest.raises(Exception) as exc:
        with ecriture_master(master, poste="POSTE-A"):
            entre.append(True)

    assert not isinstance(exc.value, ERREURS_DE_PROGRAMMATION), repr(exc.value)
    assert "fermez le classeur" in str(exc.value).lower(), str(exc.value)
    assert entre == []
    assert _sauvegardes(master) == []
    assert not master.with_suffix(".verrou").exists()
    assert master.read_bytes() == avant


def test_t_mst_10_refus_si_un_autre_poste_ecrit(master):
    """T-MST-10 — pendant qu'un poste écrit, son verrou `<ENTITE>_MASTER.verrou` porte son nom et l'heure ;
    un autre poste est refusé, avec le nom du poste qui écrit, sans sauvegarde ni verrou supprimé."""
    ecriture_master = _ecriture_master()
    verrou = master.with_suffix(".verrou")
    entre = []

    with ecriture_master(master, poste="POSTE-B"):
        assert verrou.exists()
        contenu = verrou.read_text(encoding="utf-8")
        assert "POSTE-B" in contenu and re.search(r"\d{1,2}[:h]\d{2}", contenu), contenu
        sauvegardes = _sauvegardes(master)

        with pytest.raises(Exception) as exc:
            with ecriture_master(master, poste="POSTE-A"):
                entre.append(True)

        assert not isinstance(exc.value, ERREURS_DE_PROGRAMMATION), repr(exc.value)
        assert "POSTE-B" in str(exc.value), str(exc.value)
        assert entre == []
        assert verrou.exists() and "POSTE-B" in verrou.read_text(encoding="utf-8")
        assert _sauvegardes(master) == sauvegardes

    assert not verrou.exists()


def test_t_mst_11_verrou_retire_meme_en_cas_d_erreur(master):
    """T-MST-11 — une erreur dans le bloc remonte telle quelle, et le verrou est retiré : un autre
    poste peut écrire ensuite."""
    ecriture_master = _ecriture_master()
    verrou = master.with_suffix(".verrou")

    with pytest.raises(RuntimeError, match="erreur fictive"):
        with ecriture_master(master, poste="POSTE-A"):
            assert verrou.exists()
            raise RuntimeError("erreur fictive")

    assert not verrou.exists()
    with ecriture_master(master, poste="POSTE-B"):
        _ecrire_une_seance(master)
    assert not verrou.exists()
