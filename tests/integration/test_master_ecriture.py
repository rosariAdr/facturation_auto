"""Écriture sûre du MASTER (F-11, ADR013, ADR032 ; banque de tests : T-MST-08 à T-MST-16).

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


def _refus_sans_effet(master: Path, poste: str = "POSTE-A") -> BaseException:
    """Tente une écriture qui doit être refusée : renvoie l'erreur, après avoir vérifié qu'on n'est pas
    entré dans le bloc et que ni le MASTER ni les sauvegardes n'ont changé."""
    ecriture_master = _ecriture_master()
    avant, sauvegardes, entre = master.read_bytes(), _sauvegardes(master), []
    with pytest.raises(Exception) as exc:
        with ecriture_master(master, poste=poste):
            entre.append(True)
    assert not isinstance(exc.value, ERREURS_DE_PROGRAMMATION), repr(exc.value)
    assert entre == []
    assert _sauvegardes(master) == sauvegardes
    assert master.read_bytes() == avant
    return exc.value


def test_t_mst_12_verrou_laisse_par_le_meme_poste_remplace(master):
    """T-MST-12 — un verrou laissé par ce même poste (arrêt brutal) est remplacé, avec un avertissement
    qui nomme le poste ; l'écriture se fait, avec sa sauvegarde, et le verrou est retiré à la sortie."""
    ecriture_master = _ecriture_master()
    verrou = master.with_suffix(".verrou")
    with ecriture_master(master, poste="POSTE-A"):
        orphelin = verrou.read_bytes()  # verrou authentique, sans supposer son format
    verrou.write_bytes(orphelin)  # resté en place après un arrêt brutal
    sauvegardes = _sauvegardes(master)

    with pytest.warns(UserWarning) as avertissements:
        with ecriture_master(master, poste="POSTE-A"):
            assert verrou.exists()
            _ecrire_une_seance(master)

    assert any("POSTE-A" in str(a.message) for a in avertissements), [str(a.message) for a in avertissements]
    assert len(_sauvegardes(master)) == len(sauvegardes) + 1
    assert not verrou.exists()


def test_t_mst_13_ecriture_imbriquee_refusee(master):
    """T-MST-13 — une écriture ouverte dans une autre, par le même programme et le même poste : refus
    (pas d'avertissement « verrou laissé » ni de remplacement), et le verrou de l'écriture en cours reste."""
    ecriture_master = _ecriture_master()
    verrou = master.with_suffix(".verrou")
    with ecriture_master(master, poste="POSTE-A"):
        tenu = verrou.read_bytes()
        _refus_sans_effet(master, poste="POSTE-A")
        assert verrou.read_bytes() == tenu
    assert not verrou.exists()


def test_t_mst_14_deux_sauvegardes_dans_la_meme_seconde(master, monkeypatch):
    """T-MST-14 — trois écritures dans la même seconde : sauvegardes suffixées `-1`, `-2`, aucune écrasée.

    L'horloge du module est figée (`master.datetime`) : sans cela, le test dépendrait de la vitesse du poste.
    """
    module = importlib.import_module("facturation.master")
    assert isinstance(getattr(module, "datetime", None), type), "le module ne lit plus l'heure par `datetime`"

    class _Fige(datetime):
        @classmethod
        def now(cls, tz=None):
            return cls(2026, 11, 25, 16, 0, 0, tzinfo=tz)

    monkeypatch.setattr(module, "datetime", _Fige)
    etats = []
    for i in range(3):
        etats.append(master.read_bytes())
        with module.ecriture_master(master, poste="POSTE-A"):
            wb = load_workbook(master)
            wb["Sessions"].append([None, "ABC", None, f"R_ABC00{i + 1}"])
            wb.save(master)

    base = master.parent / "sauvegardes" / "ROSARI_MASTER_20261125-160000"
    copies = [Path(f"{base}.xlsx"), Path(f"{base}-1.xlsx"), Path(f"{base}-2.xlsx")]
    assert _sauvegardes(master) == sorted(copies)
    assert [c.read_bytes() for c in copies] == etats  # chacune garde l'état d'avant son écriture
    assert len(set(etats)) == 3  # les trois états diffèrent : un écrasement se verrait


def test_t_mst_15_refus_si_libreoffice_a_le_classeur_ouvert(master):
    """T-MST-15 — un fichier `.~lock.<nom>#` de LibreOffice à côté du MASTER : refus « fermez le classeur »,
    sans sauvegarde ni verrou laissé."""
    (master.parent / f".~lock.{master.name}#").write_text("Personne A,poste-fictif,01.01.2026 10:00,", encoding="utf-8")
    erreur = _refus_sans_effet(master)
    assert "fermez le classeur" in str(erreur).lower(), str(erreur)
    assert not master.with_suffix(".verrou").exists()


@pytest.mark.parametrize(
    "contenu",
    [b"", b"\xff\xfe\x00\x81\x00 illisible", b"verrou fictif sans nom de poste\n"],
    ids=["vide", "illisible", "sans-poste"],
)
def test_t_mst_16_verrou_vide_ou_illisible_refuse(master, contenu):
    """T-MST-16 — un verrou vide, illisible ou sans nom de poste : refus (on ne sait pas qui écrit), qui
    nomme le fichier verrou à supprimer s'il est orphelin ; le verrou est laissé tel quel."""
    verrou = master.with_suffix(".verrou")
    verrou.write_bytes(contenu)
    erreur = _refus_sans_effet(master)
    assert verrou.name in str(erreur), str(erreur)
    assert verrou.read_bytes() == contenu
