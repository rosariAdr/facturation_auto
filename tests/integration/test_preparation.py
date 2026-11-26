"""Préparation multi-écoles : lancement, fichiers produits et commandes `facturation preparer` /
`facturation apercu` (F-16a ; banque de tests : T-PRP-04, 09, 11, 18, 19 et T-PRP-38 à T-PRP-44, les
autres dans `tests/unit/test_preparation.py`).

Tests d'acceptation écrits avant l'implémentation : chaque test importe `facturation.preparation`
(ou `facturation.cli`) lui-même, pour que l'absence du module fasse échouer les tests sans bloquer la
collecte des autres. Racine de données, MASTER et profils fictifs créés dans `tmp_path`, hors de tout
dépôt ; l'agenda est un contenu .ics fictif injecté (`contenu_ics` ou `--fichier`) : tout accès réseau
fait échouer le test. Les commandes sont lancées dans le processus (`cli.main`), avec `--profil`.

Contrat (contrats des vagues 3 à 5, § 1.4 à 1.7 et 4) :
- `lancer_preparation` refuse une racine placée dans un dépôt ; racine absente →
  `PreparationImpossible` (« dossier de données introuvable … Google Drive ») ; tout le reste se fait
  sous `verrou_ecriture` : `charger_entrees`, `preparer`, `differences` avec la préparation
  précédente, copie de l'ancien `preparation.json` dans `historique/`, écritures atomiques de
  `preparation.json`, des aperçus et de `a_classer.txt`, événement `preparation` / `ok` ;
- en cas d'échec (agenda, MASTER, verrou…) : événement `preparation` / `échec` au message masqué, puis
  `PreparationImpossible` levée depuis l'exception d'origine ; la préparation précédente reste
  identique octet pour octet ;
- fichiers produits sous `_sorties/<mois>/` seulement : `preparation.json`, `historique/`,
  `apercu.txt`, `apercus/`, `a_classer.txt`, `journal.jsonl` ; aucune écriture dans le MASTER ;
- `facturation preparer [--mois] [--ecoles] [--fichier] [--profil]` : codes 3 (facture à valider ou
  bloquée, créneau à classer), 0 sinon, 1 sur `PreparationImpossible` ou profil invalide ;
- `facturation apercu [--mois] [--facture ID|F<n>] [--ouvrir] [--profil]` : lecture seule, sans relire
  l'agenda ; 1 sans préparation (« lancez `facturation preparer --mois …` »), 0 sinon.

Lectures retenues (`💡 [proposé — à valider]`) :
- aperçu en texte seulement (Q-D7, valeur « Recommandé ») : aucun PDF ni XLSX sous `_sorties/` ;
- une racine dans un dépôt est refusée avant toute écriture, que le refus soit une `ValueError` ou
  une `PreparationImpossible` (le contrat ne range pas ce refus parmi les échecs journalisés) ;
- les champs `brouillons`, `bloques` et `a_classer` de l'événement `preparation` se lisent comme un
  nombre ou comme une liste (le schéma de P-0b ne fixe pas leur type) ;
- l'état d'une facture validée (`journal.etat_validation`) ne compte plus pour le code 3.
"""

from __future__ import annotations

import json
import os
import re
import shutil
import urllib.request
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pytest

from facturation import chemins, fichiers, journal
from facturation import pcf_format as fmt
from tests.aide_preparation import (
    MAINTENANT,
    MOIS,
    SEANCES_STANDARD,
    SeanceFictive,
    arborescence,
    brouillon,
    cle,
    ecrire_decision,
    ecrire_piece,
    ecrire_validation,
    ics,
    identifiants,
    installer,
    master_standard,
    profil,
    prp,
    regle,
    seances_sauf,
    standard,
)

pytestmark = pytest.mark.integration


@pytest.fixture(autouse=True)
def _isolement(monkeypatch):
    """Aucun accès réseau (l'agenda est injecté) ; racine de `pcf_format` restaurée après le test."""

    def refuser(*args, **kwargs):
        raise AssertionError("accès réseau interdit dans les tests : l'agenda est injecté (contenu_ics)")

    monkeypatch.setattr(urllib.request, "urlopen", refuser)
    ancienne = fmt._DATA_ROOT
    yield
    fmt._DATA_ROOT = ancienne
    fmt._MASTER_DATA_CACHE.clear()


def _lancer(p, contenu: bytes | None = None, **options):
    options.setdefault("maintenant", MAINTENANT)
    return prp().lancer_preparation(p, MOIS, contenu_ics=contenu if contenu is not None else ics(), **options)


def _dossier(p) -> Path:
    return chemins.dossier_mois(p, MOIS)


def _contenu_du_dossier(dossier: Path, *, sauf: tuple[str, ...] = ()) -> dict[str, bytes]:
    return {
        f.relative_to(dossier).as_posix(): f.read_bytes()
        for f in sorted(dossier.rglob("*")) if f.is_file() and f.name not in sauf
    }


def _cli(*arguments: str) -> int:
    from facturation import cli

    return cli.main(list(arguments))


def _compte(valeur) -> int:
    return valeur if isinstance(valeur, int) else len(valeur)


# ──────────────────────────────────────────────────────────────────────
# Lancement : relance, échec, entités, écritures
# ──────────────────────────────────────────────────────────────────────


def test_t_prp_04_relance_apres_deplacement_ajout_et_annulation(tmp_path):
    """T-PRP-04 — Relance après un déplacement, un ajout et une annulation dans l'agenda : mêmes
    identifiants, différences exactes (aucun « total modifié » : les heures sont inchangées),
    ancienne préparation dans `historique/`, aucun fichier en double."""
    p = installer(tmp_path)
    premiere = _lancer(p)
    dossier = _dossier(p)
    octets_premiere = (dossier / "preparation.json").read_bytes()
    assert list(premiere.differences) == []

    deplacee = standard("abc-2")._replace(jour="2026-12-17")
    ajoutee = SeanceFictive("abc-4", "2026-12-14", "09:00", "12:00", "Statistiques M1", "#ABC-M1-STATS")
    agenda_v2 = [deplacee if s.uid == "abc-2" else s for s in SEANCES_STANDARD if s.uid != "abc-1"] + [ajoutee]
    seconde = _lancer(p, ics(agenda_v2))

    assert identifiants(seconde) == identifiants(premiere) == {"ABC-2026-12", "ECO-2026-12"}
    assert sorted((d.identifiant, d.nature) for d in seconde.differences) == [
        ("ABC-2026-12", "séance ajoutée"), ("ABC-2026-12", "séance déplacée"), ("ABC-2026-12", "séance retirée")]
    historique = sorted((dossier / "historique").iterdir())
    assert len(historique) == 1
    assert re.fullmatch(r"preparation_[0-9]{8}-[0-9]{6}(-[0-9]+)?\.json", historique[0].name)
    assert historique[0].read_bytes() == octets_premiere
    assert sorted(f.name for f in (dossier / "apercus").iterdir()) == ["ABC-2026-12.txt", "ECO-2026-12.txt"]
    assert sorted(f.name for f in dossier.iterdir() if f.is_file()) == [
        "a_classer.txt", "apercu.txt", "journal.jsonl", "preparation.json"]
    assert not [f for f in dossier.rglob("*") if ".tmp" in f.name]
    assert prp().lire_preparation(p, MOIS).empreinte_agenda == seconde.empreinte_agenda


def test_t_prp_09_agenda_indisponible(tmp_path, monkeypatch):
    """T-PRP-09 — Agenda indisponible : `PreparationImpossible` levée depuis `AgendaIndisponible`,
    événement `échec` sans `https://` (message masqué, garde : « adresse masquée » y figure),
    préparation précédente identique octet pour octet."""
    from facturation import agenda

    P = prp()
    p = installer(tmp_path)
    _lancer(p)
    dossier = _dossier(p)
    avant = _contenu_du_dossier(dossier, sauf=("journal.jsonl",))

    def indisponible(profil_):
        raise agenda.AgendaIndisponible("agenda injoignable (https://agenda.example.invalid/prive/0123abcd/basic.ics)")

    monkeypatch.setattr(agenda, "telecharger_ics", indisponible)
    monkeypatch.setattr(P, "telecharger_ics", indisponible, raising=False)
    with pytest.raises(P.PreparationImpossible) as exc:
        P.lancer_preparation(p, MOIS, contenu_ics=None, maintenant=MAINTENANT)
    assert isinstance(exc.value.__cause__, agenda.AgendaIndisponible)

    assert _contenu_du_dossier(dossier, sauf=("journal.jsonl",)) == avant
    evenements = journal.lire_evenements(chemins.journal_mois(p, MOIS), types=["preparation"])
    assert [e["statut"] for e in evenements] == ["ok", "échec"]
    assert evenements[-1]["erreur"] and "adresse masquée" in evenements[-1]["erreur"]
    assert "https://" not in (dossier / "journal.jsonl").read_text(encoding="utf-8")


def test_t_prp_11_profil_kaur_isole(tmp_path):
    """T-PRP-11 — Un profil KAUR n'écrit et ne lit que sous `KAUR/` : l'état de ROSARI (pièce
    d'émission, décision) ne déteint pas, rien ne change sous `ROSARI/` ; numéros `K_`, classeur
    `KAUR/ABC/KAUR_ABC.xlsx`."""
    racine = tmp_path / "donnees"
    master_standard(racine, "ROSARI")
    master_standard(racine, "KAUR")
    rosari, kaur = profil(tmp_path, racine, "ROSARI"), profil(tmp_path, racine, "KAUR")
    ecrire_piece(rosari, "R_ABC001", identifiant="ABC-2026-12")
    ecrire_decision(rosari, "exclure", cle(standard("abc-1")), {"raison": "fictive"})
    avant = arborescence(racine)

    prep = _lancer(kaur)
    apres = arborescence(racine)
    assert {k: v for k, v in apres.items() if k.startswith("ROSARI/")} == {
        k: v for k, v in avant.items() if k.startswith("ROSARI/")}
    nouveaux = set(apres) - set(avant)
    assert nouveaux and all(k.startswith("KAUR/_sorties/2026-12/") for k in nouveaux)

    assert prep.entite == "KAUR"
    abc = brouillon(prep, "ABC-2026-12")
    assert "abc-1" in {ligne.uid for ligne in abc.lignes}
    assert abc.numero_probable == "K_ABC001"
    assert abc.emetteur["nom"] == "FORMATEUR FICTIF KAUR EI"
    assert abc.spec("K_ABC001")["fichier_sortie"] == "KAUR/ABC/KAUR_ABC.xlsx"
    assert prp().lire_preparation(rosari, MOIS) is None
    assert prp().lire_preparation(kaur, MOIS).entite == "KAUR"


def test_t_prp_11_preparation_d_une_autre_entite_refusee(tmp_path):
    """T-PRP-11 — Un `preparation.json` dont le champ `entite` diffère du profil est refusé
    (`ValueError` qui nomme l'entité trouvée, § 1.1)."""
    racine = tmp_path / "donnees"
    master_standard(racine, "ROSARI")
    master_standard(racine, "KAUR")
    rosari, kaur = profil(tmp_path, racine, "ROSARI"), profil(tmp_path, racine, "KAUR")
    _lancer(rosari)
    cible = chemins.dossier_mois(kaur, MOIS) / "preparation.json"
    cible.parent.mkdir(parents=True)
    shutil.copyfile(chemins.dossier_mois(rosari, MOIS) / "preparation.json", cible)
    with pytest.raises(ValueError, match="ROSARI"):
        prp().lire_preparation(kaur, MOIS)


def test_t_prp_18_aucune_ecriture_hors_des_sorties_du_mois(tmp_path):
    """T-PRP-18 — L'empreinte du MASTER est identique avant et après ; rien n'est écrit hors de
    `_sorties/<mois>/` (ni sauvegarde, ni verrou laissé) ; aperçu en texte seulement (Q-D7)."""
    p = installer(tmp_path)
    avant = arborescence(p.racine)
    empreinte_master = fichiers.empreinte_fichier(p.master)
    _lancer(p)
    apres = arborescence(p.racine)

    assert fichiers.empreinte_fichier(p.master) == empreinte_master
    assert {k: apres.get(k) for k in avant} == avant
    nouveaux = set(apres) - set(avant)
    assert nouveaux and all(k.startswith("ROSARI/_sorties/2026-12/") for k in nouveaux)
    assert {Path(k).suffix for k in nouveaux} <= {".json", ".jsonl", ".txt"}
    assert not p.master.with_suffix(".verrou").exists()
    assert not (p.racine / "ROSARI" / "sauvegardes").exists()


# ──────────────────────────────────────────────────────────────────────
# Commande `facturation preparer`
# ──────────────────────────────────────────────────────────────────────


def test_t_prp_19_commande_preparer(tmp_path, capsys):
    """T-PRP-19 — `facturation preparer --fichier <fixture.ics> --mois 2026-12` : code 3, « À FAIRE »
    affiché, avec l'en-tête du mois et le dossier des sorties."""
    p = installer(tmp_path)
    fichier = tmp_path / "agenda_fictif.ics"
    fichier.write_bytes(ics())

    code = _cli("preparer", "--fichier", str(fichier), "--mois", "2026-12", "--profil", str(p.chemin))
    sortie = capsys.readouterr().out
    assert code == 3
    assert "À FAIRE" in sortie
    assert "Mois : décembre 2026 — entité ROSARI" in sortie
    assert "_sorties" in sortie
    assert (_dossier(p) / "preparation.json").is_file()


def test_t_prp_41_codes_de_la_commande_preparer(tmp_path, capsys):
    """T-PRP-41 — Codes de `preparer` : 3 tant qu'une facture est à valider ; 0 quand toutes sont
    validées (empreinte courante) et que rien n'est à classer ; 0 sur un agenda vide ; 1 sur
    `PreparationImpossible` (MASTER introuvable) ou un profil invalide."""
    p = installer(tmp_path / "pilote")
    fichier = tmp_path / "agenda_sans_a_classer.ics"
    fichier.write_bytes(ics(seances_sauf("inconnu-1", "srg-1")))
    options = ("--fichier", str(fichier), "--mois", "2026-12", "--profil", str(p.chemin))

    assert _cli("preparer", *options) == 3
    for b in prp().lire_preparation(p, MOIS).brouillons:
        ecrire_validation(p, b.identifiant, b.empreinte)
    assert _cli("preparer", *options) == 0

    vide = tmp_path / "agenda_vide.ics"
    vide.write_bytes(ics([]))
    p_vide = installer(tmp_path / "vide")
    assert _cli("preparer", "--fichier", str(vide), "--mois", "2026-12", "--profil", str(p_vide.chemin)) == 0

    p_sans_master = installer(tmp_path / "sans_master")
    p_sans_master.master.unlink()
    assert _cli("preparer", "--fichier", str(fichier), "--mois", "2026-12", "--profil", str(p_sans_master.chemin)) == 1
    assert _cli("preparer", "--fichier", str(fichier), "--mois", "2026-12",
                "--profil", str(tmp_path / "profil_absent.toml")) == 1


# ──────────────────────────────────────────────────────────────────────
# Commande `facturation apercu`
# ──────────────────────────────────────────────────────────────────────


def test_t_prp_42_commande_apercu_en_lecture_seule(tmp_path, capsys, monkeypatch):
    """T-PRP-42 — `facturation apercu` relit `preparation.json` sans relire l'agenda (fichier .ics
    supprimé, réseau interdit) et sans rien écrire : aperçu global, aperçu d'une facture par numéro
    court (`F1`) ou par identifiant, `--ouvrir` ouvre le dossier du mois."""
    P = prp()
    p = installer(tmp_path)
    fichier = tmp_path / "agenda_fictif.ics"
    fichier.write_bytes(ics())
    assert _cli("preparer", "--fichier", str(fichier), "--mois", "2026-12", "--profil", str(p.chemin)) == 3
    fichier.unlink()
    capsys.readouterr()
    avant = arborescence(p.racine)

    assert _cli("apercu", "--mois", "2026-12", "--profil", str(p.chemin)) == 0
    sortie = capsys.readouterr().out
    assert "À FAIRE" in sortie and "Mois : décembre 2026 — entité ROSARI" in sortie

    relue = P.lire_preparation(p, MOIS)
    f1 = brouillon(relue, P.numeros_courts(relue)["F1"])
    assert _cli("apercu", "--mois", "2026-12", "--facture", "F1", "--profil", str(p.chemin)) == 0
    assert f"Empreinte : {f1.empreinte[:12]}" in capsys.readouterr().out
    eco = brouillon(relue, "ECO-2026-12")
    assert _cli("apercu", "--mois", "2026-12", "--facture", "ECO-2026-12", "--profil", str(p.chemin)) == 0
    assert f"Empreinte : {eco.empreinte[:12]}" in capsys.readouterr().out

    ouverts = []
    monkeypatch.setattr(os, "startfile", lambda chemin, *a, **k: ouverts.append(Path(chemin)), raising=False)
    assert _cli("apercu", "--mois", "2026-12", "--ouvrir", "--profil", str(p.chemin)) == 0
    assert [o.resolve() for o in ouverts] == [_dossier(p).resolve()]
    assert arborescence(p.racine) == avant


def test_t_prp_42_apercu_sans_preparation(tmp_path, capsys):
    """T-PRP-42 — Sans préparation du mois : code 1 et « lancez `facturation preparer --mois …` » ;
    sans `--mois`, le mois par défaut du profil (§ 1.6)."""
    from facturation import commandes

    p = installer(tmp_path)
    assert _cli("apercu", "--mois", "2026-12", "--profil", str(p.chemin)) == 1
    sortie = capsys.readouterr()
    assert "facturation preparer --mois 2026-12" in sortie.out + sortie.err

    mois = commandes.mois_par_defaut(p, date.today())
    assert _cli("apercu", "--profil", str(p.chemin)) == 1
    sortie = capsys.readouterr()
    assert f"facturation preparer --mois {mois}" in sortie.out + sortie.err


# ──────────────────────────────────────────────────────────────────────
# T-PRP-38 à T-PRP-40, T-PRP-43, T-PRP-44 : relecture, refus, journal, aperçus
# ──────────────────────────────────────────────────────────────────────


def _resume(preparation) -> list:
    return sorted(
        (b.identifiant, b.empreinte, b.statut, b.total_ht, b.numero_probable, tuple(b.blocages), tuple(b.avertissements))
        for b in preparation.brouillons
    )


def test_t_prp_38_lire_preparation(tmp_path):
    """T-PRP-38 — `lire_preparation` : None sans préparation ; sinon la préparation relue à l'identique
    (brouillons, empreintes recalculables, types des lignes, file à classer, numéros courts)."""
    P = prp()
    p = installer(tmp_path)
    assert P.lire_preparation(p, MOIS) is None
    prep = _lancer(p)
    relue = P.lire_preparation(p, MOIS)

    assert (relue.format, relue.entite, relue.mois) == (1, "ROSARI", MOIS)
    assert (relue.fenetre, relue.empreinte_agenda, relue.prepare_le) == (prep.fenetre, prep.empreinte_agenda, MAINTENANT)
    assert _resume(relue) == _resume(prep)
    for b in relue.brouillons:
        assert P.empreinte_brouillon(b) == b.empreinte
        for ligne in b.lignes:
            assert isinstance(ligne.date, date) and isinstance(ligne.debut, datetime)
            assert isinstance(ligne.montant, Decimal)
    assert {a.seance.uid for a in relue.a_classer} == {"inconnu-1"}
    assert P.numeros_courts(relue) == P.numeros_courts(prep)


def test_t_prp_38_preparation_abimee_ou_de_format_inconnu(tmp_path):
    """T-PRP-38 — `preparation.json` de format inconnu, ou tronqué : `PreparationIllisible` (une
    `ValueError`) qui nomme le fichier."""
    P = prp()
    p = installer(tmp_path)
    _lancer(p)
    chemin = _dossier(p) / "preparation.json"
    donnees = json.loads(chemin.read_text(encoding="utf-8"))
    donnees["format"] = 99
    chemin.write_text(json.dumps(donnees, ensure_ascii=False), encoding="utf-8")
    with pytest.raises(P.PreparationIllisible):
        P.lire_preparation(p, MOIS)

    chemin.write_text('{"format": 1, "entite": "ROSARI", "brouill', encoding="utf-8")
    with pytest.raises(P.PreparationIllisible) as exc:
        P.lire_preparation(p, MOIS)
    assert isinstance(exc.value, ValueError)
    assert "preparation.json" in str(exc.value)


def test_t_prp_39_racine_absente(tmp_path):
    """T-PRP-39 — Racine absente : `PreparationImpossible` (« dossier de données introuvable : … ; si
    c'est un lecteur Google Drive, vérifiez qu'il est lancé »), rien n'est créé."""
    P = prp()
    absente = tmp_path / "lecteur_absent"
    p = profil(tmp_path, absente)
    with pytest.raises(P.PreparationImpossible, match="dossier de données introuvable") as exc:
        P.lancer_preparation(p, MOIS, contenu_ics=ics(), maintenant=MAINTENANT)
    assert "Google Drive" in str(exc.value)
    assert not absente.exists()


def test_t_prp_39_racine_dans_un_depot_refusee(tmp_path):
    """T-PRP-39 — Racine placée dans un dépôt git : refus avant toute écriture (aucune donnée réelle
    dans un dépôt, § 1.1)."""
    P = prp()
    depot = tmp_path / "depot"
    (depot / ".git").mkdir(parents=True)
    racine = depot / "donnees"
    master_standard(racine)
    p = profil(tmp_path, racine)
    avant = arborescence(racine)
    with pytest.raises((ValueError, P.PreparationImpossible)) as exc:
        P.lancer_preparation(p, MOIS, contenu_ics=ics(), maintenant=MAINTENANT)
    assert "dépôt git" in f"{exc.value} {exc.value.__cause__ or ''}"
    assert arborescence(racine) == avant


def test_t_prp_40_verrou_d_un_autre_poste(tmp_path):
    """T-PRP-40 — Verrou d'entité tenu par un autre poste : `PreparationImpossible` levée depuis
    `MasterVerrouille` ; préparation précédente et verrou intacts."""
    from facturation.master import MasterVerrouille

    P = prp()
    p = installer(tmp_path)
    _lancer(p)
    precedente = (_dossier(p) / "preparation.json").read_bytes()
    verrou = p.master.with_suffix(".verrou")
    contenu = "poste = AUTRE-POSTE-FICTIF\ndepuis = 2026-12-20 15:59:00\npid = 4242\n"
    verrou.write_text(contenu, encoding="utf-8")

    with pytest.raises(P.PreparationImpossible) as exc:
        _lancer(p)
    assert isinstance(exc.value.__cause__, MasterVerrouille)
    assert (_dossier(p) / "preparation.json").read_bytes() == precedente
    assert verrou.read_text(encoding="utf-8") == contenu


def test_t_prp_43_evenement_preparation_ok(tmp_path):
    """T-PRP-43 — Événement `preparation` / `ok` au journal du mois : mois, nombre de brouillons, de
    brouillons bloqués et de créneaux à classer, `erreur` nulle."""
    p = installer(tmp_path, regles=[regle("ABC", **{"Bon de commande requis": "oui"}), regle("ECO"),
                                    regle("FUT", "Plus tard")])
    _lancer(p)
    evenements = journal.lire_evenements(chemins.journal_mois(p, MOIS), types=["preparation"])
    assert len(evenements) == 1
    ev = evenements[0]
    assert (ev["statut"], ev["mois"], ev["entite"], ev["erreur"]) == ("ok", MOIS, "ROSARI", None)
    assert (_compte(ev["brouillons"]), _compte(ev["bloques"]), _compte(ev["a_classer"])) == (2, 1, 1)


def test_t_prp_44_apercus_ecrits(tmp_path):
    """T-PRP-44 — Aperçus écrits : `apercu.txt` (global, « À FAIRE »), un `apercus/<identifiant>.txt`
    par brouillon (empreinte), `a_classer.txt` (créneaux de la file)."""
    p = installer(tmp_path)
    prep = _lancer(p)
    dossier = _dossier(p)

    global_ = (dossier / "apercu.txt").read_text(encoding="utf-8")
    assert global_.startswith("Mois : décembre 2026 — entité ROSARI") and "À FAIRE" in global_
    for b in prep.brouillons:
        texte = (dossier / "apercus" / f"{b.identifiant}.txt").read_text(encoding="utf-8")
        assert f"Empreinte : {b.empreinte[:12]}" in texte
    assert "Réunion pédagogique" in (dossier / "a_classer.txt").read_text(encoding="utf-8")
