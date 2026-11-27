"""Audit agenda ↔ facturé de bout en bout : `lancer_audit` et la commande `facturation audit` (F-19 ;
banque de tests : T-AUD-10, T-AUD-11, T-AUD-33 à T-AUD-44 et T-AUD-60 ; le reste dans
tests/unit/test_audit.py). T-AUD-60 et le renfort de T-AUD-39 viennent de la revue de F-19.

Tests d'acceptation écrits avant l'implémentation : chaque test importe `facturation.audit` lui-même,
pour que l'absence du module fasse échouer les tests sans bloquer la collecte des autres. Racines de
données, MASTER, profils, agendas (.ics) et préparations entièrement fictifs, créés dans `tmp_path`.
La commande est lancée en sous-processus avec `PYTHONPATH` pointé sur ce dépôt et un `APPDATA`
temporaire : elle ne lit jamais le profil réel, et ne touche jamais au réseau (`--fichier`, ou
adresse iCal absente).

Contrat (contrats des vagues 3 à 5, § 6, § 1.1, § 1.6, § 1.7) :
- `lancer_audit(profil, du, au, *, contenu_ics=None, maintenant=None, ecoles=None)` lit le MASTER et
  l'agenda sur [du, au], lit les préparations des mois concernés (en ignorant leurs erreurs, en le
  notant), écrit `_sorties/audits/audit_<du>_<au>_<AAAAMMJJ-HHMMSS>.txt` et `.csv` de façon atomique,
  ne prend aucun verrou et ne modifie pas le MASTER ; rend `(RapportAudit, Path)` ;
- `refuser_racine_dans_un_depot(profil.racine)` au début de chaque commande qui écrit ;
- `facturation audit [--mois AAAA-MM | --du JJ/MM/AAAA --au JJ/MM/AAAA] [--ecoles ABC,ECO]
  [--fichier AGENDA.ics] [--profil]` ; période par défaut 💡 : du 1er du mois M−3 à hier (Q-I2) ;
  codes 0 sans écart, 3 avec écarts, 1 en cas d'erreur, 2 pour une erreur d'usage (argparse).

Lectures retenues (`💡 [proposé — à valider]`) :
- T-AUD-10 : F-16a n'est pas encore développé ; `_sorties/<AAAA-MM>/preparation.json` est simulé par
  un fichier fictif au format du § 4 (la `Preparation` passée par `fichiers.en_json`) ; un brouillon y
  contient la séance par sa `cle_seance` ;
- « les mois concernés » sont les mois civils que touche [du, au] : chacun voit sa préparation lue
  (T-AUD-44) ;
- un brouillon dont la pièce d'émission est « émise » (pièce finale présente) n'est pas « non émis » ;
- une préparation illisible, ou d'une autre entité, est ignorée et notée : le nom `preparation.json`
  figure dans le rapport texte ou dans un avertissement (`warnings`) ;
- le `Path` rendu par `lancer_audit` est sous `_sorties/audits/` ; `<du>` et `<au>` du nom de fichier
  s'écrivent AAAA-MM-JJ ou AAAAMMJJ ;
- `--du` sans `--au` (ou l'inverse) et `--mois` avec `--du` sont des erreurs d'usage (code 2) qui
  nomment l'option ; une période inversée est refusée (code 1 ou 2) avec un message qui le dit ;
- les erreurs (profil, MASTER, fichier .ics, agenda illisible ou indisponible, racine dans un dépôt)
  rendent le code 1 avec un message qui nomme la cause, sans trace Python ni rapport écrit.
"""

from __future__ import annotations

import csv
import hashlib
import importlib
import io
import json
import os
import re
import subprocess
import sys
import warnings
from datetime import date, datetime, time, timedelta
from pathlib import Path
from types import SimpleNamespace
from zoneinfo import ZoneInfo

import pytest

from facturation import chemins, config, fichiers, journal, master
from facturation import pcf_format as fmt
from facturation.chemins import cle_seance
from facturation.master import ONGLETS_PREVUS
from tests.conftest import NDA_FICTIF, creer_master

pytestmark = pytest.mark.integration

DEPOT = Path(__file__).resolve().parents[2]
PARIS = ZoneInfo("Europe/Paris")
DU, AU = date(2026, 10, 1), date(2026, 10, 31)
MAINTENANT = datetime(2026, 11, 26, 16, 0, tzinfo=PARIS)
BOM = b"\xef\xbb\xbf"
TENUE_NON_FACTUREE = "tenue non facturée"

I_ABC = "Cours fictif ABC M1"
I_ECO = "Cours fictif ECO M2"

ECOLES = [
    {"École": "École Fictive A", "Code facturation": "ABC"},
    {"École": "Institut Fictif B", "Code facturation": "ECO"},
]


def _aud():
    """Module testé, importé dans le test : son absence fait échouer le test, pas la collecte."""
    return importlib.import_module("facturation.audit")


@pytest.fixture(autouse=True)
def _racine_pcf_format_restauree():
    """`lancer_audit` peut appeler `fmt.set_data_root(profil.racine)` (§ 1.1) : la racine et le cache
    de `pcf_format` sont rendus tels quels après chaque test, pour ne pas fuir vers les suivants."""
    ancienne = fmt._DATA_ROOT
    yield
    fmt._DATA_ROOT = ancienne
    fmt._MASTER_DATA_CACHE.clear()


# ──────────────────────────────────────────────────────────────────────
# Données fictives
# ──────────────────────────────────────────────────────────────────────


def _regle(ecole: str, statut: str = "Lot 1") -> dict:
    ligne = dict.fromkeys(ONGLETS_PREVUS["Règles_facturation"])
    ligne.update({"École": ecole, "Statut lot": statut, "Bon de commande requis": "non"})
    return ligne


def _corresp(cle: str, ecole: str, module: str, promotion: str) -> dict:
    ligne = dict.fromkeys(ONGLETS_PREVUS["Correspondances"])
    ligne.update({"Clé": cle, "Type de clé": "exact", "École": ecole, "Module": module, "Promotion": promotion,
                  "Clé tarif": f"{ecole}|{module}|{promotion}"})
    return ligne


def _ses(ecole: str, jour: date, debut: str, fin: str, duree: float, *, n: str = "R_ABC001",
         promotion: str = "M1", module: str = "Module A") -> dict:
    return {
        "Date_log": "2026-11-02", "École": ecole, "Période_facture": "Octobre 2026", "N°_Facture": n,
        "Date_séance": jour, "Début": debut, "Fin": fin, "Durée_h": duree, "Promotion": promotion,
        "Module / UE": module, "Taux": 45, "Total HT": None, "Summary calendrier verbatim": "",
    }


REGLES = [_regle("ABC"), _regle("ECO")]
CORRESPONDANCES = [_corresp(I_ABC, "ABC", "Module A", "M1"), _corresp(I_ECO, "ECO", "Module B", "M2")]


def _dt(jour: date, heure: str) -> datetime:
    return datetime.combine(jour, time.fromisoformat(heure), tzinfo=PARIS)


def _uid(jour: date, debut: str, intitule: str) -> str:
    return f"{jour:%Y%m%d}-{debut.replace(':', '')}-{re.sub(r'[^0-9A-Za-z]+', '-', intitule).lower()}@exemple.invalid"


def _ics(evenements) -> bytes:
    """Agenda .ics fictif ; chaque événement est (jour, début, fin, intitulé), heures locales flottantes
    (rattachées au fuseau du profil, Europe/Paris)."""
    lignes = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Exemple fictif//Audit facturation_auto//FR"]
    for jour, debut, fin, intitule in evenements:
        lignes += [
            "BEGIN:VEVENT", f"UID:{_uid(jour, debut, intitule)}", "DTSTAMP:20260901T080000Z",
            f"DTSTART:{_dt(jour, debut):%Y%m%dT%H%M%S}", f"DTEND:{_dt(jour, fin):%Y%m%dT%H%M%S}",
            f"SUMMARY:{intitule}", "DESCRIPTION:Séance fictive", "END:VEVENT",
        ]
    lignes.append("END:VCALENDAR")
    return ("\r\n".join(lignes) + "\r\n").encode("utf-8")


S_06 = (date(2026, 10, 6), "09:00", "12:30", I_ABC)    # journalisée
S_13 = (date(2026, 10, 13), "09:00", "12:30", I_ABC)
S_20 = (date(2026, 10, 20), "09:00", "12:30", I_ABC)
SESSIONS_OCT = [_ses("ABC", S_06[0], "09:00", "12:30", 3.5)]
ICS_OCT = _ics([S_06, S_13, S_20])


def _installer(tmp_path: Path, *, entite: str = "ROSARI", sessions=(), racine: Path | None = None) -> config.Profil:
    """MASTER fictif sous `<racine>/<entite>/` et profil fictif qui le désigne."""
    racine = racine or tmp_path / "donnees"
    creer_master(racine, entite, regles=REGLES, correspondances=CORRESPONDANCES, ecoles=ECOLES,
                 sessions=list(sessions), factures=[])
    chemin = tmp_path / f"profil_{entite}.toml"
    chemin.write_text(config.rediger("Personne A", entite, str(racine)), encoding="utf-8")
    return config.charger(chemin)


def _instantane(dossier: Path) -> dict[str, str]:
    """Empreinte de chaque fichier sous `dossier`, par chemin relatif."""
    return {
        p.relative_to(dossier).as_posix(): hashlib.sha256(p.read_bytes()).hexdigest()
        for p in dossier.rglob("*") if p.is_file()
    }


def _rapports(profil: config.Profil, extension: str = "txt") -> list[Path]:
    dossier = chemins.dossier_audits(profil)
    return sorted(dossier.glob(f"audit_*.{extension}")) if dossier.is_dir() else []


def _periode_du_nom(fichier: Path) -> tuple[date, date]:
    nom = re.fullmatch(r"audit_(\d{4}-?\d{2}-?\d{2})_(\d{4}-?\d{2}-?\d{2})_(\d{8}-\d{6})(?:-\d+)?", fichier.stem)
    assert nom, fichier.name
    return tuple(datetime.strptime(nom[i].replace("-", ""), "%Y%m%d").date() for i in (1, 2))


def _non_facturees(rapport) -> dict[date, object]:
    return {e.date: e for e in rapport.ecarts if e.nature == TENUE_NON_FACTUREE}


def _ligne_brouillon(jour: date, debut: str, fin: str, intitule: str) -> dict:
    d, f = _dt(jour, debut), _dt(jour, fin)
    uid = _uid(jour, debut, intitule)
    return {
        "cle_seance": cle_seance(SimpleNamespace(uid=uid, debut=d)), "uid": uid, "date": jour.isoformat(),
        "debut": d.isoformat(), "fin": f.isoformat(), "intitule": intitule, "ecole": "ABC", "promotion": "M1",
        "module": "Module A", "libelle": "Module A", "cle_tarif": "ABC|Module A|M1", "classee_par": "exact",
        "duree_agenda_h": 3.5, "heures": 3.5, "taux": 45.0, "montant": "157.5", "ref_contrat": None, "tenue": True,
    }


def _brouillon(identifiant: str, evenements, *, entite: str = "ROSARI", mois: str = "2026-10") -> dict:
    """Brouillon fictif au format du § 4 (`Brouillon` passé par `fichiers.en_json`)."""
    lignes = [_ligne_brouillon(*e) for e in evenements]
    return {
        "identifiant": identifiant, "entite": entite, "ecole": "ABC", "mois": mois,
        "periode": ["2026-10-01", "2026-10-31"], "periode_label": "Oct 2026", "ecole_label": "École Fictive A",
        "contrat": None, "bon_commande": None,
        "destinataire": {"nom": "ÉCOLE FICTIVE A", "adresse_l1": "10 avenue Test", "adresse_l2": "00000 TESTVILLE",
                         "adresse_l3": "France", "email": "compta@ecole-a.example"},
        "emetteur": {"nom": f"FORMATEUR FICTIF {entite} EI", "adresse_l1": "1 rue de l'Exemple,",
                     "adresse_l2": "00000 Villefictive", "email": "formateur@example.org", "tel": "+33600000000"},
        "info": {"SIRET": "00000000000000", "IBAN": "FR76 0000 0000 0000 0000 0000 000", "BIC": "TESTFRPP"},
        "nda": NDA_FICTIF, "date_emission": "2026-11-01", "echeance": "2026-12-01", "masquer_jour": False,
        "regroupement": "séance", "lignes": lignes, "total_heures": 3.5 * len(lignes),
        "total_ht": f"{157.5 * len(lignes):.2f}", "numero_probable": "R_ABC002", "blocages": [],
        "avertissements": [], "empreinte": "0" * 64,
    }


def _ecrire_preparation(profil: config.Profil, mois: str, brouillons, *, entite: str | None = None) -> Path:
    """`_sorties/<mois>/preparation.json` fictif au format du § 4 (`Preparation` passée par `en_json`)."""
    debut, fin = chemins.mois_valide(mois)
    preparation = {
        "format": 1, "entite": entite or profil.entite, "mois": mois, "prepare_le": "2026-10-20T16:00:00+02:00",
        "version_outil": "0.0.0", "fenetre": [debut.isoformat(), fin.isoformat()], "empreinte_agenda": "0" * 64,
        "brouillons": list(brouillons), "a_classer": [], "hors_lot": [], "ecoles_sans_regle": [],
        "deja_facturees": [], "deja_emises": [], "non_facturees_periode_emise": [], "ecoles_sans_seance": [],
        "ecoles_non_preparees": [], "differences": [],
    }
    chemin = chemins.dossier_mois(profil, mois) / "preparation.json"
    chemin.parent.mkdir(parents=True, exist_ok=True)
    chemin.write_text(json.dumps(preparation, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")
    return chemin


# ──────────────────────────────────────────────────────────────────────
# lancer_audit
# ──────────────────────────────────────────────────────────────────────


def test_t_aud_11_lecture_seule_sans_verrou(tmp_path):
    """T-AUD-11 — L'audit laisse le MASTER identique (empreinte) et n'écrit que son rapport .txt et .csv
    sous `_sorties/audits/` ; il ne prend aucun verrou (il tourne pendant qu'un verrou d'écriture est
    tenu par ce programme, et MASTER ouvert dans Excel) et ne fait aucune sauvegarde."""
    profil = _installer(tmp_path, sessions=SESSIONS_OCT)
    (profil.dossier_entite / f"~${profil.master.name}").write_bytes(b"verrou fictif d'Excel")
    empreinte_master = fichiers.empreinte_fichier(profil.master)
    avant = _instantane(profil.racine)

    with master.verrou_ecriture(profil.master, poste=master.poste_courant()):
        rapport, chemin = _aud().lancer_audit(profil, DU, AU, contenu_ics=ICS_OCT, maintenant=MAINTENANT)

    assert sorted(_non_facturees(rapport)) == [S_13[0], S_20[0]]
    assert fichiers.empreinte_fichier(profil.master) == empreinte_master
    apres = _instantane(profil.racine)
    assert {k: v for k, v in apres.items() if k in avant} == avant
    audits = chemins.dossier_audits(profil)
    nouveaux = sorted(profil.racine / k for k in set(apres) - set(avant))
    assert [p.parent for p in nouveaux] == [audits, audits], nouveaux
    assert sorted(p.suffix for p in nouveaux) == [".csv", ".txt"]
    assert not (profil.dossier_entite / "sauvegardes").exists()
    assert chemin.exists() and audits in (chemin, *chemin.parents)
    assert all(_periode_du_nom(p) == (DU, AU) for p in nouveaux)


def test_t_aud_10_seance_dans_un_brouillon_non_emis_de_la_preparation(tmp_path):
    """T-AUD-10 — Séance présente dans un brouillon non émis (`preparation.json` fictif du mois, F-16a
    n'étant pas encore développé) : l'écart « tenue non facturée » porte l'identifiant du brouillon,
    que le rapport texte cite ; la séance absente du brouillon n'est pas annotée."""
    profil = _installer(tmp_path, sessions=SESSIONS_OCT)
    _ecrire_preparation(profil, "2026-10", [_brouillon("ABC-2026-10", [S_13])])
    rapport, _ = _aud().lancer_audit(profil, DU, AU, contenu_ics=ICS_OCT, maintenant=MAINTENANT)
    non_facturees = _non_facturees(rapport)
    assert sorted(non_facturees) == [S_13[0], S_20[0]]
    assert non_facturees[S_13[0]].dans_brouillon == "ABC-2026-10"
    assert non_facturees[S_20[0]].dans_brouillon is None
    (texte,) = [p.read_text(encoding="utf-8-sig") for p in _rapports(profil)]
    assert "ABC-2026-10" in texte


def test_t_aud_44_preparations_de_chaque_mois_de_la_periode(tmp_path):
    """T-AUD-44 — Sur une période de deux mois, l'audit lit la préparation de chacun des mois concernés :
    la séance de septembre porte le brouillon de septembre, celle d'octobre celui d'octobre ; une séance
    absente des deux brouillons n'est pas annotée."""
    s_sept = (date(2026, 9, 15), "09:00", "12:30", I_ABC)
    profil = _installer(tmp_path, sessions=SESSIONS_OCT)
    _ecrire_preparation(profil, "2026-09", [_brouillon("ABC-2026-09", [s_sept], mois="2026-09")])
    _ecrire_preparation(profil, "2026-10", [_brouillon("ABC-2026-10", [S_13])])
    rapport, _ = _aud().lancer_audit(profil, date(2026, 9, 1), AU, contenu_ics=_ics([s_sept, S_06, S_13, S_20]),
                                     maintenant=MAINTENANT)
    non_facturees = _non_facturees(rapport)
    assert sorted(non_facturees) == [s_sept[0], S_13[0], S_20[0]]
    assert {jour: e.dans_brouillon for jour, e in non_facturees.items()} == {
        s_sept[0]: "ABC-2026-09", S_13[0]: "ABC-2026-10", S_20[0]: None,
    }


def test_t_aud_33_brouillon_emis_non_annote(tmp_path):
    """T-AUD-33 — Un brouillon dont la pièce d'émission est « émise » n'est plus un brouillon non émis :
    la séance qu'il contient, absente de `Sessions`, reste « tenue non facturée », sans annotation."""
    profil = _installer(tmp_path, sessions=SESSIONS_OCT)
    brouillon = _brouillon("ABC-2026-10", [S_13])
    _ecrire_preparation(profil, "2026-10", [brouillon])
    emissions = chemins.dossier_emissions(profil)
    emissions.mkdir(parents=True)
    piece = {
        "format": 1, "entite": "ROSARI", "n_facture": "R_ABC002", "identifiant": "ABC-2026-10", "mois": "2026-10",
        "empreinte": "0" * 64, "onglet": "R_ABC002 Oct 2026", "classeur": "ABC/ROSARI_ABC.xlsx",
        "validation": {"horodatage": "2026-10-31T16:00:00+01:00", "poste": "POSTE-FICTIF"}, "brouillon": brouillon,
    }
    finale = {
        "format": 1, "n_facture": "R_ABC002", "pdf": "ABC/R_ABC002 Oct 2026.pdf", "pdf_sha256": "0" * 64,
        "sauvegardes": [], "emise_le": "2026-11-01T10:00:00+01:00",
    }
    (emissions / "R_ABC002.json").write_text(json.dumps(piece, ensure_ascii=False), encoding="utf-8")
    (emissions / "R_ABC002.final.json").write_text(json.dumps(finale, ensure_ascii=False), encoding="utf-8")
    assert [(p.identifiant, p.statut) for p in journal.pieces_emission(profil)] == [("ABC-2026-10", "émise")]

    rapport, _ = _aud().lancer_audit(profil, DU, AU, contenu_ics=ICS_OCT, maintenant=MAINTENANT)
    non_facturees = _non_facturees(rapport)
    assert sorted(non_facturees) == [S_13[0], S_20[0]]
    assert [e.dans_brouillon for e in non_facturees.values()] == [None, None]


@pytest.mark.parametrize("defaut", ["json-illisible", "autre-entite"])
def test_t_aud_34_preparation_illisible_ignoree_et_notee(tmp_path, defaut):
    """T-AUD-34 — Une préparation illisible, ou d'une autre entité, ne bloque pas l'audit : elle est
    ignorée (aucune annotation) et notée (`preparation.json` cité dans le rapport ou un avertissement)."""
    profil = _installer(tmp_path, sessions=SESSIONS_OCT)
    if defaut == "json-illisible":
        chemin = chemins.dossier_mois(profil, "2026-10") / "preparation.json"
        chemin.parent.mkdir(parents=True)
        chemin.write_text("{ pas du json", encoding="utf-8")
    else:
        _ecrire_preparation(profil, "2026-10", [_brouillon("ABC-2026-10", [S_13], entite="KAUR")], entite="KAUR")

    with warnings.catch_warnings(record=True) as avertissements:
        warnings.simplefilter("always")
        rapport, _ = _aud().lancer_audit(profil, DU, AU, contenu_ics=ICS_OCT, maintenant=MAINTENANT)

    non_facturees = _non_facturees(rapport)
    assert sorted(non_facturees) == [S_13[0], S_20[0]]
    assert [e.dans_brouillon for e in non_facturees.values()] == [None, None]
    (texte,) = [p.read_text(encoding="utf-8-sig") for p in _rapports(profil)]
    notes = texte + "\n".join(str(a.message) for a in avertissements)
    assert "preparation.json" in notes


def test_t_aud_35_entite_kaur_ne_lit_et_n_ecrit_que_sous_kaur(tmp_path):
    """T-AUD-35 — Un profil KAUR audite le MASTER de KAUR et n'écrit que sous `KAUR/_sorties/audits/` ;
    la racine ROSARI voisine reste intacte (les entités ne se mélangent jamais)."""
    racine = tmp_path / "donnees"
    creer_master(racine, "ROSARI", regles=REGLES, correspondances=CORRESPONDANCES, ecoles=ECOLES,
                 sessions=[_ses("ABC", s[0], "09:00", "12:30", 3.5) for s in (S_06, S_13, S_20)], factures=[])
    avant_rosari = _instantane(racine / "ROSARI")
    profil = _installer(tmp_path, entite="KAUR", sessions=SESSIONS_OCT, racine=racine)

    rapport, chemin = _aud().lancer_audit(profil, DU, AU, contenu_ics=ICS_OCT, maintenant=MAINTENANT)
    assert rapport.entite == "KAUR"
    assert sorted(_non_facturees(rapport)) == [S_13[0], S_20[0]]
    assert _instantane(racine / "ROSARI") == avant_rosari
    assert chemins.dossier_audits(profil) == racine / "KAUR" / "_sorties" / "audits"
    assert len(_rapports(profil)) == 1 and len(_rapports(profil, "csv")) == 1
    assert (racine / "KAUR") in chemin.parents


def test_t_aud_36_racine_dans_un_depot_refusee(tmp_path):
    """T-AUD-36 — Une racine de données placée dans un dépôt git est refusée (`ValueError` qui le dit)
    avant toute écriture : aucun rapport n'est écrit."""
    depot = tmp_path / "depot"
    (depot / ".git").mkdir(parents=True)
    profil = _installer(tmp_path, sessions=SESSIONS_OCT, racine=depot / "donnees")
    avant = _instantane(profil.racine)
    with pytest.raises(ValueError, match="dépôt"):
        _aud().lancer_audit(profil, DU, AU, contenu_ics=ICS_OCT, maintenant=MAINTENANT)
    assert _instantane(profil.racine) == avant
    assert not chemins.dossier_audits(profil).exists()


# ──────────────────────────────────────────────────────────────────────
# Commande `facturation audit`
# ──────────────────────────────────────────────────────────────────────


def _lancer(tmp_path: Path, *args: str) -> subprocess.CompletedProcess:
    """`python -m facturation audit …` sur le paquet de ce dépôt, isolé du profil réel."""
    env = {
        **os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONPATH": str(DEPOT),
        "APPDATA": str(tmp_path / "appdata"), "XDG_CONFIG_HOME": str(tmp_path / "config"),
        "FACTURATION_PROFIL": str(tmp_path / "aucun_profil.toml"),
    }
    env.pop("PCF_DATA_ROOT", None)
    return subprocess.run(
        [sys.executable, "-m", "facturation", "audit", *args],
        capture_output=True, text=True, encoding="utf-8", env=env, cwd=tmp_path,
    )


def _mois_decale(aujourd_hui: date, decalage: int) -> tuple[date, date]:
    """Premier et dernier jour du mois situé `decalage` mois avant celui d'`aujourd_hui`."""
    annee, mois = aujourd_hui.year, aujourd_hui.month - decalage
    while mois < 1:
        annee, mois = annee - 1, mois + 12
    premier = date(annee, mois, 1)
    suivant = date(annee + 1, 1, 1) if mois == 12 else date(annee, mois + 1, 1)
    return premier, suivant - timedelta(days=1)


def _sha(chemin: Path) -> str:
    return hashlib.sha256(chemin.read_bytes()).hexdigest()


def test_t_aud_37_commande_avec_ecart_code_3(tmp_path):
    """T-AUD-37 — `facturation audit --mois … --fichier … --profil …` avec une séance tenue non facturée :
    code 3, un rapport .txt (qui cite l'écart) et un .csv (avec BOM) pour la période du mois, MASTER
    inchangé, aucune trace Python."""
    premier, dernier = _mois_decale(date.today(), 2)
    j1, j2 = premier + timedelta(days=9), premier + timedelta(days=16)
    profil = _installer(tmp_path, sessions=[_ses("ABC", j1, "09:00", "12:30", 3.5)])
    agenda = tmp_path / "agenda.ics"
    agenda.write_bytes(_ics([(j1, "09:00", "12:30", I_ABC), (j2, "09:00", "12:30", I_ABC)]))
    empreinte = _sha(profil.master)

    res = _lancer(tmp_path, "--mois", f"{premier:%Y-%m}", "--fichier", str(agenda), "--profil", str(profil.chemin))
    sortie = res.stdout + res.stderr
    assert res.returncode == 3, sortie
    assert "Traceback" not in sortie
    (txt,) = _rapports(profil)
    (csv_,) = _rapports(profil, "csv")
    assert _periode_du_nom(txt) == (premier, dernier)
    texte = txt.read_text(encoding="utf-8-sig")
    assert TENUE_NON_FACTUREE in texte and f"{j2:%d/%m/%Y}" in texte
    assert csv_.read_bytes().startswith(BOM)
    assert _sha(profil.master) == empreinte


def test_t_aud_38_commande_sans_ecart_code_0_et_filtre_ecoles(tmp_path):
    """T-AUD-38 — Sans écart, code 0 ; `--ecoles ABC` écarte la séance non facturée d'ECO (code 0), que
    l'audit de toutes les écoles, ou de `ABC,ECO`, signale (code 3). Chaque lancement écrit son rapport."""
    premier, _ = _mois_decale(date.today(), 2)
    j1, j2 = premier + timedelta(days=9), premier + timedelta(days=16)
    profil = _installer(tmp_path, sessions=[_ses("ABC", j, "09:00", "12:30", 3.5) for j in (j1, j2)])
    agenda = tmp_path / "agenda.ics"
    agenda.write_bytes(_ics([(j1, "09:00", "12:30", I_ABC), (j2, "09:00", "12:30", I_ABC),
                             (j1, "14:00", "17:00", I_ECO)]))
    base = ("--mois", f"{premier:%Y-%m}", "--fichier", str(agenda), "--profil", str(profil.chemin))

    toutes = _lancer(tmp_path, *base)
    assert toutes.returncode == 3, toutes.stdout + toutes.stderr
    abc = _lancer(tmp_path, *base, "--ecoles", "ABC")
    assert abc.returncode == 0, abc.stdout + abc.stderr
    deux = _lancer(tmp_path, *base, "--ecoles", "ABC,ECO")
    assert deux.returncode == 3, deux.stdout + deux.stderr
    assert len(_rapports(profil)) == 3 and len(_rapports(profil, "csv")) == 3


def test_t_aud_39_periode_par_defaut_du_premier_de_m_moins_3_a_hier(tmp_path):
    """T-AUD-39 — 💡 Sans `--mois` ni `--du`/`--au`, la période va du 1er du mois M−3 à hier (Q-I2) : les
    séances de ces deux jours sont auditées, celles d'avant ne le sont pas (garde cassée : `--mois`
    du mois d'avant les signale).

    Revue de F-19 : les séances des deux bornes sont cherchées parmi les écarts du .csv (nature et date),
    et non dans le texte, dont l'en-tête « Période : du … au … » porte déjà ces deux dates (l'assertion
    passait à vide : un agenda lu jusqu'à la veille de « au » laissait le test vert)."""
    aujourd_hui = date.today()
    hier = aujourd_hui - timedelta(days=1)
    du, _ = _mois_decale(aujourd_hui, 3)
    veille, plus_tot = du - timedelta(days=1), du - timedelta(days=10)
    profil = _installer(tmp_path)
    agenda = tmp_path / "agenda.ics"
    agenda.write_bytes(_ics([
        (plus_tot, "09:00", "12:00", I_ABC), (veille, "09:00", "12:00", I_ABC), (du, "09:00", "12:00", I_ABC),
        (hier, "09:00", "12:00", I_ABC), (aujourd_hui, "00:00", "00:01", I_ABC),
    ]))
    base = ("--fichier", str(agenda), "--profil", str(profil.chemin))

    res = _lancer(tmp_path, *base)
    assert res.returncode == 3, res.stdout + res.stderr
    (txt,) = _rapports(profil)
    assert _periode_du_nom(txt) == (du, hier)
    texte = txt.read_text(encoding="utf-8-sig")
    (csv_,) = _rapports(profil, "csv")
    lignes = list(csv.reader(io.StringIO(csv_.read_bytes().decode("utf-8-sig"), newline=""), delimiter=";"))
    nature, date_ecart = lignes[0].index("Nature"), lignes[0].index("Date")
    ecarts = sorted((ligne[nature], ligne[date_ecart]) for ligne in lignes[1:])
    assert ecarts == sorted([(TENUE_NON_FACTUREE, f"{du:%d/%m/%Y}"), (TENUE_NON_FACTUREE, f"{hier:%d/%m/%Y}")]), ecarts
    assert f"{veille:%d/%m/%Y}" not in texte and f"{plus_tot:%d/%m/%Y}" not in texte

    garde = _lancer(tmp_path, *base, "--mois", f"{plus_tot:%Y-%m}")
    assert garde.returncode == 3, garde.stdout + garde.stderr
    (txt_garde,) = [p for p in _rapports(profil) if p != txt]
    assert f"{plus_tot:%d/%m/%Y}" in txt_garde.read_text(encoding="utf-8-sig")


@pytest.mark.parametrize("forme", ["%d/%m/%Y", "%Y-%m-%d"], ids=["jj-mm-aaaa", "aaaa-mm-jj"])
def test_t_aud_40_commande_du_au(tmp_path, forme):
    """T-AUD-40 — `--du` et `--au` (JJ/MM/AAAA ou AAAA-MM-JJ) fixent la période du rapport."""
    premier, dernier = _mois_decale(date.today(), 2)
    profil = _installer(tmp_path)
    agenda = tmp_path / "agenda.ics"
    agenda.write_bytes(_ics([(premier + timedelta(days=9), "09:00", "12:30", I_ABC)]))
    res = _lancer(tmp_path, "--du", f"{premier:{forme}}", "--au", f"{dernier:{forme}}", "--fichier", str(agenda),
                  "--profil", str(profil.chemin))
    assert res.returncode == 3, res.stdout + res.stderr
    (txt,) = _rapports(profil)
    assert _periode_du_nom(txt) == (premier, dernier)


@pytest.mark.parametrize(
    ("cas", "attendu"),
    [
        ("profil-absent", "Profil"),
        ("master-absent", "MASTER"),
        ("fichier-absent", "absent.ics"),
        ("agenda-illisible", "agenda"),
        ("agenda-indisponible", "agenda"),
        ("racine-dans-un-depot", "dépôt"),
    ],
)
def test_t_aud_41_commande_erreurs_code_1(tmp_path, cas, attendu):
    """T-AUD-41 — Profil absent, MASTER absent, fichier .ics absent ou illisible, agenda indisponible
    (aucune adresse iCal), racine dans un dépôt : code 1, message qui nomme la cause (jamais une adresse),
    sans trace Python ni rapport écrit."""
    premier, _ = _mois_decale(date.today(), 2)
    racine = tmp_path / "donnees"
    if cas == "racine-dans-un-depot":
        (tmp_path / "depot" / ".git").mkdir(parents=True)
        racine = tmp_path / "depot" / "donnees"
    if cas == "master-absent":
        (racine / "ROSARI").mkdir(parents=True)
        chemin_profil = tmp_path / "profil_ROSARI.toml"
        chemin_profil.write_text(config.rediger("Personne A", "ROSARI", str(racine)), encoding="utf-8")
        profil = config.charger(chemin_profil)
    else:
        profil = _installer(tmp_path, sessions=[], racine=racine)
    agenda = tmp_path / "agenda.ics"
    agenda.write_bytes(_ics([(premier + timedelta(days=9), "09:00", "12:30", I_ABC)]))

    args = ["--mois", f"{premier:%Y-%m}", "--fichier", str(agenda), "--profil", str(profil.chemin)]
    if cas == "profil-absent":
        args[-1] = str(tmp_path / "absent.toml")
    elif cas == "fichier-absent":
        args[3] = str(tmp_path / "absent.ics")
    elif cas == "agenda-illisible":
        agenda.write_text("ceci n'est pas un agenda", encoding="utf-8")
    elif cas == "agenda-indisponible":
        del args[2:4]

    res = _lancer(tmp_path, *args)
    sortie = res.stdout + res.stderr
    assert res.returncode == 1, sortie
    assert attendu.casefold() in sortie.casefold(), sortie
    assert "Traceback" not in sortie and "https://" not in sortie
    assert _rapports(profil) == [] and _rapports(profil, "csv") == []


@pytest.mark.parametrize(
    ("options", "nommees"),
    [
        (("--mois", "{mois}", "--du", "{premier}", "--au", "{dernier}"), ("--mois", "--du")),
        (("--du", "{premier}"), ("--au",)),
        (("--au", "{dernier}"), ("--du",)),
        (("--mois", "2026-13"), ("2026-13",)),
        (("--du", "31/02/2026", "--au", "{dernier}"), ("31/02/2026",)),
    ],
    ids=["mois-et-du", "du-sans-au", "au-sans-du", "mois-invalide", "date-impossible"],
)
def test_t_aud_42_commande_erreurs_d_usage_code_2(tmp_path, options, nommees):
    """T-AUD-42 — `--mois` avec `--du`, `--du` sans `--au` (ou l'inverse), mois ou date invalides : erreur
    d'usage, code 2, message qui nomme l'option ou la valeur, sans rapport écrit."""
    premier, dernier = _mois_decale(date.today(), 2)
    profil = _installer(tmp_path)
    agenda = tmp_path / "agenda.ics"
    agenda.write_bytes(_ics([(premier + timedelta(days=9), "09:00", "12:30", I_ABC)]))
    valeurs = {"mois": f"{premier:%Y-%m}", "premier": f"{premier:%d/%m/%Y}", "dernier": f"{dernier:%d/%m/%Y}"}
    args = [o.format(**valeurs) for o in options]
    res = _lancer(tmp_path, *args, "--fichier", str(agenda), "--profil", str(profil.chemin))
    assert res.returncode == 2, res.stdout + res.stderr
    assert "invalid choice" not in res.stderr, res.stderr
    assert any(n in res.stderr for n in nommees), res.stderr
    assert _rapports(profil) == []


def test_t_aud_43_commande_periode_inversee_refusee(tmp_path):
    """T-AUD-43 — `--du` postérieur à `--au` : refus (code 1 ou 2) qui dit que la période est inversée, ou
    cite les deux dates, sans trace Python ni rapport écrit."""
    premier, dernier = _mois_decale(date.today(), 2)
    profil = _installer(tmp_path)
    agenda = tmp_path / "agenda.ics"
    agenda.write_bytes(_ics([(premier + timedelta(days=9), "09:00", "12:30", I_ABC)]))
    res = _lancer(tmp_path, "--du", f"{dernier:%d/%m/%Y}", "--au", f"{premier:%d/%m/%Y}", "--fichier", str(agenda),
                  "--profil", str(profil.chemin))
    sortie = res.stdout + res.stderr
    assert res.returncode in (1, 2), sortie
    assert "invalid choice" not in sortie
    assert "invers" in sortie.lower() or (f"{dernier:%d/%m/%Y}" in sortie and f"{premier:%d/%m/%Y}" in sortie), sortie
    assert "Traceback" not in sortie
    assert _rapports(profil) == []


def test_t_aud_60_commande_ecole_inconnue_code_1(tmp_path):
    """T-AUD-60 — Revue de F-19 (garde du refus d'une école inconnue) : `--ecoles ABX` (faute de frappe)
    rend le code 1 avec un message qui nomme « ABX », jamais le code 0 d'un audit vide « sans écart » ;
    aucun rapport n'est écrit. Garde cassée : `--ecoles ABC` audite (code 3, une séance non facturée)."""
    premier, _ = _mois_decale(date.today(), 2)
    profil = _installer(tmp_path)
    agenda = tmp_path / "agenda.ics"
    agenda.write_bytes(_ics([(premier + timedelta(days=9), "09:00", "12:30", I_ABC)]))
    base = ("--mois", f"{premier:%Y-%m}", "--fichier", str(agenda), "--profil", str(profil.chemin))

    res = _lancer(tmp_path, *base, "--ecoles", "ABX")
    sortie = res.stdout + res.stderr
    assert res.returncode == 1, sortie
    assert "ABX" in sortie and "Traceback" not in sortie
    assert _rapports(profil) == []

    abc = _lancer(tmp_path, *base, "--ecoles", "ABC")
    assert abc.returncode == 3, abc.stdout + abc.stderr
