"""Registre des factures dans le MASTER et commandes `facturation registre` (F-25 ; banque de tests :
T-RGF-01, T-RGF-03 à T-RGF-09, T-RGF-23 à T-RGF-37).

Tests d'acceptation écrits avant l'implémentation : chaque test importe `facturation.registre` (ou
`facturation.cli`) lui-même, pour que l'absence du module fasse échouer les tests sans bloquer la
collecte des autres. Données fictives uniquement : MASTER et profils créés dans `tmp_path`
(`tests.conftest.creer_master`), écoles ABC et ECO, montants inventés.

Les commandes sont lancées dans le processus (`cli.main`) : l'entrée standard est simulée
(`io.StringIO`, une chaîne vide valant fin d'entrée), le profil est désigné par la variable
`FACTURATION_PROFIL` (jamais le profil réel), et la racine de `pcf_format` est rétablie après
chaque test.

Contrat (contrats des vagues 3 à 5, § 7, avec les § 15 et 16 qui priment ; § 1.4 à 1.6) :
- `lire_registre` : onglet absent → [] ; dates `date`, `datetime`, `JJ/MM/AAAA`, `AAAA-MM-JJ` ;
  montants en nombre ou en texte « 1 234,56 » ; toute autre valeur non vide lève `RegistreInvalide`
  (« onglet « Factures », ligne 7, colonne « Date de paiement » : « fin nov » n'est pas une date
  (JJ/MM/AAAA) ») ; un statut hors des statuts saisis aussi, sauf « échue », toléré et traité comme
  un statut vide ; un N° en double (sans tenir compte de la casse) aussi ;
- `lignes_historiques` : une ligne par N° de `Sessions` absent du registre (école, période, Σ Total HT,
  statut « émise », note « repris de Sessions le JJ/MM/AAAA : dates à compléter ») ; date d'émission
  et échéance vides, jamais inventées ;
- fonctions de premier niveau (`mettre_a_jour_registre`, `noter_manquant`, `initialiser_registre`) :
  `ValueError("validation explicite requise")` sans confirmation ; écriture sous `ecriture_master`
  (refus d'un MASTER ouvert, verrou, sauvegarde), `load_workbook(data_only=False)`, enregistrement
  atomique et relecture : seules les cellules annoncées changent ; un événement `registre` au journal
  des écritures ; elles rendent la sauvegarde ;
- `section_apercu` : « FACTURES ÉCHUES NON PAYÉES : N », avec le N° et le retard en jours ; None
  sinon ;
- `facturation registre [--ecole X] [--mois AAAA-MM] [--echues] [--trous] [--profil]` : tableau,
  échues en tête, code 3 s'il y a une facture échue ou une anomalie, 0 sinon ; onglet absent :
  « ajoutez l'onglet Factures avec `facturation registre initialiser` », code 1 ;
- `registre initialiser [--reprendre-historique] [--oui]`, `registre envoyee N° --le DATE
  [--remplacer] [--oui]`, `registre payee N° --le DATE …`, `registre note N° "…"`,
  `registre manquant N° --raison "…"` : chacune affiche l'avant et l'après, puis demande « oui »
  (ou `--oui`).

Lectures retenues (`💡 [proposé — à valider]`) :
- T-RGF-03 « une seule cellule modifiée » : la saisie d'un paiement change la date de paiement et,
  par la règle de `modifier_ligne`, le statut qui passe à « payée » ; aucune autre cellule d'aucun
  onglet ne change (la formule du pied de `Devis`, § 15, est gardée). Une note seule change
  exactement une cellule (T-RGF-27) ;
- les dates lues sont des `date` (pas des `datetime`), les montants des `Decimal` égaux au nombre tel
  qu'il s'écrit ; une cellule vide donne None (dates, montant, contrat) ou "" (statut, note) ;
- `lignes_historiques` rend des dictionnaires {colonne de `COLONNES` : valeur}, à passer à
  `ajouter_ligne` ; l'école y est le code (libellé de `Sessions` traduit par `Écoles`, § 15) ;
- `initialiser_registre(reprendre_historique=True)` sur un onglet `Factures` qui existe déjà ajoute
  les lignes manquantes sans toucher aux autres : c'est ce que propose l'anomalie « N° absent du
  registre » ; sans l'onglet, il le crée (en fin de classeur ou ailleurs, l'ordre des autres onglets
  ne changeant pas) ;
- `noter_manquant` déduit l'école du N° et du préfixe du profil (« R_ABC002 » → ABC) ;
- `section_apercu` : une ligne par facture échue, qui porte le N° et le nombre de jours de retard ;
- la commande : `--profil` est accepté après l'action, comme pour `profil verifier` et `agenda lire` ;
  le tableau écrit les dates en JJ/MM/AAAA ; une réponse autre que « oui » (« non », fin d'entrée, ou
  un mot qui n'est pas proposé, `demander` reposant alors la question jusqu'à la fin de l'entrée)
  n'écrit rien, et le code de retour n'est pas imposé dans ce cas ; un refus (N° absent, date
  antérieure à l'émission, date déjà saisie sans `--remplacer`, MASTER ouvert) rend le code 1 avec
  un message qui nomme la cause ;
- `--mois` filtre sur un mois ; les tests n'y mettent que des factures dont la période et la date
  d'émission tombent dans le même mois, pour ne pas trancher entre les deux lectures.
"""

from __future__ import annotations

import importlib
import io
import json
import re
import sys
from datetime import date, datetime, timedelta
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from facturation import config
from facturation import pcf_format as fmt
from facturation.master import ONGLETS_PREVUS
from tests.conftest import ENTETES, creer_master

pytestmark = pytest.mark.integration

COLONNES_FACTURES = ONGLETS_PREVUS["Factures"]
ECOLES = [["École Fictive A", "ABC"], ["École Fictive B", "ECO"]]
AUJOURD_HUI = date.today()


def _reg():
    """Module testé, importé dans le test : son absence fait échouer le test, pas la collecte."""
    return importlib.import_module("facturation.registre")


def _master():
    return importlib.import_module("facturation.master")


@pytest.fixture(autouse=True)
def _racine_pcf_retablie(monkeypatch):
    """Les commandes fixent la racine de `pcf_format` : elle est rétablie après chaque test."""
    monkeypatch.delenv("PCF_DATA_ROOT", raising=False)
    ancienne = fmt._DATA_ROOT
    yield
    fmt._DATA_ROOT = ancienne
    fmt._MASTER_DATA_CACHE.clear()


# ──────────────────────────────────────────────────────────────────────
# Données fictives
# ──────────────────────────────────────────────────────────────────────


def _jour(valeur):
    return valeur.date() if isinstance(valeur, datetime) else valeur


def _jj(jour: date) -> str:
    return jour.strftime("%d/%m/%Y")


def _facture(n_facture: str, **valeurs) -> dict:
    """Ligne de `Factures` {en-tête : valeur} ; l'école par défaut est celle du N° (« R_ABC001 » → ABC)."""
    noms = {
        "ecole": "École", "contrat": "Contrat", "periode": "Période", "emission": "Date d'émission",
        "echeance": "Échéance", "total": "Total HT", "envoi": "Date d'envoi", "paiement": "Date de paiement",
        "statut": "Statut", "note": "Note",
    }
    ligne = {"N° facture": n_facture, "École": n_facture[2:5], "Période": "Novembre 2026"}
    for cle, valeur in valeurs.items():
        ligne[noms[cle]] = valeur
    return ligne


def _session(n_facture: str, total, *, ecole: str = "ABC", periode: str = "Novembre 2026",
             jour: date = date(2026, 11, 3)) -> list:
    """Ligne de `Sessions` (ordre de MASTER_SCHEMA) : une séance de 3,5 h facturée sous `n_facture`."""
    return ["2026-12-01", ecole, periode, n_facture, jour, "09:00", "12:30", 3.5, "M1", "Module fictif", 45, total,
            "#ABC-M1 Cours fictif"]


def _installer(tmp_path: Path, *, factures=None, sessions=None, ecoles=ECOLES, formule_devis: bool = True):
    """MASTER fictif (avec, par défaut, une formule au pied de `Devis`, comme le MASTER réel, § 15) et
    profil ROSARI (préfixe R_) qui le désigne."""
    racine = tmp_path / "donnees"
    chemin = creer_master(racine, factures=factures, sessions=sessions, ecoles=ecoles)
    if formule_devis:
        wb = load_workbook(chemin)
        devis = wb["Devis"]
        devis.append(["ABC", "Module fictif", "D-0001", date(2026, 10, 1), 500, "Accepté", "", ""])
        devis.append(["", "", "", "", "=SUM(E2:E2)", "", "", ""])
        wb.save(chemin)
    chemin_profil = tmp_path / "profil.toml"
    chemin_profil.write_text(config.rediger("Personne A", "ROSARI", str(racine)), encoding="utf-8")
    return config.charger(chemin_profil)


def _lire(profil):
    """Registre relu depuis le fichier, {N° : LigneRegistre}."""
    lignes = _reg().lire_registre(_master().lire_master(profil.master))
    return {ligne.n_facture: ligne for ligne in lignes}


def _instantane(chemin: Path) -> dict[str, dict[str, object]]:
    """Toutes les cellules non vides de tous les onglets, formules comprises (`data_only=False`)."""
    wb = load_workbook(chemin, data_only=False)
    try:
        return {
            ws.title: {c.coordinate: c.value for rangee in ws.iter_rows() for c in rangee if c.value is not None}
            for ws in wb.worksheets
        }
    finally:
        wb.close()


def _differences(avant: dict, apres: dict) -> set[tuple[str, str]]:
    changees = set()
    for onglet in avant.keys() | apres.keys():
        a, b = avant.get(onglet, {}), apres.get(onglet, {})
        changees |= {(onglet, c) for c in a.keys() | b.keys() if a.get(c) != b.get(c)}
    return changees


def _cellule(colonne: str, rang: int) -> tuple[str, str]:
    return "Factures", f"{get_column_letter(COLONNES_FACTURES.index(colonne) + 1)}{rang}"


def _sauvegardes(profil) -> list[Path]:
    return sorted((profil.master.parent / "sauvegardes").glob("*.xlsx"))


def _evenements(profil) -> list[dict]:
    from facturation import chemins, journal

    return journal.lire_evenements(chemins.journal_ecritures(profil), types={"registre"})


def _master_inchange(profil, octets: bytes):
    """MASTER identique, aucun événement `registre`, aucun verrou laissé (une sauvegarde faite à
    l'entrée d'`ecriture_master` avant un refus reste permise)."""
    assert profil.master.read_bytes() == octets
    assert _evenements(profil) == []
    assert not profil.master.with_suffix(".verrou").exists()


def _rien_n_a_ete_ecrit(profil, octets: bytes):
    """Ni écriture, ni sauvegarde, ni journal, ni verrou."""
    _master_inchange(profil, octets)
    assert _sauvegardes(profil) == []


@pytest.fixture
def commande(monkeypatch, capsys):
    """Lance `facturation …` dans le processus ; rend (code, sortie standard, sortie d'erreur)."""

    def lancer(profil, *args: str, entree: str = ""):
        from facturation import cli

        if profil is not None:
            monkeypatch.setenv("FACTURATION_PROFIL", str(profil.chemin))
        monkeypatch.setattr(sys, "stdin", io.StringIO(entree))
        capsys.readouterr()
        try:
            code = cli.main(list(args))
        except SystemExit as exc:
            code = exc.code if isinstance(exc.code, int) else (0 if exc.code is None else 1)
        sortie = capsys.readouterr()
        return code, sortie.out, sortie.err

    return lancer


# ──────────────────────────────────────────────────────────────────────
# Lecture du registre
# ──────────────────────────────────────────────────────────────────────


def test_t_rgf_23_lire_registre_formes_acceptees(tmp_path):
    """T-RGF-23 — `lire_registre` lit les dates (`date`, `datetime`, JJ/MM/AAAA, AAAA-MM-JJ) en `date`,
    les montants (nombre ou texte « 1 234,56 ») en `Decimal`, donne à chaque ligne son rang Excel
    (lignes vides comprises) ; « échue », toléré, se lit comme un statut vide ; un onglet absent ou
    vide donne []."""
    reg = _reg()
    profil = _installer(tmp_path, factures=[
        _facture("R_ABC001", contrat="CT0001", emission=date(2026, 12, 1), echeance=datetime(2027, 1, 31),
                 total=1234.56, envoi="02/12/2026", paiement="2026-12-20", statut="payée", note="Payée par virement"),
        None,
        _facture("R_ECO001", emission="01/12/2026", echeance="2027-01-15", total="1 234,56", statut="échue"),
        _facture("R_ABC002", total=90, statut=None),
    ])
    lignes = reg.lire_registre(_master().lire_master(profil.master))

    assert [ligne.rang for ligne in lignes] == [2, 4, 5]
    premiere, seconde, troisieme = lignes
    assert isinstance(premiere, reg.LigneRegistre)
    assert (premiere.n_facture, premiere.ecole, premiere.contrat, premiere.periode) == (
        "R_ABC001", "ABC", "CT0001", "Novembre 2026")
    for valeur, attendue in ((premiere.date_emission, date(2026, 12, 1)), (premiere.echeance, date(2027, 1, 31)),
                             (premiere.date_envoi, date(2026, 12, 2)), (premiere.date_paiement, date(2026, 12, 20)),
                             (seconde.date_emission, date(2026, 12, 1)), (seconde.echeance, date(2027, 1, 15))):
        assert type(valeur) is date and valeur == attendue, valeur
    assert isinstance(premiere.total_ht, Decimal) and premiere.total_ht == Decimal("1234.56")
    assert (premiere.statut_saisi, premiere.note) == ("payée", "Payée par virement")
    assert seconde.total_ht == Decimal("1234.56")
    assert (seconde.statut_saisi, seconde.contrat, seconde.note) == ("", None, "")
    assert troisieme.total_ht == Decimal("90") and troisieme.statut_saisi == ""
    assert (troisieme.date_emission, troisieme.echeance, troisieme.date_envoi, troisieme.date_paiement) == (
        None, None, None, None)

    assert reg.lire_registre(_master().lire_master(creer_master(tmp_path / "sans_onglet"))) == []
    assert reg.lire_registre(_master().lire_master(creer_master(tmp_path / "vide", factures=[]))) == []


CAS_ILLISIBLES = [
    # (identifiant, colonne, valeur fautive, mots attendus dans le message)
    ("date-en-clair", "Date de paiement", "fin nov", ("Date de paiement", "fin nov", "JJ/MM/AAAA")),
    ("date-impossible", "Échéance", "31/02/2027", ("Échéance", "31/02/2027")),
    ("nombre-dans-une-date", "Date d'envoi", 45000, ("Date d'envoi", "45000")),
    ("montant-en-clair", "Total HT", "beaucoup", ("Total HT", "beaucoup")),
    ("statut-inconnu", "Statut", "réglée", ("Statut", "réglée")),
]


@pytest.mark.parametrize(("colonne", "valeur", "mots"), [c[1:] for c in CAS_ILLISIBLES], ids=[c[0] for c in CAS_ILLISIBLES])
def test_t_rgf_06_cellule_illisible_nommee(tmp_path, colonne, valeur, mots):
    """T-RGF-06 — une date illisible (et de même un montant ou un statut) lève `RegistreInvalide`, dont le
    message nomme l'onglet, la ligne Excel et la colonne, et cite la valeur."""
    reg = _reg()
    fautive = _facture("R_ABC005", emission=date(2026, 12, 1), echeance=date(2027, 1, 31), total=90, statut="émise")
    fautive[colonne] = valeur
    profil = _installer(tmp_path, factures=[
        _facture("R_ABC001", emission=date(2026, 12, 1), total=90, statut="émise"),
        _facture("R_ABC002", emission=date(2026, 12, 1), total=90, statut="émise"),
        None,
        _facture("R_ABC003", emission=date(2026, 12, 1), total=90, statut="émise"),
        _facture("R_ABC004", emission=date(2026, 12, 1), total=90, statut="émise"),
        fautive,
    ])
    with pytest.raises(reg.RegistreInvalide) as exc:
        reg.lire_registre(_master().lire_master(profil.master))
    message = str(exc.value)
    assert "Factures" in message and "ligne 7" in message, message
    for mot in mots:
        assert mot in message, message


def test_t_rgf_04b_lire_registre_refuse_un_numero_en_double(tmp_path):
    """T-RGF-04b — deux lignes au même N°, à la casse près, lèvent `RegistreInvalide` qui nomme le N°."""
    reg = _reg()
    profil = _installer(tmp_path, factures=[
        _facture("R_ABC001", emission=date(2026, 12, 1), total=90, statut="émise"),
        _facture("R_ABC002", emission=date(2026, 12, 1), total=90, statut="émise"),
        _facture("r_abc001", emission=date(2026, 12, 1), total=90, statut="émise"),
    ])
    with pytest.raises(reg.RegistreInvalide) as exc:
        reg.lire_registre(_master().lire_master(profil.master))
    assert "R_ABC001" in str(exc.value).upper(), str(exc.value)


def test_t_rgf_01_ligne_d_emission_lisible_de_bout_en_bout(tmp_path):
    """T-RGF-01 — `ligne_pour_emission` remplit les 11 colonnes (statut « en émission », envoi et paiement
    vides) ; écrite par `ajouter_ligne`, complétée par `modifier_ligne` puis enregistrée, la facture
    porte une échéance, une date d'envoi et une date de paiement que `lire_registre` relit."""
    reg = _reg()
    profil = _installer(tmp_path, factures=[])
    ligne = reg.ligne_pour_emission(
        n_facture="R_ABC001", ecole="ABC", contrat="CT0001", periode="Novembre 2026",
        date_emission=date(2026, 12, 1), echeance=date(2027, 1, 31), total_ht=Decimal("1800.00"),
    )
    assert set(ligne) == set(COLONNES_FACTURES) and len(ligne) == 11
    assert (ligne["N° facture"], ligne["École"], ligne["Contrat"], ligne["Période"]) == (
        "R_ABC001", "ABC", "CT0001", "Novembre 2026")
    assert _jour(ligne["Date d'émission"]) == date(2026, 12, 1) and _jour(ligne["Échéance"]) == date(2027, 1, 31)
    assert Decimal(str(ligne["Total HT"])) == Decimal("1800")
    assert ligne["Statut"] == "en émission"
    assert ligne["Date d'envoi"] in (None, "") and ligne["Date de paiement"] in (None, "")

    wb = load_workbook(profil.master)
    ws = wb["Factures"]
    assert reg.ajouter_ligne(ws, ligne) == 2
    reg.modifier_ligne(ws, "R_ABC001", statut="émise")
    reg.modifier_ligne(ws, "R_ABC001", date_envoi=date(2026, 12, 2))
    reg.modifier_ligne(ws, "R_ABC001", date_paiement=date(2027, 1, 10))
    wb.save(profil.master)

    relue = _lire(profil)["R_ABC001"]
    assert (relue.date_emission, relue.echeance, relue.date_envoi, relue.date_paiement) == (
        date(2026, 12, 1), date(2027, 1, 31), date(2026, 12, 2), date(2027, 1, 10))
    assert relue.total_ht == Decimal("1800") and relue.statut_saisi == "payée" and relue.contrat == "CT0001"
    assert reg.statut_affiche(relue, aujourd_hui=date(2027, 2, 15)) == "payée"


# ──────────────────────────────────────────────────────────────────────
# Écritures de premier niveau
# ──────────────────────────────────────────────────────────────────────


def _registre_courant():
    factures = [
        _facture("R_ABC001", emission=date(2026, 12, 1), echeance=date(2027, 1, 31), total=157.5,
                 envoi=date(2026, 12, 2), statut="envoyée"),
        _facture("R_ABC002", emission=date(2026, 12, 1), echeance=date(2027, 1, 31), total=90, statut="émise"),
    ]
    sessions = [_session("R_ABC001", 157.5), _session("R_ABC002", 90)]
    return factures, sessions


def test_t_rgf_03_saisie_d_un_paiement(tmp_path):
    """T-RGF-03 — saisie d'un paiement : la sauvegarde rendue est la copie intacte d'avant l'écriture ;
    comparaison cellule à cellule de tous les onglets : seules la date de paiement et le statut
    (« payée ») de la facture changent, la formule de `Devis` est gardée ; aucun fichier temporaire
    ne reste ; un événement `registre` est journalisé, avec le N°, l'après et la sauvegarde."""
    reg = _reg()
    factures, sessions = _registre_courant()
    profil = _installer(tmp_path, factures=factures, sessions=sessions)
    octets = profil.master.read_bytes()
    avant = _instantane(profil.master)

    sauvegarde = reg.mettre_a_jour_registre(profil, "R_ABC001", date_paiement=date(2026, 12, 20), confirme=True)

    apres = _instantane(profil.master)
    paiement, statut = _cellule("Date de paiement", 2), _cellule("Statut", 2)
    assert _differences(avant, apres) == {paiement, statut}
    assert _jour(apres["Factures"][paiement[1]]) == date(2026, 12, 20)
    assert apres["Factures"][statut[1]] == "payée"
    assert apres["Devis"]["E3"] == "=SUM(E2:E2)"
    sauvegarde = Path(sauvegarde)
    assert sauvegarde in _sauvegardes(profil) and sauvegarde.read_bytes() == octets
    assert not list(profil.master.parent.glob("*.tmp*"))
    assert not profil.master.with_suffix(".verrou").exists()
    (evenement,) = _evenements(profil)
    assert evenement["n_facture"] == "R_ABC001" and evenement["entite"] == "ROSARI"
    assert Path(evenement["sauvegarde"]).name == sauvegarde.name
    texte_apres = json.dumps(evenement["apres"], ensure_ascii=False)
    assert any(forme in texte_apres for forme in ("2026-12-20", "20/12/2026")), texte_apres
    assert "payée" in texte_apres, texte_apres


def test_t_rgf_27_une_note_ne_change_qu_une_cellule(tmp_path):
    """T-RGF-27 — `mettre_a_jour_registre(note=…)` change exactement une cellule du MASTER, la note ;
    l'événement `registre` est journalisé."""
    reg = _reg()
    factures, sessions = _registre_courant()
    profil = _installer(tmp_path, factures=factures, sessions=sessions)
    avant = _instantane(profil.master)

    reg.mettre_a_jour_registre(profil, "R_ABC002", note="Relance envoyée le 05/01/2027", confirme=True)

    apres = _instantane(profil.master)
    note = _cellule("Note", 3)
    assert _differences(avant, apres) == {note}
    assert apres["Factures"][note[1]] == "Relance envoyée le 05/01/2027"
    assert [e["n_facture"] for e in _evenements(profil)] == ["R_ABC002"]


APPELS = {
    "mettre_a_jour_registre": lambda reg, profil, confirme: reg.mettre_a_jour_registre(
        profil, "R_ABC002", date_envoi=date(2026, 12, 5), confirme=confirme),
    "noter_manquant": lambda reg, profil, confirme: reg.noter_manquant(
        profil, "R_ABC003", raison="Numéro sauté à la saisie", confirme=confirme),
    "initialiser_registre": lambda reg, profil, confirme: reg.initialiser_registre(
        profil, reprendre_historique=True, confirme=confirme),
}


def _profil_pour(tmp_path: Path, appel: str):
    factures, sessions = _registre_courant()
    if appel == "initialiser_registre":
        factures = None
    return _installer(tmp_path, factures=factures, sessions=sessions)


@pytest.mark.parametrize("appel", list(APPELS))
def test_t_rgf_24_validation_explicite_requise(tmp_path, appel):
    """T-RGF-24 — sans confirmation, chaque fonction de premier niveau lève
    `ValueError("validation explicite requise")` : ni écriture, ni sauvegarde, ni journal, ni verrou."""
    reg = _reg()
    profil = _profil_pour(tmp_path, appel)
    octets = profil.master.read_bytes()
    with pytest.raises(ValueError, match="validation explicite requise"):
        APPELS[appel](reg, profil, False)
    _rien_n_a_ete_ecrit(profil, octets)


@pytest.mark.parametrize("appel", list(APPELS))
def test_t_rgf_07_master_ouvert_rien_n_est_ecrit(tmp_path, appel):
    """T-RGF-07 — MASTER ouvert dans Excel (fichier `~$` fictif à côté) : `MasterVerrouille` qui dit de
    fermer le classeur ; ni écriture, ni sauvegarde, ni journal, ni verrou laissé."""
    reg = _reg()
    profil = _profil_pour(tmp_path, appel)
    (profil.master.parent / f"~${profil.master.name}").write_bytes(b"")
    octets = profil.master.read_bytes()
    with pytest.raises(_master().MasterVerrouille) as exc:
        APPELS[appel](reg, profil, True)
    assert "fermez" in str(exc.value).casefold(), str(exc.value)
    _rien_n_a_ete_ecrit(profil, octets)


def test_t_rgf_05_numero_manquant_note(tmp_path):
    """T-RGF-05 — un numéro manquant noté (ligne « abandonnée », raison en note, école déduite du N°) :
    `trous_sans_raison` ne le liste plus, et `prochain_numero` saute un numéro noté au-delà du plus
    grand ; chaque note est journalisée ; un N° déjà au registre est refusé sans rien écrire."""
    reg = _reg()
    num = importlib.import_module("facturation.numerotation")
    profil = _installer(tmp_path, factures=[
        _facture("R_ABC001", emission=date(2026, 10, 1), echeance=date(2026, 11, 30), total=90, statut="émise"),
        _facture("R_ABC003", emission=date(2026, 12, 1), echeance=date(2027, 1, 31), total=90, statut="émise"),
    ], sessions=[_session("R_ABC001", 90), _session("R_ABC003", 90)])

    def trous():
        lignes = list(_lire(profil).values())
        return reg.trous_sans_raison(lignes, num.lire_numeros(profil.master), prefixe="R_", ecoles=["ABC", "ECO"])

    def prochain():
        return num.prochain_numero(num.lire_numeros(profil.master), prefixe="R_", code_ecole="ABC")

    assert trous() == ["R_ABC002"] and prochain() == "R_ABC004"

    reg.noter_manquant(profil, "R_ABC002", raison="Numéro sauté à la saisie", confirme=True)
    reg.noter_manquant(profil, "R_ABC004", raison="Facture perdue avant envoi", confirme=True)

    lignes = _lire(profil)
    notee = lignes["R_ABC002"]
    assert (notee.ecole, notee.statut_saisi) == ("ABC", "abandonnée")
    assert "Numéro sauté à la saisie" in notee.note
    assert reg.statut_affiche(notee, aujourd_hui=AUJOURD_HUI) == "abandonnée"
    assert trous() == []
    assert prochain() == "R_ABC005"
    assert [e["n_facture"] for e in _evenements(profil)] == ["R_ABC002", "R_ABC004"]

    octets = profil.master.read_bytes()
    with pytest.raises(reg.RegistreInvalide):
        reg.noter_manquant(profil, "R_ABC001", raison="Doublon", confirme=True)
    assert profil.master.read_bytes() == octets
    assert len(_evenements(profil)) == 2


def _historique():
    sessions = [
        _session("R_ABC001", 100.1, ecole="École Fictive A", periode="Octobre 2026", jour=date(2026, 10, 6)),
        _session("R_ABC001", 200.2, ecole="ABC", periode="Octobre 2026", jour=date(2026, 10, 13)),
        _session("R_ECO004", 157.5, ecole="ECO", periode="Novembre 2026"),
        _session("R_ABC002", 90),
        _session("", 45, jour=date(2026, 11, 24)),
    ]
    factures = [_facture("R_ABC002", emission=date(2026, 12, 1), echeance=date(2027, 1, 31), total=90, statut="émise")]
    return factures, sessions


def test_t_rgf_09a_lignes_historiques(tmp_path):
    """T-RGF-09a — `lignes_historiques` propose une ligne par N° de `Sessions` absent du registre : école
    (en code), période, Σ Total HT exacte, statut « émise », note « repris de Sessions le JJ/MM/AAAA :
    dates à compléter » ; date d'émission, échéance, envoi et paiement vides (jamais inventés) ; ni le
    N° déjà au registre ni la ligne sans N°."""
    reg = _reg()
    factures, sessions = _historique()
    profil = _installer(tmp_path, factures=factures, sessions=sessions)

    proposees = {ligne["N° facture"]: ligne for ligne in reg.lignes_historiques(_master().lire_master(profil.master))}

    assert set(proposees) == {"R_ABC001", "R_ECO004"}
    abc = proposees["R_ABC001"]
    assert set(abc) <= set(COLONNES_FACTURES)
    assert (abc["École"], abc["Période"], abc["Statut"]) == ("ABC", "Octobre 2026", "émise")
    assert Decimal(str(abc["Total HT"])) == Decimal("300.3")
    for vide in ("Date d'émission", "Échéance", "Date d'envoi", "Date de paiement"):
        assert abc.get(vide) in (None, ""), vide
    note = re.fullmatch(r"repris de Sessions le (\d{2}/\d{2}/\d{4})\s:\sdates à compléter", abc["Note"])
    assert note, abc["Note"]
    assert datetime.strptime(note.group(1), "%d/%m/%Y").date() in (AUJOURD_HUI, date.today())
    assert proposees["R_ECO004"]["École"] == "ECO"
    assert Decimal(str(proposees["R_ECO004"]["Total HT"])) == Decimal("157.5")


def test_t_rgf_09b_reprise_de_l_historique(tmp_path):
    """T-RGF-09b — `initialiser_registre(reprendre_historique=True)` sur un registre existant ajoute les
    lignes reprises (dates vides, total = Σ Sessions) sans toucher aux autres ni à `Sessions` ; la
    facture reprise s'affiche « émise », jamais échue ; l'anomalie « absent du registre » disparaît."""
    reg = _reg()
    factures, sessions = _historique()
    profil = _installer(tmp_path, factures=factures, sessions=sessions)
    avant = _instantane(profil.master)
    assert any("reprendre-historique" in a for a in reg.anomalies(
        list(_lire(profil).values()), _sessions_lues(profil)))

    sauvegarde = reg.initialiser_registre(profil, reprendre_historique=True, confirme=True)

    assert Path(sauvegarde) in _sauvegardes(profil)
    apres = _instantane(profil.master)
    assert apres["Sessions"] == avant["Sessions"]
    assert {c: v for c, v in apres["Factures"].items() if re.fullmatch(r"[A-Z]+[12]", c)} == avant["Factures"]
    lignes = _lire(profil)
    assert set(lignes) == {"R_ABC001", "R_ABC002", "R_ECO004"}
    reprise = lignes["R_ABC001"]
    assert (reprise.date_emission, reprise.echeance, reprise.date_envoi, reprise.date_paiement) == (
        None, None, None, None)
    assert reprise.total_ht == Decimal("300.3") and reprise.statut_saisi == "émise" and reprise.ecole == "ABC"
    assert reg.statut_affiche(reprise, aujourd_hui=date(2030, 1, 1)) == "émise"
    assert lignes["R_ECO004"].total_ht == Decimal("157.5")
    assert len(_evenements(profil)) >= 1
    restantes = reg.anomalies(list(lignes.values()), _sessions_lues(profil))
    assert not any("reprendre-historique" in a for a in restantes), restantes


def _sessions_lues(profil):
    from facturation import sessions

    lignes, erreurs = sessions.lire_sessions(_master().lire_master(profil.master))
    assert erreurs == []
    return lignes


def test_t_rgf_25_initialiser_cree_l_onglet(tmp_path):
    """T-RGF-25 — sans onglet `Factures`, `initialiser_registre(reprendre_historique=False)` le crée avec
    ses 11 en-têtes et aucune ligne, sans changer les autres onglets ni leur ordre ; la sauvegarde est
    rendue et l'écriture journalisée."""
    reg = _reg()
    profil = _installer(tmp_path, sessions=[_session("R_ABC001", 90)])
    avant = _instantane(profil.master)

    sauvegarde = reg.initialiser_registre(profil, reprendre_historique=False, confirme=True)

    assert Path(sauvegarde) in _sauvegardes(profil)
    wb = load_workbook(profil.master)
    try:
        assert [nom for nom in wb.sheetnames if nom != "Factures"] == list(avant)
        assert [c.value for c in wb["Factures"][1]] == list(COLONNES_FACTURES)
        assert wb["Factures"].max_row == 1
    finally:
        wb.close()
    apres = _instantane(profil.master)
    assert {onglet: apres[onglet] for onglet in avant} == avant
    assert _lire(profil) == {}
    assert len(_evenements(profil)) >= 1


def test_t_rgf_26_section_de_l_apercu(tmp_path):
    """T-RGF-26 — `section_apercu` : « FACTURES ÉCHUES NON PAYÉES : N », une ligne par facture échue avec
    son N° et son retard en jours ; ni la facture payée ni celle au jour de son échéance ; None quand
    rien n'est échu, ou sans onglet `Factures`."""
    reg = _reg()
    profil = _installer(tmp_path, factures=[
        _facture("R_ABC001", emission=date(2026, 10, 1), echeance=date(2026, 11, 15), total=90,
                 envoi=date(2026, 10, 2), statut="envoyée"),
        _facture("R_ECO002", emission=date(2026, 11, 1), echeance=date(2026, 12, 14), total=90, statut="émise"),
        _facture("R_ABC002", emission=date(2026, 10, 1), echeance=date(2026, 11, 15), total=90,
                 envoi=date(2026, 10, 2), paiement=date(2026, 11, 10), statut="payée"),
        _facture("R_ABC003", emission=date(2026, 12, 1), echeance=date(2026, 12, 15), total=90, statut="émise"),
    ], sessions=[_session(n, 90) for n in ("R_ABC001", "R_ABC002", "R_ABC003")] + [_session("R_ECO002", 90, ecole="ECO")])

    titre, lignes = reg.section_apercu(profil, aujourd_hui=date(2026, 12, 15))

    assert titre == "FACTURES ÉCHUES NON PAYÉES : 2"
    assert any("R_ABC001" in l and re.search(r"(?<![\d/,.])30(?![\d/,.])", l) for l in lignes), lignes
    assert any("R_ECO002" in l and re.search(r"(?<![\d/,.])1(?![\d/,.])", l) for l in lignes), lignes
    texte = "\n".join(lignes)
    assert "R_ABC002" not in texte and "R_ABC003" not in texte, texte
    assert reg.section_apercu(profil, aujourd_hui=date(2026, 11, 1)) is None
    assert reg.section_apercu(_installer(tmp_path / "sans_onglet"), aujourd_hui=date(2026, 12, 15)) is None


# ──────────────────────────────────────────────────────────────────────
# Commande `facturation registre`
# ──────────────────────────────────────────────────────────────────────


def _date(decalage: int) -> date:
    return AUJOURD_HUI + timedelta(days=decalage)


def _registre_du_jour():
    """Registre relatif à aujourd'hui : une payée, une envoyée non échue, une échue, une en émission."""
    factures = [
        _facture("R_ABC001", emission=_date(-90), echeance=_date(-40), total=90, envoi=_date(-89),
                 paiement=_date(-50), statut="payée"),
        _facture("R_ABC002", emission=_date(-20), echeance=_date(25), total=90, envoi=_date(-19), statut="envoyée"),
        _facture("R_ECO001", ecole="ECO", emission=_date(-60), echeance=_date(-10), total=90, statut="émise"),
        _facture("R_ABC003", emission=_date(-1), echeance=_date(44), total=90, statut="en émission"),
    ]
    sessions = [_session(n, 90) for n in ("R_ABC001", "R_ABC002", "R_ABC003")] + [_session("R_ECO001", 90, ecole="ECO")]
    return factures, sessions


def test_t_rgf_08_registre_echues_code_3(tmp_path, commande):
    """T-RGF-08 — `registre --echues` : seules les factures échues, code 3."""
    factures, sessions = _registre_du_jour()
    profil = _installer(tmp_path, factures=factures, sessions=sessions)
    code, sortie, erreur = commande(profil, "registre", "--echues")
    assert code == 3, sortie + erreur
    assert "R_ECO001" in sortie, sortie
    for autre in ("R_ABC001", "R_ABC002", "R_ABC003"):
        assert autre not in sortie, sortie


def test_t_rgf_28_registre_tableau_echues_en_tete(tmp_path, commande):
    """T-RGF-28 — `registre` : tableau de toutes les factures (n°, école, période, émission, échéance,
    total, envoi, paiement, statut affiché, anomalies), échues en tête, dates en JJ/MM/AAAA ; code 3
    avec une facture échue."""
    factures, sessions = _registre_du_jour()
    profil = _installer(tmp_path, factures=factures, sessions=sessions)
    code, sortie, erreur = commande(profil, "registre")
    assert code == 3, sortie + erreur
    positions = {n: sortie.find(n) for n in ("R_ABC001", "R_ABC002", "R_ABC003", "R_ECO001")}
    assert all(p >= 0 for p in positions.values()), sortie
    assert positions["R_ECO001"] < min(positions[n] for n in ("R_ABC001", "R_ABC002", "R_ABC003")), sortie
    minuscule = sortie.casefold()
    for mot in ("échéance", "envoi", "paiement", "statut", "total", "payée", "envoyée", "échue", "en émission"):
        assert mot in minuscule, (mot, sortie)
    assert _jj(_date(-10)) in sortie and _jj(_date(25)) in sortie, sortie


def test_t_rgf_28b_code_selon_echues_et_anomalies(tmp_path, commande):
    """T-RGF-28b — `registre` rend 0 sur un registre sans facture échue ni anomalie (garde cassée : la
    même commande rend 3 dès qu'une anomalie, un paiement sans envoi, apparaît)."""
    factures, sessions = _registre_du_jour()
    factures = [f for f in factures if f["N° facture"] != "R_ECO001"]
    sessions = [s for s in sessions if s[3] != "R_ECO001"]
    profil = _installer(tmp_path / "sain", factures=factures, sessions=sessions)
    code, sortie, erreur = commande(profil, "registre")
    assert code == 0, sortie + erreur
    assert "R_ABC002" in sortie, sortie

    factures[1] = _facture("R_ABC002", emission=_date(-20), echeance=_date(25), total=90, paiement=_date(-2),
                           statut="payée")
    profil = _installer(tmp_path / "anomalie", factures=factures, sessions=sessions)
    code, sortie, erreur = commande(profil, "registre")
    assert code == 3, sortie + erreur
    assert "R_ABC002" in sortie, sortie


def test_t_rgf_29_registre_sans_onglet(tmp_path, commande):
    """T-RGF-29 — sans onglet `Factures`, `registre` dit de l'ajouter avec `facturation registre
    initialiser` et rend 1."""
    profil = _installer(tmp_path)
    code, sortie, erreur = commande(profil, "registre")
    assert code == 1, sortie + erreur
    assert "facturation registre initialiser" in sortie + erreur, sortie + erreur
    assert "Factures" in sortie + erreur, sortie + erreur


def test_t_rgf_30_registre_filtres_ecole_et_mois(tmp_path, commande):
    """T-RGF-30 — `--ecole ECO` ne montre que les factures de ECO ; `--mois` ne montre que celles du mois
    (période et émission dans le même mois) ; un mois invalide est une erreur d'usage (code 2) qui
    cite la valeur."""
    factures = [
        _facture("R_ABC001", periode="Septembre 2026", emission=date(2026, 9, 30), echeance=date(2026, 11, 14),
                 total=90, envoi=date(2026, 10, 1), paiement=date(2026, 10, 20), statut="payée"),
        _facture("R_ECO001", ecole="ECO", periode="Novembre 2026", emission=date(2026, 11, 30),
                 echeance=date(2027, 1, 14), total=90, envoi=date(2026, 12, 1), paiement=date(2026, 12, 10),
                 statut="payée"),
    ]
    sessions = [_session("R_ABC001", 90), _session("R_ECO001", 90, ecole="ECO")]
    profil = _installer(tmp_path, factures=factures, sessions=sessions)

    code, sortie, erreur = commande(profil, "registre", "--ecole", "ECO")
    assert code in (0, 3), sortie + erreur
    assert "R_ECO001" in sortie and "R_ABC001" not in sortie, sortie

    code, sortie, erreur = commande(profil, "registre", "--mois", "2026-11")
    assert code in (0, 3), sortie + erreur
    assert "R_ECO001" in sortie and "R_ABC001" not in sortie, sortie

    code, sortie, erreur = commande(profil, "registre", "--mois", "2026-13")
    assert code == 2 and "2026-13" in erreur, sortie + erreur


def test_t_rgf_31_registre_trous(tmp_path, commande):
    """T-RGF-31 — `registre --trous` liste les numéros manquants sans raison notée."""
    profil = _installer(tmp_path, factures=[
        _facture("R_ABC001", emission=date(2026, 10, 1), echeance=date(2026, 11, 30), total=90,
                 envoi=date(2026, 10, 2), paiement=date(2026, 10, 20), statut="payée"),
        _facture("R_ABC003", emission=date(2026, 11, 1), echeance=date(2026, 12, 31), total=90,
                 envoi=date(2026, 11, 2), paiement=date(2026, 11, 20), statut="payée"),
    ], sessions=[_session("R_ABC001", 90), _session("R_ABC003", 90)])
    code, sortie, erreur = commande(profil, "registre", "--trous")
    assert "R_ABC002" in sortie, sortie + erreur


REPONSES = [
    # (identifiant, entrée standard, écrit ?)
    ("oui", "oui\n", True),
    ("non", "non\n", False),
    ("fin-d-entree", "", False),
    ("debut-de-mot", "o\n", False),
    ("autre-mot", "peut-être\n", False),
]


@pytest.mark.parametrize(("entree", "ecrit"), [r[1:] for r in REPONSES], ids=[r[0] for r in REPONSES])
def test_t_rgf_32_envoyee_demande_oui(tmp_path, commande, entree, ecrit):
    """T-RGF-32 — `registre envoyee N° --le DATE` affiche l'avant et l'après (N°, nouvelle date, statut
    « envoyée »), puis demande « oui » : seul « oui » écrit (date d'envoi, statut, sauvegarde,
    journal, code 0) ; « non », la fin de l'entrée ou un autre mot n'écrivent rien."""
    factures, sessions = _registre_courant()
    profil = _installer(tmp_path, factures=factures, sessions=sessions)
    octets = profil.master.read_bytes()

    code, sortie, erreur = commande(profil, "registre", "envoyee", "R_ABC002", "--le", "05/12/2026", entree=entree)

    assert "R_ABC002" in sortie, sortie + erreur
    assert "05/12/2026" in sortie or "2026-12-05" in sortie, sortie
    assert "envoyée" in sortie and "oui" in sortie, sortie
    if ecrit:
        assert code == 0, sortie + erreur
        ligne = _lire(profil)["R_ABC002"]
        assert ligne.date_envoi == date(2026, 12, 5) and ligne.statut_saisi == "envoyée"
        assert len(_sauvegardes(profil)) == 1
        assert [e["n_facture"] for e in _evenements(profil)] == ["R_ABC002"]
    else:
        _rien_n_a_ete_ecrit(profil, octets)


def test_t_rgf_33_option_oui_sans_question(tmp_path, commande):
    """T-RGF-33 — avec `--oui`, l'écriture se fait sans lire l'entrée standard (fin d'entrée immédiate)."""
    factures, sessions = _registre_courant()
    profil = _installer(tmp_path, factures=factures, sessions=sessions)
    code, sortie, erreur = commande(profil, "registre", "envoyee", "R_ABC002", "--le", "2026-12-05", "--oui", entree="")
    assert code == 0, sortie + erreur
    assert _lire(profil)["R_ABC002"].date_envoi == date(2026, 12, 5)


def test_t_rgf_34_payee_note_manquant(tmp_path, commande):
    """T-RGF-34 — `registre payee`, `registre note` et `registre manquant`, confirmés par `--oui`,
    écrivent la date de paiement (statut « payée »), la note, et la ligne du numéro manquant
    (« abandonnée », école déduite du N°, raison en note) ; trois événements sont journalisés."""
    factures, sessions = _registre_courant()
    profil = _installer(tmp_path, factures=factures, sessions=sessions)

    assert commande(profil, "registre", "payee", "R_ABC001", "--le", "20/12/2026", "--oui")[0] == 0
    assert commande(profil, "registre", "note", "R_ABC002", "Relance envoyée", "--oui")[0] == 0
    assert commande(profil, "registre", "manquant", "R_ABC003", "--raison", "Numéro sauté à la saisie", "--oui")[0] == 0

    lignes = _lire(profil)
    assert lignes["R_ABC001"].date_paiement == date(2026, 12, 20) and lignes["R_ABC001"].statut_saisi == "payée"
    assert lignes["R_ABC002"].note == "Relance envoyée"
    manquant = lignes["R_ABC003"]
    assert (manquant.ecole, manquant.statut_saisi) == ("ABC", "abandonnée")
    assert "Numéro sauté à la saisie" in manquant.note
    assert [e["n_facture"] for e in _evenements(profil)] == ["R_ABC001", "R_ABC002", "R_ABC003"]


CAS_REFUS_COMMANDE = [
    # (identifiant, arguments, mot attendu dans les sorties)
    ("numero-absent", ("envoyee", "R_ABC099", "--le", "05/12/2026", "--oui"), "R_ABC099"),
    ("date-deja-saisie", ("envoyee", "R_ABC001", "--le", "04/12/2026", "--oui"), "remplacer"),
    ("paiement-avant-emission", ("payee", "R_ABC002", "--le", "30/11/2026", "--oui"), "émission"),
]


@pytest.mark.parametrize(("arguments", "mot"), [c[1:] for c in CAS_REFUS_COMMANDE], ids=[c[0] for c in CAS_REFUS_COMMANDE])
def test_t_rgf_35a_refus_de_la_commande(tmp_path, commande, arguments, mot):
    """T-RGF-35a — un N° absent, une date déjà saisie sans `--remplacer`, une date antérieure à l'émission :
    code 1, message qui nomme la cause, rien n'est écrit."""
    factures, sessions = _registre_courant()
    profil = _installer(tmp_path, factures=factures, sessions=sessions)
    octets = profil.master.read_bytes()
    code, sortie, erreur = commande(profil, "registre", *arguments)
    assert code == 1, sortie + erreur
    assert mot.casefold() in (sortie + erreur).casefold(), sortie + erreur
    _master_inchange(profil, octets)


def test_t_rgf_35b_remplacer_master_ouvert_date_invalide(tmp_path, commande):
    """T-RGF-35b — `--remplacer` remplace une date déjà saisie ; un MASTER ouvert (fichier `~$` fictif)
    rend 1, dit de fermer le classeur et n'écrit rien ; une date invalide est une erreur d'usage
    (code 2) qui cite la valeur."""
    factures, sessions = _registre_courant()
    profil = _installer(tmp_path, factures=factures, sessions=sessions)

    code, sortie, erreur = commande(profil, "registre", "envoyee", "R_ABC001", "--le", "04/12/2026", "--remplacer",
                                    "--oui")
    assert code == 0, sortie + erreur
    assert _lire(profil)["R_ABC001"].date_envoi == date(2026, 12, 4)

    code, sortie, erreur = commande(profil, "registre", "envoyee", "R_ABC002", "--le", "demain", "--oui")
    assert code == 2 and "demain" in erreur, sortie + erreur

    (profil.master.parent / f"~${profil.master.name}").write_bytes(b"")
    octets = profil.master.read_bytes()
    sauvegardes = _sauvegardes(profil)
    code, sortie, erreur = commande(profil, "registre", "envoyee", "R_ABC002", "--le", "05/12/2026", "--oui")
    assert code == 1, sortie + erreur
    assert "fermez" in (sortie + erreur).casefold(), sortie + erreur
    assert profil.master.read_bytes() == octets and _sauvegardes(profil) == sauvegardes


@pytest.mark.parametrize(("entree", "ecrit"), [("oui\n", True), ("", False)], ids=["oui", "fin-d-entree"])
def test_t_rgf_36_initialiser_par_la_commande(tmp_path, commande, entree, ecrit):
    """T-RGF-36 — `registre initialiser --reprendre-historique` montre les lignes reprises puis demande
    « oui » : « oui » crée l'onglet et la ligne reprise (dates vides) ; la fin de l'entrée n'écrit
    rien."""
    profil = _installer(tmp_path, sessions=[_session("R_ABC001", 90)])
    octets = profil.master.read_bytes()

    code, sortie, erreur = commande(profil, "registre", "initialiser", "--reprendre-historique", entree=entree)

    assert "R_ABC001" in sortie and "oui" in sortie, sortie + erreur
    if ecrit:
        assert code == 0, sortie + erreur
        ligne = _lire(profil)["R_ABC001"]
        assert ligne.date_emission is None and ligne.echeance is None and ligne.total_ht == Decimal("90")
    else:
        _rien_n_a_ete_ecrit(profil, octets)


def test_t_rgf_37_option_profil(tmp_path, commande, monkeypatch):
    """T-RGF-37 — `--profil` désigne le profil, après l'action comme pour les autres commandes ; sans
    profil lisible, la commande rend 1."""
    factures, sessions = _registre_courant()
    profil = _installer(tmp_path, factures=factures, sessions=sessions)
    monkeypatch.setenv("FACTURATION_PROFIL", str(tmp_path / "aucun_profil.toml"))

    code, sortie, erreur = commande(None, "registre")
    assert code == 1, sortie + erreur

    code, sortie, erreur = commande(None, "registre", "--profil", str(profil.chemin))
    assert code in (0, 3) and "R_ABC001" in sortie, sortie + erreur

    code, sortie, erreur = commande(None, "registre", "note", "R_ABC001", "Relance envoyée", "--profil",
                                    str(profil.chemin), "--oui")
    assert code == 0, sortie + erreur
    assert _lire(profil)["R_ABC001"].note == "Relance envoyée"
