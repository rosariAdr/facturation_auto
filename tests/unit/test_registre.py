"""Registre des factures : fonctions pures et feuille ouverte (F-25, `facturation/registre.py` ;
banque de tests : T-RGF-02, T-RGF-04a, T-RGF-10 à T-RGF-22 ; constats de revue : T-RGF-01b,
T-RGF-04c, T-RGF-13d, T-RGF-14b, T-RGF-14c, T-RGF-17b, T-RGF-19c à T-RGF-19e, T-RGF-22b).

Tests d'acceptation écrits avant l'implémentation : chaque test importe `facturation.registre`
lui-même, pour que l'absence du module fasse échouer les tests sans bloquer la collecte des autres.
Données fictives uniquement (écoles ABC et ECO, entité ROSARI, montants inventés). Les feuilles sont
des classeurs openpyxl en mémoire : aucune lecture ni écriture de fichier.

Contrat (contrats des vagues 3 à 5, § 7, avec les § 15 et 16 qui priment) :
- `COLONNES = master.ONGLETS_PREVUS["Factures"]` ; `STATUTS_SAISIS` = en émission, émise, envoyée,
  payée, abandonnée ; `STATUTS_AFFICHES` = `STATUTS_SAISIS` + échue, à compléter (calculés, jamais
  écrits) ; `RegistreInvalide(ValueError)` ; `LigneRegistre` (dataclass gelée) ;
- `statut_affiche`, dans cet ordre : abandonnée ; date de paiement → payée ; statut « en émission »
  (jamais échue) ; échéance antérieure à aujourd'hui → échue (le jour de l'échéance n'est pas échu) ;
  date d'envoi → envoyée ; date d'émission ou statut « émise » → émise ; sinon à compléter ;
- `echues` trie par échéance, puis par N° ;
- `anomalies` : N° de `Sessions` absent du registre (avec la proposition
  `registre initialiser --reprendre-historique`) ; envoi, paiement ou échéance antérieurs à
  l'émission ; paiement sans envoi ; `Total HT` différent de Σ `Sessions.Total HT` du N° (en
  `Decimal`, au centime) ; statut saisi incohérent avec les dates ;
- `trous_sans_raison` : `numeros_manquants` école par école, moins les N° qui ont une ligne au
  registre (ADR029) ;
- `ligne_pour_emission` remplit les 11 colonnes (statut « en émission », envoi et paiement vides) ;
- `creer_onglet`, `ajouter_ligne` (par nom de colonne, après `derniere_ligne_remplie`, N° déjà
  présent refusé, dates au format DD/MM/YYYY, rend la ligne), `modifier_ligne` (champs admis :
  date_envoi, date_paiement, statut réservé à F-17 — émise, abandonnée —, note ; refus d'un N°
  absent ou ambigu, d'une ligne abandonnée sauf la note, d'une date antérieure à l'émission, d'une
  cellule déjà remplie d'une autre valeur sans `remplacer=True` ; envoi : émise → envoyée ;
  paiement → payée ; « échue » jamais écrit ; rend {colonne : (avant, après)}),
  `noter_numero_manquant` (ligne au statut abandonnée, raison en note) : aucune ne prend de verrou ;
- § 1.4 : une cellule cible qui contient une formule n'est jamais écrasée.

Lectures retenues (`💡 [proposé — à valider]`) :
- `retard_jours` rend None quand la facture n'est pas échue (payée, abandonnée, en émission, sans
  échéance, ou le jour même de l'échéance), et le nombre de jours depuis l'échéance sinon ;
- un refus de `modifier_ligne` ou d'`ajouter_ligne` est une `ValueError` (`RegistreInvalide`
  convient) dont le message nomme la cause (le N°, le champ, « remplacer », « formule ») ; la feuille
  n'est pas modifiée ;
- le N° se compare sans tenir compte de la casse dans `ajouter_ligne` comme dans `lire_registre`
  (un N° en double « sans tenir compte de la casse » est refusé) ;
- une même date déjà saisie, redonnée sans `remplacer`, n'est pas un refus ;
- dans `anomalies`, chaque anomalie est une chaîne qui nomme le N° ; le contrôle du total compare
  au centime, sans tolérance : les cas testés n'ont que des montants de ligne à deux décimales au
  plus, pour ne pas dépendre de la règle du § 1.2 (« aucun contrôle ne compare une somme de lignes
  arrondies au total arrondi ») ;
- `trous_sans_raison` ne liste pas deux fois un trou quand une école figure à la fois dans
  `Règles_facturation` et dans `Écoles`.
"""

from __future__ import annotations

import dataclasses
import importlib
import re
from datetime import date, datetime, time
from decimal import Decimal

import pytest
from openpyxl import Workbook

from facturation.master import ONGLETS_PREVUS

pytestmark = pytest.mark.unit

COLONNES_FACTURES = ONGLETS_PREVUS["Factures"]
AUJOURD_HUI = date(2026, 12, 15)


def _reg():
    """Module testé, importé dans le test : son absence fait échouer le test, pas la collecte."""
    return importlib.import_module("facturation.registre")


def _jour(valeur):
    """Une date, que la cellule rende une `date` ou un `datetime` (openpyxl)."""
    return valeur.date() if isinstance(valeur, datetime) else valeur


def _ligne(**champs):
    """`LigneRegistre` fictive : facture émise le 01/12/2026, échéance au 31/01/2027."""
    valeurs = dict(
        rang=2, n_facture="R_ABC001", ecole="ABC", contrat=None, periode="Novembre 2026",
        date_emission=date(2026, 12, 1), echeance=date(2027, 1, 31), total_ht=Decimal("1800.00"),
        date_envoi=None, date_paiement=None, statut_saisi="émise", note="",
    )
    valeurs.update(champs)
    return _reg().LigneRegistre(**valeurs)


def _seance(n_facture: str, total, *, ecole: str = "ABC", rang: int = 2, abandonnee: bool = False):
    """`LigneSessions` fictive d'une séance de 3,5 h facturée sous `n_facture`."""
    from facturation.sessions import LigneSessions

    return LigneSessions(
        rang=rang, ecole=ecole, n_facture=n_facture, periode="Novembre 2026", date_seance=date(2026, 11, 3),
        debut=time(9, 0), fin=time(12, 30), duree_h=3.5, promotion="M1", module="Module fictif",
        taux=45.0, total_ht=None if total is None else Decimal(str(total)), summary="#ABC-M1 Cours fictif",
        abandonnee=abandonnee,
    )


def _facture(n_facture: str, **valeurs) -> dict:
    """Ligne de l'onglet `Factures` sous forme {en-tête : valeur} (en-têtes absents = cellule vide)."""
    noms = {
        "ecole": "École", "contrat": "Contrat", "periode": "Période", "emission": "Date d'émission",
        "echeance": "Échéance", "total": "Total HT", "envoi": "Date d'envoi", "paiement": "Date de paiement",
        "statut": "Statut", "note": "Note",
    }
    ligne = {"N° facture": n_facture, "École": n_facture[2:5], "Période": "Novembre 2026"}
    for cle, valeur in valeurs.items():
        ligne[noms[cle]] = valeur
    return ligne


def _feuille(lignes=(), colonnes=COLONNES_FACTURES):
    """Classeur en mémoire avec un onglet `Factures` : en-têtes en ligne 1, puis les lignes (None = vide)."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Factures"
    entetes = list(colonnes)
    ws.append(entetes)
    for ligne in lignes:
        ws.append([] if ligne is None else [ligne.get(entete) for entete in entetes])
    return wb, ws


def _index(ws) -> dict[str, int]:
    return {cellule.value: cellule.column for cellule in ws[1] if cellule.value is not None}


def _valeur(ws, rang: int, colonne: str):
    return ws.cell(row=rang, column=_index(ws)[colonne]).value


def _instantane(ws) -> dict[str, object]:
    return {
        cellule.coordinate: cellule.value
        for rangee in ws.iter_rows() for cellule in rangee if cellule.value is not None
    }


# ──────────────────────────────────────────────────────────────────────
# Constantes et dataclass
# ──────────────────────────────────────────────────────────────────────


def test_t_rgf_10_constantes_et_ligne_du_registre():
    """T-RGF-10 — `COLONNES` reprend les 11 colonnes de `Factures` (MASTER_SCHEMA) ; les statuts saisis
    et affichés sont ceux du contrat, « échue » et « à compléter » n'étant jamais saisis ;
    `RegistreInvalide` est une `ValueError` ; `LigneRegistre` est une dataclass gelée."""
    reg = _reg()
    assert tuple(reg.COLONNES) == COLONNES_FACTURES and len(reg.COLONNES) == 11
    assert tuple(reg.STATUTS_SAISIS) == ("en émission", "émise", "envoyée", "payée", "abandonnée")
    assert tuple(reg.STATUTS_AFFICHES) == tuple(reg.STATUTS_SAISIS) + ("échue", "à compléter")
    assert issubclass(reg.RegistreInvalide, ValueError)
    assert dataclasses.is_dataclass(reg.LigneRegistre)
    assert [f.name for f in dataclasses.fields(reg.LigneRegistre)] == [
        "rang", "n_facture", "ecole", "contrat", "periode", "date_emission", "echeance", "total_ht",
        "date_envoi", "date_paiement", "statut_saisi", "note",
    ]
    ligne = _ligne()
    with pytest.raises(dataclasses.FrozenInstanceError):
        ligne.statut_saisi = "payée"


# ──────────────────────────────────────────────────────────────────────
# Statut affiché, retard, échues
# ──────────────────────────────────────────────────────────────────────

CAS_STATUT = [
    # (identifiant, champs de la ligne, statut affiché le 15/12/2026, retard en jours)
    ("abandonnee-prime-tout", dict(statut_saisi="abandonnée", date_envoi=date(2026, 12, 2),
                                   date_paiement=date(2026, 12, 10), echeance=date(2026, 12, 5)), "abandonnée", None),
    ("payee-prime-sur-echue", dict(statut_saisi="envoyée", date_envoi=date(2026, 12, 2),
                                   date_paiement=date(2026, 12, 10), echeance=date(2026, 12, 5)), "payée", None),
    ("payee-par-la-date-seule", dict(statut_saisi="", date_paiement=date(2026, 12, 10)), "payée", None),
    ("en-emission-jamais-echue", dict(statut_saisi="en émission", echeance=date(2026, 12, 1)), "en émission", None),
    ("echue-le-lendemain", dict(statut_saisi="envoyée", date_envoi=date(2026, 12, 2),
                                echeance=date(2026, 12, 14)), "échue", 1),
    ("pas-echue-le-jour-meme", dict(statut_saisi="envoyée", date_envoi=date(2026, 12, 2),
                                    echeance=date(2026, 12, 15)), "envoyée", None),
    ("echue-sans-envoi", dict(statut_saisi="émise", echeance=date(2026, 12, 5)), "échue", 10),
    ("envoyee", dict(statut_saisi="envoyée", date_envoi=date(2026, 12, 2)), "envoyée", None),
    ("emise-par-la-date", dict(statut_saisi=""), "émise", None),
    ("emise-par-le-statut", dict(statut_saisi="émise", date_emission=None, echeance=None), "émise", None),
    ("sans-echeance-jamais-echue", dict(statut_saisi="émise", date_emission=date(2020, 1, 1), echeance=None),
     "émise", None),
    ("a-completer", dict(statut_saisi="", date_emission=None, echeance=None, total_ht=None), "à compléter", None),
    ("payee-avant-echeance", dict(statut_saisi="payée", date_envoi=date(2026, 12, 2),
                                  date_paiement=date(2026, 12, 5), echeance=date(2026, 12, 10)), "payée", None),
]


@pytest.mark.parametrize(("champs", "attendu", "retard"), [c[1:] for c in CAS_STATUT], ids=[c[0] for c in CAS_STATUT])
def test_t_rgf_02_ordre_des_statuts(champs, attendu, retard):
    """T-RGF-02 — ordre des statuts affichés. Non échue le jour de l'échéance, échue le lendemain
    (`retard_jours` = 1) ; payée avant l'échéance : non signalée ; envoyée et échue : « échue »
    (l'échéance passe avant l'envoi) ; « en émission » : jamais échue."""
    reg = _reg()
    ligne = _ligne(**champs)
    assert reg.statut_affiche(ligne, aujourd_hui=AUJOURD_HUI) == attendu
    assert attendu in reg.STATUTS_AFFICHES
    assert reg.retard_jours(ligne, aujourd_hui=AUJOURD_HUI) == retard


def test_t_rgf_11_retard_compte_les_jours_depuis_l_echeance():
    """T-RGF-11 — `retard_jours` compte les jours depuis l'échéance, changement d'année compris ; il rend
    None pour une facture qui n'est pas échue (💡)."""
    reg = _reg()
    echue = _ligne(statut_saisi="envoyée", date_envoi=date(2026, 12, 2), echeance=date(2026, 12, 31))
    assert reg.retard_jours(echue, aujourd_hui=date(2027, 1, 30)) == 30
    assert reg.retard_jours(echue, aujourd_hui=date(2026, 12, 31)) is None
    payee = dataclasses.replace(echue, date_paiement=date(2027, 1, 20), statut_saisi="payée")
    assert reg.retard_jours(payee, aujourd_hui=date(2027, 1, 30)) is None


def test_t_rgf_12_echues_triees_par_echeance_puis_par_numero():
    """T-RGF-12 — `echues` ne garde que les factures échues, triées par échéance puis par N° ; une facture
    payée, abandonnée, en émission ou au jour de son échéance n'y figure pas."""
    reg = _reg()
    lignes = [
        _ligne(n_facture="R_ECO002", ecole="ECO", statut_saisi="émise", echeance=date(2026, 12, 5)),
        _ligne(n_facture="R_ABC002", statut_saisi="payée", date_envoi=date(2026, 12, 2),
               date_paiement=date(2026, 12, 4), echeance=date(2026, 12, 5)),
        _ligne(n_facture="R_ABC004", statut_saisi="envoyée", date_envoi=date(2026, 12, 2), echeance=date(2026, 12, 5)),
        _ligne(n_facture="R_ABC003", statut_saisi="en émission", echeance=date(2026, 12, 1)),
        _ligne(n_facture="R_ABC005", statut_saisi="émise", echeance=date(2026, 12, 15)),
        _ligne(n_facture="R_ABC006", statut_saisi="abandonnée", echeance=date(2026, 12, 1)),
        _ligne(n_facture="R_ABC001", statut_saisi="émise", echeance=date(2026, 12, 1)),
    ]
    assert [ligne.n_facture for ligne in reg.echues(lignes, aujourd_hui=AUJOURD_HUI)] == [
        "R_ABC001", "R_ABC004", "R_ECO002",
    ]


# ──────────────────────────────────────────────────────────────────────
# Anomalies
# ──────────────────────────────────────────────────────────────────────


def _registre_sain():
    """Registre cohérent avec `Sessions` : aucune anomalie."""
    lignes = [
        _ligne(n_facture="R_ABC001", statut_saisi="payée", date_emission=date(2026, 11, 1),
               echeance=date(2026, 12, 31), total_ht=Decimal("157.50"), date_envoi=date(2026, 11, 2),
               date_paiement=date(2026, 11, 20)),
        _ligne(n_facture="R_ABC002", statut_saisi="envoyée", total_ht=Decimal("315.00"), date_envoi=date(2026, 12, 2)),
        _ligne(n_facture="R_ABC003", statut_saisi="émise", total_ht=Decimal("90.00")),
        _ligne(n_facture="R_ABC004", statut_saisi="abandonnée", date_emission=None, echeance=None, total_ht=None,
               note="Numéro sauté à la saisie"),
        _ligne(n_facture="R_ECO001", ecole="ECO", statut_saisi="en émission", echeance=date(2027, 1, 15),
               total_ht=Decimal("200")),
    ]
    sessions = [
        _seance("R_ABC001", 157.5, rang=2),
        _seance("R_ABC002", 157.5, rang=3),
        _seance("R_ABC002", 157.5, rang=4),
        _seance("R_ABC003", 90, rang=5),
        _seance("R_ECO001", 200, ecole="ECO", rang=6),
    ]
    return lignes, sessions


def _remplacer(lignes, n_facture, **champs):
    return [dataclasses.replace(l, **champs) if l.n_facture == n_facture else l for l in lignes]


def test_t_rgf_13a_registre_sain_sans_anomalie():
    """T-RGF-13a — un registre cohérent avec `Sessions` (tous les statuts saisis représentés) ne donne
    aucune anomalie (garde cassée : chaque cas de T-RGF-13b en ajoute une)."""
    lignes, sessions = _registre_sain()
    assert _reg().anomalies(lignes, sessions) == []


CAS_ANOMALIE = [
    # (identifiant, N° concerné, modification de la ligne, mot attendu dans le message)
    ("envoi-avant-emission", "R_ABC003", dict(statut_saisi="envoyée", date_envoi=date(2026, 11, 28)), "envoi"),
    ("paiement-avant-emission", "R_ABC002", dict(statut_saisi="payée", date_paiement=date(2026, 11, 30)), "paiement"),
    ("echeance-avant-emission", "R_ABC003", dict(echeance=date(2026, 11, 15)), "échéance"),
    ("paiement-sans-envoi", "R_ABC003", dict(statut_saisi="payée", date_paiement=date(2026, 12, 20)), "envoi"),
    ("total-different-de-sessions", "R_ABC002", dict(total_ht=Decimal("300.00")), "total"),
    ("total-a-un-centime-pres", "R_ABC003", dict(total_ht=Decimal("90.01")), "total"),
    ("statut-paye-sans-paiement", "R_ABC003", dict(statut_saisi="payée"), "statut"),
    ("statut-envoye-sans-envoi", "R_ABC003", dict(statut_saisi="envoyée"), "statut"),
]


@pytest.mark.parametrize(("n_facture", "champs", "mot"), [c[1:] for c in CAS_ANOMALIE], ids=[c[0] for c in CAS_ANOMALIE])
def test_t_rgf_13b_une_anomalie_par_incoherence(n_facture, champs, mot):
    """T-RGF-13b — chaque incohérence ajoutée au registre sain donne exactement une anomalie, qui nomme
    le N° et la cause."""
    lignes, sessions = _registre_sain()
    trouvees = _reg().anomalies(_remplacer(lignes, n_facture, **champs), sessions)
    assert len(trouvees) == 1, trouvees
    assert n_facture in trouvees[0] and mot in trouvees[0].casefold(), trouvees


def test_t_rgf_13c_numero_de_sessions_absent_du_registre():
    """T-RGF-13c — un N° présent dans `Sessions` et absent du registre est une anomalie qui propose
    `registre initialiser --reprendre-historique`."""
    lignes, sessions = _registre_sain()
    trouvees = _reg().anomalies(lignes, [*sessions, _seance("R_ABC005", 100, rang=7)])
    assert len(trouvees) == 1, trouvees
    assert "R_ABC005" in trouvees[0], trouvees
    assert "registre initialiser --reprendre-historique" in trouvees[0], trouvees


def test_t_rgf_14_total_compare_en_decimal():
    """T-RGF-14 — le total se compare en `Decimal` : 100,10 + 200,20 font exactement 300,30 (en virgule
    flottante, 300,29999999999995), donc aucune anomalie ; 157,49 contre 157,50 en est une."""
    reg = _reg()
    lignes = [_ligne(n_facture="R_ABC001", total_ht=Decimal("300.30"))]
    sessions = [_seance("R_ABC001", 100.1, rang=2), _seance("R_ABC001", 200.2, rang=3)]
    assert reg.anomalies(lignes, sessions) == []
    lignes = [_ligne(n_facture="R_ABC001", total_ht=Decimal("157.50"))]
    sessions = [_seance("R_ABC001", 100, rang=2), _seance("R_ABC001", 57.49, rang=3)]
    trouvees = reg.anomalies(lignes, sessions)
    assert len(trouvees) == 1 and "R_ABC001" in trouvees[0], trouvees


# ──────────────────────────────────────────────────────────────────────
# Trous de numérotation
# ──────────────────────────────────────────────────────────────────────


def test_t_rgf_15_trous_sans_raison():
    """T-RGF-15 — `trous_sans_raison` : les trous de chaque école de la liste, moins les N° qui ont une
    ligne au registre (ADR029) ; les autres préfixes et les écoles hors liste sont ignorés ; une école
    donnée deux fois ne double pas ses trous."""
    reg = _reg()
    numeros = ["R_ABC001", "R_ABC004", "R_ECO003", "K_ABC005", "R_XYZ002"]
    lignes = [_ligne(n_facture="R_ABC002", statut_saisi="abandonnée", note="Numéro sauté")]
    trous = reg.trous_sans_raison(lignes, numeros, prefixe="R_", ecoles=["ABC", "ECO", "ABC"])
    assert sorted(trous) == ["R_ABC003", "R_ECO001", "R_ECO002"] and len(trous) == 3
    assert reg.trous_sans_raison([], ["R_ABC001", "R_ABC002"], prefixe="R_", ecoles=["ABC"]) == []


# ──────────────────────────────────────────────────────────────────────
# Écriture sur une feuille ouverte
# ──────────────────────────────────────────────────────────────────────


def _pour_emission(n_facture="R_ABC003", **surcharges):
    valeurs = dict(n_facture=n_facture, ecole="ABC", contrat="CT0001", periode="Novembre 2026",
                   date_emission=date(2026, 12, 1), echeance=date(2027, 1, 31), total_ht=Decimal("1800.00"))
    valeurs.update(surcharges)
    return _reg().ligne_pour_emission(**valeurs)


def test_t_rgf_16_creer_onglet():
    """T-RGF-16 — `creer_onglet` ajoute au classeur l'onglet `Factures`, avec les 11 en-têtes en ligne 1
    et aucune autre ligne."""
    reg = _reg()
    wb = Workbook()
    ws = reg.creer_onglet(wb)
    assert ws.title == "Factures" and wb["Factures"] is ws
    assert [cellule.value for cellule in ws[1]] == list(COLONNES_FACTURES)
    assert ws.max_row == 1


def test_t_rgf_17_ajouter_ligne_par_nom_de_colonne():
    """T-RGF-17 — `ajouter_ligne` écrit chaque valeur sous son en-tête, quel que soit l'ordre des colonnes
    (une colonne en plus reste vide), sur la ligne qui suit la dernière ligne remplie (une ligne vide
    au milieu est sautée), et rend ce numéro de ligne ; les dates prennent le format DD/MM/YYYY ; les
    lignes existantes ne changent pas."""
    reg = _reg()
    colonnes = list(reversed(COLONNES_FACTURES))
    colonnes.insert(4, "Commentaire libre")
    wb, ws = _feuille(
        [_facture("R_ABC001", statut="émise", total=90), None, _facture("R_ABC002", statut="émise", total=45)],
        colonnes=colonnes,
    )
    avant = _instantane(ws)
    ligne = _pour_emission()

    rang = reg.ajouter_ligne(ws, ligne)

    assert rang == 5
    assert _valeur(ws, 5, "N° facture") == "R_ABC003"
    assert _valeur(ws, 5, "École") == "ABC"
    assert _valeur(ws, 5, "Contrat") == "CT0001"
    assert _valeur(ws, 5, "Période") == "Novembre 2026"
    assert _jour(_valeur(ws, 5, "Date d'émission")) == date(2026, 12, 1)
    assert _jour(_valeur(ws, 5, "Échéance")) == date(2027, 1, 31)
    assert Decimal(str(_valeur(ws, 5, "Total HT"))) == Decimal("1800")
    assert _valeur(ws, 5, "Statut") == "en émission"
    for vide in ("Date d'envoi", "Date de paiement", "Commentaire libre"):
        assert _valeur(ws, 5, vide) in (None, ""), vide
    for colonne in ("Date d'émission", "Échéance"):
        cellule = ws.cell(row=5, column=_index(ws)[colonne])
        assert cellule.number_format.upper() == "DD/MM/YYYY", (colonne, cellule.number_format)
    apres = _instantane(ws)
    assert {c: v for c, v in apres.items() if not c.endswith("5")} == avant


@pytest.mark.parametrize("numero", ["R_ABC001", "r_abc001"], ids=["meme-casse", "autre-casse"])
def test_t_rgf_04a_ajouter_ligne_refuse_un_numero_deja_present(numero):
    """T-RGF-04a — `ajouter_ligne` refuse un N° déjà au registre (sans tenir compte de la casse) :
    `RegistreInvalide` qui le nomme, feuille inchangée."""
    reg = _reg()
    wb, ws = _feuille([_facture("R_ABC001", statut="émise", total=90)])
    avant = _instantane(ws)
    with pytest.raises(reg.RegistreInvalide) as exc:
        reg.ajouter_ligne(ws, _pour_emission(numero))
    assert "R_ABC001" in str(exc.value).upper(), str(exc.value)
    assert _instantane(ws) == avant


def test_t_rgf_18_modifier_ligne_envoi_puis_paiement():
    """T-RGF-18 — une date d'envoi fait passer « émise » à « envoyée », une date de paiement fait passer
    à « payée » ; `modifier_ligne` rend {colonne : (avant, après)} ; une note seule ne touche qu'à la
    note ; « échue » n'est jamais écrit, même pour une facture dont l'échéance est passée."""
    reg = _reg()
    wb, ws = _feuille([
        _facture("R_ABC001", statut="émise", emission=date(2026, 12, 1), echeance=date(2027, 1, 31), total=90),
        _facture("R_ABC002", statut="émise", emission=date(2020, 1, 1), echeance=date(2020, 1, 31), total=45),
        _facture("R_ABC003", statut="émise", emission=date(2026, 12, 1), echeance=date(2027, 1, 31), total=45),
    ])

    changements = reg.modifier_ligne(ws, "R_ABC001", date_envoi=date(2026, 12, 2))
    assert set(changements) == {"Date d'envoi", "Statut"}
    avant, apres = changements["Date d'envoi"]
    assert avant in (None, "") and _jour(apres) == date(2026, 12, 2)
    assert changements["Statut"] == ("émise", "envoyée")
    assert _jour(_valeur(ws, 2, "Date d'envoi")) == date(2026, 12, 2) and _valeur(ws, 2, "Statut") == "envoyée"

    changements = reg.modifier_ligne(ws, "R_ABC001", date_paiement=date(2026, 12, 20))
    assert set(changements) == {"Date de paiement", "Statut"}
    assert changements["Statut"] == ("envoyée", "payée")
    assert _jour(_valeur(ws, 2, "Date de paiement")) == date(2026, 12, 20) and _valeur(ws, 2, "Statut") == "payée"

    reg.modifier_ligne(ws, "R_ABC002", date_envoi=date(2020, 2, 15))
    assert _valeur(ws, 3, "Statut") == "envoyée"

    changements = reg.modifier_ligne(ws, "R_ABC003", date_paiement=date(2026, 12, 20))
    assert changements["Statut"] == ("émise", "payée")

    changements = reg.modifier_ligne(ws, "R_ABC002", note="Relance faite le 05/01/2027")
    assert set(changements) == {"Note"} and changements["Note"][1] == "Relance faite le 05/01/2027"
    assert _valeur(ws, 3, "Note") == "Relance faite le 05/01/2027" and _valeur(ws, 3, "Statut") == "envoyée"


def _feuille_des_refus():
    return _feuille([
        _facture("R_ABC001", statut="émise", emission=date(2026, 12, 1), echeance=date(2027, 1, 31), total=90),
        _facture("R_ABC002", statut="envoyée", emission=date(2026, 12, 1), echeance=date(2027, 1, 31), total=90,
                 envoi=date(2026, 12, 3)),
        _facture("R_ABC003", statut="abandonnée", note="Erreur de saisie"),
        _facture("R_ECO001", statut="émise", emission=date(2026, 12, 1), total=50),
        _facture("r_eco001", statut="émise", emission=date(2026, 12, 1), total=50),
    ])


CAS_REFUS = [
    # (identifiant, N°, champs, mot attendu dans le message)
    ("champ-inconnu", "R_ABC001", dict(montant=100), "montant"),
    ("statut-reserve-envoyee", "R_ABC001", dict(statut="envoyée"), "statut"),
    ("statut-echue-jamais-ecrit", "R_ABC001", dict(statut="échue"), "statut"),
    ("numero-absent", "R_ABC099", dict(date_envoi=date(2026, 12, 5)), "R_ABC099"),
    ("numero-ambigu", "R_ECO001", dict(date_envoi=date(2026, 12, 5)), "R_ECO001"),
    ("ligne-abandonnee", "R_ABC003", dict(date_envoi=date(2026, 12, 5)), "R_ABC003"),
    ("envoi-avant-emission", "R_ABC001", dict(date_envoi=date(2026, 11, 30)), "émission"),
    ("paiement-avant-emission", "R_ABC001", dict(date_paiement=date(2026, 11, 30)), "émission"),
    ("date-deja-saisie", "R_ABC002", dict(date_envoi=date(2026, 12, 4)), "remplacer"),
]


@pytest.mark.parametrize(("n_facture", "champs", "mot"), [c[1:] for c in CAS_REFUS], ids=[c[0] for c in CAS_REFUS])
def test_t_rgf_19a_modifier_ligne_refuse(n_facture, champs, mot):
    """T-RGF-19a — `modifier_ligne` refuse, par une `ValueError` qui nomme la cause, un champ inconnu, un
    statut autre que « émise » ou « abandonnée », un N° absent ou ambigu, une ligne abandonnée, une
    date antérieure à l'émission, une cellule déjà remplie d'une autre valeur sans `remplacer` ; la
    feuille n'est pas modifiée."""
    reg = _reg()
    wb, ws = _feuille_des_refus()
    avant = _instantane(ws)
    with pytest.raises(ValueError) as exc:
        reg.modifier_ligne(ws, n_facture, **champs)
    assert mot.casefold() in str(exc.value).casefold(), str(exc.value)
    assert _instantane(ws) == avant


def test_t_rgf_19b_modifier_ligne_ce_qui_reste_permis():
    """T-RGF-19b — la note d'une ligne abandonnée reste modifiable ; `remplacer=True` remplace une date
    déjà saisie ; la même date redonnée sans `remplacer` n'est pas un refus et ne change rien."""
    reg = _reg()
    wb, ws = _feuille_des_refus()

    changements = reg.modifier_ligne(ws, "R_ABC003", note="Erreur de saisie, numéro jamais envoyé", remplacer=True)
    assert changements["Note"] == ("Erreur de saisie", "Erreur de saisie, numéro jamais envoyé")
    assert _valeur(ws, 4, "Statut") == "abandonnée"

    avant = _instantane(ws)
    reg.modifier_ligne(ws, "R_ABC002", date_envoi=date(2026, 12, 3))
    assert _instantane(ws) == avant

    changements = reg.modifier_ligne(ws, "R_ABC002", date_envoi=date(2026, 12, 4), remplacer=True)
    avant_envoi, apres_envoi = changements["Date d'envoi"]
    assert _jour(avant_envoi) == date(2026, 12, 3) and _jour(apres_envoi) == date(2026, 12, 4)
    assert _jour(_valeur(ws, 3, "Date d'envoi")) == date(2026, 12, 4)
    assert _valeur(ws, 3, "Statut") == "envoyée"


def test_t_rgf_20_modifier_ligne_n_ecrase_jamais_une_formule():
    """T-RGF-20 — une cellule cible qui contient une formule n'est jamais écrasée, même avec
    `remplacer=True` (§ 1.4) : `ValueError` qui parle de formule, feuille inchangée."""
    reg = _reg()
    wb, ws = _feuille([
        _facture("R_ABC001", statut="envoyée", emission=date(2026, 12, 1), echeance=date(2027, 1, 31), total=90,
                 envoi=date(2026, 12, 2), paiement="=TODAY()"),
    ])
    avant = _instantane(ws)
    with pytest.raises(ValueError) as exc:
        reg.modifier_ligne(ws, "R_ABC001", date_paiement=date(2026, 12, 20), remplacer=True)
    assert "formule" in str(exc.value).casefold(), str(exc.value)
    assert _instantane(ws) == avant
    assert _valeur(ws, 2, "Date de paiement") == "=TODAY()"


def test_t_rgf_21_noter_numero_manquant_sur_la_feuille():
    """T-RGF-21 — `noter_numero_manquant` ajoute une ligne au N°, avec l'école, le statut « abandonnée »
    et la raison en note, sans date ni montant, et rend son numéro de ligne ; un N° déjà au registre
    est refusé (`RegistreInvalide`)."""
    reg = _reg()
    wb, ws = _feuille([_facture("R_ABC001", statut="émise", total=90)])

    rang = reg.noter_numero_manquant(ws, "R_ABC002", ecole="ABC", raison="Numéro sauté à la saisie")

    assert rang == 3
    assert _valeur(ws, 3, "N° facture") == "R_ABC002"
    assert _valeur(ws, 3, "École") == "ABC"
    assert _valeur(ws, 3, "Statut") == "abandonnée"
    assert "Numéro sauté à la saisie" in str(_valeur(ws, 3, "Note"))
    for vide in ("Date d'émission", "Échéance", "Total HT", "Date d'envoi", "Date de paiement"):
        assert _valeur(ws, 3, vide) in (None, ""), vide
    avant = _instantane(ws)
    with pytest.raises(reg.RegistreInvalide):
        reg.noter_numero_manquant(ws, "R_ABC001", ecole="ABC", raison="Doublon")
    assert _instantane(ws) == avant


def test_t_rgf_22_statut_ecrit_par_l_emission():
    """T-RGF-22 — F-17 passe une ligne « en émission » à « émise » (contrôle vert) ou à « abandonnée »
    par `modifier_ligne(statut=…)`, qui rend le changement de statut."""
    reg = _reg()
    wb, ws = _feuille([])
    reg.ajouter_ligne(ws, _pour_emission("R_ABC001"))
    reg.ajouter_ligne(ws, _pour_emission("R_ABC002"))

    assert reg.modifier_ligne(ws, "R_ABC001", statut="émise") == {"Statut": ("en émission", "émise")}
    assert reg.modifier_ligne(ws, "R_ABC002", statut="abandonnée") == {"Statut": ("en émission", "abandonnée")}
    assert _valeur(ws, 2, "Statut") == "émise" and _valeur(ws, 3, "Statut") == "abandonnée"


# ──────────────────────────────────────────────────────────────────────
# Constats de revue de F-25
# ──────────────────────────────────────────────────────────────────────
#
# Rouges sur le code du commit 4e696a6 : T-RGF-01b, T-RGF-04c, T-RGF-13d, T-RGF-14b, T-RGF-17b (cas
# « caractère de contrôle »), T-RGF-19d, T-RGF-19e, T-RGF-22b (statuts « émise » et vide). Les autres
# cas gardent un comportement déjà correct que rien ne testait : chacun a été vu rouge, code cassé
# exprès (AGENTS.md § 5).


@pytest.mark.parametrize("total", [Decimal("NaN"), Decimal("Infinity")], ids=["nan", "infini"])
def test_t_rgf_01b_ligne_pour_emission_refuse_un_total_qui_n_est_pas_un_montant(total):
    """T-RGF-01b — `ligne_pour_emission` refuse un total `Decimal` qui n'est pas un montant (NaN,
    infini) : `RegistreInvalide`. Aujourd'hui, NaN est accepté (la cellule écrite se relit vide) et
    l'infini lève `decimal.InvalidOperation`, qui n'est pas une `ValueError`."""
    reg = _reg()
    with pytest.raises(reg.RegistreInvalide):
        _pour_emission(total_ht=total)


@pytest.mark.parametrize("second", ["R_ABC_001", "R_ABC001​"], ids=["tiret-bas-avant-le-rang", "espace-invisible"])
def test_t_rgf_04c_doublon_au_meme_numero_de_la_serie(second):
    """T-RGF-04c — deux lignes qui désignent le même numéro lèvent `RegistreInvalide` : « R_ABC_001 »,
    que `numerotation` compte au même rang que « R_ABC001 » (`_rangs`), ou « R_ABC001 » suivi d'un
    espace de largeur nulle (U+200B), invisible dans Excel. Sinon le registre porte deux lignes pour
    une même facture, et la reprise de l'historique peut la reprendre une seconde fois."""
    from facturation.master import Master

    reg = _reg()
    master = Master(factures=[{"N° facture": "R_ABC001", "Statut": "émise"}, {"N° facture": second, "Statut": "émise"}])
    with pytest.raises(reg.RegistreInvalide):
        reg.lire_registre(master)


CAS_STATUT_INCOHERENT = [
    # (identifiant, N° du registre sain, dates ajoutées, statut que l'anomalie doit citer)
    ("en-emission-envoyee", "R_ECO001", dict(date_envoi=date(2026, 12, 2)), "en émission"),
    ("en-emission-payee", "R_ECO001", dict(date_envoi=date(2026, 12, 2), date_paiement=date(2026, 12, 20)),
     "en émission"),
    ("abandonnee-payee", "R_ABC004", dict(date_paiement=date(2026, 12, 20)), "abandonnée"),
]


@pytest.mark.parametrize(("n_facture", "champs", "statut"), [c[1:] for c in CAS_STATUT_INCOHERENT],
                         ids=[c[0] for c in CAS_STATUT_INCOHERENT])
def test_t_rgf_13d_statut_incoherent_avec_les_dates(n_facture, champs, statut):
    """T-RGF-13d — « statut saisi incohérent avec les dates » couvre aussi une ligne « en émission » qui
    porte une date d'envoi ou de paiement (saisie dans Excel ; `modifier_ligne` la refuse, et elle
    s'afficherait « payée »), et une ligne « abandonnée » qui porte une date de paiement : une anomalie
    qui nomme le N° et le statut saisi."""
    lignes, sessions = _registre_sain()
    trouvees = _reg().anomalies(_remplacer(lignes, n_facture, **champs), sessions)
    assert any(n_facture in a and statut in a.casefold() for a in trouvees), trouvees


@pytest.mark.parametrize("total", ["157.50", "315.00"], ids=["total-des-seules-lignes-remplies", "vrai-total"])
def test_t_rgf_14b_ligne_sessions_sans_total(total):
    """T-RGF-14b — une ligne `Sessions` du N° sans Total HT (ligne 3) : le total ne peut pas être
    rapproché. `anomalies` le signale en nommant la ligne `Sessions` fautive, que le registre porte
    le total des seules lignes remplies (157,50 €, aujourd'hui accepté sans bruit) ou le vrai total
    (315,00 €, aujourd'hui « différent de la somme », sans dire pourquoi)."""
    reg = _reg()
    sessions = [_seance("R_ABC001", 157.5, rang=2), _seance("R_ABC001", None, rang=3)]
    trouvees = reg.anomalies([_ligne(n_facture="R_ABC001", total_ht=Decimal(total))], sessions)
    assert any("R_ABC001" in a and "Sessions" in a and re.search(r"\b3\b", a) for a in trouvees), trouvees


@pytest.mark.parametrize(("total", "anomalie"), [("99.99", False), ("100.00", False), ("100.01", True)],
                         ids=["total-de-la-facture", "somme-des-lignes", "un-centime-de-trop"])
def test_t_rgf_14c_total_arrondi_une_seule_fois(total, anomalie):
    """T-RGF-14c — deux séances de 1,5 h à 33,33 € : chaque ligne `Sessions` porte 50,00 € (49,995
    arrondi), la facture 99,99 € (somme exacte 99,990 arrondie une fois, § 1.2). Le registre est
    cohérent à 99,99 € comme à 100,00 € ; à 100,01 €, c'est une anomalie."""
    reg = _reg()
    seances = [dataclasses.replace(_seance("R_ABC001", 50, rang=rang), duree_h=1.5, taux=33.33) for rang in (2, 3)]
    trouvees = reg.anomalies([_ligne(n_facture="R_ABC001", total_ht=Decimal(total))], seances)
    assert bool(trouvees) is anomalie, trouvees


@pytest.mark.parametrize("note", ["=1+1", "bip\x07"], ids=["formule", "caractere-de-controle"])
def test_t_rgf_17b_valeur_refusee_ne_laisse_aucune_cellule(note):
    """T-RGF-17b — `ajouter_ligne` et `modifier_ligne` refusent une note qui serait une formule
    (« =1+1 ») ou qui porte un caractère de contrôle qu'un classeur ne peut pas enregistrer (BEL) :
    `RegistreInvalide`, et la feuille n'est pas modifiée. Aujourd'hui, le caractère de contrôle lève
    `IllegalCharacterError` (qui n'est pas une `ValueError`) après que le N° et l'école de la ligne
    ajoutée ont été écrits."""
    reg = _reg()
    wb, ws = _feuille([_facture("R_ABC001", statut="émise", emission=date(2026, 12, 1), total=90)])
    avant = _instantane(ws)
    with pytest.raises(reg.RegistreInvalide):
        reg.ajouter_ligne(ws, {"N° facture": "R_ABC002", "École": "ABC", "Note": note})
    assert _instantane(ws) == avant
    with pytest.raises(reg.RegistreInvalide):
        reg.modifier_ligne(ws, "R_ABC001", note=note)
    assert _instantane(ws) == avant


@pytest.mark.parametrize("champ", ["date_envoi", "date_paiement"])
def test_t_rgf_19c_envoi_et_paiement_refuses_en_emission(champ):
    """T-RGF-19c — une facture « en émission » (son contrôle après écriture n'est pas vert) ne peut pas
    être notée envoyée ni payée : `RegistreInvalide` qui cite « en émission », feuille inchangée."""
    reg = _reg()
    wb, ws = _feuille([_facture("R_ABC001", statut="en émission", emission=date(2026, 12, 1),
                                echeance=date(2027, 1, 31), total=90)])
    avant = _instantane(ws)
    with pytest.raises(reg.RegistreInvalide) as exc:
        reg.modifier_ligne(ws, "R_ABC001", **{champ: date(2026, 12, 2)})
    assert "en émission" in str(exc.value), str(exc.value)
    assert _instantane(ws) == avant


def test_t_rgf_19d_envoi_saisi_apres_le_paiement_ne_recule_pas_le_statut():
    """T-RGF-19d — une facture payée dans Excel sans toucher au statut (ADR031 : date de paiement
    remplie, statut vide) : la date d'envoi notée ensuite ne fait pas écrire « envoyée » (un statut
    en retrait sur la date de paiement) ; le statut reste vide ou devient « payée »."""
    reg = _reg()
    wb, ws = _feuille([_facture("R_ABC001", statut=None, emission=date(2026, 12, 1), echeance=date(2027, 1, 31),
                                total=90, paiement=date(2026, 12, 20))])
    changements = reg.modifier_ligne(ws, "R_ABC001", date_envoi=date(2026, 12, 2))
    assert _jour(_valeur(ws, 2, "Date d'envoi")) == date(2026, 12, 2)
    assert _valeur(ws, 2, "Statut") in (None, "", "payée"), changements


def test_t_rgf_19e_refus_du_statut_emise_ne_parle_pas_d_abandon():
    """T-RGF-19e — demander « émise » pour une facture « envoyée » est refusé par un message qui parle
    de ce passage-là, pas d'abandon (aujourd'hui : « une facture envoyée ou payée ne s'abandonne
    pas »)."""
    reg = _reg()
    wb, ws = _feuille_des_refus()
    with pytest.raises(reg.RegistreInvalide) as exc:
        reg.modifier_ligne(ws, "R_ABC002", statut="émise")
    message = str(exc.value)
    assert "R_ABC002" in message and "émise" in message, message
    assert "abandon" not in message.casefold(), message


CAS_ABANDON_REFUSE = [
    # (identifiant, cellules de la ligne R_ABC001, émise le 01/12/2026)
    ("emise", dict(statut="émise")),
    ("emise-envoyee-et-payee-dans-excel", dict(statut="émise", envoi=date(2026, 12, 2), paiement=date(2026, 12, 20))),
    ("statut-vide", dict(statut=None)),
    ("statut-vide-envoyee-et-payee-dans-excel", dict(statut=None, envoi=date(2026, 12, 2),
                                                     paiement=date(2026, 12, 20))),
    ("envoyee", dict(statut="envoyée", envoi=date(2026, 12, 2))),
    ("payee", dict(statut="payée", envoi=date(2026, 12, 2), paiement=date(2026, 12, 20))),
]


@pytest.mark.parametrize("cellules", [c[1] for c in CAS_ABANDON_REFUSE], ids=[c[0] for c in CAS_ABANDON_REFUSE])
def test_t_rgf_22b_abandon_seulement_en_emission(cellules):
    """T-RGF-22b — `modifier_ligne(statut="abandonnée")` n'abandonne qu'une facture « en émission »
    (§ 8.7 : « jamais pour une facture émise » ; Q-E2 et Q-E3, § 16 : une facture émise se corrige par
    un avoir). Une facture émise, au statut vide, envoyée ou payée, y compris quand ses dates ont été
    saisies dans Excel sans toucher au statut (ADR031), est refusée (`RegistreInvalide` qui cite le
    N°), feuille inchangée. Sinon ses lignes `Sessions` sortiraient de l'anti-doublon et ses séances
    seraient refacturées. L'abandon d'une facture « en émission » reste permis (T-RGF-22)."""
    reg = _reg()
    wb, ws = _feuille([_facture("R_ABC001", emission=date(2026, 12, 1), echeance=date(2027, 1, 31), total=90,
                                **cellules)])
    avant = _instantane(ws)
    with pytest.raises(reg.RegistreInvalide) as exc:
        reg.modifier_ligne(ws, "R_ABC001", statut="abandonnée")
    assert "R_ABC001" in str(exc.value), str(exc.value)
    assert _instantane(ws) == avant
