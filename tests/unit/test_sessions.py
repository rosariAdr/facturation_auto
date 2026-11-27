"""Lecture de l'onglet `Sessions` (P-0b, `facturation/sessions.py` ; banque de tests : T-SES-01 à T-SES-18).

Tests d'acceptation écrits avant l'implémentation : chaque test importe `facturation.sessions`
(et `facturation.master`) lui-même, pour que l'absence du module fasse échouer les tests sans
bloquer la collecte des autres. MASTER fictifs créés dans `tmp_path` (`tests.conftest.creer_master`).

Contrat (contrats des vagues 3 à 5, § 3.3), et décisions du § 15 qui priment :
- `code_ecole` : vide → None ; un `Code facturation` de `Écoles` (sans tenir compte de la casse,
  des espaces ni de la forme Unicode) → ce code ; un nom de la colonne `École` → son code ;
- `lire_sessions` lit les dates (`date`, `datetime`, `JJ/MM/AAAA`, `AAAA-MM-JJ`), les heures
  (`time`, `datetime`, `HH:MM`, `HHhMM`, et les formes historiques « 9h », « 9h30 », « 09:00 »
  et le vide, § 15), les nombres (valeur numérique ou texte « 3,5 ») ; une cellule illisible place
  la ligne dans la seconde liste (rang Excel, message qui nomme la colonne et la valeur), sans
  plantage ;
- `chercher` rend les lignes exactes (même école, même date, même début et même fin à la minute),
  puis les lignes du même jour sans heure, hors lignes d'un N° abandonné.

Lectures retenues (`💡 [proposé — à valider]`) :
- le § 15 (« un libellé inconnu donne une erreur qui nomme l'onglet, la ligne et la valeur »)
  prime sur le « sinon → la valeur nettoyée, en majuscules » du § 3.3 pour un texte qui n'a pas
  la forme d'un code école (majuscules et chiffres, ADR029) : `code_ecole` lève alors une
  `ValueError` qui cite la valeur et l'onglet `Écoles` ; une valeur qui a la forme d'un code
  (« xyz », « XYZ ») reste rendue nettoyée, en majuscules, même absente de `Écoles` ;
- dans `lire_sessions`, ce refus place la ligne dans la seconde liste, avec son rang ;
- une heure vide vaut None et ne rend pas la ligne illisible ; « sans heure » (pour `chercher`)
  veut dire qu'au moins une des deux heures manque (revue de P-0b, T-SES-17) ;
- `Total HT` devient un `Decimal` égal au nombre tel qu'il s'écrit (157.5 → Decimal("157.5")) ;
- le statut « abandonnée » de `Factures` est reconnu sans tenir compte de la casse ni des espaces,
  ni de l'accent ; le n° abandonné se compare à celui de `Sessions` sans casse ni espaces (revue de
  P-0b, T-SES-18).
"""

from __future__ import annotations

import dataclasses
import importlib
import unicodedata
from datetime import date, datetime, time
from decimal import Decimal
from zoneinfo import ZoneInfo

import pytest

from tests.conftest import creer_master

pytestmark = pytest.mark.unit

PARIS = ZoneInfo("Europe/Paris")

ECOLES = [
    {"École": "École Fictive A", "Code facturation": "ABC", "Entité légale": "SAS FICTIVE A"},
    {"École": "Institut Fictif B", "Code facturation": "ECO", "Entité légale": "ASSOCIATION FICTIVE B"},
]

CHAMPS = (
    "rang", "ecole", "n_facture", "periode", "date_seance", "debut", "fin", "duree_h", "promotion",
    "module", "taux", "total_ht", "summary", "abandonnee",
)


def _ses():
    """Module testé, importé dans le test : son absence fait échouer le test, pas la collecte."""
    return importlib.import_module("facturation.sessions")


def _session(**valeurs) -> dict:
    """Ligne fictive de `Sessions` : une séance d'ABC du 15/09/2026, 09:00–12:30, sauf surcharge.

    Les arguments nomment les colonnes sans accents : `ecole`, `periode`, `n_facture`, `date`,
    `debut`, `fin`, `duree`, `promotion`, `module`, `taux`, `total`, `summary`.
    """
    ligne = {
        "Date_log": "2026-10-01", "École": "ABC", "Période_facture": "Septembre 2026",
        "N°_Facture": "R_ABC001", "Date_séance": date(2026, 9, 15), "Début": "09:00", "Fin": "12:30",
        "Durée_h": 3.5, "Promotion": "M1 S1", "Module / UE": "Module fictif A", "Taux": 45,
        "Total HT": 157.5, "Summary calendrier verbatim": "#ABC-M1 Cours fictif",
    }
    alias = {"ecole": "École", "periode": "Période_facture", "n_facture": "N°_Facture",
             "date": "Date_séance", "debut": "Début", "fin": "Fin", "duree": "Durée_h",
             "promotion": "Promotion", "module": "Module / UE", "taux": "Taux", "total": "Total HT",
             "summary": "Summary calendrier verbatim"}
    for cle, valeur in valeurs.items():
        ligne[alias[cle]] = valeur
    return ligne


def _lire(tmp_path, sessions, *, factures=None, ecoles=ECOLES):
    """(lignes, erreurs) de `lire_sessions` sur un MASTER fictif."""
    chemin = creer_master(tmp_path / "donnees", sessions=sessions, factures=factures, ecoles=ecoles)
    master = importlib.import_module("facturation.master").lire_master(chemin)
    return _ses().lire_sessions(master)


def _une(tmp_path, **valeurs):
    """La seule ligne lue d'un MASTER à une séance ; échoue si elle est jugée illisible."""
    lignes, erreurs = _lire(tmp_path, [_session(**valeurs)])
    assert erreurs == [], erreurs
    (ligne,) = lignes
    return ligne


# ──────────────────────────────────────────────────────────────────────
# code_ecole
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    "valeur, attendu",
    [("École Fictive A", "ABC"), ("  école fictive a  ", "ABC"), (unicodedata.normalize("NFD", "École Fictive A"), "ABC"),
     ("ÉCOLE FICTIVE A", "ABC"), ("Institut Fictif B", "ECO"), ("ABC", "ABC"), (" abc ", "ABC"), ("eco", "ECO")],
    ids=["libelle", "libelle-casse-espaces", "libelle-nfd", "libelle-majuscules", "autre-libelle", "code",
         "code-minuscules-espaces", "autre-code-minuscules"],
)
def test_t_ses_01_code_ecole_traduit_un_nom_en_code(valeur, attendu):
    """T-SES-01 — `code_ecole` rend le code d'un libellé de la colonne `École` ou d'un `Code
    facturation`, sans tenir compte de la casse, des espaces ni de la forme Unicode."""
    assert _ses().code_ecole(valeur, ECOLES) == attendu


@pytest.mark.parametrize("valeur", [None, "", "   "], ids=["None", "vide", "espaces"])
def test_t_ses_04_code_ecole_vide(valeur):
    """T-SES-04 — une cellule vide n'a pas d'école : None."""
    assert _ses().code_ecole(valeur, ECOLES) is None


@pytest.mark.parametrize("valeur, attendu", [("XYZ", "XYZ"), (" xyz ", "XYZ"), ("EC2", "EC2")])
def test_t_ses_05_code_absent_de_ecoles_rendu_nettoye(valeur, attendu):
    """T-SES-05 — une valeur qui a la forme d'un code école, même absente de `Écoles` (onglet vide
    compris), est rendue nettoyée, en majuscules (§ 3.3)."""
    ses = _ses()
    assert ses.code_ecole(valeur, ECOLES) == attendu
    assert ses.code_ecole(valeur, []) == attendu


@pytest.mark.parametrize("valeur", ["École Fictive Inconnue", "Institut Fictif C", "Ecole fictive A Paris"])
def test_t_ses_06_libelle_inconnu_refuse(valeur):
    """T-SES-06 — un libellé inconnu de l'onglet `Écoles` lève `ValueError` qui cite la valeur et
    l'onglet `Écoles` (§ 15) : une école n'est jamais devinée."""
    with pytest.raises(ValueError) as erreur:
        _ses().code_ecole(valeur, ECOLES)
    message = str(erreur.value)
    assert valeur in message
    assert "Écoles" in message


# ──────────────────────────────────────────────────────────────────────
# lire_sessions
# ──────────────────────────────────────────────────────────────────────


def test_t_ses_07_ligne_sessions_dataclass_figee():
    """T-SES-07 — `LigneSessions` est une dataclass figée, aux champs du contrat, dans l'ordre."""
    ses = _ses()
    assert dataclasses.is_dataclass(ses.LigneSessions)
    assert tuple(f.name for f in dataclasses.fields(ses.LigneSessions)) == CHAMPS


def test_t_ses_08_une_ligne_lue_champ_par_champ(tmp_path):
    """T-SES-08 — une ligne complète : rang Excel, code école, textes, date, heures, nombres et
    montant `Decimal` ; ligne non abandonnée."""
    ligne = _une(tmp_path)
    assert ligne.rang == 2
    assert (ligne.ecole, ligne.n_facture, ligne.periode) == ("ABC", "R_ABC001", "Septembre 2026")
    assert ligne.date_seance == date(2026, 9, 15)
    assert (ligne.debut, ligne.fin) == (time(9, 0), time(12, 30))
    assert ligne.duree_h == 3.5 and isinstance(ligne.duree_h, float)
    assert ligne.taux == 45.0 and isinstance(ligne.taux, float)
    assert ligne.total_ht == Decimal("157.5") and isinstance(ligne.total_ht, Decimal)
    assert (ligne.promotion, ligne.module, ligne.summary) == ("M1 S1", "Module fictif A", "#ABC-M1 Cours fictif")
    assert ligne.abandonnee is False
    with pytest.raises(dataclasses.FrozenInstanceError):
        ligne.ecole = "ECO"


def test_t_ses_02_date_illisible_signalee_avec_son_rang(tmp_path):
    """T-SES-02 — une date illisible place la ligne dans la seconde liste avec son rang Excel (ligne
    vide comprise dans le compte) et un message qui nomme la colonne et la valeur ; les autres
    lignes sont lues, sans plantage."""
    lignes, erreurs = _lire(tmp_path, [
        _session(n_facture="R_ABC001"),
        None,  # ligne vide : ignorée, mais comptée dans les rangs
        _session(n_facture="R_ABC002", date="fin nov"),
        _session(n_facture="R_ABC003", date=date(2026, 9, 22)),
    ])
    assert [(l.rang, l.n_facture) for l in lignes] == [(2, "R_ABC001"), (5, "R_ABC003")]
    assert len(erreurs) == 1, erreurs
    rang, message = erreurs[0]
    assert rang == 4
    assert "Date_séance" in message and "fin nov" in message


@pytest.mark.parametrize(
    "saisie",
    [date(2026, 9, 15), datetime(2026, 9, 15, 0, 0), "15/09/2026", "2026-09-15", " 15/09/2026 "],
    ids=["date", "datetime", "jj-mm-aaaa", "aaaa-mm-jj", "texte-espaces"],
)
def test_t_ses_09_formes_de_date(tmp_path, saisie):
    """T-SES-09 — `Date_séance` se lit en date, datetime, JJ/MM/AAAA ou AAAA-MM-JJ."""
    assert _une(tmp_path, date=saisie).date_seance == date(2026, 9, 15)


@pytest.mark.parametrize(
    "debut, fin, attendu",
    [
        ("09:00", "12:30", (time(9, 0), time(12, 30))),
        ("9h", "12h", (time(9, 0), time(12, 0))),
        ("9h30", "12h45", (time(9, 30), time(12, 45))),
        ("09h30", "14h05", (time(9, 30), time(14, 5))),
        (time(9, 0), time(12, 30), (time(9, 0), time(12, 30))),
        (datetime(2026, 9, 15, 9, 0), datetime(2026, 9, 15, 12, 30), (time(9, 0), time(12, 30))),
        (" 9h ", " 12:30 ", (time(9, 0), time(12, 30))),
        (None, None, (None, None)),
        ("", "   ", (None, None)),
    ],
    ids=["hh-mm", "9h", "9h30", "hhhmm", "time", "datetime", "espaces", "vides-none", "vides-texte"],
)
def test_t_ses_10_formes_d_heure_historiques(tmp_path, debut, fin, attendu):
    """T-SES-10 — `Début` et `Fin` : formes historiques du MASTER réel (« 09:00 », « 9h », « 9h30 »,
    vide ; § 15), plus `time`, `datetime` et HHhMM ; une heure vide vaut None, sans erreur."""
    ligne = _une(tmp_path, debut=debut, fin=fin)
    assert (ligne.debut, ligne.fin) == attendu


@pytest.mark.parametrize(
    "colonne, alias, valeur",
    [("Début", "debut", "matin"), ("Fin", "fin", "25h"), ("Durée_h", "duree", "trois heures"),
     ("Total HT", "total", "à voir")],
)
def test_t_ses_11_cellule_illisible_signalee(tmp_path, colonne, alias, valeur):
    """T-SES-11 — une heure ou un nombre illisible : la ligne va dans la seconde liste, avec son rang
    et un message qui nomme la colonne et la valeur."""
    lignes, erreurs = _lire(tmp_path, [_session(**{alias: valeur})])
    assert lignes == []
    ((rang, message),) = erreurs
    assert rang == 2
    assert colonne in message and valeur in message


@pytest.mark.parametrize(
    "duree, taux, total, attendu",
    [
        (3.5, 45, 157.5, (3.5, 45.0, Decimal("157.5"))),
        ("3,5", "45", "157,50", (3.5, 45.0, Decimal("157.50"))),
        ("3.5", "45,5", "159,25", (3.5, 45.5, Decimal("159.25"))),
        (2, 60.0, 120, (2.0, 60.0, Decimal("120"))),
        (None, None, None, (None, None, None)),
    ],
    ids=["nombres", "textes-virgule", "textes-point", "entiers", "vides"],
)
def test_t_ses_12_nombres_en_valeur_ou_en_texte(tmp_path, duree, taux, total, attendu):
    """T-SES-12 — `Durée_h`, `Taux` et `Total HT` se lisent en nombre ou en texte (« 3,5 ») ; vides → None."""
    ligne = _une(tmp_path, duree=duree, taux=taux, total=total)
    assert (ligne.duree_h, ligne.taux, ligne.total_ht) == attendu


@pytest.mark.parametrize(
    "periode",
    ["Décembre 2026", "01/12–15/12/2026", "Nov-Déc 2026", "Décembre 2026 (complément)", "Déc 2026"],
)
def test_t_ses_13_periodes_historiques_tolerees(tmp_path, periode):
    """T-SES-13 — `Période_facture` est un texte libre : les formes historiques (intervalle,
    « Mois-Mois AAAA », parenthèses ; § 15) se lisent telles quelles, sans erreur."""
    assert _une(tmp_path, periode=f" {periode} ").periode == periode


def test_t_ses_14_ecole_en_libelle_ou_inconnue(tmp_path):
    """T-SES-14 — la colonne `École` porte le code ou le libellé (§ 15) : le libellé devient le code ;
    un libellé inconnu de `Écoles` place la ligne dans la seconde liste, avec son rang, la valeur
    et l'onglet `Écoles`."""
    lignes, erreurs = _lire(tmp_path, [
        _session(ecole="Institut Fictif B", n_facture="R_ECO001"),
        _session(ecole=" abc ", n_facture="R_ABC001"),
        _session(ecole="École Fictive Inconnue", n_facture="R_XXX001"),
    ])
    assert [(l.rang, l.ecole) for l in lignes] == [(2, "ECO"), (3, "ABC")]
    ((rang, message),) = erreurs
    assert rang == 4
    assert "École Fictive Inconnue" in message and "Écoles" in message


# ──────────────────────────────────────────────────────────────────────
# Factures abandonnées et recherche
# ──────────────────────────────────────────────────────────────────────


def _facture(numero: str, statut) -> dict:
    return {"N° facture": numero, "École": "ABC", "Période": "Sept 2026", "Statut": statut}


def test_t_ses_15_numeros_abandonnes():
    """T-SES-15 — `numeros_abandonnes` : les N° au statut « abandonnée » de `Factures`, sans tenir
    compte de la casse ni des espaces ; aucun autre statut."""
    factures = [
        _facture("R_ABC001", "émise"), _facture("R_ABC002", "abandonnée"), _facture("R_ABC003", " Abandonnée "),
        _facture("R_ABC004", "payée"), _facture("R_ABC005", None), _facture("R_ABC006", "envoyée"),
    ]
    assert _ses().numeros_abandonnes(factures) == {"R_ABC002", "R_ABC003"}
    assert _ses().numeros_abandonnes([]) == set()


def _ligne(**valeurs):
    """Une `LigneSessions` construite à la main : ABC, 15/09/2026, 09:00–12:30, sauf surcharge."""
    champs = {
        "rang": 2, "ecole": "ABC", "n_facture": "R_ABC001", "periode": "Septembre 2026",
        "date_seance": date(2026, 9, 15), "debut": time(9, 0), "fin": time(12, 30), "duree_h": 3.5,
        "promotion": "M1 S1", "module": "Module fictif A", "taux": 45.0, "total_ht": Decimal("157.5"),
        "summary": "#ABC-M1 Cours fictif", "abandonnee": False,
    }
    champs.update(valeurs)
    return _ses().LigneSessions(**champs)


DEBUT = datetime(2026, 9, 15, 9, 0, tzinfo=PARIS)
FIN = datetime(2026, 9, 15, 12, 30, tzinfo=PARIS)


def test_t_ses_03_chercher_ignore_un_numero_abandonne(tmp_path):
    """T-SES-03 — `chercher` ignore la ligne d'un N° abandonné (dans `Factures`), exacte ou du même
    jour sans heure ; la même ligne, non abandonnée, est trouvée (garde cassée)."""
    ses = _ses()
    lignes, erreurs = _lire(
        tmp_path,
        [_session(n_facture="R_ABC001"), _session(n_facture="R_ABC002", debut=None, fin=None)],
        factures=[_facture("R_ABC001", "abandonnée"), _facture("R_ABC002", "Abandonnée")],
    )
    assert erreurs == []
    assert [l.abandonnee for l in lignes] == [True, True]
    assert ses.chercher("ABC", DEBUT, FIN, lignes) == ([], [])
    vivantes = [dataclasses.replace(l, abandonnee=False) for l in lignes]
    exactes, meme_jour = ses.chercher("ABC", DEBUT, FIN, vivantes)
    assert [l.n_facture for l in exactes] == ["R_ABC001"]
    assert [l.n_facture for l in meme_jour] == ["R_ABC002"]


def test_t_ses_16_chercher_exactes_puis_meme_jour_sans_heure():
    """T-SES-16 — `chercher` : lignes exactes (même école, date, début et fin à la minute), puis
    lignes du même jour sans heure ; une autre école, un autre jour ou d'autres heures ne sortent
    dans aucune liste."""
    ses = _ses()
    lignes = [
        _ligne(n_facture="EXACTE"),
        _ligne(n_facture="SANS-HEURE", debut=None, fin=None),
        _ligne(n_facture="AUTRES-HEURES", debut=time(14, 0), fin=time(17, 30)),
        _ligne(n_facture="FIN-DIFFERENTE", fin=time(12, 0)),
        _ligne(n_facture="AUTRE-ECOLE", ecole="ECO"),
        _ligne(n_facture="AUTRE-ECOLE-SANS-HEURE", ecole="ECO", debut=None, fin=None),
        _ligne(n_facture="AUTRE-JOUR", date_seance=date(2026, 9, 16)),
        _ligne(n_facture="AUTRE-JOUR-SANS-HEURE", date_seance=date(2026, 9, 16), debut=None, fin=None),
    ]
    exactes, meme_jour = ses.chercher("ABC", DEBUT, FIN, lignes)
    assert [l.n_facture for l in exactes] == ["EXACTE"]
    assert [l.n_facture for l in meme_jour] == ["SANS-HEURE"]


# ──────────────────────────────────────────────────────────────────────
# Revue de P-0b (T-SES-17, T-SES-18)
# ──────────────────────────────────────────────────────────────────────
#
# Constats des deux relecteurs, écrits rouges avant le correctif :
# - T-SES-17 : une ligne du même jour et de la même école qui n'a qu'une seule heure (Début « 9h »
#   et Fin vide, ou l'inverse ; le § 15 constate des heures vides dans le MASTER réel) ne sortait
#   dans aucune liste de `chercher`, ni exacte ni « sans heure » : la séance déjà facturée serait
#   reprise dans un nouveau brouillon (double facturation). Lecture retenue (`💡 [proposé — à
#   valider]`) : toute ligne du même jour, non exacte, dont une heure manque va dans la seconde
#   liste, avec les lignes sans heure, pour être montrée comme un doublon possible ; elle n'est
#   jamais « exacte ». Elle remplace la lecture d'origine (« sans heure » = début et fin vides).
# - T-SES-18 : `numeros_abandonnes` tenait compte de la casse du n° (« r_abc001 » saisi à la main
#   dans `Factures` ne désignait pas « R_ABC001 » de `Sessions`) et ne reconnaissait pas
#   « abandonnee » sans accent : les lignes du n° abandonné comptaient comme facturées, et la
#   séance n'aurait jamais été refacturée. Lecture retenue (`💡 [proposé — à valider]`) : le n° se
#   compare sans tenir compte de la casse ni des espaces ; le statut est reconnu sans accent.


@pytest.mark.parametrize(
    "debut, fin",
    [(time(9, 0), None), (None, time(12, 30)), (time(14, 0), None), (None, time(17, 0))],
    ids=["debut-seul-de-la-seance", "fin-seule-de-la-seance", "debut-seul-autre", "fin-seule-autre"],
)
def test_t_ses_17a_ligne_a_une_seule_heure_dans_la_seconde_liste(debut, fin):
    """T-SES-17 — une ligne du même jour et de la même école dont une seule heure est connue n'est
    jamais exacte, mais sort dans la seconde liste (doublon possible, 💡), qu'elle ait l'heure de la
    séance ou une autre : une séance déjà facturée ne disparaît pas de la recherche. Une autre
    école ou un autre jour ne sortent toujours pas."""
    ses = _ses()
    lignes = [
        _ligne(n_facture="UNE-HEURE", debut=debut, fin=fin),
        _ligne(n_facture="AUTRE-ECOLE", ecole="ECO", debut=debut, fin=fin),
        _ligne(n_facture="AUTRE-JOUR", date_seance=date(2026, 9, 16), debut=debut, fin=fin),
    ]
    exactes, meme_jour = ses.chercher("ABC", DEBUT, FIN, lignes)
    assert exactes == []
    assert [l.n_facture for l in meme_jour] == ["UNE-HEURE"]


def test_t_ses_17b_heure_vide_dans_le_master(tmp_path):
    """T-SES-17 — même constat sur un MASTER fictif, aux formes du § 15 : « 9h » et une Fin vide, puis
    un Début vide et « 12h30 » ; chaque ligne lue sort dans la seconde liste de `chercher`."""
    ses = _ses()
    lignes, erreurs = _lire(tmp_path, [
        _session(n_facture="R_ABC001", debut="9h", fin=""),
        _session(n_facture="R_ABC002", debut=None, fin="12h30"),
    ])
    assert erreurs == []
    assert [(l.debut, l.fin) for l in lignes] == [(time(9, 0), None), (None, time(12, 30))]
    exactes, meme_jour = ses.chercher("ABC", DEBUT, FIN, lignes)
    assert exactes == []
    assert [l.n_facture for l in meme_jour] == ["R_ABC001", "R_ABC002"]


@pytest.mark.parametrize(
    "n_sessions, n_factures, statut",
    [("R_ABC001", "r_abc001", "abandonnée"), ("R_ABC001", " R_abc001 ", "Abandonnée"),
     ("r_abc001", "R_ABC001", "abandonnée"), ("R_ABC001", "R_ABC001", "abandonnee"),
     ("R_ABC001", "R_ABC001", " ABANDONNEE ")],
    ids=["numero-minuscules", "numero-casse-espaces", "sessions-minuscules", "statut-sans-accent",
         "statut-majuscules-sans-accent"],
)
def test_t_ses_18_numero_abandonne_saisi_a_la_main(tmp_path, n_sessions, n_factures, statut):
    """T-SES-18 — un n° abandonné saisi à la main dans `Factures`, à la casse près, ou au statut écrit
    sans accent (💡), marque abandonnée la ligne de `Sessions` qui porte ce n° : `chercher` l'ignore,
    et la séance pourra être refacturée."""
    ses = _ses()
    lignes, erreurs = _lire(tmp_path, [_session(n_facture=n_sessions)], factures=[_facture(n_factures, statut)])
    assert erreurs == []
    assert [l.abandonnee for l in lignes] == [True]
    assert ses.chercher("ABC", DEBUT, FIN, lignes) == ([], [])
