"""Lecture de l'agenda iCal (F-10, ADR019 ; banque de tests : T-AGD-02 à T-AGD-11).

T-AGD-01 est déjà réservé par ADR019 (délai de mise à jour du flux iCal, contrôle manuel).

Tests d'acceptation écrits avant l'implémentation : chaque test importe `facturation.agenda`
lui-même, pour que l'absence du module fasse échouer les tests sans bloquer la collecte des
autres. Données fictives (`tests/fixtures/agenda_fictif.ics`). Aucun accès réseau : le
téléchargement est simulé, et toute connexion réelle est interdite pendant ces tests.

Contenu de la fixture, sur octobre et novembre 2026 (heure de Paris) :
- cours hebdomadaire du mardi 9 h – 12 h 30, du 06/10 au 24/11 (8 occurrences) ; exception le
  20/10 (EXDATE) ; occurrence du 03/11 déplacée au 05/11, 14 h – 17 h 30 (RECURRENCE-ID) ;
  date ajoutée le samedi 28/11, 10 h (RDATE) ;
- cours annulé le 14/10 (STATUS:CANCELLED) ; journée entière le 11/11 ;
- cours de 1 h à 4 h la nuit du changement d'heure (25/10) : 4 heures réelles ;
- cours noté en UTC (12/11, 23 h 30 Z), soit le 13/11 à 0 h 30 à Paris ;
- cours du soir le 30/11, 20 h – 22 h ; cours hors période le 15/12.
"""

from __future__ import annotations

import dataclasses
import importlib
import io
import socket
import urllib.error
import urllib.request
from datetime import date, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import pytest

from facturation import config

pytestmark = pytest.mark.unit

FIXTURE = Path(__file__).resolve().parents[1] / "fixtures" / "agenda_fictif.ics"
PARIS = ZoneInfo("Europe/Paris")
OCTOBRE, FIN_NOVEMBRE = date(2026, 10, 1), date(2026, 11, 30)
JETON = "private-0123456789abcdef"
URL_FICTIVE = f"https://calendar.example.org/calendar/ical/personne%40example.org/{JETON}/basic.ics"
FRAGMENTS_SECRETS = (URL_FICTIVE, JETON, "calendar.example.org")

ATTENDUES = [  # (date locale, début, fin) à Paris, triées par début
    (date(2026, 10, 6), "09:00", "12:30"),
    (date(2026, 10, 13), "09:00", "12:30"),
    (date(2026, 10, 25), "01:00", "04:00"),
    (date(2026, 10, 27), "09:00", "12:30"),
    (date(2026, 11, 5), "14:00", "17:30"),
    (date(2026, 11, 10), "09:00", "12:30"),
    (date(2026, 11, 13), "00:30", "02:30"),
    (date(2026, 11, 17), "09:00", "12:30"),
    (date(2026, 11, 24), "09:00", "12:30"),
    (date(2026, 11, 28), "10:00", "13:30"),
    (date(2026, 11, 30), "20:00", "22:00"),
]


def _agenda():
    """Module testé, importé dans le test : son absence fait échouer le test, pas la collecte."""
    return importlib.import_module("facturation.agenda")


def _lire(debut=OCTOBRE, fin=FIN_NOVEMBRE, **options):
    return _agenda().lire_ics(FIXTURE.read_bytes(), debut=debut, fin=fin, **options)


def _creneau(seance, zone=PARIS) -> tuple:
    return (seance.date, seance.debut.astimezone(zone).strftime("%H:%M"), seance.fin.astimezone(zone).strftime("%H:%M"))


def _par_uid(seances, uid: str) -> list:
    return [s for s in seances if s.uid == uid]


# ──────────────────────────────────────────────────────────────────────
# lire_ics
# ──────────────────────────────────────────────────────────────────────


def test_t_agd_02_seances_de_la_periode_triees():
    """T-AGD-02 — les séances d'octobre et novembre 2026, triées par début, exactement."""
    assert [_creneau(s) for s in _lire()] == ATTENDUES


def test_t_agd_03_recurrence_exception_deplacement_et_date_ajoutee():
    """T-AGD-03 — le cours hebdomadaire : sans le 20/10 (EXDATE), le 03/11 déplacé au 05/11 à 14 h,
    plus le 28/11 (RDATE)."""
    cours = _par_uid(_lire(), "cours-hebdo-0001@exemple.invalid")
    assert [s.date for s in cours] == [
        date(2026, 10, 6), date(2026, 10, 13), date(2026, 10, 27), date(2026, 11, 5),
        date(2026, 11, 10), date(2026, 11, 17), date(2026, 11, 24), date(2026, 11, 28),
    ]
    deplace = next(s for s in cours if s.date == date(2026, 11, 5))
    assert _creneau(deplace) == (date(2026, 11, 5), "14:00", "17:30")
    assert deplace.intitule == "#ABC-M1 Cours fictif - Module A (déplacé)"


def test_t_agd_04_annule_et_journee_entiere_exclus():
    """T-AGD-04 — un événement annulé et un événement « journée entière » ne sont pas des séances."""
    texte = FIXTURE.read_text(encoding="utf-8")
    exclus = ("cours-annule-0002@exemple.invalid", "journee-entiere-0004@exemple.invalid")
    assert all(uid in texte for uid in exclus)  # ils sont bien dans la fixture : la garde ne passe pas à vide
    seances = _lire()
    assert [s for s in seances if s.uid in exclus] == []
    assert not {date(2026, 10, 14), date(2026, 11, 11)} & {s.date for s in seances}


def test_t_agd_05_duree_reelle_le_jour_du_changement_d_heure():
    """T-AGD-05 — 1 h à 4 h la nuit du 25/10/2026 : 4 heures réelles, pas 3 ; les autres durées sont exactes."""
    seances = _lire()
    (nuit,) = _par_uid(seances, "cours-changement-heure-0003@exemple.invalid")
    assert nuit.duree_h == pytest.approx(4.0)
    assert nuit.debut.astimezone(PARIS).utcoffset() == timedelta(hours=2)
    assert nuit.fin.astimezone(PARIS).utcoffset() == timedelta(hours=1)
    durees = {s.date: s.duree_h for s in seances}
    assert durees[date(2026, 10, 27)] == pytest.approx(3.5)  # 9 h – 12 h 30, après le changement d'heure
    assert durees[date(2026, 11, 13)] == pytest.approx(2.0)
    assert all(isinstance(s.duree_h, float) for s in seances)


def test_t_agd_06_dates_locales_et_heures_avec_fuseau():
    """T-AGD-06 — début et fin portent un fuseau ; la date est la date locale du fuseau demandé."""
    seances = _lire()
    assert all(s.debut.utcoffset() is not None and s.fin.utcoffset() is not None for s in seances)
    (utc,) = _par_uid(seances, "cours-utc-0006@exemple.invalid")
    assert _creneau(utc) == (date(2026, 11, 13), "00:30", "02:30")  # 23 h 30 Z le 12/11
    (en_utc,) = _par_uid(_lire(fuseau="UTC"), "cours-utc-0006@exemple.invalid")
    assert _creneau(en_utc, ZoneInfo("UTC")) == (date(2026, 11, 12), "23:30", "01:30")


def test_t_agd_07_bornes_incluses():
    """T-AGD-07 — les deux bornes sont des dates incluses : le dernier jour compte jusqu'au soir."""
    assert [_creneau(s) for s in _lire(date(2026, 11, 24), date(2026, 11, 24))] == [
        (date(2026, 11, 24), "09:00", "12:30")
    ]
    assert [_creneau(s) for s in _lire(date(2026, 11, 30), date(2026, 11, 30))] == [
        (date(2026, 11, 30), "20:00", "22:00")
    ]
    assert [_creneau(s) for s in _lire(OCTOBRE, date(2026, 11, 29))] == ATTENDUES[:-1]
    assert [_creneau(s) for s in _lire(date(2026, 12, 1), date(2026, 12, 31))] == [
        (date(2026, 12, 15), "09:00", "12:00")
    ]


def test_t_agd_08_seance_immuable_et_ses_champs():
    """T-AGD-08 — `Seance` est une dataclass figée : uid, date, debut, fin, duree_h, intitule, description."""
    agenda = _agenda()
    premiere = _lire()[0]
    assert isinstance(premiere, agenda.Seance)
    assert [champ.name for champ in dataclasses.fields(agenda.Seance)] == [
        "uid", "date", "debut", "fin", "duree_h", "intitule", "description",
    ]
    assert (premiere.uid, premiere.intitule, premiere.description) == (
        "cours-hebdo-0001@exemple.invalid", "#ABC-M1 Cours fictif - Module A", "Séance fictive de test",
    )
    with pytest.raises(dataclasses.FrozenInstanceError):
        premiere.intitule = "modifié"


def test_t_agd_09_contenu_texte_ou_octets():
    """T-AGD-09 — le contenu du .ics peut être donné en `str` comme en `bytes`."""
    agenda = _agenda()
    en_texte = agenda.lire_ics(FIXTURE.read_text(encoding="utf-8"), debut=OCTOBRE, fin=FIN_NOVEMBRE)
    assert [(s.uid, _creneau(s)) for s in en_texte] == [(s.uid, _creneau(s)) for s in _lire()]


# ──────────────────────────────────────────────────────────────────────
# telecharger_ics — téléchargement simulé, adresse jamais affichée
# ──────────────────────────────────────────────────────────────────────


@pytest.fixture
def sans_reseau(monkeypatch):
    """Interdit toute connexion réelle pendant le test."""
    def interdit(*args, **kwargs):
        raise AssertionError("accès réseau interdit dans les tests")

    monkeypatch.setattr(socket.socket, "connect", interdit)
    monkeypatch.setattr(socket, "create_connection", interdit)
    monkeypatch.setattr(socket, "getaddrinfo", interdit)


class _Reponse(io.BytesIO):
    status = 200

    def getcode(self) -> int:
        return 200


def _simuler(monkeypatch, agenda, reponse) -> list:
    """Remplace `urllib.request.urlopen` : renvoie `reponse` (octets) ou la lève (exception).

    Renvoie la liste des appels (adresse, délai d'attente).
    """
    appels = []

    def faux_urlopen(url, data=None, timeout=None, *args, **kwargs):
        appels.append((getattr(url, "full_url", url), timeout))
        if isinstance(reponse, BaseException):
            raise reponse
        return _Reponse(reponse)

    monkeypatch.setattr(urllib.request, "urlopen", faux_urlopen)
    if hasattr(agenda, "urlopen"):  # `from urllib.request import urlopen` dans le module
        monkeypatch.setattr(agenda, "urlopen", faux_urlopen)
    return appels


def _profil(tmp_path: Path, contenu_secret: str | None) -> config.Profil:
    """Profil fictif dont le fichier `agenda.secret` contient `contenu_secret` (absent si None)."""
    secret = tmp_path / "agenda.secret"
    if contenu_secret is not None:
        secret.write_text(contenu_secret, encoding="utf-8")
    return config.Profil(
        chemin=tmp_path / "profil.toml", nom="Personne A", entite="ROSARI", prefixe_facture="R_",
        agenda_source="ical", agenda_fichier_url_secrete=secret, fuseau="Europe/Paris",
        racine=tmp_path / "donnees", lancement_jour=25, lancement_heure="16:00",
        date_emission="premier_du_mois_suivant", pdf_moteur="libreoffice",
    )


def _messages_affiches(exc: BaseException) -> list[str]:
    """Textes qu'une trace d'erreur afficherait : l'exception et les causes non masquées."""
    textes = []
    while exc is not None:
        textes += [str(exc), repr(exc)]
        exc = exc.__cause__ if exc.__cause__ is not None else (None if exc.__suppress_context__ else exc.__context__)
    return textes


def test_t_agd_10_telechargement_simule(tmp_path, monkeypatch, sans_reseau):
    """T-AGD-10 — l'adresse est lue dans `agenda.secret` et téléchargée avec un délai d'attente."""
    agenda = _agenda()
    appels = _simuler(monkeypatch, agenda, FIXTURE.read_bytes())
    contenu = agenda.telecharger_ics(_profil(tmp_path, URL_FICTIVE + "\n"))
    assert contenu == FIXTURE.read_bytes()
    assert len(appels) == 1
    adresse, delai = appels[0]
    assert adresse == URL_FICTIVE
    assert delai is not None and 0 < delai <= 120, f"délai d'attente : {delai!r}"


@pytest.mark.parametrize(
    ("contenu_secret", "reponse"),
    [
        (None, b""),
        ("http://calendar.example.org/" + JETON + "/basic", b""),
        (URL_FICTIVE, urllib.error.URLError(f"échec de connexion à {URL_FICTIVE}")),
        (URL_FICTIVE, TimeoutError(f"délai dépassé pour {URL_FICTIVE}")),
        (URL_FICTIVE, urllib.error.HTTPError(URL_FICTIVE, 404, "Not Found", None, None)),
    ],
    ids=["fichier-absent", "adresse-invalide", "injoignable", "delai-depasse", "http-404"],
)
def test_t_agd_11_erreur_claire_sans_jamais_afficher_l_adresse(tmp_path, monkeypatch, capsys, caplog, sans_reseau,
                                                                contenu_secret, reponse):
    """T-AGD-11 — adresse absente, invalide ou injoignable : une erreur claire, et l'adresse n'apparaît
    ni dans le message, ni dans les causes affichées, ni à l'écran, ni dans le journal."""
    agenda = _agenda()
    _simuler(monkeypatch, agenda, reponse)
    with pytest.raises(Exception) as exc:
        agenda.telecharger_ics(_profil(tmp_path, contenu_secret))
    erreurs_de_programmation = (AttributeError, TypeError, NameError, ImportError, KeyError, IndexError, AssertionError)
    assert not isinstance(exc.value, erreurs_de_programmation), repr(exc.value)
    if contenu_secret is None:
        assert "agenda.secret" in str(exc.value)
    sortie = capsys.readouterr()
    affiche = [*_messages_affiches(exc.value), sortie.out, sortie.err, caplog.text]
    fuites = [f for f in FRAGMENTS_SECRETS for texte in affiche if f in texte]
    assert not fuites, f"adresse affichée : {sorted(set(fuites))}"
