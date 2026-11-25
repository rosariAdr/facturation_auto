"""agenda.py — Lecture de l'agenda des cours par son adresse iCal (F-10, ADR019).

Deux étapes, séparées pour que la lecture se teste sans réseau :

- `telecharger_ics(profil)` lit l'adresse iCal secrète dans `agenda.secret` et télécharge
  le fichier .ics. L'adresse est un secret : elle n'apparaît dans aucun message, aucune
  cause d'erreur, aucun journal ;
- `lire_ics(contenu, debut=…, fin=…)` transforme ce contenu en séances : récurrences
  développées (RRULE, RDATE), exceptions retirées (EXDATE), occurrences déplacées prises
  à leur nouvelle date (RECURRENCE-ID), événements annulés et « journée entière » exclus.

Les récurrences sont développées par `recurring-ical-events`, au-dessus d'`icalendar`.
"""

from __future__ import annotations

import contextlib
import http.client
import urllib.error
import urllib.request
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from typing import Union
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import icalendar
import recurring_ical_events

from facturation import __version__, config

FUSEAU_PAR_DEFAUT = "Europe/Paris"
DELAI_S = 30  # délai d'attente du téléchargement, en secondes

# Marge autour de la période demandée : un événement noté dans un autre fuseau peut
# tomber, à l'heure locale, la veille ou le lendemain de sa date d'origine.
_MARGE = timedelta(days=2)
_MESSAGE_ADRESSE = (
    "copiez « Adresse secrète au format iCal » dans les paramètres de l'agenda de cours, "
    "puis lancez `facturation profil agenda`"
)


class AgendaIndisponible(RuntimeError):
    """Agenda impossible à télécharger. Le message dit quoi faire, sans jamais citer l'adresse."""


@dataclass(frozen=True)
class Seance:
    """Une séance de l'agenda.

    `date` est la date locale du début, dans le fuseau demandé ; `debut` et `fin` portent
    ce fuseau ; `duree_h` est la durée réelle en heures décimales (4.0 pour 1 h – 4 h la
    nuit du passage à l'heure d'hiver).
    """

    uid: str
    date: date
    debut: datetime
    fin: datetime
    duree_h: float
    intitule: str
    description: str


# ──────────────────────────────────────────────────────────────────────
# Lecture d'un fichier .ics
# ──────────────────────────────────────────────────────────────────────


def lire_ics(
    contenu: Union[str, bytes],
    *,
    debut: date,
    fin: date,
    fuseau: str = FUSEAU_PAR_DEFAUT,
) -> list[Seance]:
    """Séances dont le début tombe entre `debut` et `fin` (dates locales incluses), triées par début.

    Args:
        contenu: texte ou octets d'un fichier .ics.
        debut, fin: bornes de la période, incluses (le dernier jour compte jusqu'à minuit).
        fuseau: fuseau des dates locales et des heures rendues (ex. Europe/Paris).

    Raises:
        ValueError: fuseau inconnu, période inversée ou contenu iCal illisible.
    """
    zone = _zone(fuseau)
    debut, fin = _jour(debut), _jour(fin)
    if debut > fin:
        raise ValueError(f"période inversée : du {debut:%d/%m/%Y} au {fin:%d/%m/%Y}")

    calendrier = _calendrier(contenu)
    try:
        occurrences = recurring_ical_events.of(calendrier).between(debut - _MARGE, fin + _MARGE)
    except Exception as exc:  # règle de récurrence mal formée, date invalide…
        raise ValueError(f"agenda illisible : récurrence ou date invalide ({type(exc).__name__})") from exc

    seances = []
    for evenement in occurrences:
        seance = _seance(evenement, zone)
        if seance is not None and debut <= seance.date <= fin:
            seances.append(seance)
    seances.sort(key=lambda s: (s.debut, s.fin, s.uid))
    return seances


def _zone(fuseau: str) -> ZoneInfo:
    try:
        return ZoneInfo(fuseau)
    except (ZoneInfoNotFoundError, ValueError):
        raise ValueError(f"fuseau inconnu : {fuseau!r} (ex. Europe/Paris)") from None


def _jour(valeur: date) -> date:
    """Une date, même si l'appelant passe un datetime."""
    return valeur.date() if isinstance(valeur, datetime) else valeur


def _calendrier(contenu: Union[str, bytes]) -> icalendar.Calendar:
    if isinstance(contenu, bytes):
        try:
            contenu = contenu.decode("utf-8-sig")
        except UnicodeDecodeError:
            raise ValueError("agenda illisible : le fichier .ics n'est pas encodé en UTF-8") from None
    contenu = contenu.lstrip("\ufeff")
    try:
        calendrier = icalendar.Calendar.from_ical(contenu)
    except Exception as exc:
        raise ValueError(f"agenda illisible : ce n'est pas un fichier iCal valide ({type(exc).__name__})") from exc
    if getattr(calendrier, "name", None) != "VCALENDAR":
        raise ValueError("agenda illisible : aucun bloc BEGIN:VCALENDAR")
    return calendrier


def _seance(evenement, zone: ZoneInfo) -> Seance | None:
    """La séance d'une occurrence, ou None si ce n'est pas une séance (annulée, journée entière)."""
    if str(evenement.get("STATUS", "")).strip().upper() == "CANCELLED":
        return None
    propriete_debut = evenement.get("DTSTART")
    if propriete_debut is None or not isinstance(propriete_debut.dt, datetime):
        return None  # sans heure : événement « journée entière »
    debut = _heure_locale(propriete_debut.dt, zone)

    propriete_fin = evenement.get("DTEND")
    if propriete_fin is not None and isinstance(propriete_fin.dt, datetime):
        fin = _heure_locale(propriete_fin.dt, zone)
    elif evenement.get("DURATION") is not None:
        fin = (debut.astimezone(ZoneInfo("UTC")) + evenement.get("DURATION").dt).astimezone(zone)
    else:
        fin = debut  # RFC 5545 : sans fin ni durée, un événement daté à l'heure dure zéro

    # Durée réelle : calculée sur les instants, pas sur l'heure affichée (changement d'heure).
    duree_h = (fin.timestamp() - debut.timestamp()) / 3600
    return Seance(
        uid=str(evenement.get("UID", "")).strip(),
        date=debut.date(),
        debut=debut,
        fin=fin,
        duree_h=float(duree_h),
        intitule=str(evenement.get("SUMMARY", "")).strip(),
        description=str(evenement.get("DESCRIPTION", "")).strip(),
    )


def _heure_locale(moment: datetime, zone: ZoneInfo) -> datetime:
    """Le moment dans le fuseau demandé ; une heure « flottante » (sans fuseau) y est rattachée."""
    if moment.tzinfo is None or moment.utcoffset() is None:
        return moment.replace(tzinfo=zone)
    return moment.astimezone(zone)


# ──────────────────────────────────────────────────────────────────────
# Téléchargement — l'adresse n'est jamais affichée
# ──────────────────────────────────────────────────────────────────────


def telecharger_ics(profil: config.Profil) -> bytes:
    """Télécharge l'agenda du profil, à partir de l'adresse rangée dans `agenda.secret`.

    Raises:
        AgendaIndisponible: adresse absente ou invalide, agenda injoignable, délai dépassé,
            refus du serveur ou réponse qui n'est pas un agenda. Le message ne cite jamais
            l'adresse, et l'erreur d'origine (qui pourrait la contenir) n'est pas enchaînée.
    """
    adresse = _adresse_ical(profil)
    requete = urllib.request.Request(adresse, headers={"User-Agent": f"facturation-auto/{__version__}"})
    try:
        with urllib.request.urlopen(requete, timeout=DELAI_S) as reponse:
            contenu = reponse.read()
    except urllib.error.HTTPError as exc:
        code = exc.code
        with contextlib.suppress(Exception):
            exc.close()  # libère la connexion portée par la réponse d'erreur
        if code in (401, 403, 404, 410):
            raise AgendaIndisponible(
                f"l'agenda a refusé l'adresse iCal (HTTP {code}) : elle a sans doute été réinitialisée ; "
                f"{_MESSAGE_ADRESSE} (adresse non affichée)"
            ) from None
        raise AgendaIndisponible(
            f"le serveur de l'agenda a répondu HTTP {code} : réessayez dans quelques minutes"
        ) from None
    except TimeoutError:
        raise _delai_depasse() from None
    except urllib.error.URLError as exc:
        if isinstance(exc.reason, TimeoutError):
            raise _delai_depasse() from None
        raise AgendaIndisponible(
            f"agenda injoignable{_nature(exc.reason)} : vérifiez la connexion Internet "
            "(ou le proxy), puis réessayez"
        ) from None
    except (OSError, http.client.HTTPException) as exc:
        raise AgendaIndisponible(
            f"téléchargement de l'agenda interrompu{_nature(exc)} : réessayez"
        ) from None
    except ValueError:
        raise AgendaIndisponible(
            f"l'adresse iCal rangée dans {_nom_secret(profil)} est inutilisable : {_MESSAGE_ADRESSE} "
            "(adresse non affichée)"
        ) from None

    if not contenu.lstrip(b"\xef\xbb\xbf \t\r\n").startswith(b"BEGIN:VCALENDAR"):
        raise AgendaIndisponible(
            "la réponse de l'agenda n'est pas un fichier iCal : vérifiez l'adresse ; "
            f"{_MESSAGE_ADRESSE} (adresse non affichée)"
        )
    return contenu


def _adresse_ical(profil: config.Profil) -> str:
    """L'adresse rangée dans le fichier secret du profil, contrôlée mais jamais affichée."""
    if profil.agenda_source != "ical":
        raise AgendaIndisponible(
            f"la source d'agenda « {profil.agenda_source} » n'est pas encore disponible : "
            'mettez source = "ical" dans la section [agenda] du profil'
        )
    fichier = profil.agenda_fichier_url_secrete
    if fichier is None:
        raise AgendaIndisponible(
            f"aucune adresse iCal déclarée dans le profil ({config.NOM_FICHIER_ICAL} absent) : "
            f"{_MESSAGE_ADRESSE}"
        )
    # Seul le nom du fichier s'affiche : le chemin complet pourrait contenir l'adresse (FA-10).
    if not fichier.is_file():
        raise AgendaIndisponible(f"fichier d'adresse iCal introuvable : {fichier.name} ; {_MESSAGE_ADRESSE}")
    try:
        # utf-8-sig : un fichier retouché au Bloc-notes ou sous PowerShell 5 peut commencer par un BOM.
        adresse = fichier.read_text(encoding="utf-8-sig").strip()
    except (OSError, UnicodeDecodeError):
        raise AgendaIndisponible(
            f"impossible de lire {fichier.name} : vérifiez le fichier, ou {_MESSAGE_ADRESSE}"
        ) from None
    if not config.adresse_ical_valide(adresse):
        raise AgendaIndisponible(
            f"le contenu de {fichier.name} n'est pas une adresse iCal (https://….ics) : "
            f"{_MESSAGE_ADRESSE} (adresse non affichée)"
        )
    return adresse


def _nom_secret(profil: config.Profil) -> str:
    fichier = profil.agenda_fichier_url_secrete
    return fichier.name if fichier is not None else config.NOM_FICHIER_ICAL


def _delai_depasse() -> AgendaIndisponible:
    return AgendaIndisponible(
        f"l'agenda n'a pas répondu en {DELAI_S} s : vérifiez la connexion Internet, puis réessayez"
    )


def _nature(cause: object) -> str:
    """Nature de l'erreur réseau (nom de sa classe), sans son texte, qui peut citer l'adresse."""
    return f" ({type(cause).__name__})" if isinstance(cause, BaseException) else ""
