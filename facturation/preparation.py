"""preparation.py — Préparation multi-écoles des brouillons de facture (F-16a, PROC00 étapes 1 à 4).

Trois temps, séparés pour que la logique se teste sans disque ni réseau :

- `charger_entrees(profil, mois)` lit, sans rien écrire ni verrouiller : le MASTER (règles,
  correspondances, tarifs, écoles, `Sessions`, `Factures`), les coordonnées (`pcf_format`), l'agenda
  de la fenêtre de lecture, les décisions actives du mois, les pièces d'émission et les numéros ;
- `preparer(...)` est pure et déterministe : un brouillon par école du lot (ou par contrat), des
  blocages cumulés sans rien lever, les créneaux non reconnus jamais facturés (ADR028),
  l'anti-doublon séance par séance contre `Sessions` (ADR005), les brouillons engagés (pièce
  d'émission présente) jamais régénérés. Une erreur propre à une école devient un blocage et ne
  retient jamais les autres (ADR031) ;
- `lancer_preparation(profil, mois)` fait le tout sous le verrou d'entité et écrit, de façon
  atomique, sous `_sorties/<AAAA-MM>/` seulement : `preparation.json` (l'ancien dans `historique/`),
  `apercu.txt`, `apercus/<identifiant>.txt`, `a_classer.txt` et l'événement `preparation` du
  journal du mois. Aucune écriture dans le MASTER.

Montants exacts en `Decimal` (`heures.montant`), total arrondi une seule fois au centime
(`heures.arrondi_centime`) : il égale le total de `facture.apercu` et le `=SUM` du gabarit.
L'empreinte d'un brouillon (`empreinte_brouillon`) couvre tout ce qui sera écrit sur la facture,
jamais la marque `tenue`, le numéro probable, les avertissements, les blocages ni la date de
préparation : une validation porte sur une empreinte (§ 1.3).

Contrat : contrats des vagues 3 à 5, § 1 et 4, avec les § 15 et 16, qui priment (une ligne de
`Règles_facturation` par contrat ; le n° imprimé vient de cette ligne, `Tarifs.Réf. contrat` y relie
la ligne de tarif).

Lectures retenues (`💡 [proposé — à valider]`), là où le contrat se tait :
- `EntreesPreparation.regles` est le dictionnaire de `regles.regles_par_ecole`, clé (école, n° de
  contrat) ; une école à plusieurs lignes de règles est facturée par contrat, et bloquée si l'une
  de ses lignes ne dit pas « Une facture par contrat = oui » ; une `Réf. contrat` de `Tarifs` sans
  ligne de règles est un blocage ;
- suffixe d'un identifiant par contrat : la référence en majuscules, réduite à [A-Z0-9] et à 20
  caractères ; deux références réduites au même suffixe reçoivent chacune « - » et les 4 premiers
  caractères hexadécimaux du sha256 (UTF-8) de leur référence ;
- sans « Une facture par contrat », le contrat commun est la seule `Réf. contrat` renseignée sur
  les lignes de tarif (les lignes sans référence ne comptent pas) ;
- un brouillon engagé (pièce « en émission » ou « émise ») n'est pas dans `brouillons` : il est
  listé dans `deja_emises`, comparé au brouillon recalculé sans les lignes `Sessions` de son propre
  N° ; une pièce abandonnée n'engage rien et son N° n'est jamais réutilisé ;
- une décision `inclure` sur une séance ne réintègre que le cas « même jour sans heure » ;
- « période déjà facturée hors outil » : `Factures.Période` lue sous le libellé de l'onglet
  (« Déc 2026 ») ou sous le nom du mois en entier (« Décembre 2026 », § 15), sans tenir compte de
  la casse ni des accents ; une ligne « abandonnée » ne compte pas ;
- une séance et une ligne `Sessions` de la même école, le même jour, à des heures qui se
  chevauchent sans être identiques : avertissement « possible doublon avec Sessions, ligne N »,
  repris dans « À FAIRE », sans exclusion ni blocage ;
- `ecole_label` vaut le code quand l'école n'a pas de libellé dans `Écoles` (avertissement) : la
  spec n'a jamais de libellé vide (§ 18) ;
- numéros courts : « F<n> » dans l'ordre des brouillons, « S<n> » dans l'ordre chronologique des
  séances des brouillons, des séances déjà facturées et des créneaux à classer ; `apercu_facture`
  et `apercu_global` les reçoivent par `numeros=` (sans eux, une facture seule est numérotée à part) ;
- « N créneau(x) à classer dans la période » : les créneaux dont la date tombe dans la période du
  brouillon ;
- `Preparation.precedente_le` (champ ajouté, facultatif) date la préparation précédente, pour
  « CHANGEMENTS DEPUIS LE … » ; `EntreesPreparation.sessions_illisibles` (ajouté, facultatif)
  porte les lignes de `Sessions` qui ne se lisent pas ;
- un `preparation.json` illisible (abîmé, d'un autre format ou d'une autre entité) n'empêche pas
  de préparer : il est rangé dans `historique/` et la préparation repart de zéro ; un refus du
  verrou (autre poste) ne s'écrit pas au journal, que l'outil ne tient pas.

Constats de revue de F-16a (`💡 [proposé — à valider]`) :
- les séances d'une pièce engagée ne repartent jamais dans un autre brouillon, même si son
  identifiant a changé depuis (règles modifiées) : elles sont retirées, avec un avertissement ; une
  pièce dont l'identifiant n'est plus celui d'aucun groupe bloque les brouillons de son école (pas
  de facture complémentaire automatique, Q-D3) ;
- chaque ligne de contrat porte son « Statut lot » (§ 16) : seules les lignes du lot ont une
  période ; la séance d'un contrat « Plus tard » ou « Hors facturation » va dans `hors_lot`, et la
  ligne hors lot en « fin de module » ne retient pas l'école ;
- une valeur vide d'une mention obligatoire (SIRET, IBAN, BIC ; adresse L1 et L2, email et
  téléphone de l'émetteur, ADR030 §1) bloque tous les brouillons ;
- une ligne de `Sessions` illisible, ou sans école ou sans date, est un avertissement « ligne de
  Sessions à corriger » des brouillons qu'elle peut concerner, repris dans « À FAIRE » ;
- sans « Une facture par contrat », une « Réf. contrat » de Tarifs différente du « N° de
  contrat » de la règle est un blocage ; une clé tarif qui désigne la ligne d'une autre école est un
  blocage (jamais le taux d'une autre école) ; 0 h à facturer est un blocage (heures non
  calculables) ;
- la nature « contenu modifié » s'ajoute à celles du contrat : elle nomme chaque champ de
  l'empreinte que les autres natures ne disent pas (intitulé, libellé de la ligne ou de l'école,
  bon de commande…), et un filet garantit qu'une empreinte changée a toujours sa différence ;
- `Factures.Période` saisie en date (Excel convertit « Déc 2026 » en 01/12/2026) vaut le mois de
  la date ; une ligne de `Factures` de la même période dont le contrat n'est renseigné que d'un côté
  donne l'avertissement « période peut-être déjà facturée hors outil », repris dans « À FAIRE » ;
- `Preparation.exclues` (champ ajouté, facultatif) garde les séances exclues par décision, reconnues
  ou non, montrées dans l'aperçu global ; une décision `classer` inapplicable (école inconnue)
  renvoie sa seule séance dans la file à classer, avec la raison ;
- `lancer_preparation(…, ecoles=…)` ne recalcule que ces écoles : les autres gardent leur état de
  la préparation précédente ; un échec d'écriture ou du journal remet chaque fichier du mois tel
  qu'il était ; une cellule que JSON ne porte pas (durée) est écrite en texte.
"""

from __future__ import annotations

import hashlib
import importlib
import math
import os
import re
import unicodedata
from collections import Counter
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass, field, replace
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import TYPE_CHECKING, Any
from zoneinfo import ZoneInfo

from facturation import __version__, agenda, chemins, facture, fichiers, journal, master
from facturation import pcf_format as fmt
from facturation import regles as regles_facturation
from facturation.agenda import AgendaIndisponible, Seance
from facturation.classement import SeanceAClasser, SeanceClassee, classer
from facturation.decisions import BON_COMMANDE, CLASSER, EXCLURE, INCLURE, Decision, decisions_actives
from facturation.heures import arrondi_centime, heures_facturables, ligne_tarif, taux_horaire
from facturation.heures import montant as montant_exact
from facturation.journal import PieceEmission
from facturation.master import Master, MasterInvalide, MasterVerrouille
from facturation.numerotation import lire_numeros, numeros_manquants, prochain_numero
from facturation.regles import RegleEcole, ecoles_du_lot, echeance_selon_regle, periode_a_facturer, regles_par_ecole
from facturation.sessions import LigneSessions, chercher, code_ecole, lire_sessions, numeros_abandonnes

if TYPE_CHECKING:
    from facturation.config import Profil

FORMAT_PREPARATION = 1
LOT_PAR_DEFAUT = "Lot 1"

PRET, BLOQUE = "prêt", "bloqué"
EXACTE, MEME_JOUR = "exacte", "même jour sans heure"
PAR_DECISION = "décision"
SANS_CONTRAT = "SANS-CONTRAT"

NOM_PREPARATION = "preparation.json"
NOM_APERCU = "apercu.txt"
NOM_A_CLASSER = "a_classer.txt"
DOSSIER_APERCUS = "apercus"
DOSSIER_HISTORIQUE = "historique"

NATURES_DIFFERENCES = (
    "facture nouvelle", "facture disparue", "séance ajoutée", "séance retirée", "séance déplacée",
    "heures modifiées", "taux modifié", "coordonnées modifiées", "échéance modifiée",
    "contrat modifié", "contenu modifié", "nouveau blocage", "blocage levé", "total modifié",
)
# Débuts d'avertissement repris dans « À FAIRE » (une action d'Adrian est attendue).
POSSIBLE_DOUBLON = "possible doublon avec Sessions"
SESSIONS_A_CORRIGER = "ligne de Sessions à corriger"
PERIODE_PEUT_ETRE_FACTUREE = "période peut-être déjà facturée hors outil"
# Modules qui peuvent ajouter une section à l'aperçu global (`section_apercu`), dans cet ordre.
MODULES_ANNEXES = (
    ("facturation.registre", "REGISTRE DES FACTURES"),
    ("facturation.apres_emission", "APRÈS ÉMISSION"),
)
# États qui demandent une action d'Adrian (code 3 de `facturation preparer`).
ETATS_A_TRAITER = (journal.A_VALIDER, journal.A_REVALIDER, journal.A_CORRIGER, journal.BLOQUEE)

_JOURS = ("lun.", "mar.", "mer.", "jeu.", "ven.", "sam.", "dim.")
# Mentions obligatoires de l'émetteur (ADR030 §1) : clé de `pcf_format`, colonne de `Coordonnées`.
_CHAMPS_EMETTEUR = (("adresse_l1", "Adresse L1"), ("adresse_l2", "Adresse L2"), ("email", "Email"),
                    ("tel", "Téléphone"))
_INFOS_OBLIGATOIRES = ("SIRET", "IBAN", "BIC")
_MINUTE = "%Y-%m-%dT%H:%M%z"
_VALEUR_DECISION_RE = re.compile(r"[0-9]+(?:[.,][0-9]+)?")


class PreparationImpossible(RuntimeError):
    """La préparation n'a pas pu se faire ; l'exception d'origine est enchaînée (`__cause__`)."""


class PreparationIllisible(ValueError):
    """`preparation.json` absent là où il est exigé, abîmé, d'un format inconnu ou d'une autre entité."""


# ──────────────────────────────────────────────────────────────────────
# Dataclasses
# ──────────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class LigneBrouillon:
    """Une séance facturée par un brouillon."""

    cle_seance: str
    uid: str
    date: date
    debut: datetime
    fin: datetime
    intitule: str                    # SUMMARY verbatim → Sessions.« Summary calendrier verbatim »
    ecole: str
    promotion: str
    module: str
    libelle: str
    cle_tarif: str
    classee_par: str                 # "code" | "exact" | "préfixe" | "décision"
    duree_agenda_h: float
    heures: float | None             # None : heures non calculables (blocage)
    taux: float | None               # None : taux introuvable (blocage) ; jamais inventé
    montant: Decimal | None          # heures × taux, exact
    ref_contrat: str | None          # Tarifs.« Réf. contrat »
    tenue: bool                      # fin <= maintenant ; hors empreinte


@dataclass(frozen=True)
class Brouillon:
    """Brouillon d'une facture : tout ce qui sera écrit, et ce qui empêche de l'émettre."""

    identifiant: str
    entite: str
    ecole: str
    mois: str
    periode: tuple[date, date]
    periode_label: str
    ecole_label: str
    contrat: str | None
    bon_commande: str | None
    destinataire: dict[str, str] | None
    emetteur: dict[str, str] | None
    info: dict[str, str]
    nda: str | None
    date_emission: date
    echeance: date | None
    masquer_jour: bool
    regroupement: str
    lignes: tuple[LigneBrouillon, ...]
    total_heures: float | None
    total_ht: Decimal | None
    numero_probable: str | None
    blocages: tuple[str, ...]
    avertissements: tuple[str, ...]
    empreinte: str

    @property
    def statut(self) -> str:
        """« bloqué » s'il reste un blocage, sinon « prêt »."""
        return BLOQUE if self.blocages else PRET

    def groupes(self) -> list[tuple[str, str, list[tuple[date, float, float]]]]:
        """Lignes de la facture : (promotion, libellé, [(date, heures, taux)]).

        Groupes dans l'ordre (promotion, libellé), séances par début. En regroupement « jour », les
        heures d'une même (date, promotion, libellé, taux) s'additionnent en `Decimal`, puis sont
        rendues en `float` ; les lignes `Sessions` restent une par séance. Une heure ou un taux
        inconnu reste None (brouillon bloqué).
        """
        par_groupe: dict[tuple[str, str], list[LigneBrouillon]] = {}
        for ligne in sorted(self.lignes, key=lambda l: (l.promotion, l.libelle, l.debut, l.cle_seance)):
            par_groupe.setdefault((ligne.promotion, ligne.libelle), []).append(ligne)
        resultat = []
        for (promotion, libelle), lignes in par_groupe.items():
            if self.regroupement != "jour":
                resultat.append((promotion, libelle, [(l.date, l.heures, l.taux) for l in lignes]))
                continue
            jours: dict[tuple[date, Any], list[float | None]] = {}
            for l in lignes:
                jours.setdefault((l.date, l.taux), []).append(l.heures)
            seances = []
            for (jour, taux), heures in jours.items():
                total = None if any(h is None for h in heures) else float(sum(Decimal(repr(h)) for h in heures))
                seances.append((jour, total, taux))
            resultat.append((promotion, libelle, seances))
        return resultat

    def spec(self, n_facture: str) -> dict[str, Any]:
        """Spec de `facturation.facture` pour le N° donné (dates en ISO, sérialisable en JSON).

        Raises:
            ValueError: brouillon bloqué (le message cite le premier blocage), ou nom d'onglet
                invalide (`chemins.onglet_facture`).
        """
        if self.statut != PRET:
            raise ValueError(
                f"Brouillon {self.identifiant} bloqué : {self.blocages[0]} — aucune spec tant qu'il reste "
                f"{len(self.blocages)} blocage(s)."
            )
        entite, ecole = self.entite, self.ecole
        return {
            "entite": entite,
            "ecole": ecole,
            "ecole_label": self.ecole_label or ecole,
            "fichier_sortie": f"{entite}/{ecole}/{entite}_{ecole}.xlsx",
            "onglet": chemins.onglet_facture(n_facture, self.periode_label),
            "n_facture": n_facture,
            "date_emission": self.date_emission.isoformat(),
            "echeance": self.echeance.isoformat() if self.echeance else None,
            "ref_bon_commande": self.bon_commande,
            "ref_contrat": self.contrat,
            "masquer_jour": self.masquer_jour,
            "lignes": [
                {
                    "promotion": promotion,
                    "libelle": libelle,
                    "seances": [{"date": jour.isoformat(), "heures": h, "taux": t} for jour, h, t in seances],
                }
                for promotion, libelle, seances in self.groupes()
            ],
        }


@dataclass(frozen=True)
class Difference:
    """Changement d'un brouillon depuis la préparation précédente (`NATURES_DIFFERENCES`)."""

    identifiant: str
    nature: str
    detail: str


@dataclass(frozen=True)
class SeanceDejaFacturee:
    """Séance de l'agenda retrouvée dans `Sessions`, retirée des brouillons (ADR005)."""

    cle_seance: str
    ecole: str
    date: date
    n_facture: str
    certitude: str   # "exacte" | "même jour sans heure"


@dataclass(frozen=True)
class EcoleNonPreparee:
    ecole: str
    raison: str


@dataclass(frozen=True)
class SeanceExclue:
    """Séance retirée par une décision `exclure` : ni facturée ni à classer, mais toujours montrée."""

    cle_seance: str
    ecole: str        # "" : séance non reconnue par `Correspondances`
    seance: Seance
    raison: str
    decision: str     # id de la décision


@dataclass(frozen=True)
class EntreesPreparation:
    """Tout ce que lit la préparation, rassemblé par `charger_entrees`."""

    seances: tuple[Seance, ...]
    master: Master
    regles: Mapping[Any, RegleEcole]
    sessions: tuple[LigneSessions, ...]
    emetteur: dict[str, str] | None
    info: dict[str, str]
    nda: str | None
    erreur_coordonnees: str | None
    destinataires: dict[str, dict | None]
    decisions: tuple[Decision, ...]
    engagees: dict[str, PieceEmission]   # identifiant → pièce « en émission » ou « émise »
    numeros: tuple[str, ...]
    fenetre: tuple[date, date]
    empreinte_agenda: str
    sessions_illisibles: tuple[tuple[int, str], ...] = ()


@dataclass(frozen=True)
class Preparation:
    """Résultat d'une préparation du mois (`preparation.json`)."""

    format: int
    entite: str
    mois: str
    prepare_le: datetime
    version_outil: str
    fenetre: tuple[date, date]
    empreinte_agenda: str
    brouillons: tuple[Brouillon, ...]
    a_classer: tuple[SeanceAClasser, ...]
    hors_lot: tuple[SeanceClassee, ...]
    ecoles_sans_regle: tuple[str, ...]
    deja_facturees: tuple[SeanceDejaFacturee, ...]
    deja_emises: tuple[tuple[str, str, bool], ...]   # (identifiant, N°, contenu changé depuis l'émission)
    non_facturees_periode_emise: tuple[SeanceClassee, ...]
    ecoles_sans_seance: tuple[str, ...]
    ecoles_non_preparees: tuple[EcoleNonPreparee, ...]
    differences: tuple[Difference, ...]
    precedente_le: datetime | None = None
    exclues: tuple[SeanceExclue, ...] = ()


# ──────────────────────────────────────────────────────────────────────
# Petits outils
# ──────────────────────────────────────────────────────────────────────


def _texte(valeur: Any) -> str:
    """Contenu d'une cellule en texte NFC, sans espaces autour ; "" si vide ; 3.0 s'écrit 3."""
    if valeur is None:
        return ""
    if isinstance(valeur, float) and valeur.is_integer():
        valeur = int(valeur)
    return unicodedata.normalize("NFC", str(valeur)).strip()


def _forme(valeur: Any) -> str:
    """Forme de comparaison : NFC, sans casse, espaces réduits."""
    return " ".join(_texte(valeur).casefold().split())


def _sans_accents(texte: str) -> str:
    decompose = unicodedata.normalize("NFD", texte)
    return "".join(c for c in decompose if not unicodedata.combining(c))


def _code(valeur: Any) -> str:
    return _texte(valeur).upper()


def _uniques(valeurs: Iterable[str]) -> list[str]:
    vus: list[str] = []
    for valeur in valeurs:
        if valeur not in vus:
            vus.append(valeur)
    return vus


def _date_fr(jour: date) -> str:
    return f"{jour:%d/%m/%Y}"


def _jour_fr(jour: date) -> str:
    return f"{_JOURS[jour.weekday()]} {jour:%d/%m/%Y}"


def _horaire(debut: datetime, fin: datetime) -> str:
    return f"{debut:%H:%M}–{fin:%H:%M}"


def _h(valeur: float | None) -> str:
    return "—" if valeur is None else f"{valeur} h"


def _euros(valeur: Decimal | None) -> str:
    return "—" if valeur is None else f"{arrondi_centime(valeur)} €"


def _taux_txt(valeur: float | None) -> str:
    return "—" if valeur is None else f"{valeur:.2f} €/h"


def _pluriel(n: int, singulier: str, pluriel: str | None = None) -> str:
    return f"{n} {singulier if n <= 1 else (pluriel or singulier + 's')}"


def _cle_chrono(cle: str) -> tuple[str, str]:
    """Ordre chronologique d'une `cle_seance` (`uid@AAAA-MM-JJTHH:MM±hhmm`)."""
    return cle.rsplit("@", 1)[-1][:16], cle


def date_emission_par_defaut(mois: str) -> date:
    """Date d'émission par défaut : le 1er du mois suivant, pour toutes les écoles (ADR027).

    Raises:
        ValueError: mois invalide (le message le cite).
    """
    _, fin = chemins.mois_valide(mois)
    return fin + timedelta(days=1)


def _annee_mois(mois: str) -> tuple[int, int]:
    debut, _ = chemins.mois_valide(mois)
    return debut.year, debut.month


def fenetre_de_lecture(regles: Mapping[Any, RegleEcole], mois: str, ecoles: Sequence[str]) -> tuple[date, date]:
    """Dates de l'agenda à lire : union des périodes des écoles données et du mois civil.

    Le rythme « fin de module » (non pris en charge) et une règle dont la période ne se calcule
    pas sont ignorés ici : `preparer` en fait une école non préparée.
    """
    debut, fin = chemins.mois_valide(mois)
    annee, numero = debut.year, debut.month
    codes = {_code(e) for e in ecoles}
    for regle in regles.values():
        if regle.ecole not in codes:
            continue
        try:
            du, au = periode_a_facturer(regle, annee, numero)
        except (NotImplementedError, ValueError):
            continue
        debut, fin = min(debut, du), max(fin, au)
    return debut, fin


# ──────────────────────────────────────────────────────────────────────
# Entrées
# ──────────────────────────────────────────────────────────────────────


def _en_textes(valeurs: Mapping[str, Any]) -> dict[str, str]:
    return {str(cle): "" if valeur is None else str(valeur) for cle, valeur in valeurs.items()}


def _coordonnees_du_master(lu: Master, entite: str) -> tuple[dict | None, dict, str | None, dict]:
    """Émetteur, infos, NDA et destinataires lus dans `Coordonnées` comme `pcf_format` les lit.

    Ne sert que lorsque `pcf_format` refuse les coordonnées (émetteur, SIRET, IBAN ou BIC absent) :
    l'aperçu montre alors ce qui est présent, et chaque brouillon est bloqué.
    """
    emetteur: dict | None = None
    info: dict[str, str] = {}
    destinataires: dict[str, dict | None] = {}
    for ligne in lu.coordonnees:
        type_, code = ligne.get("Type"), ligne.get("Code")
        nom, l1, l2, l3 = (ligne.get(c) for c in ("Nom / Raison sociale", "Adresse L1", "Adresse L2", "Adresse L3"))
        email, tel = ligne.get("Email"), ligne.get("Téléphone")
        if type_ == "Émetteur" and code == entite:
            emetteur = _en_textes({"nom": nom or "", "adresse_l1": l1 or "", "adresse_l2": l2 or "",
                                   "email": email or "", "tel": tel or ""})
        elif type_ == "Destinataire" and code:
            if not nom and not l1 and not l2 and not email:
                destinataires[str(code)] = None
            else:
                destinataires[str(code)] = _en_textes({"nom": nom or "", "adresse_l1": l1 or "", "adresse_l2": l2 or "",
                                                        "adresse_l3": l3 or "", "email": email or ""})
        elif type_ == "Info" and code == entite and nom:
            info[str(nom)] = "" if l1 is None else str(l1)
    nda = info.get("NDA") or None
    return emetteur, {k: info[k] for k in ("SIRET", "IBAN", "BIC") if k in info}, nda, destinataires


def _coordonnees(entite: str, lu: Master) -> tuple[dict | None, dict, str | None, str | None, dict]:
    """(émetteur, SIRET/IBAN/BIC, NDA, erreur de coordonnées, destinataires), par `pcf_format`."""
    try:
        donnees = fmt._load_master_data(entite)
    except RuntimeError as exc:
        emetteur, info, nda, destinataires = _coordonnees_du_master(lu, entite)
        return emetteur, info, nda, str(exc), destinataires
    emetteur = _en_textes(donnees["emetteur"])
    info = {k: str(donnees["info"][k]) for k in ("SIRET", "IBAN", "BIC") if k in donnees["info"]}
    destinataires = {str(code): (_en_textes(dest) if dest else None) for code, dest in donnees["destinataires"].items()}
    try:
        nda = str(fmt.get_nda(entite=entite))
    except ValueError:
        nda = None
    return emetteur, info, nda, None, destinataires


def charger_entrees(profil: "Profil", mois: str, *, contenu_ics: bytes | None = None,
                    ecoles: Sequence[str] | None = None) -> EntreesPreparation:
    """Lit tout ce dont la préparation a besoin ; n'écrit rien et ne prend aucun verrou.

    Suppose `fmt.set_data_root(profil.racine)` déjà fait (§ 1.1). Une coordonnée manquante
    (émetteur, SIRET, IBAN, BIC) devient `erreur_coordonnees`, jamais une erreur globale ; un NDA
    absent donne `nda=None`. La fenêtre de lecture couvre les écoles `ecoles`, sinon celles du lot.

    Raises:
        FileNotFoundError, MasterInvalide, MasterVerrouille: MASTER introuvable, mal formé ou fermé.
        ValueError: mois invalide, règles illisibles ou en double, agenda illisible.
        agenda.AgendaIndisponible: agenda impossible à télécharger (sans `contenu_ics`).
        fichiers.FichierAbime: journal du mois ou pièce d'émission abîmés.
    """
    chemins.mois_valide(mois)
    lu = master.lire_master(profil.master)
    regles = regles_par_ecole(lu.regles_facturation)
    emetteur, info, nda, erreur, destinataires = _coordonnees(profil.entite, lu)

    codes = list(ecoles) if ecoles is not None else ecoles_du_lot(regles, LOT_PAR_DEFAUT)
    fenetre = fenetre_de_lecture(regles, mois, codes)
    contenu = contenu_ics if contenu_ics is not None else agenda.telecharger_ics(profil)
    seances = tuple(agenda.lire_ics(contenu, debut=fenetre[0], fin=fenetre[1], fuseau=profil.fuseau))
    empreinte_agenda = fichiers.empreinte(
        [[s.uid, s.debut, s.fin, s.intitule, s.description] for s in seances]
    )

    decisions = tuple(decisions_actives(profil, mois))
    pieces = journal.pieces_emission(profil)
    engagees = {
        p.identifiant: p for p in pieces if p.mois == mois and p.statut in (journal.EMISE, journal.EN_EMISSION)
    }
    sessions, illisibles = lire_sessions(lu)
    numeros = tuple(sorted(set(lire_numeros(profil.master)) | {p.n_facture for p in pieces}))
    return EntreesPreparation(
        seances=seances, master=lu, regles=regles, sessions=tuple(sessions),
        emetteur=emetteur, info=info, nda=nda, erreur_coordonnees=erreur, destinataires=destinataires,
        decisions=decisions, engagees=engagees, numeros=numeros, fenetre=fenetre,
        empreinte_agenda=empreinte_agenda, sessions_illisibles=tuple(illisibles),
    )


# ──────────────────────────────────────────────────────────────────────
# Préparation (pure)
# ──────────────────────────────────────────────────────────────────────


@dataclass
class _Calcul:
    """Une séance classée d'une école préparée, avec ses heures, son taux et son contrat."""

    classee: SeanceClassee
    cle: str
    heures: float | None = None
    erreur_heures: str | None = None
    taux: float | None = None
    erreur_taux: str | None = None
    erreur_cle: str | None = None    # clé tarif d'une autre école : jamais son taux
    ref_contrat: str | None = None

    @property
    def sans_taux(self) -> bool:
        return self.erreur_taux is not None or self.erreur_cle is not None


@dataclass
class _Groupe:
    """Séances d'un futur brouillon (une école, ou un contrat de l'école)."""

    identifiant: str
    ecole: str
    regle: RegleEcole
    periode: tuple[date, date]
    contrat: str | None
    par_contrat: bool
    calculs: list[_Calcul] = field(default_factory=list)
    blocages: list[str] = field(default_factory=list)


@dataclass
class _Decisions:
    exclure: dict[str, Decision] = field(default_factory=dict)
    classer: dict[str, Decision] = field(default_factory=dict)
    inclure: set[str] = field(default_factory=set)
    bon_commande: dict[str, str] = field(default_factory=dict)

    @classmethod
    def depuis(cls, decisions: Iterable[Decision]) -> "_Decisions":
        resultat = cls()
        for d in decisions:  # dans l'ordre du journal : la dernière décision sur une cible l'emporte
            if d.nature == EXCLURE:
                resultat.exclure[d.cible] = d
            elif d.nature == CLASSER:
                resultat.classer[d.cible] = d
            elif d.nature == INCLURE:
                resultat.inclure.add(d.cible)
            elif d.nature == BON_COMMANDE:
                reference = _texte(d.valeurs.get("bon_commande"))
                if reference:
                    resultat.bon_commande[d.cible] = reference
                else:
                    resultat.bon_commande.pop(d.cible, None)
        return resultat


def _selection(ecoles: Sequence[str] | None, du_lot: list[str], lot: str) -> list[str]:
    if ecoles is None:
        return list(du_lot)
    codes = _uniques(_code(e) for e in ecoles if _code(e))
    for code in codes:
        if code not in du_lot:
            raise ValueError(f"école « {code} » inconnue ou hors du {lot} (onglet Règles_facturation)")
    return sorted(codes)


def _avec_code(classee: SeanceClassee, ecoles: Sequence[Mapping[str, Any]]) -> SeanceClassee:
    """La séance classée, son école ramenée au code (`Correspondances.École` peut porter le libellé, § 15)."""
    try:
        code = code_ecole(classee.ecole, ecoles)
    except ValueError as exc:
        raise ValueError(
            f"Correspondances, clé « {_texte(classee.correspondance.get('Clé'))} », colonne « École » : {exc}"
        ) from exc
    return replace(classee, ecole=code or _code(classee.ecole))


def _valeur_de_decision(texte: str) -> Any:
    texte = _texte(texte)
    if not texte:
        return None
    if _VALEUR_DECISION_RE.fullmatch(texte):
        return float(texte.replace(",", "."))
    return texte  # heures_facturables la refusera : « heures non calculables »


def _classee_par_decision(seance: Seance, decision: Decision, ecoles: Sequence[Mapping[str, Any]]) -> SeanceClassee:
    valeurs = decision.valeurs
    try:
        code = code_ecole(valeurs.get("ecole"), ecoles)
    except ValueError as exc:
        raise ValueError(f"Décision {decision.id} (classer, journal du mois) : {exc}") from exc
    if not code:
        raise ValueError(f"Décision {decision.id} (classer, journal du mois) : école absente.")
    correspondance = {
        "Clé": f"décision {decision.id}",
        "Mode heures": _texte(valeurs.get("mode_heures")) or None,
        "Valeur": _valeur_de_decision(valeurs.get("valeur", "")),
    }
    return SeanceClassee(
        seance=seance, ecole=code, module=_texte(valeurs.get("module")),
        promotion=_texte(valeurs.get("promotion")), cle_tarif=_texte(valeurs.get("cle_tarif")),
        correspondance=correspondance, par=PAR_DECISION,
    )


def _origine(classee: SeanceClassee, colonne: str) -> str:
    """Où corriger le classement d'une séance : la ligne de `Correspondances`, ou la décision."""
    if classee.par == PAR_DECISION:
        return f"{_texte(classee.correspondance.get('Clé'))} du journal du mois, valeur « {colonne} »"
    return f"onglet Correspondances, clé « {_texte(classee.correspondance.get('Clé'))} », colonne « {colonne} »"


def _calculer(classee: SeanceClassee, lu: Master) -> _Calcul:
    calcul = _Calcul(classee=classee, cle=chemins.cle_seance(classee.seance))
    try:
        calcul.heures = heures_facturables(classee.seance, classee.correspondance)
    except (ValueError, TypeError) as exc:
        calcul.erreur_heures = str(exc)
    if calcul.heures is not None and calcul.heures <= 0:
        # `heures_facturables` accepte une Valeur nulle ; une ligne de facture à 0 h serait refusée
        # par `facture.charger_spec` (heures ≤ 0) : on bloque ici, sans rien deviner.
        calcul.erreur_heures = (
            f"{_desc(classee.seance)} : {calcul.heures} h à facturer (mode « "
            f"{_texte(classee.correspondance.get('Mode heures')) or 'agenda'} », valeur "
            f"« {_texte(classee.correspondance.get('Valeur'))} ») ; une ligne de facture porte des heures "
            "strictement positives"
        )
        calcul.heures = None
    try:
        ligne = ligne_tarif(classee.cle_tarif, lu.tarifs, ecoles=lu.ecoles)
    except ValueError as exc:
        invalide = lu.suivi_invalide.get("Tarifs")
        calcul.erreur_taux = f"{exc}{f' ({invalide})' if invalide else ''}"
        return calcul
    try:
        ecole_tarif = code_ecole(ligne.get("École"), lu.ecoles) or ""
    except ValueError:
        ecole_tarif = _code(ligne.get("École"))
    if ecole_tarif != classee.ecole:
        calcul.erreur_cle = (
            f"clé tarif « {classee.cle_tarif} » : ligne de Tarifs de l'école {ecole_tarif or '(sans école)'}, alors "
            f"que la séance du {_jour_fr(classee.seance.date)} « {classee.seance.intitule} » est classée pour "
            f"{classee.ecole} ({_origine(classee, 'Clé tarif')}) : jamais le taux d'une autre école ; corrigez la "
            "clé tarif"
        )
        return calcul
    try:
        calcul.taux = taux_horaire(classee.cle_tarif, lu.tarifs, ecoles=lu.ecoles)
        calcul.ref_contrat = _texte(ligne.get("Réf. contrat")) or None
    except ValueError as exc:
        invalide = lu.suivi_invalide.get("Tarifs")
        calcul.erreur_taux = f"{exc}{f' ({invalide})' if invalide else ''}"
    return calcul


def _reduire(reference: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", reference.upper())[:20]


def _suffixes(references: Iterable[str]) -> dict[str, str]:
    """Suffixe d'identifiant de chaque référence de contrat, sans collision."""
    reduites = {ref: _reduire(ref) or "CONTRAT" for ref in references}
    compte = Counter(reduites.values())
    return {
        ref: suffixe if compte[suffixe] == 1 else f"{suffixe}-{hashlib.sha256(ref.encode('utf-8')).hexdigest()[:4]}"
        for ref, suffixe in reduites.items()
    }


def _forme_contrat(contrat: str | None) -> str:
    return unicodedata.normalize("NFC", contrat or "").strip().casefold()


def _groupes_de_l_ecole(code: str, regles_e: list[RegleEcole], periodes: dict[str, tuple[date, date]],
                        calculs: list[_Calcul], mois: str) -> tuple[list[_Groupe], list[_Calcul]]:
    """Répartit les séances d'une école en groupes (un par contrat, ou un seul), période comprise.

    `regles_e` : toutes les lignes de l'école ; `periodes` : celles des seules lignes du lot. La
    séance d'un contrat dont la ligne est hors du lot (« Plus tard », « Hors facturation », § 16)
    n'entre dans aucun groupe : elle est rendue à part (hors lot), jamais facturée ici.
    """
    du_lot = [r for r in regles_e if r.n_contrat in periodes]
    base = du_lot[0]
    par_contrat = len(regles_e) > 1 or base.une_facture_par_contrat
    if not par_contrat:
        groupe = _Groupe(identifiant=f"{code}-{mois}", ecole=code, regle=base, periode=periodes[base.n_contrat],
                         contrat=None, par_contrat=False)
        groupe.calculs = [c for c in calculs if groupe.periode[0] <= c.classee.seance.date <= groupe.periode[1]]
        return ([groupe] if groupe.calculs else []), []

    regle_du_contrat = {_forme_contrat(r.n_contrat): r for r in regles_e if r.n_contrat}
    graphies: dict[str, str] = {}  # une référence écrite de deux façons (casse) reste un seul contrat
    affectations: list[tuple[_Calcul, str | None, RegleEcole, str | None]] = []
    hors_lot: list[_Calcul] = []
    for c in sorted(calculs, key=lambda c: (c.classee.seance.debut, c.cle)):
        if c.ref_contrat is not None:
            c.ref_contrat = graphies.setdefault(_forme_contrat(c.ref_contrat), c.ref_contrat)
        cle_tarif = c.classee.cle_tarif
        if c.sans_taux:
            affectations.append((c, None, base, None))
        elif c.ref_contrat is None:
            affectations.append((c, None, base, f"référence de contrat absente (Tarifs, colonne « Réf. contrat », clé "
                                                 f"« {cle_tarif} ») : la facture d'une école facturée par contrat cite "
                                                 "son contrat (ADR030 §3)"))
        elif _forme_contrat(c.ref_contrat) in regle_du_contrat:
            regle = regle_du_contrat[_forme_contrat(c.ref_contrat)]
            if regle.n_contrat not in periodes:
                hors_lot.append(c)
                continue
            affectations.append((c, regle.n_contrat, regle, None))
        elif len(regles_e) == 1 and not base.n_contrat:
            affectations.append((c, c.ref_contrat, base, None))
        else:
            affectations.append((c, c.ref_contrat, base,
                                 f"« Réf. contrat » « {c.ref_contrat} » (Tarifs, clé « {cle_tarif} ») sans ligne dans "
                                 f"Règles_facturation : ajoutez la ligne de ce contrat (colonne « N° de contrat »)"))

    references = {r.n_contrat for r in regles_e if r.n_contrat} | {cle for _, cle, _, _ in affectations if cle}
    suffixes = _suffixes(sorted(references))
    groupes: dict[str | None, _Groupe] = {}
    for c, cle, regle, blocage in affectations:
        if cle not in groupes:
            suffixe = SANS_CONTRAT if cle is None else suffixes[cle]
            groupes[cle] = _Groupe(identifiant=f"{code}-{mois}-{suffixe}", ecole=code, regle=regle,
                                   periode=periodes.get(regle.n_contrat, periodes[base.n_contrat]),
                                   contrat=cle, par_contrat=True)
        groupe = groupes[cle]
        if not groupe.periode[0] <= c.classee.seance.date <= groupe.periode[1]:
            continue
        groupe.calculs.append(c)
        if blocage:
            groupe.blocages.append(blocage)
    if len(regles_e) > 1 and not all(r.une_facture_par_contrat for r in regles_e):
        numeros = ", ".join(r.n_contrat for r in regles_e)
        for groupe in groupes.values():
            groupe.blocages.append(
                f"école {code} : plusieurs lignes de contrat dans Règles_facturation ({numeros}) sans "
                "« Une facture par contrat = oui » sur chacune : renseignez-la, une facture par contrat (ADR030 §3)"
            )
    return [g for g in sorted(groupes.values(), key=lambda g: g.identifiant) if g.calculs], hors_lot


def _destinataire(destinataires: Mapping[str, dict | None], code: str) -> dict | None:
    if code in destinataires:
        return destinataires[code]
    for cle, dest in destinataires.items():
        if _code(cle) == code:
            return dest
    return None


def _blocages_coordonnees(entrees: EntreesPreparation, code: str) -> list[str]:
    """Mentions et données obligatoires communes (ADR030 §10)."""
    blocages = []
    dest = _destinataire(entrees.destinataires, code)
    if dest is None:
        blocages.append(
            f"destinataire absent : ajoutez la ligne « Destinataire » / {code} de l'onglet Coordonnées (colonnes "
            "« Nom / Raison sociale », « Adresse L1 » et « Adresse L2 »)"
        )
    else:
        manquants = fmt.champs_destinataire_manquants(dest)
        if manquants:
            blocages.append(
                f"destinataire {code} incomplet : colonne{'s' if len(manquants) > 1 else ''} "
                f"{', '.join(f'« {m} »' for m in manquants)} vide{'s' if len(manquants) > 1 else ''} dans l'onglet "
                f"Coordonnées (ligne Destinataire / {code})"
            )
    if entrees.erreur_coordonnees:
        blocages.append(f"coordonnées de l'émetteur incomplètes (onglet Coordonnées) : {entrees.erreur_coordonnees}")
    if entrees.emetteur is None:
        if not entrees.erreur_coordonnees:
            blocages.append("émetteur absent : ajoutez la ligne « Émetteur » de l'entité dans l'onglet Coordonnées")
    else:
        if not fmt.emetteur_porte_ei(entrees.emetteur.get("nom")):
            blocages.append(
                f"le nom de l'émetteur « {entrees.emetteur.get('nom')} » ne porte pas la mention « EI » ni "
                "« Entrepreneur individuel » (onglet Coordonnées, ligne Émetteur, colonne « Nom / Raison sociale ») : "
                "mention obligatoire (ADR030)"
            )
        vides = [colonne for cle, colonne in _CHAMPS_EMETTEUR if not _texte(entrees.emetteur.get(cle))]
        if vides:
            blocages.append(
                f"émetteur incomplet : colonne{'s' if len(vides) > 1 else ''} {', '.join(f'« {v} »' for v in vides)} "
                f"vide{'s' if len(vides) > 1 else ''} dans l'onglet Coordonnées (ligne Émetteur) : mention "
                "obligatoire de la facture (ADR030 §1)"
            )
    for nom in _INFOS_OBLIGATOIRES:
        if nom in entrees.info and not _texte(entrees.info.get(nom)):
            blocages.append(
                f"{nom} vide : renseignez la colonne « Adresse L1 » de la ligne Info / {nom} de l'onglet Coordonnées "
                "(mention obligatoire de la facture, ADR030 §1 et §10)"
            )
    if not entrees.nda:
        blocages.append(
            "n° de déclaration d'activité (NDA) absent : ajoutez la ligne Info / NDA de l'onglet Coordonnées "
            "(colonne « Adresse L1 » = texte imprimé sur la facture, ADR030)"
        )
    return blocages


def _libelle_ecole(code: str, ecoles: Sequence[Mapping[str, Any]]) -> str:
    for ligne in ecoles:
        if _code(ligne.get("Code facturation")) == code and _texte(ligne.get("École")):
            return _texte(ligne.get("École"))
    return ""


def _cle_periode(valeur: Any) -> str:
    return _sans_accents(_forme(valeur))


def _periode_lisible(valeur: Any) -> str:
    if isinstance(valeur, (date, datetime)):
        return _date_fr(valeur.date() if isinstance(valeur, datetime) else valeur)
    return _texte(valeur)


def _meme_periode(valeur: Any, formes: set[str], annee: int, numero: int) -> bool:
    """`Factures.Période` désigne-t-elle le mois ? Libellé (« Déc 2026 », « Décembre 2026 »), ou date du
    mois : Excel en français convertit la saisie « Déc 2026 » en date (01/12/2026)."""
    if isinstance(valeur, (date, datetime)):
        return (valeur.year, valeur.month) == (annee, numero)
    return _cle_periode(valeur) in formes


def _periode_deja_facturee(code: str, contrat: str | None, mois: str, lu: Master, engagees: Iterable[PieceEmission],
                           identifiant: str) -> tuple[list[str], list[str]]:
    """Lignes de `Factures` de la même école et de la même période, faites hors de l'outil.

    Même contrat (vide égal vide) : blocage. Contrat renseigné d'un seul côté (ligne historique de
    `Factures` sans « Contrat », Q-D6) : avertissement, pour que la garde ne se taise jamais.
    """
    annee, numero = _annee_mois(mois)
    formes = {_cle_periode(chemins.libelle_periode(mois)), _cle_periode(f"{chemins.NOMS_MOIS[numero - 1]} {annee}")}
    abandonnes = numeros_abandonnes(lu.factures)
    pieces = {_code(p.n_facture) for p in engagees}
    rangs = list(lu.rangs.get("factures") or [])
    if len(rangs) != len(lu.factures):
        rangs = list(range(2, len(lu.factures) + 2))
    blocages: list[str] = []
    avertissements: list[str] = []
    for rang, ligne in zip(rangs, lu.factures):
        try:
            ecole = code_ecole(ligne.get("École"), lu.ecoles) or ""
        except ValueError:
            ecole = _code(ligne.get("École"))
        n = _texte(ligne.get("N° facture"))
        if ecole != code or not _meme_periode(ligne.get("Période"), formes, annee, numero):
            continue
        if "".join(n.split()).upper() in abandonnes or _code(n) in pieces:
            continue
        contrat_ligne = _texte(ligne.get("Contrat"))
        ou = f"onglet Factures, ligne {rang}, colonne « Période » « {_periode_lisible(ligne.get('Période'))} »"
        if _forme_contrat(contrat_ligne) == _forme_contrat(contrat):
            blocages.append(
                f"période déjà facturée hors outil : la facture {n or '(sans n°)'} ({ou}) couvre déjà "
                f"{chemins.libelle_periode(mois)} pour l'école {code}{f' et le contrat {contrat}' if contrat else ''}, "
                f"sans pièce d'émission de l'outil ; si elle ne couvre pas ces séances, levez ce blocage par une "
                f"décision « inclure » sur {identifiant} (`facturation corriger --inclure`)"
            )
        elif not contrat_ligne or not contrat:
            cote_factures = f"pour le contrat {contrat_ligne}" if contrat_ligne else "sans « Contrat »"
            cote_brouillon = f"cite le contrat {contrat}" if contrat else "ne cite aucun contrat"
            avertissements.append(
                f"{PERIODE_PEUT_ETRE_FACTUREE} : la facture {n or '(sans n°)'} ({ou}) couvre déjà "
                f"{chemins.libelle_periode(mois)} pour l'école {code}, {cote_factures}, alors que ce brouillon "
                f"{cote_brouillon} ; vérifiez qu'elle ne couvre pas ces séances (renseignez sa colonne « Contrat »), "
                f"ou confirmez par une décision « inclure » sur {identifiant}"
            )
    return blocages, avertissements


def _seances_identiques(a: LigneBrouillon | Seance, b: LigneBrouillon | Seance) -> bool:
    return a.debut == b.debut and a.fin == b.fin


def _se_chevauchent(a, b) -> bool:
    return a.debut < b.fin and b.debut < a.fin


def _desc(ligne) -> str:
    return f"{_jour_fr(ligne.date)} {_horaire(ligne.debut, ligne.fin)} « {ligne.intitule} »"


def _possibles_doublons(code: str, seance: Seance, sessions: Sequence[LigneSessions]) -> list[str]:
    """Lignes `Sessions` du même jour et de la même école, aux heures qui chevauchent sans être identiques."""
    messages = []
    debut, fin = seance.debut.time().replace(tzinfo=None), seance.fin.time().replace(tzinfo=None)
    for ligne in sessions:
        if ligne.abandonnee or ligne.ecole.upper() != code or ligne.date_seance != seance.date:
            continue
        if ligne.debut is None or ligne.fin is None:
            continue
        l_debut, l_fin = ligne.debut.replace(second=0, microsecond=0), ligne.fin.replace(second=0, microsecond=0)
        if (l_debut, l_fin) == (debut.replace(second=0, microsecond=0), fin.replace(second=0, microsecond=0)):
            continue
        if l_debut < fin and debut < l_fin:
            messages.append(
                f"{POSSIBLE_DOUBLON}, ligne {ligne.rang} ({_date_fr(ligne.date_seance)} "
                f"{l_debut:%H:%M}–{l_fin:%H:%M}, {ligne.n_facture or 'sans n°'}) : séance du "
                f"{_jour_fr(seance.date)} {_horaire(seance.debut, seance.fin)} « {seance.intitule} » — vérifiez "
                "qu'elle n'est pas déjà facturée"
            )
    return messages


def _heures_sessions(ligne: LigneSessions) -> str:
    if ligne.debut is None and ligne.fin is None:
        return "sans heure"
    return f"{f'{ligne.debut:%H:%M}' if ligne.debut else '?'}–{f'{ligne.fin:%H:%M}' if ligne.fin else '?'}"


def _sessions_a_corriger(code: str, jours: set[date], sessions: Sequence[LigneSessions],
                         illisibles: Sequence[tuple[int, str]]) -> list[str]:
    """Lignes de `Sessions` que l'anti-doublon (ADR005) ne peut pas rapprocher des séances du brouillon.

    Une ligne illisible, ou sans école ni date, vaut pour chaque brouillon ; une ligne sans école
    vaut pour les brouillons qui ont une séance ce jour-là ; une ligne sans date, pour ceux de son
    école. Avertissement repris dans « À FAIRE » : une séance déjà facturée pourrait être refacturée.
    """
    messages = []
    for ligne in sessions:
        if ligne.abandonnee or (ligne.ecole and ligne.date_seance is not None):
            continue
        n = ligne.n_facture or "sans n°"
        if not ligne.ecole and ligne.date_seance is None:
            manque, concerne = "colonnes « École » et « Date_séance » vides", True
        elif not ligne.ecole:
            manque, concerne = "colonne « École » vide", ligne.date_seance in jours
        else:
            manque, concerne = "colonne « Date_séance » vide", ligne.ecole.upper() == code
        if not concerne:
            continue
        jour = _date_fr(ligne.date_seance) if ligne.date_seance else "date ?"
        messages.append(
            f"{SESSIONS_A_CORRIGER} : onglet « Sessions », ligne {ligne.rang}, {manque} ({ligne.ecole or 'école ?'}, "
            f"{jour}, {_heures_sessions(ligne)}, {n}) — l'anti-doublon ne la rapproche d'aucune séance : "
            "complétez-la, une séance déjà facturée pourrait être refacturée"
        )
    for _, message in illisibles:
        messages.append(
            f"{SESSIONS_A_CORRIGER} : {message} — l'anti-doublon ne lit pas cette ligne : corrigez-la, une séance "
            "déjà facturée pourrait être refacturée"
        )
    return messages


def _blocages_heures_taux(calculs: Iterable[_Calcul]) -> list[str]:
    blocages = []
    for c in calculs:
        if c.erreur_heures is not None:
            ou = "décision du journal du mois" if c.classee.par == PAR_DECISION else "onglet Correspondances"
            blocages.append(f"heures non calculables : {c.erreur_heures} ({ou})")
        if c.erreur_taux is not None:
            blocages.append(
                f"taux introuvable pour la clé « {c.classee.cle_tarif} » (onglet Tarifs) : ne jamais inventer un taux "
                f"— {c.erreur_taux}"
            )
        if c.erreur_cle is not None:
            blocages.append(c.erreur_cle)
    return blocages


def _contrat_commun(groupe: _Groupe, calculs: Iterable[_Calcul]) -> tuple[str | None, list[str]]:
    """Contrat d'un brouillon d'école sans « Une facture par contrat » (ADR030 §3).

    Le « N° de contrat » de la règle l'emporte ; une « Réf. contrat » de Tarifs qui en diffère est
    un blocage (la mention du contrat ne peut pas être fausse, comme en facturation par contrat).
    """
    if groupe.regle.n_contrat:
        n = groupe.regle.n_contrat
        ecarts: dict[str, list[str]] = {}
        for c in calculs:
            if c.ref_contrat and _forme_contrat(c.ref_contrat) != _forme_contrat(n):
                cles = ecarts.setdefault(c.ref_contrat, [])
                if c.classee.cle_tarif not in cles:
                    cles.append(c.classee.cle_tarif)
        if not ecarts:
            return n, []
        refs = ", ".join(f"« {ref} » (clé{'s' if len(cles) > 1 else ''} {', '.join(f'« {k} »' for k in cles)})"
                         for ref, cles in sorted(ecarts.items()))
        return n, [
            f"« Réf. contrat » {refs} de l'onglet Tarifs différente du « N° de contrat » « {n} » de l'école "
            f"{groupe.ecole} (onglet Règles_facturation) : la facture citerait un autre contrat que celui de ses "
            "tarifs (ADR030 §3) ; corrigez l'une des deux références, ou ajoutez la ligne de chaque contrat avec "
            "« Une facture par contrat = oui »"
        ]
    references = sorted({c.ref_contrat for c in calculs if c.ref_contrat}, key=_forme_contrat)
    distinctes = _uniques(references) if len({_forme_contrat(r) for r in references}) > 1 else references[:1]
    if len(distinctes) <= 1:
        return (distinctes[0] if distinctes else None), []
    return None, [
        f"plusieurs contrats ({', '.join(distinctes)}) : renseignez « N° de contrat » ou « Une facture par contrat = "
        "oui » (Règles_facturation) ; la facture cite son contrat (ADR030 §3)"
    ]


@dataclass
class _Resultat:
    """Brouillon construit pour un groupe, et ce que l'anti-doublon en a retiré."""

    brouillon: Brouillon | None
    deja: list[SeanceDejaFacturee]
    retenus: list[_Calcul]


def _construire(groupe: _Groupe, *, entite: str, mois: str, entrees: EntreesPreparation, decisions: _Decisions,
                maintenant: datetime, a_classer: Sequence[SeanceAClasser], exclues: Sequence[SeanceExclue],
                ignorer_numero: str | None = None, cles_pieces: Mapping[str, set[str]] | None = None,
                orphelines: Sequence[PieceEmission] = ()) -> _Resultat:
    """Brouillon d'un groupe (sans numéro probable ni empreinte), après l'anti-doublon.

    `cles_pieces` : `cle_seance` de chaque pièce engagée de l'école. Une séance d'une autre pièce que
    celle du groupe n'est jamais reprise (la pièce n'a peut-être pas encore écrit ses lignes dans
    `Sessions`). `orphelines` : pièces de l'école dont l'identifiant n'est plus celui d'aucun groupe
    (règles changées depuis) ; elles bloquent le brouillon : pas de facture complémentaire
    automatique d'une période déjà émise (Q-D3).
    """
    lu = entrees.master
    code = groupe.ecole
    cles_pieces = cles_pieces or {}
    sessions = [
        l for l in entrees.sessions
        if ignorer_numero is None or "".join(l.n_facture.split()).upper() != ignorer_numero.upper()
    ]
    deja: list[SeanceDejaFacturee] = []
    retenus: list[_Calcul] = []
    avertissements: list[str] = []
    dans_une_piece: list[tuple[_Calcul, str]] = []
    for c in sorted(groupe.calculs, key=lambda c: (c.classee.seance.debut, c.cle)):
        autre = next((i for i, cles in sorted(cles_pieces.items()) if i != groupe.identifiant and c.cle in cles), None)
        if autre is not None:
            dans_une_piece.append((c, autre))
            continue
        s = c.classee.seance
        exactes, sans_heure = chercher(code, s.debut, s.fin, sessions)
        if exactes:
            deja.append(SeanceDejaFacturee(c.cle, code, s.date, exactes[0].n_facture, EXACTE))
            continue
        if sans_heure and c.cle not in decisions.inclure:
            deja.append(SeanceDejaFacturee(c.cle, code, s.date, sans_heure[0].n_facture, MEME_JOUR))
            continue
        retenus.append(c)
        avertissements += _possibles_doublons(code, s, sessions)
    if not retenus:
        return _Resultat(None, deja, retenus)

    regle = groupe.regle
    lignes = tuple(
        LigneBrouillon(
            cle_seance=c.cle, uid=c.classee.seance.uid, date=c.classee.seance.date, debut=c.classee.seance.debut,
            fin=c.classee.seance.fin, intitule=c.classee.seance.intitule, ecole=code,
            promotion=c.classee.promotion, module=c.classee.module, libelle=c.classee.module,
            cle_tarif=c.classee.cle_tarif, classee_par=c.classee.par, duree_agenda_h=c.classee.seance.duree_h,
            heures=c.heures, taux=c.taux,
            montant=montant_exact(c.heures, c.taux) if c.heures is not None and c.taux is not None else None,
            ref_contrat=c.ref_contrat, tenue=c.classee.seance.fin <= maintenant,
        )
        for c in retenus
    )

    if groupe.par_contrat:
        contrat, blocages_contrat = groupe.contrat, list(groupe.blocages)
    else:
        contrat, blocages_contrat = _contrat_commun(groupe, retenus)
        blocages_contrat = list(groupe.blocages) + blocages_contrat

    identifiant = groupe.identifiant
    bon_commande = decisions.bon_commande.get(identifiant)
    blocages = _blocages_coordonnees(entrees, code)
    for raison in regles_facturation.blocages(regle, ref_bon_commande=bon_commande):
        if regle.bc_requis and not bon_commande and "bon de commande" in raison.casefold():
            raison += (
                f" Saisissez sa référence par une décision « bon_commande » sur {identifiant} (`facturation corriger "
                "--bon-commande`), ou, pour une facture par ligne du bon de commande, voyez "
                "examples/factures_par_bon_de_commande.py (Q-C1)."
            )
        blocages.append(raison)
    blocages += blocages_contrat
    blocages += _blocages_heures_taux(retenus)
    for piece in orphelines:
        blocages.append(
            f"période déjà émise sous un autre identifiant : la pièce {piece.n_facture} ({piece.identifiant}) couvre "
            f"déjà {chemins.libelle_periode(mois)} pour l'école {code}, et les règles ont changé depuis (onglet "
            "Règles_facturation) ; ces séances n'y sont pas : facturez-les à la main si elles sont dues, jamais par "
            "une facture complémentaire automatique (Q-D3)"
        )

    date_emission = date_emission_par_defaut(mois)
    echeance: date | None
    try:
        echeance = fmt.calculer_echeance(date_emission, echeance_selon_regle(date_emission, regle.delai_paiement))
    except (ValueError, TypeError) as exc:
        echeance = None
        blocages.append(f"échéance impossible à calculer (Règles_facturation, colonne « Délai de paiement ») : {exc}")
    if identifiant not in decisions.inclure:
        deja_hors_outil, peut_etre = _periode_deja_facturee(code, contrat, mois, lu, entrees.engagees.values(),
                                                            identifiant)
        blocages += deja_hors_outil
        avertissements += peut_etre

    if regle.rythme == "cycle":
        avertissements.insert(0, "école à cycle : date d'émission au 1er du mois suivant (ADR027) — voir Q-A2")
    ecole_label = _libelle_ecole(code, lu.ecoles)
    if not ecole_label:
        ecole_label = code
        avertissements.append(
            f"libellé de l'école absent de l'onglet Écoles (colonne « École ») : le code {code} est utilisé"
        )
    a_venir = sum(1 for l in lignes if not l.tenue)
    if a_venir:
        avertissements.append(f"{_pluriel(a_venir, 'séance(s)', 'séance(s)')} à venir : recontrôle avant envoi (F-23)")
    du, au = groupe.periode
    creneaux = sum(1 for a in a_classer if du <= a.seance.date <= au)
    if creneaux:
        avertissements.append(
            f"{creneaux} créneau(x) à classer dans la période : jamais facturé(s) tant qu'il(s) ne sont pas classé(s) "
            "(facturation classer)"
        )
    for exclue in exclues:
        if exclue.ecole == code and du <= exclue.seance.date <= au:
            raison = f" ({exclue.raison})" if exclue.raison else ""
            avertissements.append(f"séance exclue par décision : {_desc(exclue.seance)}{raison}")
    for c, autre in dans_une_piece:
        piece_n = next((p.n_facture for p in entrees.engagees.values() if p.identifiant == autre), "?")
        avertissements.append(
            f"séance déjà dans la pièce {piece_n} ({autre}, en émission ou émise) : {_desc(c.classee.seance)} — "
            "retirée de ce brouillon, jamais refacturée"
        )
    avertissements += _sessions_a_corriger(code, {l.date for l in lignes}, sessions, entrees.sessions_illisibles)

    heures_lignes = [l.heures for l in lignes]
    total_heures = None
    if all(h is not None for h in heures_lignes):
        total_heures = float(sum(Decimal(repr(h)) for h in heures_lignes))
    montants = [l.montant for l in lignes]
    total_ht = None if any(m is None for m in montants) else arrondi_centime(sum(montants, Decimal(0)))

    brouillon = Brouillon(
        identifiant=identifiant, entite=entite, ecole=code, mois=mois, periode=groupe.periode,
        periode_label=chemins.libelle_periode(mois), ecole_label=ecole_label, contrat=contrat,
        bon_commande=bon_commande,
        destinataire=dict(dest) if (dest := _destinataire(entrees.destinataires, code)) is not None else None,
        emetteur=dict(entrees.emetteur) if entrees.emetteur is not None else None,
        info=dict(entrees.info), nda=entrees.nda, date_emission=date_emission, echeance=echeance,
        masquer_jour=regle.colonne_jour == "masquée", regroupement=regle.regroupement or "séance",
        lignes=lignes, total_heures=total_heures, total_ht=total_ht, numero_probable=None,
        blocages=tuple(blocages), avertissements=tuple(avertissements), empreinte="",
    )
    return _Resultat(brouillon, deja, retenus)


def _cles_de_la_piece(piece: PieceEmission) -> set[str]:
    """`cle_seance` des lignes du brouillon figé dans la pièce d'émission."""
    lignes = piece.brouillon.get("lignes") if isinstance(piece.brouillon, dict) else None
    return {str(l["cle_seance"]) for l in lignes or [] if isinstance(l, dict) and l.get("cle_seance")}


def _valeur_stockable(valeur: Any) -> Any:
    """Valeur de cellule écrite telle quelle dans `preparation.json` si JSON la porte, sinon en texte
    (une durée [h]:mm lue en `timedelta`, par exemple) : jamais d'échec d'écriture pour une cellule."""
    if valeur is None or isinstance(valeur, (str, bool, int)):
        return valeur
    if isinstance(valeur, float):
        return valeur if math.isfinite(valeur) else str(valeur)
    try:
        return str(fichiers.en_json(valeur))
    except TypeError:
        return str(valeur)


def _stockable(classee: SeanceClassee) -> SeanceClassee:
    """La séance classée, sa ligne de `Correspondances` réduite à des valeurs que JSON porte."""
    return replace(classee, correspondance={str(k): _valeur_stockable(v) for k, v in classee.correspondance.items()})


def _gardes_de_l_ecole(resultats: list[_Resultat]) -> None:
    """Doublons d'agenda (blocage) et chevauchements (avertissement) entre séances d'une même école."""
    lignes = sorted(
        ((r_index, l) for r_index, r in enumerate(resultats) if r.brouillon for l in r.brouillon.lignes),
        key=lambda x: (x[1].debut, x[1].cle_seance),
    )
    ajouts: dict[int, tuple[list[str], list[str]]] = {}
    for i, (ri, a) in enumerate(lignes):
        for rj, b in lignes[i + 1:]:
            if b.debut >= a.fin and not _seances_identiques(a, b):
                break
            if _seances_identiques(a, b):
                message = (f"doublon d'agenda : deux séances identiques le {_jour_fr(a.date)} de {a.debut:%H:%M} à "
                           f"{a.fin:%H:%M} (« {a.intitule} », uid {a.uid} et {b.uid}) : supprimez l'une d'elles dans "
                           "l'agenda")
                for index in {ri, rj}:
                    ajouts.setdefault(index, ([], []))[0].append(message)
            elif _se_chevauchent(a, b):
                message = (f"séances qui se chevauchent : {_desc(a)} et {_horaire(b.debut, b.fin)} « {b.intitule} » — "
                           "vérifiez l'agenda")
                for index in {ri, rj}:
                    ajouts.setdefault(index, ([], []))[1].append(message)
    for index, (blocages, avertissements) in ajouts.items():
        b = resultats[index].brouillon
        resultats[index].brouillon = replace(
            b,
            blocages=tuple(_uniques([*b.blocages, *blocages])),
            avertissements=tuple(_uniques([*b.avertissements, *avertissements])),
        )


def preparer(*, entite: str, mois: str, entrees: EntreesPreparation, prefixe: str, maintenant: datetime,
             ecoles: Sequence[str] | None = None, lot: str = LOT_PAR_DEFAUT) -> Preparation:
    """Brouillons du mois à partir des entrées ; fonction pure et déterministe.

    Raises:
        ValueError: mois invalide ; école de `ecoles` inconnue ou hors du lot ; ligne de
            `Correspondances` mal formée (erreur globale, le message nomme l'onglet). Une décision
            `classer` inapplicable (école inconnue) renvoie seulement sa séance dans la file à classer.
    """
    debut_mois, fin_mois = chemins.mois_valide(mois)
    annee, numero = debut_mois.year, debut_mois.month
    lu = entrees.master
    decisions = _Decisions.depuis(entrees.decisions)

    # 1. Écoles ; seules les lignes du lot (« Statut lot » de chaque ligne de contrat, § 16) ont une période
    par_code: dict[str, list[RegleEcole]] = {}
    for regle in entrees.regles.values():
        par_code.setdefault(regle.ecole, []).append(regle)
    for liste in par_code.values():
        liste.sort(key=lambda r: r.n_contrat)
    du_lot = ecoles_du_lot(entrees.regles, lot)
    selection = _selection(ecoles, du_lot, lot)
    periodes_ecole: dict[str, dict[str, tuple[date, date]]] = {}
    non_preparees: list[EcoleNonPreparee] = []
    for code in selection:
        lignes_du_lot = [r for r in par_code[code] if _forme(r.statut_lot) == _forme(lot)]
        try:
            periodes_ecole[code] = {r.n_contrat: periode_a_facturer(r, annee, numero) for r in lignes_du_lot}
        except NotImplementedError:
            non_preparees.append(EcoleNonPreparee(code, "rythme « fin de module » : à préparer à la main"))
        except ValueError as exc:
            non_preparees.append(EcoleNonPreparee(code, str(exc)))

    # 2. Classement, puis décisions du mois
    classees, a_classer_brut = classer(entrees.seances, lu.correspondances)
    reconnues = {id(c.seance): c for c in classees}
    file = {id(a.seance): a for a in a_classer_brut}
    retenues: list[SeanceClassee] = []
    a_classer: list[SeanceAClasser] = []
    exclues: list[SeanceExclue] = []
    for seance in entrees.seances:
        cle = chemins.cle_seance(seance)
        if cle in decisions.exclure:
            d = decisions.exclure[cle]
            origine = reconnues.get(id(seance))
            ecole = ""
            if origine is not None:
                try:
                    ecole = code_ecole(origine.ecole, lu.ecoles) or ""
                except ValueError:
                    ecole = ""
            exclues.append(SeanceExclue(cle, ecole, seance, _texte(d.valeurs.get("raison")) or _texte(d.commentaire),
                                        d.id))
            continue
        if cle in decisions.classer:
            d = decisions.classer[cle]
            try:
                retenues.append(_classee_par_decision(seance, d, lu.ecoles))
            except ValueError as exc:  # école inconnue : la séance n'est jamais facturée, elle revient à classer
                a_classer.append(SeanceAClasser(
                    seance=seance,
                    raison=f"décision {d.id} (classer) inapplicable : {exc} ; annulez-la ou classez à nouveau",
                ))
        elif id(seance) in reconnues:
            retenues.append(_avec_code(reconnues[id(seance)], lu.ecoles))
        else:
            a_classer.append(file[id(seance)])

    def dans_une_periode(jour: date) -> bool:
        if debut_mois <= jour <= fin_mois:
            return True
        return any(du <= jour <= au for periodes in periodes_ecole.values() for du, au in periodes.values())

    a_classer = [a for a in a_classer if dans_une_periode(a.seance.date)]
    par_ecole: dict[str, list[SeanceClassee]] = {}
    hors_lot: list[SeanceClassee] = []
    sans_regle: set[str] = set()
    for classee in retenues:
        if classee.ecole in periodes_ecole:
            par_ecole.setdefault(classee.ecole, []).append(classee)
        elif classee.ecole not in par_code:
            if debut_mois <= classee.seance.date <= fin_mois:
                sans_regle.add(classee.ecole)
        elif classee.ecole not in du_lot and debut_mois <= classee.seance.date <= fin_mois:
            hors_lot.append(classee)

    # 3 à 12. Groupes, brouillons, brouillons engagés
    brouillons: list[Brouillon] = []
    deja_facturees: list[SeanceDejaFacturee] = []
    deja_emises: list[tuple[str, str, bool]] = []
    non_facturees: list[SeanceClassee] = []
    sans_seance: list[str] = []
    for code in sorted(periodes_ecole):
        calculs = [_calculer(c, lu) for c in par_ecole.get(code, [])]
        groupes, contrats_hors_lot = _groupes_de_l_ecole(code, par_code[code], periodes_ecole[code], calculs, mois)
        hors_lot += [c.classee for c in contrats_hors_lot if debut_mois <= c.classee.seance.date <= fin_mois]
        engagees = {i: p for i, p in entrees.engagees.items() if i.split("-")[0] == code}
        if not groupes and not engagees:
            sans_seance.append(code)
            continue
        cles_pieces = {i: _cles_de_la_piece(p) for i, p in engagees.items()}
        toutes_cles = set().union(*cles_pieces.values()) if cles_pieces else set()
        identifiants_groupes = {g.identifiant for g in groupes}
        orphelines = [p for i, p in sorted(engagees.items()) if i not in identifiants_groupes]
        resultats: list[_Resultat] = []
        for groupe in groupes:
            piece = engagees.get(groupe.identifiant)
            commun = dict(entite=entite, mois=mois, entrees=entrees, decisions=decisions, maintenant=maintenant,
                          a_classer=a_classer, exclues=exclues, cles_pieces=cles_pieces)
            if piece is None:
                resultat = _construire(groupe, orphelines=orphelines, **commun)
                deja_facturees += resultat.deja
                if resultat.brouillon is not None:
                    resultats.append(resultat)
                continue
            resultat = _construire(groupe, ignorer_numero=piece.n_facture, **commun)
            recalcule = resultat.brouillon
            change = recalcule is None or empreinte_brouillon(recalcule) != piece.empreinte
            deja_emises.append((groupe.identifiant, piece.n_facture, change))
            for c in groupe.calculs:
                s = c.classee.seance
                if c.cle in toutes_cles:
                    continue
                exactes, sans_heure = chercher(code, s.debut, s.fin, entrees.sessions)
                if not exactes and not sans_heure:
                    non_facturees.append(c.classee)
        for piece in orphelines:
            deja_emises.append((piece.identifiant, piece.n_facture, True))
        _gardes_de_l_ecole(resultats)
        brouillons += [r.brouillon for r in resultats]

    # Numéros probables, empreintes
    numeros = list(entrees.numeros)
    finals: list[Brouillon] = []
    for b in sorted(brouillons, key=lambda b: b.identifiant):
        blocages, avertissements = list(b.blocages), list(b.avertissements)
        try:
            probable = prochain_numero(numeros, prefixe=prefixe, code_ecole=b.ecole)
            numeros.append(probable)
        except ValueError as exc:
            probable = None
            blocages.append(f"{exc} (numérotation, ADR029)")
        try:
            manquants = numeros_manquants(entrees.numeros, prefixe=prefixe, code_ecole=b.ecole)
        except ValueError:
            manquants = []
        if manquants:
            avertissements.append(
                f"numéros manquants dans la série {prefixe}{b.ecole}, sans ligne au registre : "
                f"{', '.join(manquants)} — notez leur raison dans l'onglet Factures ; ils ne sont jamais réutilisés "
                "(ADR029)"
            )
        b = replace(b, numero_probable=probable, blocages=tuple(_uniques(blocages)),
                    avertissements=tuple(_uniques(avertissements)))
        finals.append(replace(b, empreinte=empreinte_brouillon(b)))

    return Preparation(
        format=FORMAT_PREPARATION, entite=entite, mois=mois, prepare_le=maintenant, version_outil=__version__,
        fenetre=tuple(entrees.fenetre), empreinte_agenda=entrees.empreinte_agenda,
        brouillons=tuple(finals),
        a_classer=tuple(a_classer),
        hors_lot=tuple(_stockable(c) for c in sorted(hors_lot, key=lambda c: (c.seance.debut,
                                                                                chemins.cle_seance(c.seance)))),
        ecoles_sans_regle=tuple(sorted(sans_regle)),
        deja_facturees=tuple(sorted(deja_facturees, key=lambda d: _cle_chrono(d.cle_seance))),
        deja_emises=tuple(sorted(deja_emises)),
        non_facturees_periode_emise=tuple(
            _stockable(c) for c in sorted(non_facturees, key=lambda c: (c.seance.debut, chemins.cle_seance(c.seance)))
        ),
        ecoles_sans_seance=tuple(sans_seance),
        ecoles_non_preparees=tuple(non_preparees),
        differences=(),
        exclues=tuple(sorted(exclues, key=lambda e: _cle_chrono(e.cle_seance))),
    )


# ──────────────────────────────────────────────────────────────────────
# Empreinte, différences, numéros courts
# ──────────────────────────────────────────────────────────────────────


def _repr(valeur: float | None) -> str | None:
    return None if valeur is None else repr(float(valeur))


def empreinte_brouillon(brouillon: Brouillon) -> str:
    """sha256 du contenu écrit sur la facture (§ 1.3) ; heures et taux en texte.

    Hors empreinte : `tenue`, `numero_probable`, avertissements, blocages, date de préparation.
    """
    b = brouillon
    return fichiers.empreinte({
        "entite": b.entite, "identifiant": b.identifiant, "ecole": b.ecole, "mois": b.mois,
        "periode": [b.periode[0], b.periode[1]], "periode_label": b.periode_label, "ecole_label": b.ecole_label,
        "contrat": b.contrat, "bon_commande": b.bon_commande, "destinataire": b.destinataire, "emetteur": b.emetteur,
        "siret": b.info.get("SIRET"), "iban": b.info.get("IBAN"), "bic": b.info.get("BIC"), "nda": b.nda,
        "date_emission": b.date_emission, "echeance": b.echeance, "masquer_jour": b.masquer_jour,
        "regroupement": b.regroupement,
        "lignes": [
            {
                "cle_seance": l.cle_seance, "date": l.date,
                "debut": f"{l.debut:{_MINUTE}}", "fin": f"{l.fin:{_MINUTE}}",
                "heures": _repr(l.heures), "taux": _repr(l.taux), "promotion": l.promotion, "libelle": l.libelle,
                "intitule": l.intitule, "cle_tarif": l.cle_tarif,
            }
            for l in b.lignes
        ],
    })


def _ligne_lisible(l: LigneBrouillon) -> str:
    return f"{_jour_fr(l.date)} {_horaire(l.debut, l.fin)} « {l.intitule} »"


def _differences_brouillon(a: Brouillon, n: Brouillon) -> list[Difference]:
    ident = n.identifiant
    par_nature: dict[str, list[str]] = {nature: [] for nature in NATURES_DIFFERENCES}
    anciennes = {l.cle_seance: l for l in a.lignes}
    nouvelles = {l.cle_seance: l for l in n.lignes}
    paires = [(anciennes[c], nouvelles[c]) for c in sorted(set(anciennes) & set(nouvelles), key=_cle_chrono)]
    retirees = [anciennes[c] for c in sorted(set(anciennes) - set(nouvelles), key=_cle_chrono)]
    ajoutees = [nouvelles[c] for c in sorted(set(nouvelles) - set(anciennes), key=_cle_chrono)]
    for uid in sorted({l.uid for l in retirees} & {l.uid for l in ajoutees}):
        avant = [l for l in retirees if l.uid == uid]
        apres = [l for l in ajoutees if l.uid == uid]
        if len(avant) == 1 and len(apres) == 1:
            par_nature["séance déplacée"].append(f"{_ligne_lisible(avant[0])} → {_ligne_lisible(apres[0])}")
            retirees.remove(avant[0])
            ajoutees.remove(apres[0])
            paires.append((avant[0], apres[0]))
    for l in ajoutees:
        par_nature["séance ajoutée"].append(f"{_ligne_lisible(l)}, {_h(l.heures)}")
    for l in retirees:
        par_nature["séance retirée"].append(f"{_ligne_lisible(l)}, {_h(l.heures)}")
    taux: dict[tuple[float | None, float | None], int] = {}
    for avant, apres in paires:
        if avant.heures != apres.heures:
            par_nature["heures modifiées"].append(f"{_ligne_lisible(apres)} : {_h(avant.heures)} → {_h(apres.heures)}")
        if avant.taux != apres.taux:
            taux[(avant.taux, apres.taux)] = taux.get((avant.taux, apres.taux), 0) + 1
    for (avant, apres), nombre in taux.items():
        par_nature["taux modifié"].append(f"{_taux_txt(avant)} → {_taux_txt(apres)} ({_pluriel(nombre, 'séance')})")
    par_nature["contenu modifié"] += _contenu_modifie(a, n, paires)
    coordonnees = []
    if a.destinataire != n.destinataire:
        coordonnees.append("destinataire")
    if a.emetteur != n.emetteur:
        coordonnees.append("émetteur")
    for cle in ("SIRET", "IBAN", "BIC"):
        if a.info.get(cle) != n.info.get(cle):
            coordonnees.append(cle)
    if a.nda != n.nda:
        coordonnees.append("NDA")
    if coordonnees:
        par_nature["coordonnées modifiées"].append(", ".join(coordonnees) + " (onglet Coordonnées)")
    if a.echeance != n.echeance:
        par_nature["échéance modifiée"].append(
            f"{_date_fr(a.echeance) if a.echeance else '—'} → {_date_fr(n.echeance) if n.echeance else '—'}"
        )
    if a.contrat != n.contrat:
        par_nature["contrat modifié"].append(f"{a.contrat or '—'} → {n.contrat or '—'}")
    par_nature["nouveau blocage"] += [b for b in n.blocages if b not in a.blocages]
    par_nature["blocage levé"] += [b for b in a.blocages if b not in n.blocages]
    if a.total_ht != n.total_ht:
        par_nature["total modifié"].append(f"{_euros(a.total_ht)} → {_euros(n.total_ht)}")
    resultat = [Difference(ident, nature, detail) for nature in NATURES_DIFFERENCES for detail in par_nature[nature]]
    hors_empreinte = {"nouveau blocage", "blocage levé"}
    if a.empreinte != n.empreinte and all(d.nature in hors_empreinte for d in resultat):
        # filet : une empreinte changée (donc « à revalider ») a toujours sa différence
        detail = f"contenu de la facture modifié (empreinte {a.empreinte[:12]} → {n.empreinte[:12]}) : relisez l'aperçu"
        resultat.append(Difference(ident, "contenu modifié", detail))
    return resultat


def _citation(valeur: Any) -> str:
    return "—" if valeur in (None, "") else f"« {valeur} »"


def _contenu_modifie(a: Brouillon, n: Brouillon, paires: Sequence[tuple[LigneBrouillon, LigneBrouillon]]) -> list[str]:
    """Champs écrits sur la facture (dans l'empreinte, § 4.13) que les autres natures ne disent pas."""
    details = []
    if a.ecole_label != n.ecole_label:
        details.append(f"libellé de l'école : {_citation(a.ecole_label)} → {_citation(n.ecole_label)} (onglet Écoles)")
    if (a.periode, a.periode_label) != (n.periode, n.periode_label):
        details.append(f"période : du {_date_fr(a.periode[0])} au {_date_fr(a.periode[1])} « {a.periode_label} » → du "
                       f"{_date_fr(n.periode[0])} au {_date_fr(n.periode[1])} « {n.periode_label} »")
    if a.date_emission != n.date_emission:
        details.append(f"date d'émission : {_date_fr(a.date_emission)} → {_date_fr(n.date_emission)}")
    if a.bon_commande != n.bon_commande:
        details.append(f"bon de commande : {_citation(a.bon_commande)} → {_citation(n.bon_commande)}")
    if a.masquer_jour != n.masquer_jour:
        etat = {True: "masquée", False: "affichée"}
        details.append(f"colonne Jour : {etat[a.masquer_jour]} → {etat[n.masquer_jour]} (Règles_facturation)")
    if a.regroupement != n.regroupement:
        details.append(f"regroupement : {a.regroupement} → {n.regroupement} (Règles_facturation)")
    for avant, apres in paires:
        seance = f"séance du {_jour_fr(apres.date)} {_horaire(apres.debut, apres.fin)}"
        meme = avant.cle_seance == apres.cle_seance  # sinon « séance déplacée » montre déjà horaires et intitulé
        if meme and avant.intitule != apres.intitule:
            details.append(f"intitulé de la {seance} : {_citation(avant.intitule)} → {_citation(apres.intitule)}")
        if meme and avant.heures == apres.heures and (f"{avant.fin:{_MINUTE}}" != f"{apres.fin:{_MINUTE}}"):
            details.append(f"horaires de la {seance} : {_horaire(avant.debut, avant.fin)} → "
                           f"{_horaire(apres.debut, apres.fin)}")
        if avant.libelle != apres.libelle:
            details.append(f"libellé de la ligne (Module) de la {seance} : {_citation(avant.libelle)} → "
                           f"{_citation(apres.libelle)}")
        if avant.promotion != apres.promotion:
            details.append(f"promotion de la {seance} : {_citation(avant.promotion)} → {_citation(apres.promotion)}")
        if avant.cle_tarif != apres.cle_tarif:
            details.append(f"clé tarif de la {seance} : {_citation(avant.cle_tarif)} → {_citation(apres.cle_tarif)}")
    return details


def differences(ancienne: Preparation | None, nouvelle: Preparation) -> list[Difference]:
    """Changements des brouillons depuis la préparation précédente ; [] sans préparation précédente.

    Les brouillons s'apparient par identifiant, les lignes par `cle_seance` ; une ancienne ligne et
    une nouvelle, seules à porter le même `uid`, sont une « séance déplacée ». Un brouillon devenu
    engagé (émis) n'est pas une « facture disparue ». Ordre stable : identifiant, puis nature.
    """
    if ancienne is None:
        return []
    anciens = {b.identifiant: b for b in ancienne.brouillons}
    nouveaux = {b.identifiant: b for b in nouvelle.brouillons}
    emises = {d[0] for d in nouvelle.deja_emises}
    resultat: list[Difference] = []
    for ident in sorted(set(anciens) | set(nouveaux)):
        a, n = anciens.get(ident), nouveaux.get(ident)
        if a is None:
            detail = f"{_pluriel(len(n.lignes), 'séance')}, total {_euros(n.total_ht)}"
            resultat.append(Difference(ident, "facture nouvelle", detail))
        elif n is None:
            if ident not in emises:
                resultat.append(Difference(ident, "facture disparue",
                                           f"plus aucune séance à facturer (ancien total {_euros(a.total_ht)})"))
        else:
            resultat += _differences_brouillon(a, n)
    return resultat


def numeros_courts(preparation: Preparation) -> dict[str, str]:
    """« F<n> » → identifiant (ordre des brouillons) ; « S<n> » → `cle_seance` (ordre chronologique
    des séances facturées, déjà facturées et à classer), chacune une seule fois."""
    courts = {f"F{i}": b.identifiant for i, b in enumerate(preparation.brouillons, start=1)}
    cles = {l.cle_seance for b in preparation.brouillons for l in b.lignes}
    cles |= {d.cle_seance for d in preparation.deja_facturees}
    cles |= {chemins.cle_seance(a.seance) for a in preparation.a_classer}
    for i, cle in enumerate(sorted(cles, key=_cle_chrono), start=1):
        courts[f"S{i}"] = cle
    return courts


# ──────────────────────────────────────────────────────────────────────
# Aperçus
# ──────────────────────────────────────────────────────────────────────


def _etat_par_defaut(b: Brouillon) -> str:
    return journal.A_VALIDER if b.statut == PRET else journal.BLOQUEE


def _labels(b: Brouillon, numeros: Mapping[str, str] | None) -> tuple[str, dict[str, str]]:
    """(numéro court de la facture, numéro court de chaque séance)."""
    if numeros is None:
        return b.identifiant, {l.cle_seance: f"S{i}" for i, l in enumerate(b.lignes, start=1)}
    f_label = next((k for k, v in numeros.items() if k.startswith("F") and v == b.identifiant), b.identifiant)
    s_labels = {v: k for k, v in numeros.items() if k.startswith("S")}
    return f_label, s_labels


def _bloc_facture(b: Brouillon) -> list[str]:
    """Aperçu de la facture (PCF étape 6), par `facture.apercu` quand le brouillon est prêt."""
    fichier = f"{b.entite}/{b.ecole}/{b.entite}_{b.ecole}.xlsx"
    numero = f"attribué à l'émission (probable : {b.numero_probable or '—'})"
    infos = [f"SIRET         : {b.info.get('SIRET') or '—'}", f"IBAN          : {b.info.get('IBAN') or '—'}",
             f"BIC           : {b.info.get('BIC') or '—'}"]
    if b.statut == PRET and b.numero_probable and b.destinataire and b.emetteur:
        try:
            texte = facture.apercu(b.spec(b.numero_probable), b.groupes(), b.destinataire, emetteur=b.emetteur,
                                   echeance=b.echeance, nda=b.nda)
        except (ValueError, KeyError, TypeError):
            texte = None
        if texte is not None:
            lignes = []
            for ligne in texte.splitlines():
                if ligne.startswith("====="):
                    continue
                if ligne.startswith("N°"):
                    ligne = f"N°            : {numero}"
                elif ligne.startswith("Fichier"):
                    ligne = f"Fichier       : {fichier}"
                elif ligne.startswith("Onglet"):
                    ligne = "Onglet        : nommé à l'émission"
                elif ligne.startswith(" ") and lignes and lignes[-1].startswith("Onglet"):
                    continue
                lignes.append(ligne)
                if ligne.startswith("NDA"):
                    lignes += infos
            return lignes
    return _bloc_facture_bloquee(b, numero, fichier, infos)


def _personne(donnees: Mapping[str, str] | None) -> str:
    if not donnees:
        return "— (absent)"
    return f"{donnees.get('nom', '')} — {donnees.get('adresse_l1', '')}, {donnees.get('adresse_l2', '')}"


def _bloc_facture_bloquee(b: Brouillon, numero: str, fichier: str, infos: list[str]) -> list[str]:
    """Même aperçu que `facture.apercu`, montrant ce qui manque (« — ») quand le brouillon est bloqué."""
    seances = [(jour, promo, libelle, h, t) for promo, libelle, liste in b.groupes() for jour, h, t in liste]
    largeur_promo = max([12, *(len(s[1]) for s in seances)])
    largeur_libelle = max([0, *(len(s[2]) for s in seances)])
    detail = []
    for jour, promo, libelle, h, t in seances:
        montant = montant_exact(h, t) if h is not None and t is not None else None
        heures_txt = "—" if h is None else f"{h}h"
        taux_txt = "—" if t is None else f"{t:.2f} €"
        detail.append(f"  | {_date_fr(jour)} | {promo:<{largeur_promo}} | {libelle:<{largeur_libelle}} "
                      f"| {heures_txt:>5} | {taux_txt:>8} | {_euros(montant):>11}")
    return [
        f"Entité        : {b.entite}",
        f"N°            : {numero}",
        f"Date émission : {_date_fr(b.date_emission)}",
        f"Échéance      : {_date_fr(b.echeance) if b.echeance else '—'}",
        f"Émetteur      : {_personne(b.emetteur)}",
        f"NDA           : {b.nda or '—'}",
        *infos,
        f"Destinataire  : {_personne(b.destinataire)}",
        f"Bon de commande : {b.bon_commande or '—'}",
        f"Contrat       : {b.contrat or '—'}",
        f"Fichier       : {fichier}",
        "Onglet        : nommé à l'émission",
        "",
        "DÉTAIL DES SÉANCES :",
        *detail,
        "",
        f"TOTAL HEURES : {'—' if b.total_heures is None else f'{b.total_heures}h'}",
        f"TOTAL HT     : {_euros(b.total_ht)}",
    ]


def _liste(titre: str, elements: Sequence[str], vide: str = "aucun") -> list[str]:
    if not elements:
        return [f"{titre} : {vide}"]
    return [f"{titre} ({len(elements)}) :", *(f"  - {e}" for e in elements)]


def apercu_facture(brouillon: Brouillon, *, differences: Sequence[Difference] | None = (),
                   etat: str | None = None, numeros: Mapping[str, str] | None = None) -> str:
    """Aperçu d'un brouillon (PCF étape 6), à relire avant de valider.

    S'appuie sur `facture.apercu` (« N° : attribué à l'émission (probable : …) », « Onglet : nommé à
    l'émission ») et ajoute : période, SIRET, IBAN, BIC, NDA ; chaque séance avec son numéro court
    et la marque [tenue] ou [à venir] ; blocages, avertissements, différences (`None` : première
    préparation ; vide : aucune) ; état de validation ; empreinte et commande de validation.

    Args:
        numeros: les numéros courts de la préparation (`numeros_courts`) ; sans eux, les séances
            sont numérotées dans le brouillon seul, et la facture est désignée par son identifiant.
    """
    b = brouillon
    f_label, s_labels = _labels(b, numeros)
    etat = etat or _etat_par_defaut(b)
    du, au = b.periode
    lignes = [
        f"====== FACTURE {b.identifiant} ({f_label}) — {b.ecole_label} ======",
        f"État          : {etat} ({b.statut})",
        f"École         : {b.ecole_label} ({b.ecole})",
        f"Période       : du {_date_fr(du)} au {_date_fr(au)} — libellé « {b.periode_label} »",
        *_bloc_facture(b),
        "",
        f"SÉANCES DE L'AGENDA ({len(b.lignes)}) :",
    ]
    for l in b.lignes:
        marque = "[tenue]  " if l.tenue else "[à venir]"
        montant = "—" if l.montant is None else f"{arrondi_centime(l.montant)} €"
        lignes.append(
            f"  {s_labels.get(l.cle_seance, '?'):<5} {marque} {_jour_fr(l.date)} {_horaire(l.debut, l.fin)}  "
            f"{_h(l.heures)} × {_taux_txt(l.taux)} = {montant}  « {l.intitule} » (classée par {l.classee_par})"
        )
    lignes.append("")
    lignes += _liste("BLOCAGES", b.blocages)
    lignes += _liste("AVERTISSEMENTS", b.avertissements)
    if differences is None:
        lignes.append("DIFFÉRENCES : première préparation")
    else:
        lignes += _liste("DIFFÉRENCES", [f"{d.nature} : {d.detail}" for d in differences], vide="aucune")
    if b.statut == PRET:
        suite = (f"pour valider : facturation valider --mois {b.mois} --facture {f_label} --oui "
                 f"--empreinte {b.empreinte[:8]}")
    else:
        suite = "bloquée : levez les blocages, puis relancez facturation preparer"
    lignes.append(f"Empreinte : {b.empreinte[:12]} — {suite}")
    return "\n".join(lignes)


def _a_faire(preparation: Preparation, etats: Mapping[str, str], f_de: Mapping[str, str]) -> list[str]:
    mois = preparation.mois
    taches: list[str] = []

    def designation(identifiants: Sequence[str]) -> str:
        return ", ".join(f"{i} ({f_de.get(i, '?')})" for i in identifiants)

    a_valider = [b.identifiant for b in preparation.brouillons if etats.get(b.identifiant) == journal.A_VALIDER]
    if a_valider:
        taches.append(
            f"Valider {_pluriel(len(a_valider), 'facture')} : {designation(a_valider)} — relisez chaque aperçu, puis "
            f"`facturation valider --mois {mois} --facture F<n> --oui --empreinte <8 caractères>`"
        )
    for b in preparation.brouillons:
        if etats.get(b.identifiant) == journal.BLOQUEE and b.blocages:
            autres = ""
            if len(b.blocages) > 1:
                autres = f" (+ {_pluriel(len(b.blocages) - 1, 'autre blocage', 'autres blocages')})"
            taches.append(f"Débloquer {designation([b.identifiant])} : {b.blocages[0]}{autres}")
    if preparation.a_classer:
        taches.append(
            f"Classer {_pluriel(len(preparation.a_classer), 'créneau', 'créneaux')} : "
            f"`facturation classer --mois {mois}` (un créneau non classé n'est jamais facturé)"
        )
    for etat, verbe in ((journal.A_REVALIDER, "Revalider"), (journal.A_CORRIGER, "Corriger")):
        concernees = [b.identifiant for b in preparation.brouillons if etats.get(b.identifiant) == etat]
        if concernees:
            taches.append(f"{verbe} {_pluriel(len(concernees), 'facture')} : {designation(concernees)}")
    if preparation.ecoles_sans_regle:
        taches.append(
            f"Écoles sans règle : {', '.join(preparation.ecoles_sans_regle)} — ajoutez leur ligne à l'onglet "
            "Règles_facturation : leurs séances ne sont pas facturées"
        )
    for identifiant, n_facture, change in preparation.deja_emises:
        if change:
            taches.append(
                f"Recontrôler {identifiant} ({n_facture}) : séances modifiées depuis l'émission — "
                "`facturation recontroler` (F-23)"
            )
    if preparation.non_facturees_periode_emise:
        taches.append(
            f"{_pluriel(len(preparation.non_facturees_periode_emise), 'séance')} d'une période déjà émise, non "
            "facturée(s) : signalée(s), jamais refacturée(s) — voir « SÉANCES D'UNE PÉRIODE ÉMISE »"
        )
    for ecole in preparation.ecoles_non_preparees:
        taches.append(f"Préparer à la main : {ecole.ecole} ({ecole.raison})")
    for b in preparation.brouillons:
        for avertissement in b.avertissements:
            if avertissement.startswith((POSSIBLE_DOUBLON, PERIODE_PEUT_ETRE_FACTUREE)):
                taches.append(f"Vérifier {b.identifiant} : {avertissement}")
    sessions: dict[str, list[str]] = {}
    for b in preparation.brouillons:
        for avertissement in b.avertissements:
            if avertissement.startswith(SESSIONS_A_CORRIGER):
                detail = avertissement[len(SESSIONS_A_CORRIGER):].lstrip(" :")
                sessions.setdefault(detail, []).append(b.identifiant)
    for detail, concernees in sessions.items():
        taches.append(f"Corriger Sessions ({designation(concernees)}) : {detail}")
    return taches


def taches_a_faire(preparation: Preparation, *, etats: Mapping[str, str] | None = None) -> list[str]:
    """Tâches de la section « À FAIRE » (sans numérotation) ; liste vide : rien à faire.

    `facturation preparer` rend 3 dès qu'il en reste une (§ 1.6 : une action d'Adrian est requise).
    """
    etats = {b.identifiant: (etats or {}).get(b.identifiant) or _etat_par_defaut(b) for b in preparation.brouillons}
    f_de = {v: k for k, v in numeros_courts(preparation).items() if k.startswith("F")}
    return _a_faire(preparation, etats, f_de)


def apercu_global(preparation: Preparation, *, etats: Mapping[str, str] | None = None,
                  sections: Sequence[tuple[str, Sequence[str]]] = ()) -> str:
    """Aperçu du mois : « À FAIRE », tableau des factures, changements, détail de chaque facture,
    puis les sections finales et les sections annexes (`sections_annexes`)."""
    p = preparation
    etats = {b.identifiant: (etats or {}).get(b.identifiant) or _etat_par_defaut(b) for b in p.brouillons}
    numeros = numeros_courts(p)
    f_de = {v: k for k, v in numeros.items() if k.startswith("F")}
    s_de = {v: k for k, v in numeros.items() if k.startswith("S")}
    debut_mois, _ = chemins.mois_valide(p.mois)
    lignes = [
        f"Mois : {chemins.NOMS_MOIS[debut_mois.month - 1]} {debut_mois.year} — entité {p.entite} — préparé le "
        f"{p.prepare_le:%d/%m/%Y} à {p.prepare_le:%H:%M}",
        f"Agenda lu du {_date_fr(p.fenetre[0])} au {_date_fr(p.fenetre[1])}",
        "",
        "À FAIRE",
    ]
    taches = _a_faire(p, etats, f_de)
    lignes += [f"  {i}. {t}" for i, t in enumerate(taches, start=1)] or ["  Rien à faire."]

    lignes += ["", f"FACTURES ({len(p.brouillons)})"]
    if p.brouillons:
        lignes.append(f"  {'N°':<4} {'Identifiant':<28} {'Période':<25} {'Tenues':>6} {'À venir':>7} "
                      f"{'Heures':>7} {'Total HT':>13}  État")
        for b in p.brouillons:
            tenues = sum(1 for l in b.lignes if l.tenue)
            heures = "—" if b.total_heures is None else f"{b.total_heures}"
            lignes.append(
                f"  {f_de[b.identifiant]:<4} {b.identifiant:<28} {_date_fr(b.periode[0])} – {_date_fr(b.periode[1])} "
                f"{tenues:>6} {len(b.lignes) - tenues:>7} {heures:>7} {_euros(b.total_ht):>13}  {etats[b.identifiant]}"
            )
    else:
        lignes.append("  Aucune facture à préparer.")

    lignes.append("")
    if p.precedente_le is None:
        lignes.append("CHANGEMENTS : première préparation du mois")
    else:
        lignes.append(f"CHANGEMENTS DEPUIS LE {p.precedente_le:%d/%m/%Y} À {p.precedente_le:%H:%M}")
        lignes += [f"  - {d.identifiant} : {d.nature} — {d.detail}" for d in p.differences] or ["  aucun"]

    for b in p.brouillons:
        propres = [d for d in p.differences if d.identifiant == b.identifiant]
        lignes += ["", apercu_facture(b, differences=None if p.precedente_le is None else propres,
                                      etat=etats[b.identifiant], numeros=numeros)]

    lignes += ["", f"FILE À CLASSER ({len(p.a_classer)})"]
    for a in p.a_classer:
        s = a.seance
        lignes.append(f"  {s_de.get(chemins.cle_seance(s), '?'):<5} {_jour_fr(s.date)} {_horaire(s.debut, s.fin)}  "
                      f"« {s.intitule} » — {a.raison}")
    if p.a_classer:
        lignes.append(f"  → `facturation classer --mois {p.mois}` : un créneau non classé n'est jamais facturé")
    lignes += ["", f"DÉJÀ FACTURÉES ({len(p.deja_facturees)}) — retrouvées dans Sessions, retirées des brouillons"]
    for d in p.deja_facturees:
        suite = (f" — si ce n'est pas la même séance : `facturation corriger --mois {p.mois} --inclure "
                 f"{s_de.get(d.cle_seance, '?')}`") if d.certitude == MEME_JOUR else ""
        lignes.append(f"  {s_de.get(d.cle_seance, '?'):<5} {d.ecole} {_jour_fr(d.date)}  {d.n_facture or 'sans n°'} "
                      f"({d.certitude}){suite}")
    lignes += ["", f"EXCLUES PAR DÉCISION ({len(p.exclues)}) — ni facturées ni à classer"]
    for e in p.exclues:
        s = e.seance
        lignes.append(f"  {e.ecole or 'non reconnue'} {_jour_fr(s.date)} {_horaire(s.debut, s.fin)}  « {s.intitule} »"
                      f"{f' — {e.raison}' if e.raison else ''} (décision {e.decision})")
    lignes += ["", f"DÉJÀ ÉMISES ({len(p.deja_emises)}) — jamais régénérées"]
    for identifiant, n_facture, change in p.deja_emises:
        etat = ("contenu changé depuis l'émission : `facturation recontroler` (F-23)" if change
                else "inchangée depuis l'émission")
        lignes.append(f"  {identifiant} — {n_facture} — {etat}")
    lignes += ["", f"SÉANCES D'UNE PÉRIODE ÉMISE ({len(p.non_facturees_periode_emise)}) — signalées, "
                   "jamais refacturées"]
    for c in p.non_facturees_periode_emise:
        lignes.append(f"  {c.ecole} {_jour_fr(c.seance.date)} {_horaire(c.seance.debut, c.seance.fin)} "
                      f"« {c.seance.intitule} »")
    lignes += ["", f"ÉCOLES NON PRÉPARÉES ({len(p.ecoles_non_preparees)})"]
    lignes += [f"  {e.ecole} : {e.raison}" for e in p.ecoles_non_preparees]
    lignes += ["", f"ÉCOLES SANS SÉANCE ({len(p.ecoles_sans_seance)})"]
    if p.ecoles_sans_seance:
        lignes.append(f"  {', '.join(p.ecoles_sans_seance)}")
    compte = Counter(c.ecole for c in p.hors_lot)
    lignes += ["", f"HORS LOT ({len(compte)})"]
    lignes += [f"  {ecole} : {_pluriel(n, 'séance')}" for ecole, n in sorted(compte.items())]
    for titre, contenu in sections:
        lignes += ["", str(titre), *(f"  {ligne}" for ligne in contenu)]
    return "\n".join(lignes) + "\n"


def _texte_a_classer(preparation: Preparation) -> str:
    s_de = {v: k for k, v in numeros_courts(preparation).items() if k.startswith("S")}
    debut_mois, _ = chemins.mois_valide(preparation.mois)
    lignes = [
        f"Créneaux à classer — {chemins.NOMS_MOIS[debut_mois.month - 1]} {debut_mois.year} — "
        f"entité {preparation.entite}",
        "Un créneau non classé n'est jamais facturé (ADR028). "
        f"Pour les classer : facturation classer --mois {preparation.mois}",
        "",
    ]
    for a in preparation.a_classer:
        s = a.seance
        lignes.append(f"{s_de.get(chemins.cle_seance(s), '?'):<5} {_jour_fr(s.date)} {_horaire(s.debut, s.fin)}  "
                      f"« {s.intitule} » — {a.raison}")
    if not preparation.a_classer:
        lignes.append("Aucun créneau à classer.")
    return "\n".join(lignes) + "\n"


def sections_annexes(profil: "Profil", *, aujourd_hui: date) -> list[tuple[str, list[str]]]:
    """Sections ajoutées à l'aperçu global par d'autres modules, s'ils sont installés.

    Pour `facturation.registre` (F-25) et `facturation.apres_emission` (F-23) : `section_apercu(profil,
    aujourd_hui=…)` rend (titre, lignes), ou None (rien à dire). Un module absent est ignoré ; toute
    autre erreur donne la section « section indisponible : <type> ».
    """
    sections: list[tuple[str, list[str]]] = []
    for nom, titre in MODULES_ANNEXES:
        try:
            module = importlib.import_module(nom)
        except ModuleNotFoundError as exc:
            if exc.name == nom:
                continue
            sections.append((titre, [f"section indisponible : {type(exc).__name__}"]))
            continue
        except Exception as exc:  # noqa: BLE001 — une section annexe ne fait jamais échouer l'aperçu
            sections.append((titre, [f"section indisponible : {type(exc).__name__}"]))
            continue
        try:
            section = module.section_apercu(profil, aujourd_hui=aujourd_hui)
        except Exception as exc:  # noqa: BLE001
            sections.append((titre, [f"section indisponible : {type(exc).__name__}"]))
            continue
        if section is None:
            continue
        titre_section, lignes = section
        sections.append((str(titre_section), [str(l) for l in lignes]))
    return sections


# ──────────────────────────────────────────────────────────────────────
# preparation.json : écriture et relecture
# ──────────────────────────────────────────────────────────────────────


def _en_json_preparation(preparation: Preparation) -> dict[str, Any]:
    donnees = fichiers.en_json(preparation)
    for donnee, brouillon in zip(donnees["brouillons"], preparation.brouillons):
        donnee["statut"] = brouillon.statut
    return donnees


def _date(valeur: Any) -> date:
    return date.fromisoformat(str(valeur))


def _datetime(valeur: Any) -> datetime:
    return datetime.fromisoformat(str(valeur))


def _float(valeur: Any) -> float | None:
    if valeur is None:
        return None
    if isinstance(valeur, bool) or not isinstance(valeur, (int, float)):
        raise TypeError(f"nombre attendu, lu {valeur!r}")
    return float(valeur)


def _decimal(valeur: Any) -> Decimal | None:
    return None if valeur is None else Decimal(str(valeur))


def _seance_de(d: Mapping[str, Any]) -> Seance:
    return Seance(uid=str(d["uid"]), date=_date(d["date"]), debut=_datetime(d["debut"]), fin=_datetime(d["fin"]),
                  duree_h=float(d["duree_h"]), intitule=str(d["intitule"]), description=str(d["description"]))


def _classee_de(d: Mapping[str, Any]) -> SeanceClassee:
    return SeanceClassee(seance=_seance_de(d["seance"]), ecole=str(d["ecole"]), module=str(d["module"]),
                         promotion=str(d["promotion"]), cle_tarif=str(d["cle_tarif"]),
                         correspondance=dict(d["correspondance"]), par=str(d["par"]))


def _ligne_de(d: Mapping[str, Any]) -> LigneBrouillon:
    return LigneBrouillon(
        cle_seance=str(d["cle_seance"]), uid=str(d["uid"]), date=_date(d["date"]), debut=_datetime(d["debut"]),
        fin=_datetime(d["fin"]), intitule=str(d["intitule"]), ecole=str(d["ecole"]), promotion=str(d["promotion"]),
        module=str(d["module"]), libelle=str(d["libelle"]), cle_tarif=str(d["cle_tarif"]),
        classee_par=str(d["classee_par"]), duree_agenda_h=float(d["duree_agenda_h"]), heures=_float(d["heures"]),
        taux=_float(d["taux"]), montant=_decimal(d["montant"]), ref_contrat=d["ref_contrat"], tenue=bool(d["tenue"]),
    )


def _dict_ou_none(valeur: Any) -> dict | None:
    return None if valeur is None else {str(k): str(v) for k, v in dict(valeur).items()}


def _brouillon_de(d: Mapping[str, Any]) -> Brouillon:
    return Brouillon(
        identifiant=str(d["identifiant"]), entite=str(d["entite"]), ecole=str(d["ecole"]), mois=str(d["mois"]),
        periode=(_date(d["periode"][0]), _date(d["periode"][1])), periode_label=str(d["periode_label"]),
        ecole_label=str(d["ecole_label"]), contrat=d["contrat"], bon_commande=d["bon_commande"],
        destinataire=_dict_ou_none(d["destinataire"]), emetteur=_dict_ou_none(d["emetteur"]),
        info=_dict_ou_none(d["info"]) or {}, nda=d["nda"], date_emission=_date(d["date_emission"]),
        echeance=_date(d["echeance"]) if d["echeance"] else None, masquer_jour=bool(d["masquer_jour"]),
        regroupement=str(d["regroupement"]), lignes=tuple(_ligne_de(l) for l in d["lignes"]),
        total_heures=_float(d["total_heures"]), total_ht=_decimal(d["total_ht"]),
        numero_probable=d["numero_probable"], blocages=tuple(str(x) for x in d["blocages"]),
        avertissements=tuple(str(x) for x in d["avertissements"]), empreinte=str(d["empreinte"]),
    )


def _preparation_de(d: Mapping[str, Any]) -> Preparation:
    return Preparation(
        format=int(d["format"]), entite=str(d["entite"]), mois=str(d["mois"]), prepare_le=_datetime(d["prepare_le"]),
        version_outil=str(d["version_outil"]), fenetre=(_date(d["fenetre"][0]), _date(d["fenetre"][1])),
        empreinte_agenda=str(d["empreinte_agenda"]),
        brouillons=tuple(_brouillon_de(b) for b in d["brouillons"]),
        a_classer=tuple(
            SeanceAClasser(seance=_seance_de(a["seance"]), raison=str(a["raison"])) for a in d["a_classer"]
        ),
        hors_lot=tuple(_classee_de(c) for c in d["hors_lot"]),
        ecoles_sans_regle=tuple(str(e) for e in d["ecoles_sans_regle"]),
        deja_facturees=tuple(
            SeanceDejaFacturee(cle_seance=str(x["cle_seance"]), ecole=str(x["ecole"]), date=_date(x["date"]),
                               n_facture=str(x["n_facture"]), certitude=str(x["certitude"]))
            for x in d["deja_facturees"]
        ),
        deja_emises=tuple((str(i), str(n), bool(c)) for i, n, c in d["deja_emises"]),
        non_facturees_periode_emise=tuple(_classee_de(c) for c in d["non_facturees_periode_emise"]),
        ecoles_sans_seance=tuple(str(e) for e in d["ecoles_sans_seance"]),
        ecoles_non_preparees=tuple(
            EcoleNonPreparee(str(e["ecole"]), str(e["raison"])) for e in d["ecoles_non_preparees"]
        ),
        differences=tuple(
            Difference(str(x["identifiant"]), str(x["nature"]), str(x["detail"])) for x in d["differences"]
        ),
        precedente_le=_datetime(d["precedente_le"]) if d.get("precedente_le") else None,
        exclues=tuple(
            SeanceExclue(cle_seance=str(x["cle_seance"]), ecole=str(x["ecole"]), seance=_seance_de(x["seance"]),
                         raison=str(x["raison"]), decision=str(x["decision"]))
            for x in d.get("exclues") or []
        ),
    )


def lire_preparation(profil: "Profil", mois: str) -> Preparation | None:
    """La préparation du mois (`_sorties/<mois>/preparation.json`), ou None s'il n'y en a pas.

    Raises:
        PreparationIllisible: fichier abîmé ou incomplet, format inconnu, préparation d'une autre
            entité, ou brouillon dont l'empreinte ne correspond plus au contenu (fichier modifié à
            la main) ; le message nomme le fichier.
    """
    chemin = chemins.dossier_mois(profil, mois) / NOM_PREPARATION
    if not chemin.is_file():
        return None
    relancer = f"relancez `facturation preparer --mois {mois}`"
    try:
        donnees = fichiers.lire_json(chemin)
    except fichiers.FichierAbime as exc:
        raise PreparationIllisible(f"Préparation illisible : {chemin} ({exc}) : {relancer}.") from exc
    if not isinstance(donnees, dict) or donnees.get("format") != FORMAT_PREPARATION:
        trouve = donnees.get("format") if isinstance(donnees, dict) else type(donnees).__name__
        raise PreparationIllisible(
            f"Préparation {chemin} : format {trouve!r} inconnu (attendu {FORMAT_PREPARATION}) : {relancer}."
        )
    if donnees.get("entite") != profil.entite:
        raise PreparationIllisible(
            f"Préparation {chemin} : entité « {donnees.get('entite')} », alors que le profil est celui de "
            f"« {profil.entite} ». Les entités ne se mélangent jamais (ADR001)."
        )
    try:
        preparation = _preparation_de(donnees)
    except (KeyError, IndexError, TypeError, ValueError, InvalidOperation) as exc:
        raise PreparationIllisible(
            f"Préparation {chemin} incomplète ou abîmée ({type(exc).__name__} : {exc}) : {relancer}."
        ) from exc
    for b in preparation.brouillons:
        if empreinte_brouillon(b) != b.empreinte:
            raise PreparationIllisible(
                f"Préparation {chemin} : le brouillon {b.identifiant} ne correspond plus à son empreinte (fichier "
                f"modifié à la main ?) : {relancer}."
            )
    return preparation


def etats_de_validation(profil: "Profil", preparation: Preparation) -> dict[str, str]:
    """État de validation de chaque brouillon (`journal.etat_validation`), d'après les validations
    du journal du mois et les pièces d'émission."""
    evenements = journal.lire_evenements(chemins.journal_mois(profil, preparation.mois), types=["validation"])
    dernieres = journal.dernieres_validations(evenements)
    pieces = {
        p.identifiant: p for p in journal.pieces_emission(profil)
        if p.mois == preparation.mois and p.statut != journal.ABANDONNEE
    }
    return {
        b.identifiant: journal.etat_validation(bloque=b.statut == BLOQUE, empreinte=b.empreinte,
                                               derniere=dernieres.get(b.identifiant), piece=pieces.get(b.identifiant))
        for b in preparation.brouillons
    }


# ──────────────────────────────────────────────────────────────────────
# Lancement (sous verrou) et écritures
# ──────────────────────────────────────────────────────────────────────

# Un `TypeError` (cellule d'un type inattendu, par exemple) échoue proprement lui aussi : événement
# `échec`, préparation précédente intacte, jamais de trace Python brute.
_ECHECS = (AgendaIndisponible, ValueError, OSError, TypeError, MasterInvalide, MasterVerrouille)


def _remettre_octets(chemin: Path, octets: bytes) -> None:
    """Remet le contenu d'origine d'un fichier, de façon atomique (temporaire du même dossier)."""
    chemin.parent.mkdir(parents=True, exist_ok=True)
    temporaire = chemin.with_name(f"{chemin.name}.restauration-{os.getpid()}")
    try:
        with open(temporaire, "wb") as flux:
            flux.write(octets)
            flux.flush()
            os.fsync(flux.fileno())
        os.replace(temporaire, chemin)
    finally:
        if temporaire.exists():
            temporaire.unlink()


class _Instantane:
    """Contenu, avant écriture, des fichiers qu'une préparation remplace ou supprime, pour les remettre
    à l'identique, octet pour octet, si la suite échoue (§ 4 : « la préparation précédente reste
    identique octet pour octet »)."""

    def __init__(self, chemins_touches: Iterable[Path], dossiers: Iterable[Path]) -> None:
        self.avant = {chemin: (chemin.read_bytes() if chemin.is_file() else None) for chemin in chemins_touches}
        self.dossiers = [dossier for dossier in dossiers if not dossier.exists()]
        self.crees: list[Path] = []

    def restaurer(self) -> None:
        """Remet chaque fichier dans son état d'avant ; au mieux, sans jamais masquer l'erreur d'origine."""
        for chemin in reversed(self.crees):
            try:
                chemin.unlink(missing_ok=True)
            except OSError:
                pass
        for chemin, octets in self.avant.items():
            try:
                if octets is None:
                    chemin.unlink(missing_ok=True)
                elif not chemin.is_file() or chemin.read_bytes() != octets:
                    _remettre_octets(chemin, octets)
            except OSError:
                pass
        for dossier in reversed(self.dossiers):
            try:
                dossier.rmdir()  # seulement s'il est resté vide
            except OSError:
                pass


def _ecrire_sorties(profil: "Profil", preparation: Preparation, ancienne: Preparation | None,
                    etats: Mapping[str, str], annexes: Sequence[tuple[str, Sequence[str]]]) -> _Instantane:
    """Écritures atomiques sous `_sorties/<mois>/` ; `preparation.json` en dernier.

    Tout est calculé avant la première écriture. Si une écriture échoue, chaque fichier remplacé ou
    supprimé est remis tel qu'il était, et la copie ajoutée à `historique/` est retirée, avant que
    l'erreur remonte. Rend l'instantané, pour que l'appelant annule aussi sur un échec qui suit
    (événement du journal).
    """
    dossier = chemins.dossier_mois(profil, preparation.mois)
    chemin = dossier / NOM_PREPARATION
    apercus = dossier / DOSSIER_APERCUS
    historique = dossier / DOSSIER_HISTORIQUE
    numeros = numeros_courts(preparation)
    textes: dict[Path, str] = {}
    for b in preparation.brouillons:
        propres = [d for d in preparation.differences if d.identifiant == b.identifiant]
        texte = apercu_facture(b, differences=None if preparation.precedente_le is None else propres,
                               etat=etats.get(b.identifiant), numeros=numeros)
        textes[apercus / f"{b.identifiant}.txt"] = texte + "\n"
    textes[dossier / NOM_APERCU] = apercu_global(preparation, etats=etats, sections=annexes)
    textes[dossier / NOM_A_CLASSER] = _texte_a_classer(preparation)
    donnees = _en_json_preparation(preparation)
    perimes = sorted(f for f in apercus.glob("*.txt") if f not in textes) if apercus.is_dir() else []

    instantane = _Instantane([chemin, *textes, *perimes], [dossier, apercus, historique])
    try:
        dossier.mkdir(parents=True, exist_ok=True)
        if chemin.is_file():
            horodatage = ancienne.prepare_le if ancienne is not None else datetime.now()
            instantane.crees.append(fichiers.sauvegarder_fichier(chemin, historique, horodatage=horodatage))
        for cible, texte in textes.items():
            fichiers.ecrire_texte_atomique(cible, texte)
        for perime in perimes:
            perime.unlink()
        fichiers.ecrire_json_atomique(chemin, donnees)
    except BaseException:
        instantane.restaurer()
        raise
    return instantane


def _codes_selectionnes(ecoles: Sequence[str] | None) -> set[str] | None:
    return None if ecoles is None else {_code(e) for e in ecoles if _code(e)}


def _fusionner(partielle: Preparation, ancienne: Preparation | None, selection: set[str]) -> Preparation:
    """Préparation partielle (`--ecoles`) complétée par l'état précédent des autres écoles.

    Les écoles hors de la sélection gardent leurs brouillons tels quels (même empreinte : leur
    validation reste valable), leurs séances déjà facturées, leurs pièces et leurs signalements ;
    les créneaux et exclusions hors de la nouvelle fenêtre de lecture sont gardés aussi. Ce que la
    nouvelle lecture a vu (fenêtre, écoles sélectionnées) remplace l'ancien.
    """
    if ancienne is None:
        return partielle
    du, au = partielle.fenetre

    def gardee(ecole: str) -> bool:
        return ecole not in selection

    def hors_fenetre(jour: date) -> bool:
        return not du <= jour <= au

    vues = {chemins.cle_seance(a.seance) for a in partielle.a_classer}
    a_classer = [*partielle.a_classer, *(a for a in ancienne.a_classer
                                          if hors_fenetre(a.seance.date) and chemins.cle_seance(a.seance) not in vues)]
    exclues_vues = {e.cle_seance for e in partielle.exclues}
    exclues = [*partielle.exclues, *(e for e in ancienne.exclues
                                     if hors_fenetre(e.seance.date) and e.cle_seance not in exclues_vues)]
    return replace(
        partielle,
        fenetre=(min(du, ancienne.fenetre[0]), max(au, ancienne.fenetre[1])),
        brouillons=tuple(sorted([*partielle.brouillons, *(b for b in ancienne.brouillons if gardee(b.ecole))],
                                key=lambda b: b.identifiant)),
        a_classer=tuple(sorted(a_classer, key=lambda a: (a.seance.debut, chemins.cle_seance(a.seance)))),
        deja_facturees=tuple(sorted([*partielle.deja_facturees, *(d for d in ancienne.deja_facturees
                                                                   if gardee(d.ecole))],
                                    key=lambda d: _cle_chrono(d.cle_seance))),
        deja_emises=tuple(sorted([*partielle.deja_emises, *(d for d in ancienne.deja_emises
                                                             if gardee(d[0].split("-")[0]))])),
        non_facturees_periode_emise=tuple(sorted(
            [*partielle.non_facturees_periode_emise, *(c for c in ancienne.non_facturees_periode_emise
                                                        if gardee(c.ecole))],
            key=lambda c: (c.seance.debut, chemins.cle_seance(c.seance)))),
        ecoles_sans_seance=tuple(sorted({*partielle.ecoles_sans_seance,
                                         *(e for e in ancienne.ecoles_sans_seance if gardee(e))})),
        ecoles_non_preparees=tuple(sorted([*partielle.ecoles_non_preparees,
                                           *(e for e in ancienne.ecoles_non_preparees if gardee(e.ecole))],
                                          key=lambda e: e.ecole)),
        exclues=tuple(sorted(exclues, key=lambda e: _cle_chrono(e.cle_seance))),
    )


def _lancer_sous_verrou(profil: "Profil", mois: str, *, contenu_ics: bytes | None, maintenant: datetime,
                        ecoles: Sequence[str] | None, poste: str) -> Preparation:
    chemin_journal = chemins.journal_mois(profil, mois)
    try:
        entrees = charger_entrees(profil, mois, contenu_ics=contenu_ics, ecoles=ecoles)
        nouvelle = preparer(entite=profil.entite, mois=mois, entrees=entrees, prefixe=profil.prefixe_facture,
                            maintenant=maintenant, ecoles=ecoles)
        try:
            ancienne = lire_preparation(profil, mois)
        except PreparationIllisible:
            ancienne = None  # regénérée ; l'ancienne est rangée dans historique/
        selection = _codes_selectionnes(ecoles)
        if selection is not None:
            nouvelle = _fusionner(nouvelle, ancienne, selection)
        nouvelle = replace(nouvelle, differences=tuple(differences(ancienne, nouvelle)),
                           precedente_le=ancienne.prepare_le if ancienne is not None else None)
        etats = etats_de_validation(profil, nouvelle)
        annexes = sections_annexes(profil, aujourd_hui=maintenant.date())
        instantane = _ecrire_sorties(profil, nouvelle, ancienne, etats, annexes)
        try:
            journal.ajouter_evenement(
                chemin_journal, "preparation", poste=poste, entite=profil.entite, mois=mois, statut="ok",
                brouillons=[b.identifiant for b in nouvelle.brouillons],
                bloques=[b.identifiant for b in nouvelle.brouillons if b.statut == BLOQUE],
                a_classer=[chemins.cle_seance(a.seance) for a in nouvelle.a_classer], erreur=None,
            )
        except BaseException:
            instantane.restaurer()  # l'issue annoncée (échec) correspond au disque
            raise
    except _ECHECS as exc:
        message = fichiers.masquer_secrets(str(exc)) or type(exc).__name__
        try:
            journal.ajouter_evenement(chemin_journal, "preparation", poste=poste, entite=profil.entite, mois=mois,
                                      statut="échec", brouillons=[], bloques=[], a_classer=[], erreur=message)
        except (OSError, ValueError, TypeError):
            pass  # le journal ne masque jamais l'erreur d'origine
        raise PreparationImpossible(message) from exc
    return nouvelle


def lancer_preparation(profil: "Profil", mois: str, *, contenu_ics: bytes | None = None,
                       maintenant: datetime | None = None, ecoles: Sequence[str] | None = None) -> Preparation:
    """Prépare le mois sous le verrou d'entité et écrit les sorties du mois (aucune écriture dans le MASTER).

    En cas d'échec (agenda, MASTER, règles, verrou, disque, journal…), l'événement `preparation` /
    `échec` est ajouté au journal du mois (message masqué) et la préparation précédente reste
    identique octet pour octet.

    Avec `ecoles` (préparation partielle, `preparer --ecoles`), seules ces écoles sont recalculées :
    les autres gardent leur état de la préparation précédente (brouillons, empreintes, aperçus).

    Raises:
        PreparationImpossible: racine dans un dépôt git ou introuvable, `ecoles` sans aucun code,
            verrou tenu ailleurs, ou échec de la préparation ; l'exception d'origine est enchaînée.
    """
    try:
        chemins.mois_valide(mois)
        chemins.refuser_racine_dans_un_depot(profil.racine)
    except ValueError as exc:
        raise PreparationImpossible(str(exc)) from exc
    if ecoles is not None and not _codes_selectionnes(ecoles):
        raise PreparationImpossible("aucune école donnée : indiquez au moins un code (ex. --ecoles ABC,ECO)")
    if not Path(profil.racine).is_dir():
        raise PreparationImpossible(
            f"dossier de données introuvable : {profil.racine} ; si c'est un lecteur Google Drive, vérifiez qu'il "
            "est lancé et la session ouverte"
        )
    if not profil.dossier_entite.is_dir():
        raise PreparationImpossible(
            f"dossier de l'entité {profil.entite} introuvable : {profil.dossier_entite} (il porte le MASTER "
            f"{profil.master.name})"
        )
    fmt.set_data_root(profil.racine)
    if maintenant is None:
        maintenant = datetime.now(ZoneInfo(profil.fuseau))
    poste = master.poste_courant()
    try:
        with master.verrou_ecriture(profil.master, poste=poste):
            return _lancer_sous_verrou(profil, mois, contenu_ics=contenu_ics, maintenant=maintenant,
                                       ecoles=ecoles, poste=poste)
    except (MasterVerrouille, OSError) as exc:  # le verrou n'a pas pu être posé
        raise PreparationImpossible(f"préparation impossible : {exc}") from exc
