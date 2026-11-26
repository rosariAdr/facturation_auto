"""Données fictives partagées par les tests de la préparation multi-écoles (F-16a, banque T-PRP).

Aucune donnée réelle : écoles `ABC`, `ECO`, `FUT` (« Plus tard »), `SRG` (sans règle), entités
`ROSARI` et `KAUR`, coordonnées, taux et montants inventés. L'agenda est un contenu .ics fictif
construit à la volée et injecté (`contenu_ics`) : aucun accès réseau.

Jeu standard (mois 2026-12, préparé le 20/12/2026 à 16 h, Europe/Paris) :

| École | Règle | Séances de décembre | Taux | Attendu |
|---|---|---|---|---|
| ABC | Lot 1, mensuel, « 45 jours fin de mois » | 03/12, 10/12 (tenues), 22/12 (à venir), 9 h–12 h, code `#ABC-M1-STATS` | 60 | 9 h, 540,00 € |
| ECO | Lot 1, mensuel, délai vide (30 jours) | 08/12 (tenue), 28/12 (à venir), 14 h–17 h, intitulé exact | « 45 €/h » (texte, § 15) | 6 h, 270,00 € |
| FUT | Plus tard | 15/12 | 50 | hors lot |
| SRG | aucune ligne de `Règles_facturation` | 16/12 | 40 | école sans règle |
| — | — | 11/12 « Réunion pédagogique » | — | à classer |

Ce module n'importe pas `facturation.preparation` : son absence ne bloque pas la collecte.
"""

from __future__ import annotations

import hashlib
import importlib
from collections import namedtuple
from collections.abc import Iterable, Mapping, Sequence
from datetime import datetime
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

from openpyxl import load_workbook

from facturation import chemins, config, fichiers, journal
from facturation import pcf_format as fmt
from tests.conftest import NDA_FICTIF, creer_master

FUSEAU = "Europe/Paris"
PARIS = ZoneInfo(FUSEAU)
MOIS = "2026-12"
MAINTENANT = datetime(2026, 12, 20, 16, 0, tzinfo=PARIS)
POSTE_TEST = "POSTE-TEST"

SeanceFictive = namedtuple("SeanceFictive", "uid jour debut fin intitule description", defaults=("",))

SEANCES_STANDARD: tuple[SeanceFictive, ...] = (
    SeanceFictive("abc-1", "2026-12-03", "09:00", "12:00", "Statistiques M1", "#ABC-M1-STATS"),
    SeanceFictive("abc-2", "2026-12-10", "09:00", "12:00", "Statistiques M1", "#ABC-M1-STATS"),
    SeanceFictive("abc-3", "2026-12-22", "09:00", "12:00", "Statistiques M1", "#ABC-M1-STATS"),
    SeanceFictive("eco-1", "2026-12-08", "14:00", "17:00", "Économie L3"),
    SeanceFictive("eco-2", "2026-12-28", "14:00", "17:00", "Économie L3"),
    SeanceFictive("fut-1", "2026-12-15", "10:00", "12:00", "Projet FUT"),
    SeanceFictive("srg-1", "2026-12-16", "10:00", "12:00", "Atelier SRG"),
    SeanceFictive("inconnu-1", "2026-12-11", "13:00", "14:00", "Réunion pédagogique"),
)

# Coordonnées fictives écrites par `tests.conftest.creer_master`.
SIRET_FICTIF = "00000000000000"
IBAN_FICTIF = "FR76 0000 0000 0000 0000 0000 000"
BIC_FICTIF = "TESTFRPP"


def prp():
    """Module testé, importé dans le test : son absence fait échouer le test, pas la collecte."""
    return importlib.import_module("facturation.preparation")


# ──────────────────────────────────────────────────────────────────────
# Agenda fictif
# ──────────────────────────────────────────────────────────────────────

_VTIMEZONE = (
    "BEGIN:VTIMEZONE", "TZID:Europe/Paris",
    "BEGIN:DAYLIGHT", "TZOFFSETFROM:+0100", "TZOFFSETTO:+0200", "TZNAME:CEST",
    "DTSTART:19700329T020000", "RRULE:FREQ=YEARLY;BYMONTH=3;BYDAY=-1SU", "END:DAYLIGHT",
    "BEGIN:STANDARD", "TZOFFSETFROM:+0200", "TZOFFSETTO:+0100", "TZNAME:CET",
    "DTSTART:19701025T030000", "RRULE:FREQ=YEARLY;BYMONTH=10;BYDAY=-1SU", "END:STANDARD",
    "END:VTIMEZONE",
)


def _horodatage_ics(jour: str, heure: str) -> str:
    return f"{jour.replace('-', '')}T{heure.replace(':', '')}00"


def ics(seances: Iterable[SeanceFictive] = SEANCES_STANDARD) -> bytes:
    """Contenu .ics fictif (heures de Paris) des séances données."""
    lignes = ["BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Exemple fictif//Tests F-16a//FR", *_VTIMEZONE]
    for s in seances:
        lignes += [
            "BEGIN:VEVENT",
            f"UID:{s.uid}",
            "DTSTAMP:20261101T080000Z",
            f"DTSTART;TZID=Europe/Paris:{_horodatage_ics(s.jour, s.debut)}",
            f"DTEND;TZID=Europe/Paris:{_horodatage_ics(s.jour, s.fin)}",
            f"SUMMARY:{s.intitule}",
        ]
        if s.description:
            lignes.append(f"DESCRIPTION:{s.description}")
        lignes.append("END:VEVENT")
    lignes.append("END:VCALENDAR")
    return ("\r\n".join(lignes) + "\r\n").encode("utf-8")


def debut(s: SeanceFictive) -> datetime:
    return datetime.strptime(f"{s.jour} {s.debut}", "%Y-%m-%d %H:%M").replace(tzinfo=PARIS)


def cle(s: SeanceFictive) -> str:
    """`cle_seance` attendue (`chemins.cle_seance`) : `<uid>@<début AAAA-MM-JJTHH:MM±hhmm>`."""
    return f"{s.uid}@{debut(s):%Y-%m-%dT%H:%M%z}"


def standard(uid: str) -> SeanceFictive:
    return next(s for s in SEANCES_STANDARD if s.uid == uid)


def seances_sauf(*uids: str) -> list[SeanceFictive]:
    return [s for s in SEANCES_STANDARD if s.uid not in uids]


# ──────────────────────────────────────────────────────────────────────
# MASTER fictif
# ──────────────────────────────────────────────────────────────────────


def regle(ecole: str, statut: str = "Lot 1", **colonnes: Any) -> dict[str, Any]:
    """Ligne de `Règles_facturation` ; les autres colonnes se passent par leur nom :
    `regle("ABC", **{"Délai de paiement": "30 jours"})`."""
    return {"École": ecole, "Statut lot": statut, **colonnes}


def correspondance(cle_: str, type_cle: str, ecole: str, module: str, promotion: str, *,
                   cle_tarif: str | None = None, mode: str | None = None, valeur: Any = None,
                   gardee_le: Any = None) -> dict[str, Any]:
    return {
        "Clé": cle_, "Type de clé": type_cle, "École": ecole, "Module": module, "Promotion": promotion,
        "Clé tarif": cle_tarif or f"{ecole}|{module}|{promotion}", "Mode heures": mode, "Valeur": valeur,
        "Gardée le": gardee_le,
    }


def tarif(ecole: str, module: str, classe: str, taux: Any, ref_contrat: str | None = None) -> dict[str, Any]:
    return {"École": ecole, "Module / Matière": module, "Classe / Programme": classe, "Taux HT/h": taux,
            "Réf. contrat": ref_contrat}


def regles_standard() -> list[dict[str, Any]]:
    return [
        regle("ABC", **{"Délai de paiement": "45 jours fin de mois"}),
        regle("ECO"),
        regle("FUT", "Plus tard"),
    ]


def ecoles_standard() -> list[dict[str, Any]]:
    return [
        {"École": "ÉCOLE FICTIVE A", "Code facturation": "ABC"},
        {"École": "ÉCOLE FICTIVE B", "Code facturation": "ECO"},
        {"École": "ÉCOLE FICTIVE FUTURE", "Code facturation": "FUT"},
        {"École": "ÉCOLE FICTIVE SANS RÈGLE", "Code facturation": "SRG"},
    ]


def tarifs_standard() -> list[dict[str, Any]]:
    return [
        tarif("ABC", "Statistiques", "M1", 60),
        tarif("ECO", "Économie", "L3", "45 €/h"),
        tarif("FUT", "Projet", "B2", 50),
        tarif("SRG", "Atelier", "B1", 40),
    ]


def correspondances_standard() -> list[dict[str, Any]]:
    return [
        correspondance("#ABC-M1-STATS", "code", "ABC", "Statistiques", "M1"),
        correspondance("Économie L3", "exact", "ECO", "Économie", "L3"),
        correspondance("Projet FUT", "exact", "FUT", "Projet", "B2"),
        correspondance("Atelier SRG", "exact", "SRG", "Atelier", "B1"),
    ]


def master_standard(racine: Path, entite: str = "ROSARI", **surcharges: Any) -> Path:
    """MASTER fictif du jeu standard ; chaque onglet peut être remplacé (`regles=`, `tarifs=`…)."""
    donnees: dict[str, Any] = {
        "regles": regles_standard(),
        "correspondances": correspondances_standard(),
        "tarifs": tarifs_standard(),
        "ecoles": ecoles_standard(),
        "sessions": [],
        "factures": None,
        "nda": NDA_FICTIF,
    }
    donnees.update(surcharges)
    return creer_master(racine, entite, **donnees)


def profil(tmp_path: Path, racine: Path, entite: str = "ROSARI") -> config.Profil:
    """Profil fictif rangé hors de la racine de données (et hors de tout dépôt)."""
    dossier = tmp_path / f"profil_{entite}"
    dossier.mkdir(parents=True, exist_ok=True)
    chemin = dossier / "profil.toml"
    chemin.write_text(config.rediger("Personne Fictive", entite, str(racine)), encoding="utf-8")
    return config.charger(chemin)


def installer(tmp_path: Path, entite: str = "ROSARI", **surcharges: Any) -> config.Profil:
    """Racine fictive (`tmp_path/donnees`), MASTER du jeu standard et profil."""
    racine = tmp_path / "donnees"
    master_standard(racine, entite, **surcharges)
    return profil(tmp_path, racine, entite)


def _feuille_coordonnees(chemin_master: Path):
    wb = load_workbook(chemin_master)
    return wb, wb["Coordonnées"]


def changer_coordonnee(chemin_master: Path, type_: str, code: str, colonne: str, valeur: Any) -> None:
    """Modifie une cellule de la ligne (Type, Code) de `Coordonnées`, par le nom de la colonne."""
    wb, ws = _feuille_coordonnees(chemin_master)
    entetes = [c.value for c in ws[1]]
    indice = entetes.index(colonne) + 1
    for ligne in ws.iter_rows(min_row=2):
        if ligne[0].value == type_ and ligne[1].value == code:
            ws.cell(row=ligne[0].row, column=indice, value=valeur)
            break
    else:
        raise AssertionError(f"ligne {type_} / {code} absente de Coordonnées")
    wb.save(chemin_master)


def retirer_info(chemin_master: Path, entite: str, nom: str) -> None:
    """Retire la ligne `Info` / `<nom>` (SIRET, IBAN, BIC, NDA) de l'entité."""
    wb, ws = _feuille_coordonnees(chemin_master)
    for ligne in ws.iter_rows(min_row=2):
        if ligne[0].value == "Info" and ligne[1].value == entite and ligne[2].value == nom:
            ws.delete_rows(ligne[0].row)
            break
    else:
        raise AssertionError(f"ligne Info / {nom} absente de Coordonnées")
    wb.save(chemin_master)


def ligne_sessions(s: SeanceFictive, *, ecole: str, n_facture: str, debut_: Any = None, fin_: Any = None,
                   heures: float = 3, taux: float = 60) -> dict[str, Any]:
    """Ligne fictive de `Sessions` pour la séance ; `debut_`/`fin_` : texte « HH:MM » ou None."""
    return {
        "Date_log": "2026-12-05", "École": ecole, "Période_facture": "Décembre 2026", "N°_Facture": n_facture,
        "Date_séance": datetime.strptime(s.jour, "%Y-%m-%d"), "Début": debut_, "Fin": fin_, "Durée_h": heures,
        "Promotion": "M1", "Module / UE": "Statistiques", "Taux": taux, "Total HT": heures * taux,
        "Summary calendrier verbatim": s.intitule,
    }


# ──────────────────────────────────────────────────────────────────────
# Préparation
# ──────────────────────────────────────────────────────────────────────


def entrees(p: config.Profil, contenu: bytes | None = None, *, mois: str = MOIS,
            ecoles: Sequence[str] | None = None):
    """`charger_entrees` sur le contenu .ics injecté, après `set_data_root` (§ 1.1)."""
    fmt.set_data_root(p.racine)
    return prp().charger_entrees(p, mois, contenu_ics=contenu if contenu is not None else ics(), ecoles=ecoles)


def preparer(p: config.Profil, contenu: bytes | None = None, *, maintenant: datetime = MAINTENANT,
             mois: str = MOIS, ecoles: Sequence[str] | None = None):
    """`charger_entrees` puis `preparer` (pure), comme `lancer_preparation` sans les écritures."""
    e = entrees(p, contenu, mois=mois, ecoles=ecoles)
    return prp().preparer(entite=p.entite, mois=mois, entrees=e, prefixe=p.prefixe_facture,
                          maintenant=maintenant, ecoles=ecoles)


def brouillon(preparation, identifiant: str):
    for b in preparation.brouillons:
        if b.identifiant == identifiant:
            return b
    raise AssertionError(
        f"brouillon {identifiant} absent ; brouillons : {[b.identifiant for b in preparation.brouillons]}"
    )


def identifiants(preparation) -> set[str]:
    return {b.identifiant for b in preparation.brouillons}


def cles_facturees(preparation) -> set[str]:
    """`cle_seance` de toutes les lignes de tous les brouillons."""
    return {ligne.cle_seance for b in preparation.brouillons for ligne in b.lignes}


def uids_factures(preparation) -> set[str]:
    return {ligne.uid for b in preparation.brouillons for ligne in b.lignes}


def textes(valeurs: Iterable[str]) -> str:
    return "\n".join(valeurs)


def bloc(texte: str, titre: str, suivant: str) -> str:
    """Partie du texte qui va du titre (première occurrence) au titre suivant (exclu), ou à la fin."""
    debut_ = texte.index(titre)
    fin = texte.find(suivant, debut_ + len(titre))
    return texte[debut_:] if fin == -1 else texte[debut_:fin]


# ──────────────────────────────────────────────────────────────────────
# État fictif : décisions, validations, pièces d'émission
# ──────────────────────────────────────────────────────────────────────


def ecrire_decision(p: config.Profil, nature: str, cible: str, valeurs: Mapping[str, str] | None = None, *,
                    mois: str = MOIS, id_: str | None = None, commentaire: str = "") -> str:
    """Ajoute une décision au journal du mois (schéma P-0b), comme F-16b l'écrira ; rend son id."""
    valeurs = dict(valeurs or {})
    id_ = id_ or hashlib.sha256(f"{nature}|{cible}|{sorted(valeurs.items())}".encode()).hexdigest()[:12]
    journal.ajouter_evenement(
        chemins.journal_mois(p, mois), "decision", poste=POSTE_TEST, entite=p.entite,
        mois=mois, id=id_, nature=nature, cible=cible, valeurs=valeurs, commentaire=commentaire,
    )
    return id_


def ecrire_validation(p: config.Profil, identifiant: str, empreinte: str, *, decision: str = "oui",
                      mois: str = MOIS) -> None:
    journal.ajouter_evenement(
        chemins.journal_mois(p, mois), "validation", poste=POSTE_TEST, entite=p.entite,
        mois=mois, identifiant=identifiant, empreinte=empreinte, decision=decision, commentaire="",
        a_classer_acceptes=[],
    )


def ecrire_piece(p: config.Profil, n_facture: str, *, identifiant: str, mois: str = MOIS,
                 empreinte: str = "0" * 64, brouillon_: Any = None, finale: bool = False,
                 abandonnee: bool = False) -> Path:
    """Pièce d'émission fictive `<N°>.json` au format de P-0b (§ 3.4), écrite comme F-17 l'écrira."""
    ecole = identifiant.split("-")[0]
    onglet = f"{n_facture} {chemins.libelle_periode(mois)}"
    donnees = {
        "format": 1, "entite": p.entite, "n_facture": n_facture, "identifiant": identifiant, "mois": mois,
        "empreinte": empreinte, "onglet": onglet, "classeur": f"{ecole}/{p.entite}_{ecole}.xlsx",
        "validation": {"horodatage": "2026-12-31T10:00:00+01:00", "poste": POSTE_TEST},
        "brouillon": fichiers.en_json(brouillon_) if brouillon_ is not None else {},
    }
    dossier = chemins.dossier_abandonnees(p) if abandonnee else chemins.dossier_emissions(p)
    chemin = dossier / f"{n_facture}.json"
    fichiers.ecrire_json_atomique(chemin, donnees)
    if finale:
        fichiers.ecrire_json_atomique(dossier / f"{n_facture}.final.json", {
            "format": 1, "n_facture": n_facture, "pdf": f"{ecole}/{onglet}.pdf", "pdf_sha256": "0" * 64,
            "sauvegardes": [], "emise_le": "2027-01-01T10:00:00+01:00",
        })
    return chemin


def arborescence(racine: Path) -> dict[str, str]:
    """Chemin relatif (POSIX) → sha256 de chaque fichier sous la racine."""
    return {
        f.relative_to(racine).as_posix(): fichiers.empreinte_fichier(f)
        for f in sorted(racine.rglob("*")) if f.is_file()
    }
