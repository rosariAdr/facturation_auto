"""journal.py — Schéma unique des journaux d'état (P-0b, socle du lot 1 ; figé ici).

Trois journaux JSON Lines, plus les signalements, partagent ce schéma (contrats des vagues 3 à
5, § 1.7 et 3.4) :

| Journal | Emplacement | Contenu |
|---|---|---|
| du mois | `_sorties/<AAAA-MM>/journal.jsonl` | préparations, décisions, validations |
| des écritures | `_sorties/journal.jsonl` | émission, abandon, onglet refait, registre, correspondances |
| de la tâche | à côté du profil | exécutions de la tâche planifiée |
| signalements | `_sorties/signalements.jsonl` | constats, disparitions, traitements |

Chaque événement porte les champs communs `v`, `horodatage` (ISO à la seconde, avec décalage),
`poste`, `pid`, `entite` et `type`, puis les champs obligatoires de son type (`SCHEMA`). Toutes
les chaînes passent par `fichiers.masquer_secrets` : aucune adresse n'entre dans un journal.
Toute évolution du schéma passe par un ticket `chore`.

Ce module lit aussi les pièces d'émission (`<N°>.json`, `<N°>.final.json`, écrites par F-17)
et calcule l'état de validation d'une facture.

Lectures retenues (`💡 [proposé — à valider]`) :
- un champ obligatoire présent mais nul (`erreur=None` d'une préparation réussie) est accepté ;
  seul un champ absent est refusé ; rien n'est écrit quand l'événement est refusé ;
- le champ `decision` d'une validation vaut `oui` (validée), `non` (refusée) ou `corriger`
  (à corriger), les valeurs de `validation.DECISIONS` (F-17) ;
- une demande de correction dont le contenu change passe « à revalider » ; seule une facture
  refusée dont le contenu change redevient « à valider » (Q-E1) ;
- une pièce `<N°>.json` dont le champ `n_facture` diffère du nom du fichier est refusée.
"""

from __future__ import annotations

import os
from collections.abc import Collection, Mapping, Sequence
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any

from facturation import chemins
from facturation.fichiers import FichierAbime, ajouter_jsonl, en_json, lire_json, lire_jsonl, masquer_secrets

if TYPE_CHECKING:
    from facturation.config import Profil

VERSION = 1
FORMAT_PIECE = 1

COMMUNS = ("v", "horodatage", "poste", "pid", "entite", "type")

SCHEMA: dict[str, frozenset[str]] = {
    # Journal du mois
    "preparation": frozenset({"mois", "statut", "brouillons", "bloques", "a_classer", "erreur"}),
    "decision": frozenset({"mois", "id", "nature", "cible", "valeurs", "commentaire"}),
    "validation": frozenset({"mois", "identifiant", "empreinte", "decision", "commentaire", "a_classer_acceptes"}),
    # Journal des écritures
    "emission": frozenset({"execution", "mois", "identifiant", "empreinte", "n_facture", "etape", "statut", "detail"}),
    "abandon": frozenset({"n_facture", "identifiant", "raison", "sauvegarde"}),
    "onglet_refait": frozenset({"n_facture", "sauvegarde"}),
    "emise": frozenset({"n_facture", "identifiant", "pdf", "pdf_sha256"}),
    "registre": frozenset({"n_facture", "avant", "apres", "sauvegarde"}),
    "correspondance": frozenset({"ligne", "sauvegarde"}),
    # Signalements
    "constat": frozenset({"id", "signalement"}),
    "disparu": frozenset({"id"}),
    "traitement": frozenset({"id", "suite", "commentaire"}),
    # Journal de la tâche
    "demarrage": frozenset({"execution"}),
    "action": frozenset({"execution", "nature", "mois", "motif", "ecoles", "statut", "erreur"}),
    "notification": frozenset({"execution", "nature", "envoyee"}),
    "echec": frozenset({"execution", "type_erreur", "message", "trace"}),
    "fin": frozenset({"execution", "code"}),
    "installation": frozenset({"nom", "declencheurs"}),
}

# États de validation, dans l'ordre de priorité de `etat_validation`.
EMISE, EN_EMISSION, ABANDONNEE = "émise", "en émission", "abandonnée"
BLOQUEE, A_REVALIDER, VALIDEE, REFUSEE, A_CORRIGER, A_VALIDER = (
    "bloquée", "à revalider", "validée", "refusée", "à corriger", "à valider",
)
# Décision d'une validation (`validation.DECISIONS`, F-17) → état.
_ETAT_DE_LA_DECISION = {"oui": VALIDEE, "non": REFUSEE, "corriger": A_CORRIGER}
_CHAMPS_PIECE = ("format", "entite", "n_facture", "identifiant", "mois", "empreinte", "onglet", "classeur", "brouillon")


# ──────────────────────────────────────────────────────────────────────
# Événements
# ──────────────────────────────────────────────────────────────────────


def _masquer(valeur: Any) -> Any:
    """Toutes les chaînes d'une valeur JSON passées par `masquer_secrets`, imbriquées comprises."""
    if isinstance(valeur, str):
        return masquer_secrets(valeur)
    if isinstance(valeur, dict):
        return {cle: _masquer(contenu) for cle, contenu in valeur.items()}
    if isinstance(valeur, list):
        return [_masquer(element) for element in valeur]
    return valeur


def ajouter_evenement(chemin: Path, type_: str, *, poste: str, entite: str | None, **donnees: Any) -> dict[str, Any]:
    """Ajoute un événement au journal (une ligne JSON) et rend l'enregistrement écrit.

    Raises:
        ValueError: type inconnu, champ obligatoire absent, ou champ qui reprend le nom d'un champ
            commun ; le message les nomme et rien n'est écrit.
        TypeError: valeur que `fichiers.en_json` refuse ; rien n'est écrit.
    """
    if type_ not in SCHEMA:
        raise ValueError(
            f"Type d'événement inconnu : « {type_} ». Types du schéma des journaux : {', '.join(sorted(SCHEMA))}."
        )
    manquants = sorted(SCHEMA[type_] - donnees.keys())
    if manquants:
        pluriel = "s" if len(manquants) > 1 else ""
        raise ValueError(
            f"Événement « {type_} » : champ{pluriel} obligatoire{pluriel} absent{pluriel} : {', '.join(manquants)}."
        )
    reserves = sorted(set(donnees) & set(COMMUNS))
    if reserves:
        raise ValueError(f"Événement « {type_} » : champ(s) réservé(s) aux champs communs : {', '.join(reserves)}.")
    enregistrement = {
        "v": VERSION,
        "horodatage": datetime.now().astimezone().isoformat(timespec="seconds"),
        "poste": poste,
        "pid": os.getpid(),
        "entite": entite,
        "type": type_,
        **donnees,
    }
    enregistrement = _masquer(en_json(enregistrement))
    ajouter_jsonl(chemin, enregistrement)
    return enregistrement


def lire_evenements(chemin: Path, *, types: Collection[str] | None = None) -> list[dict[str, Any]]:
    """Événements du journal, dans l'ordre ; seulement ceux des `types` donnés s'il y en a.

    Fichier absent → [] ; dernière ligne tronquée ignorée avec un avertissement
    (`fichiers.lire_jsonl`).
    """
    evenements = lire_jsonl(chemin)
    if types is None:
        return evenements
    voulus = set(types)
    return [e for e in evenements if e.get("type") in voulus]


def dernieres_validations(evenements: Sequence[Mapping[str, Any]]) -> dict[str, dict[str, Any]]:
    """Dernière validation de chaque identifiant de brouillon, dans l'ordre du journal."""
    dernieres: dict[str, dict[str, Any]] = {}
    for evenement in evenements:
        if evenement.get("type") == "validation":
            dernieres[evenement["identifiant"]] = dict(evenement)
    return dernieres


# ──────────────────────────────────────────────────────────────────────
# Pièces d'émission et état de validation
# ──────────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class PieceEmission:
    """Pièce d'émission d'une facture (`_sorties/emissions/<N°>.json`), écrite par F-17.

    `classeur` est relatif à `<racine>/<ENTITE>/` ; `statut` vaut « en émission », « émise » (pièce
    finale présente) ou « abandonnée » (pièce rangée dans `abandonnees/`).
    """

    n_facture: str
    identifiant: str
    mois: str
    entite: str
    empreinte: str
    onglet: str
    classeur: str
    statut: str
    brouillon: dict
    final: dict | None


def etat_validation(*, bloque: bool, empreinte: str, derniere: Mapping[str, Any] | None,
                    piece: PieceEmission | None) -> str:
    """État d'une facture préparée, dans cet ordre de priorité :

    1. « émise » ou « en émission » : une pièce d'émission existe, hors abandon ;
    2. « bloquée » ;
    3. « à revalider » : l'empreinte de la dernière validation diffère de l'empreinte courante
       (une facture refusée dont le contenu change redevient « à valider », Q-E1) ;
    4. « validée », « refusée » ou « à corriger », selon la dernière décision ;
    5. « à valider » : aucune validation.

    Raises:
        ValueError: décision inconnue dans la dernière validation.
    """
    if piece is not None and piece.statut in (EMISE, EN_EMISSION):
        return piece.statut
    if bloque:
        return BLOQUEE
    if derniere is None:
        return A_VALIDER
    decision = str(derniere.get("decision", "")).strip().casefold()
    if decision not in _ETAT_DE_LA_DECISION:
        raise ValueError(
            f"Validation de « {derniere.get('identifiant', '?')} » : décision « {derniere.get('decision')} » "
            f"inconnue (attendu : {', '.join(_ETAT_DE_LA_DECISION)})."
        )
    if derniere.get("empreinte") != empreinte:
        return A_VALIDER if decision == "non" else A_REVALIDER
    return _ETAT_DE_LA_DECISION[decision]


def _lire_piece(fichier: Path, profil: "Profil", statut_abandon: bool) -> PieceEmission:
    donnees = lire_json(fichier)
    if not isinstance(donnees, dict):
        raise FichierAbime(f"Pièce d'émission abîmée : {fichier} (un objet JSON est attendu).")
    entite = donnees.get("entite")
    if entite != profil.entite:
        raise ValueError(
            f"Pièce d'émission {fichier} : entité « {entite} », alors que le profil est celui de "
            f"« {profil.entite} ». Les entités ne se mélangent jamais (ADR001) : rangez ce fichier sous "
            "la racine de son entité."
        )
    manquants = [champ for champ in _CHAMPS_PIECE if champ not in donnees]
    if manquants:
        raise FichierAbime(f"Pièce d'émission abîmée : {fichier}, champ(s) absent(s) : {', '.join(manquants)}.")
    if donnees["format"] != FORMAT_PIECE:
        raise FichierAbime(f"Pièce d'émission {fichier} : format {donnees['format']!r} inconnu (attendu {FORMAT_PIECE}).")
    n_facture = str(donnees["n_facture"])
    if n_facture != fichier.stem:
        raise FichierAbime(f"Pièce d'émission {fichier} : elle porte le n° « {n_facture} », pas « {fichier.stem} ».")
    finale = fichier.with_name(f"{fichier.stem}.final.json")
    final = lire_json(finale) if finale.is_file() else None
    if statut_abandon:
        statut = ABANDONNEE
    else:
        statut = EMISE if final is not None else EN_EMISSION
    return PieceEmission(
        n_facture=n_facture,
        identifiant=str(donnees["identifiant"]),
        mois=str(donnees["mois"]),
        entite=str(entite),
        empreinte=str(donnees["empreinte"]),
        onglet=str(donnees["onglet"]),
        classeur=str(donnees["classeur"]),
        statut=statut,
        brouillon=donnees["brouillon"],
        final=final,
    )


def pieces_emission(profil: "Profil") -> list[PieceEmission]:
    """Pièces d'émission de l'entité du profil : celles de `emissions/`, puis celles d'`abandonnees/`.

    Les pièces finales (`<N°>.final.json`) et les PDF en attente ne sont pas des pièces.

    Raises:
        ValueError: pièce d'une autre entité (le message nomme l'entité trouvée).
        FichierAbime: pièce illisible ou incomplète.
    """
    pieces: list[PieceEmission] = []
    for dossier, abandon in ((chemins.dossier_emissions(profil), False), (chemins.dossier_abandonnees(profil), True)):
        if not dossier.is_dir():
            continue
        for fichier in sorted(dossier.glob("*.json")):
            if fichier.name.endswith(".final.json") or not fichier.is_file():
                continue
            pieces.append(_lire_piece(fichier, profil, abandon))
    return pieces
