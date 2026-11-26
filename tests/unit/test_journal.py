"""Schéma unique des journaux d'état (P-0b, `facturation/journal.py` ; banque de tests : T-JRN-01 à T-JRN-11).

Tests d'acceptation écrits avant l'implémentation : chaque test importe `facturation.journal`
lui-même, pour que l'absence du module fasse échouer les tests sans bloquer la collecte des
autres. Journaux et pièces fictifs créés dans `tmp_path`.

Contrat (contrats des vagues 3 à 5, § 1.6, 1.7 et 3.4) :
- champs communs `v`, `horodatage` (ISO, à la seconde, avec décalage), `poste`, `pid`, `entite`,
  `type` ; toutes les chaînes passent par `masquer_secrets` ; un type inconnu ou un champ
  obligatoire absent lève `ValueError` ;
- `SCHEMA` : type → champs obligatoires (table du § 3.4, figée ici) ;
- `etat_validation` : `émise` ou `en émission` (pièce présente, hors abandon), `bloquée`,
  `à revalider` (empreinte de la dernière validation différente), `validée`, `refusée` ou
  `à corriger`, `à valider` ; une facture refusée dont le contenu change redevient « à valider »
  (💡 Q-E1) ;
- pièces d'émission `<N°>.json` et `<N°>.final.json` sous `_sorties/emissions/`, pièces abandonnées
  rangées dans `abandonnees/` ; un fichier d'état d'une autre entité est refusé (`ValueError`, § 1.1).

Lectures retenues (`💡 [proposé — à valider]`) :
- un champ obligatoire présent mais nul (`erreur=None` d'une préparation réussie) est accepté :
  seul un champ absent est refusé ; rien n'est écrit quand l'événement est refusé ;
- le champ `decision` d'une validation prend les valeurs de `validation.DECISIONS` (F-17) :
  `oui` → « validée », `non` → « refusée », `corriger` → « à corriger » ;
- une demande de correction dont le contenu change passe « à revalider » (règle générale de
  l'empreinte) ; seule la facture refusée redevient « à valider » ;
- `masquer_secrets` s'applique aussi aux chaînes imbriquées (listes, dictionnaires).
"""

from __future__ import annotations

import dataclasses
import importlib
import json
import os
import re
from datetime import datetime
from pathlib import Path

import pytest

from facturation import config

pytestmark = pytest.mark.unit

ADRESSE_FICTIVE = "https://calendar.example.org/calendar/ical/personne%40example.org/private-0123abcd/basic.ics"
EMPREINTE_A = "a" * 64
EMPREINTE_B = "b" * 64

SCHEMA_ATTENDU = {
    "preparation": {"mois", "statut", "brouillons", "bloques", "a_classer", "erreur"},
    "decision": {"mois", "id", "nature", "cible", "valeurs", "commentaire"},
    "validation": {"mois", "identifiant", "empreinte", "decision", "commentaire", "a_classer_acceptes"},
    "emission": {"execution", "mois", "identifiant", "empreinte", "n_facture", "etape", "statut", "detail"},
    "abandon": {"n_facture", "identifiant", "raison", "sauvegarde"},
    "onglet_refait": {"n_facture", "sauvegarde"},
    "emise": {"n_facture", "identifiant", "pdf", "pdf_sha256"},
    "registre": {"n_facture", "avant", "apres", "sauvegarde"},
    "correspondance": {"ligne", "sauvegarde"},
    "constat": {"id", "signalement"},
    "disparu": {"id"},
    "traitement": {"id", "suite", "commentaire"},
    "demarrage": {"execution"},
    "action": {"execution", "nature", "mois", "motif", "ecoles", "statut", "erreur"},
    "notification": {"execution", "nature", "envoyee"},
    "echec": {"execution", "type_erreur", "message", "trace"},
    "fin": {"execution", "code"},
    "installation": {"nom", "declencheurs"},
}
COMMUNS = {"v", "horodatage", "poste", "pid", "entite", "type"}
HORODATAGE = re.compile(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}[+-]\d{2}:\d{2}")


def _jrn():
    """Module testé, importé dans le test : son absence fait échouer le test, pas la collecte."""
    return importlib.import_module("facturation.journal")


def _preparation(**autres) -> dict:
    donnees = {"mois": "2026-12", "statut": "ok", "brouillons": 2, "bloques": 0, "a_classer": 1, "erreur": None}
    donnees.update(autres)
    return donnees


def _validation(identifiant: str, empreinte: str, decision: str, commentaire: str = "") -> dict:
    return {"mois": "2026-12", "identifiant": identifiant, "empreinte": empreinte, "decision": decision,
            "commentaire": commentaire, "a_classer_acceptes": 0}


# ──────────────────────────────────────────────────────────────────────
# Schéma et écriture
# ──────────────────────────────────────────────────────────────────────


def test_t_jrn_03_version_et_schema_figes():
    """T-JRN-03 — `VERSION` vaut 1 ; `SCHEMA` porte exactement les 18 types du § 3.4 et leurs champs
    obligatoires, en ensembles figés."""
    jrn = _jrn()
    assert jrn.VERSION == 1
    assert set(jrn.SCHEMA) == set(SCHEMA_ATTENDU)
    for type_, champs in SCHEMA_ATTENDU.items():
        assert isinstance(jrn.SCHEMA[type_], frozenset), type_
        assert set(jrn.SCHEMA[type_]) == champs, type_


def test_t_jrn_04_ajouter_evenement_champs_communs(tmp_path):
    """T-JRN-04 — `ajouter_evenement` écrit une ligne : champs communs (`v`, `horodatage` à la seconde
    avec décalage, `poste`, `pid`, `entite`, `type`) et données ; il rend l'enregistrement écrit,
    que `lire_evenements` relit à l'identique."""
    jrn = _jrn()
    journal = tmp_path / "_sorties" / "2026-12" / "journal.jsonl"
    rendu = jrn.ajouter_evenement(journal, "preparation", poste="POSTE-A", entite="ROSARI", **_preparation())
    assert COMMUNS <= set(rendu)
    assert (rendu["v"], rendu["poste"], rendu["pid"], rendu["entite"], rendu["type"]) == (
        1, "POSTE-A", os.getpid(), "ROSARI", "preparation",
    )
    assert HORODATAGE.fullmatch(rendu["horodatage"]), rendu["horodatage"]
    assert datetime.fromisoformat(rendu["horodatage"]).utcoffset() is not None
    for cle, valeur in _preparation().items():
        assert rendu[cle] == valeur, cle
    assert journal.read_text(encoding="utf-8").count("\n") == 1
    assert jrn.lire_evenements(journal) == [rendu]


@pytest.mark.parametrize(
    "type_, donnees, cause",
    [
        ("inconnu", {"mois": "2026-12"}, "inconnu"),
        ("preparation", {k: v for k, v in _preparation().items() if k != "brouillons"}, "brouillons"),
        ("validation", {k: v for k, v in _validation("ABC-2026-12", EMPREINTE_A, "oui").items() if k != "empreinte"},
         "empreinte"),
        ("fin", {}, "code"),
    ],
    ids=["type-inconnu", "preparation-sans-brouillons", "validation-sans-empreinte", "fin-sans-code"],
)
def test_t_jrn_01_type_inconnu_ou_champ_manquant(tmp_path, type_, donnees, cause):
    """T-JRN-01 — un type inconnu ou un champ obligatoire absent lève `ValueError` qui les nomme ;
    rien n'est écrit."""
    jrn = _jrn()
    journal = tmp_path / "journal.jsonl"
    with pytest.raises(ValueError) as erreur:
        jrn.ajouter_evenement(journal, type_, poste="POSTE-A", entite="ROSARI", **donnees)
    assert cause in str(erreur.value)
    assert not journal.exists() or journal.read_text(encoding="utf-8") == ""


def test_t_jrn_05_aucun_secret_dans_les_journaux(tmp_path):
    """T-JRN-05 — toutes les chaînes journalisées passent par `masquer_secrets`, imbriquées comprises :
    l'adresse iCal n'apparaît ni dans le fichier ni dans l'enregistrement rendu."""
    jrn = _jrn()
    journal = tmp_path / "journal_tache.jsonl"
    rendu = jrn.ajouter_evenement(
        journal, "echec", poste="POSTE-A", entite=None, execution="exec-fictive-1",
        type_erreur="AgendaIndisponible", message=f"Agenda indisponible : {ADRESSE_FICTIVE}",
        trace=["ligne 1", {"url": ADRESSE_FICTIVE}],
    )
    texte = journal.read_text(encoding="utf-8")
    for fragment in ("calendar.example.org", "private-0123abcd"):
        assert fragment not in texte, fragment
        assert fragment not in json.dumps(rendu, ensure_ascii=False), fragment
    assert "<adresse masquée>" in rendu["message"]
    assert rendu["entite"] is None


def test_t_jrn_06_lire_evenements_filtre_par_type(tmp_path):
    """T-JRN-06 — `lire_evenements` relit dans l'ordre, filtré par types au besoin ; fichier absent → []."""
    jrn = _jrn()
    journal = tmp_path / "journal.jsonl"
    assert jrn.lire_evenements(journal) == []
    jrn.ajouter_evenement(journal, "preparation", poste="POSTE-A", entite="ROSARI", **_preparation())
    jrn.ajouter_evenement(journal, "validation", poste="POSTE-A", entite="ROSARI",
                          **_validation("ABC-2026-12", EMPREINTE_A, "oui"))
    jrn.ajouter_evenement(journal, "decision", poste="POSTE-A", entite="ROSARI", mois="2026-12", id="d-1",
                          nature="exclure", cible="fictif-1@example.org", valeurs={}, commentaire="")
    assert [e["type"] for e in jrn.lire_evenements(journal)] == ["preparation", "validation", "decision"]
    assert [e["type"] for e in jrn.lire_evenements(journal, types={"validation", "decision"})] == [
        "validation", "decision",
    ]
    assert jrn.lire_evenements(journal, types=["emission"]) == []


def test_t_jrn_07_dernieres_validations(tmp_path):
    """T-JRN-07 — `dernieres_validations` : la dernière validation de chaque identifiant, dans l'ordre
    du journal ; les autres types sont ignorés."""
    jrn = _jrn()
    journal = tmp_path / "journal.jsonl"
    jrn.ajouter_evenement(journal, "validation", poste="POSTE-A", entite="ROSARI",
                          **_validation("ABC-2026-12", EMPREINTE_A, "non", "à revoir"))
    jrn.ajouter_evenement(journal, "validation", poste="POSTE-A", entite="ROSARI",
                          **_validation("ECO-2026-12", EMPREINTE_B, "oui"))
    jrn.ajouter_evenement(journal, "preparation", poste="POSTE-A", entite="ROSARI", **_preparation())
    jrn.ajouter_evenement(journal, "validation", poste="POSTE-A", entite="ROSARI",
                          **_validation("ABC-2026-12", EMPREINTE_B, "oui"))
    dernieres = jrn.dernieres_validations(jrn.lire_evenements(journal))
    assert set(dernieres) == {"ABC-2026-12", "ECO-2026-12"}
    assert (dernieres["ABC-2026-12"]["decision"], dernieres["ABC-2026-12"]["empreinte"]) == ("oui", EMPREINTE_B)
    assert dernieres["ECO-2026-12"]["decision"] == "oui"
    assert jrn.dernieres_validations([]) == {}


# ──────────────────────────────────────────────────────────────────────
# État de validation
# ──────────────────────────────────────────────────────────────────────


def _piece(statut: str, **autres):
    champs = {
        "n_facture": "R_ABC001", "identifiant": "ABC-2026-12", "mois": "2026-12", "entite": "ROSARI",
        "empreinte": EMPREINTE_A, "onglet": "R_ABC001 Déc 2026", "classeur": "ABC/ROSARI_ABC.xlsx",
        "statut": statut, "brouillon": {"ecole": "ABC"}, "final": None,
    }
    champs.update(autres)
    return _jrn().PieceEmission(**champs)


def test_t_jrn_08_piece_emission_dataclass_figee():
    """T-JRN-08 — `PieceEmission` est une dataclass figée, aux champs du contrat, dans l'ordre."""
    jrn = _jrn()
    assert tuple(f.name for f in dataclasses.fields(jrn.PieceEmission)) == (
        "n_facture", "identifiant", "mois", "entite", "empreinte", "onglet", "classeur", "statut",
        "brouillon", "final",
    )
    with pytest.raises(dataclasses.FrozenInstanceError):
        _piece("émise").statut = "abandonnée"


@pytest.mark.parametrize(
    "bloque, derniere, piece, attendu",
    [
        (False, ("oui", EMPREINTE_A), "émise", "émise"),
        (True, None, "en émission", "en émission"),
        (True, ("oui", EMPREINTE_A), None, "bloquée"),
        (False, ("oui", EMPREINTE_B), None, "à revalider"),
        (False, ("oui", EMPREINTE_A), None, "validée"),
        (False, ("non", EMPREINTE_A), None, "refusée"),
        (False, ("corriger", EMPREINTE_A), None, "à corriger"),
        (False, None, None, "à valider"),
        (False, ("non", EMPREINTE_B), None, "à valider"),
        (False, ("corriger", EMPREINTE_B), None, "à revalider"),
        (False, ("oui", EMPREINTE_A), "abandonnée", "validée"),
        (False, None, "abandonnée", "à valider"),
    ],
    ids=["emise", "en-emission-meme-bloquee", "bloquee", "a-revalider", "validee", "refusee", "a-corriger",
         "a-valider", "refusee-puis-modifiee", "correction-modifiee", "piece-abandonnee-validee",
         "piece-abandonnee-sans-validation"],
)
def test_t_jrn_02_etat_validation_couvre_ses_etats(bloque, derniere, piece, attendu):
    """T-JRN-02 — `etat_validation` couvre ses états, dans l'ordre de priorité du contrat : pièce
    émise ou en émission (une pièce abandonnée ne compte pas), bloquée, à revalider (empreinte
    changée), validée / refusée / à corriger, à valider ; une facture refusée dont le contenu
    change redevient « à valider » (💡 Q-E1)."""
    jrn = _jrn()
    validation = None if derniere is None else _validation("ABC-2026-12", derniere[1], derniere[0])
    etat = jrn.etat_validation(
        bloque=bloque, empreinte=EMPREINTE_A, derniere=validation,
        piece=None if piece is None else _piece(piece),
    )
    assert etat == attendu


# ──────────────────────────────────────────────────────────────────────
# Pièces d'émission
# ──────────────────────────────────────────────────────────────────────


def _profil(tmp_path: Path, entite: str = "ROSARI") -> config.Profil:
    chemin = tmp_path / f"profil_{entite}.toml"
    chemin.write_text(config.rediger("Personne A", entite, str(tmp_path / "donnees")), encoding="utf-8")
    return config.charger(chemin)


def _ecrire_piece(dossier: Path, numero: str, *, entite: str = "ROSARI", final: bool = False) -> dict:
    dossier.mkdir(parents=True, exist_ok=True)
    piece = {
        "format": 1, "entite": entite, "n_facture": numero, "identifiant": f"ABC-2026-12-{numero}",
        "mois": "2026-12", "empreinte": EMPREINTE_A, "onglet": f"{numero} Déc 2026",
        "classeur": "ABC/ROSARI_ABC.xlsx",
        "validation": {"horodatage": "2026-12-01T10:00:00+01:00", "poste": "POSTE-A"},
        "brouillon": {"ecole": "ABC", "lignes": [{"heures": "3.5", "taux": "45.0"}]},
    }
    (dossier / f"{numero}.json").write_text(json.dumps(piece, ensure_ascii=False), encoding="utf-8")
    contenu_final = None
    if final:
        contenu_final = {"format": 1, "n_facture": numero, "pdf": f"ABC/{numero} Déc 2026.pdf",
                         "pdf_sha256": "0" * 64, "sauvegardes": [], "emise_le": "2026-12-01T10:05:00+01:00"}
        (dossier / f"{numero}.final.json").write_text(json.dumps(contenu_final), encoding="utf-8")
    return contenu_final


def test_t_jrn_09_pieces_emission_et_leur_statut(tmp_path):
    """T-JRN-09 — `pieces_emission` lit les pièces de l'entité : « émise » avec sa pièce finale,
    « en émission » sans elle, « abandonnée » dans `abandonnees/` ; la pièce finale et le PDF en
    attente ne sont pas des pièces ; aucun dossier → liste vide."""
    jrn = _jrn()
    profil = _profil(tmp_path)
    assert jrn.pieces_emission(profil) == []
    emissions = profil.dossier_sorties / "emissions"
    final = _ecrire_piece(emissions, "R_ABC001", final=True)
    _ecrire_piece(emissions, "R_ABC002")
    (emissions / "R_ABC002.pdf").write_bytes(b"%PDF-1.4 fictif")
    _ecrire_piece(emissions / "abandonnees", "R_ABC003")

    pieces = {p.n_facture: p for p in jrn.pieces_emission(profil)}
    assert set(pieces) == {"R_ABC001", "R_ABC002", "R_ABC003"}
    assert (pieces["R_ABC001"].statut, pieces["R_ABC001"].final) == ("émise", final)
    assert (pieces["R_ABC002"].statut, pieces["R_ABC002"].final) == ("en émission", None)
    assert pieces["R_ABC003"].statut == "abandonnée"
    piece = pieces["R_ABC002"]
    assert (piece.identifiant, piece.mois, piece.entite, piece.empreinte) == (
        "ABC-2026-12-R_ABC002", "2026-12", "ROSARI", EMPREINTE_A,
    )
    assert (piece.onglet, piece.classeur) == ("R_ABC002 Déc 2026", "ABC/ROSARI_ABC.xlsx")
    assert piece.brouillon == {"ecole": "ABC", "lignes": [{"heures": "3.5", "taux": "45.0"}]}


def test_t_jrn_10_piece_d_une_autre_entite_refusee(tmp_path):
    """T-JRN-10 — une pièce dont le champ `entite` diffère du profil est refusée (`ValueError` qui
    nomme l'entité trouvée) : les entités ne se mélangent jamais (ADR001)."""
    jrn = _jrn()
    profil = _profil(tmp_path)
    _ecrire_piece(profil.dossier_sorties / "emissions", "K_ABC001", entite="KAUR")
    with pytest.raises(ValueError, match="KAUR"):
        jrn.pieces_emission(profil)


def test_t_jrn_11_preparation_reussie_sans_erreur(tmp_path):
    """T-JRN-11 — un champ obligatoire présent mais nul (`erreur` d'une préparation réussie) est
    accepté et écrit nul ; la tâche planifiée journalise sans entité."""
    jrn = _jrn()
    journal = tmp_path / "journal.jsonl"
    rendu = jrn.ajouter_evenement(journal, "preparation", poste="POSTE-A", entite="ROSARI", **_preparation(erreur=None))
    assert rendu["erreur"] is None
    tache = jrn.ajouter_evenement(tmp_path / "journal_tache.jsonl", "demarrage", poste="POSTE-A", entite=None,
                                  execution="exec-fictive-2")
    assert tache["entite"] is None and tache["type"] == "demarrage"
