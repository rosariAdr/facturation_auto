"""Fixtures partagées : données fictives uniquement (aucune donnée réelle dans le dépôt).

`racine_fictive` construit dans un dossier temporaire une racine de données complète :
`ENTITE/ENTITE_MASTER.xlsx` avec un émetteur entrepreneur individuel (nom terminé par
« EI », ADR030), SIRET, IBAN, BIC et n° de déclaration d'activité factices, deux
destinataires renseignés et un destinataire vide.

`creer_master` accepte aussi (P-0b) les lignes fictives des onglets `Sessions`, `Tarifs`,
`Écoles`, `Factures`, `Règles_facturation` et `Correspondances` : chaque ligne est une liste
dans l'ordre des en-têtes, un dictionnaire {en-tête : valeur} (en-têtes absents = cellules
vides), ou None pour laisser une ligne vide dans la feuille.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any

import pytest
from openpyxl import Workbook

from facturation import pcf_format as fmt

ENTETES = {
    "Sessions": [
        "Date_log", "École", "Période_facture", "N°_Facture", "Date_séance", "Début", "Fin",
        "Durée_h", "Promotion", "Module / UE", "Taux", "Total HT", "Summary calendrier verbatim",
    ],
    "Cours": [
        "École", "Module / Matière", "Classe / Groupe", "Volume contractuel (h)",
        "Heures dispensées (h)", "Heures facturées (h)", "Heures à facturer (h)",
        "N° Facture liée", "N° Devis lié", "Statut",
    ],
    "Tarifs": ["École", "Module / Matière", "Classe / Programme", "Taux HT/h", "Notes"],
    "Écoles": ["École", "Code facturation", "Entité légale", "Contact", "Email contact", "Statut mission"],
    "Devis": ["École", "Module", "N° Devis", "Date émission", "Montant HT", "Statut", "PO reçu", "N° Facture générée"],
}


# En-têtes de `Tarifs` quand des lignes de tarif sont données : avec `Réf. contrat` (P-0b).
ENTETES_TARIFS_CONTRAT = ["École", "Module / Matière", "Classe / Programme", "Taux HT/h", "Notes", "Réf. contrat"]

# Onglets prévus de MASTER_SCHEMA, créés seulement quand des lignes sont données.
ENTETES_PREVUS = {
    "Règles_facturation": [
        "École", "Statut lot", "Rythme", "Début de cycle", "Date limite d'envoi", "Délai de paiement",
        "N° de contrat", "Une facture par contrat", "Bon de commande requis", "Colonne Jour",
        "Regroupement", "Canal d'envoi", "Pièces jointes", "Notes",
    ],
    "Correspondances": [
        "Clé", "Type de clé", "École", "Module", "Promotion", "Clé tarif", "Mode heures", "Valeur",
        "Gardée le", "Notes",
    ],
    "Factures": [
        "N° facture", "École", "Contrat", "Période", "Date d'émission", "Échéance", "Total HT",
        "Date d'envoi", "Date de paiement", "Statut", "Note",
    ],
}

Lignes = Sequence[Sequence[Any] | Mapping[str, Any] | None]


# Texte fictif de la ligne `Info` / `NDA`, imprimé tel quel sur la facture (ADR030).
NDA_FICTIF = "Déclaration d'activité enregistrée sous le n° 00000000000 auprès du préfet de région Exemple"


def _ajouter_lignes(feuille, entetes: list[str], lignes: Lignes) -> None:
    """Ajoute des lignes fictives : liste dans l'ordre des en-têtes, dictionnaire, ou None (ligne vide)."""
    for ligne in lignes:
        if ligne is None:
            feuille.append([])
        elif isinstance(ligne, Mapping):
            inconnus = set(ligne) - set(entetes)
            assert not inconnus, f"en-têtes inconnus de l'onglet {feuille.title} : {inconnus}"
            feuille.append([ligne.get(entete) for entete in entetes])
        else:
            feuille.append(list(ligne))


def creer_master(
    racine: Path,
    entite: str = "ROSARI",
    cours: list[list] | None = None,
    *,
    nom_emetteur: str | None = None,
    nda: str | None = NDA_FICTIF,
    regles: Lignes | None = None,
    correspondances: Lignes | None = None,
    tarifs: Lignes | None = None,
    ecoles: Lignes | None = None,
    factures: Lignes | None = None,
    sessions: Lignes | None = None,
) -> Path:
    """Écrit un MASTER fictif sous `racine/entite/` et renvoie son chemin.

    Args:
        nom_emetteur: nom de l'émetteur ; par défaut « FORMATEUR FICTIF <ENTITE> EI ».
        nda: texte de la ligne `Info` / `NDA` ; None pour l'omettre.
        regles, correspondances, factures: lignes des onglets prévus du même nom, créés
            seulement si l'argument est donné (même vide).
        tarifs: lignes de `Tarifs` ; quand il est donné, l'en-tête porte `Réf. contrat`.
        ecoles, sessions: lignes des onglets `Écoles` et `Sessions`.
    """
    dossier = racine / entite
    dossier.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet("Coordonnées")
    ws.append(["Type", "Code", "Nom / Raison sociale", "Adresse L1", "Adresse L2", "Adresse L3", "Email", "Téléphone"])
    nom = nom_emetteur if nom_emetteur is not None else f"FORMATEUR FICTIF {entite} EI"
    ws.append(["Émetteur", entite, nom, "1 rue de l'Exemple,", "00000 Villefictive", "", "formateur@example.org", "+33600000000"])
    ws.append(["Info", entite, "SIRET", "00000000000000"])
    ws.append(["Info", entite, "IBAN", "FR76 0000 0000 0000 0000 0000 000"])
    ws.append(["Info", entite, "BIC", "TESTFRPP"])
    ws.append(["Info", entite, "TVA", "TVA non applicable - art. 293B CGI"])
    if nda is not None:
        ws.append(["Info", entite, "NDA", nda])
    ws.append(["Destinataire", "ABC", "ÉCOLE FICTIVE A", "10 avenue Test", "00000 TESTVILLE", "France", "compta@ecole-a.example"])
    ws.append(["Destinataire", "ECO", "ÉCOLE FICTIVE B", "20 rue Test", "00000 TESTVILLE", "", "factures@ecole-b.example"])
    ws.append(["Destinataire", "VIDE", "", "", "", "", ""])
    donnees = {"Cours": cours, "Tarifs": tarifs, "Écoles": ecoles, "Sessions": sessions}
    for nom, entetes in ENTETES.items():
        if nom == "Tarifs" and tarifs is not None:
            entetes = ENTETES_TARIFS_CONTRAT
        feuille = wb.create_sheet(nom)
        feuille.append(entetes)
        _ajouter_lignes(feuille, entetes, donnees.get(nom) or [])
    for nom, lignes in (("Règles_facturation", regles), ("Correspondances", correspondances), ("Factures", factures)):
        if lignes is not None:
            feuille = wb.create_sheet(nom)
            feuille.append(ENTETES_PREVUS[nom])
            _ajouter_lignes(feuille, ENTETES_PREVUS[nom], lignes)
    chemin = dossier / f"{entite}_MASTER.xlsx"
    wb.save(chemin)
    return chemin


@pytest.fixture
def racine_fictive(tmp_path: Path) -> Path:
    """Racine de données fictive, branchée sur `pcf_format` le temps du test."""
    racine = tmp_path / "donnees"
    creer_master(racine)
    ancienne = fmt._DATA_ROOT
    fmt.set_data_root(racine)
    yield racine
    fmt._DATA_ROOT = ancienne
    fmt._MASTER_DATA_CACHE.clear()
