# RECETTE_DOC — Vérifier que la documentation suffit

**Dernière mise à jour** : 2026-11-23

> But : prouver qu'une personne ou un agent qui découvre le projet peut l'installer, produire une facture fictive et répondre aux questions courantes **avec le dépôt seul** (ADR025, F-04). Toutes les données sont inventées ; rien ne touche aux vraies données.

---

## Qui fait la recette

1. Une personne extérieure au projet, à qui l'on ne donne que le lien du dépôt.
2. Un agent neuf (Claude Code ou Codex) lancé dans un clone du dépôt, sans `CLAUDE.md`, sans `memory/`, sans accès à Notion.

Règle commune : ne poser aucune question à Adrian pendant la recette. Chaque blocage est noté (voir « Compte rendu »).

---

## Scénario

Entité fictive `DEMO` (préfixe de facture `D_`), école fictive `FIC`. Travailler dans un dossier vide, par exemple `C:\recette`, jamais dans la racine de données réelle.

### 1. Installer
Suivre `README.md`, section « Installation ».

### 2. Créer un profil à part
Le profil est créé à un emplacement dédié, pour ne jamais remplacer le profil de la personne qui facture :

```bash
facturation profil init --nom "Camille Exemple" --entite DEMO --racine "C:/recette/donnees" --emplacement "C:/recette/profil.toml"
```

### 3. Construire un MASTER fictif
En suivant `references/MASTER_SCHEMA.md`, section « Créer un MASTER de zéro ». Pour aller vite, ce script crée le même classeur :

```python
from pathlib import Path
from openpyxl import Workbook

racine = Path("C:/recette/donnees/DEMO")
racine.mkdir(parents=True, exist_ok=True)
wb = Workbook(); wb.remove(wb.active)
ws = wb.create_sheet("Coordonnées")
ws.append(["Type", "Code", "Nom / Raison sociale", "Adresse L1", "Adresse L2", "Adresse L3", "Email", "Téléphone"])
ws.append(["Émetteur", "DEMO", "EXEMPLE Camille EI", "1 rue de l'Exemple", "00000 VILLE-EXEMPLE", "", "camille@example.org", "+33600000000"])
ws.append(["Info", "DEMO", "SIRET", "00000000000000"])
ws.append(["Info", "DEMO", "IBAN", "FR76 0000 0000 0000 0000 0000 000"])
ws.append(["Info", "DEMO", "BIC", "TESTFRPP"])
ws.append(["Info", "DEMO", "TVA", "TVA non applicable, article 293 B du Code général des impôts"])
ws.append(["Destinataire", "FIC", "ÉCOLE FICTIVE SAS", "2 avenue du Test", "00000 VILLE-TEST", "France", "compta@ecole-fictive.example"])
wb.create_sheet("Sessions").append(["Date_log", "École", "Période_facture", "N°_Facture", "Date_séance", "Début", "Fin",
                                    "Durée_h", "Promotion", "Module / UE", "Taux", "Total HT", "Summary calendrier verbatim"])
wb.save(racine / "DEMO_MASTER.xlsx")
```

### 4. Vérifier l'installation
```bash
facturation profil verifier --profil "C:/recette/profil.toml"
```
Attendu : aucune ligne `ERREUR`. Des lignes `ATTENTION` sont normales (onglets de suivi absents, adresse iCal non enregistrée, LibreOffice absent).

### 5. Décrire une facture fictive
Créer `C:/recette/spec.json` en suivant `examples/spec_exemple.json` :

```json
{
  "entite": "DEMO",
  "ecole": "FIC",
  "ecole_label": "FIC (École fictive)",
  "fichier_sortie": "DEMO/FIC/DEMO_FIC.xlsx",
  "onglet": "D_FIC001 Nov 2026",
  "n_facture": "D_FIC001",
  "date_emission": "2026-12-01",
  "masquer_jour": false,
  "lignes": [
    {"promotion": "M1", "libelle": "Statistiques — groupe A",
     "seances": [{"date": "2026-11-04", "heures": 3.5, "taux": 50},
                 {"date": "2026-11-18", "heures": 3.5, "taux": 50}]},
    {"promotion": "M2", "libelle": "Programmation — groupe B",
     "seances": [{"date": "2026-11-25", "heures": 3, "taux": 50}]}
  ]
}
```

### 6. Aperçu, puis écriture
```bash
python scripts/nouvelle_facture.py "C:/recette/spec.json" --data-root "C:/recette/donnees" --dry-run
python scripts/nouvelle_facture.py "C:/recette/spec.json" --data-root "C:/recette/donnees"
```
Attendu : l'aperçu annonce un total de 500 € HT (10 h × 50 €) et n'écrit rien ; l'écriture crée `C:/recette/donnees/DEMO/FIC/DEMO_FIC.xlsx`, onglet `D_FIC001 Nov 2026`, et le contrôle après écriture ne signale aucun écart. Relancer la même écriture doit être refusé (la facture existe déjà).

### 7. Répondre à cinq questions, avec le dépôt seul
Pour chaque réponse, citer le fichier et la section qui la donnent.

1. Quel sera le numéro de la prochaine facture de l'école `FIC` ?
2. Quelle date d'émission porteront des factures de séances de décembre ?
3. Que se passe-t-il si une séance déjà facturée est annulée après l'émission ?
4. Où change-t-on le taux d'une école, et quelle pièce le justifie ?
5. Qu'est-ce que l'outil fait aujourd'hui, et qu'est-ce qui est seulement prévu ?

---

## Corrigé (pour la personne qui évalue)

| Question | Réponse attendue | Où |
|---|---|---|
| 1 | `D_FIC002` : plus grand numéro de la série + 1, jamais de réutilisation | ADR002, ADR029 ; `MASTER_SCHEMA.md`, « Classeurs de factures et numérotation » |
| 2 | Le 1er janvier suivant (1er du mois suivant) | ADR027 |
| 3 | La séance est signalée avec la facture concernée, puis corrigée par un avoir ; la facture émise n'est jamais modifiée | ADR024 (conséquences reprises), ADR027, F-23 dans `ROADMAP.md` |
| 4 | Dans le MASTER, onglet `Tarifs` ; le contrat signé fait foi et sa référence est gardée dans le MASTER | ADR026 ; `MASTER_SCHEMA.md`, « Tarifs » |
| 5 | Aujourd'hui : aperçu et écriture d'une facture depuis une spec, contrôle après écriture, vérification de l'installation. Prévu : lecture de l'agenda, préparation du 20, PDF et MASTER par l'outil, mentions de l'ADR030 dans le gabarit | `README.md`, « Ce que l'outil fait aujourd'hui » ; `ROADMAP.md` |

---

## Critères de réussite

- Aucune commande du scénario n'échoue.
- Les cinq réponses sont justes et sourcées.
- Aucune question n'a été posée à Adrian.
- Durée indicative : moins d'une heure pour une personne, hors installation de Python.

## Compte rendu

Pour chaque blocage : l'étape, ce qui manquait ou était faux, le fichier concerné. Chaque blocage devient un ticket `docs` (branche `chore/…`, `AGENTS.md` §1).
