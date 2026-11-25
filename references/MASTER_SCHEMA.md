# MASTER_SCHEMA — Structure du classeur de données

**Dernière mise à jour** : 2026-11-23

> Le module ne contient aucune donnée : tout vient du MASTER de l'entité, lu au moment de l'exécution. Ce document décrit ce que le classeur doit contenir. Le MASTER est la source de vérité (ADR005) et le seul endroit où vivent les règles réelles des écoles (ADR026).
>
> Chaque onglet est marqué **existant** (lu ou contrôlé par l'outil aujourd'hui) ou **prévu** (lot 1). Les colonnes des onglets prévus sont une proposition : `💡 [proposé — à valider]` jusqu'à leur mise en œuvre. Les exemples sont fictifs.

---

## Emplacement

```
<racine de données>/
├── ROSARI/
│   ├── ROSARI_MASTER.xlsx          ← source de vérité de l'entité
│   ├── ABC/ROSARI_ABC.xlsx         ← un classeur par école, un onglet par facture
│   └── ECO/ROSARI_ECO.xlsx
└── KAUR/
    └── KAUR_MASTER.xlsx
```

Le chemin du MASTER est toujours `<racine>/<ENTITE>/<ENTITE>_MASTER.xlsx`.

**Résolution de la racine**
- Par les scripts et la commande `facturation` : `--data-root` > `--profil` > variable `PCF_DATA_ROOT` > profil par défaut s'il existe > dossier courant.
- Par le module seul (`facturation.pcf_format`) : `set_data_root()` > `PCF_DATA_ROOT` > dossier courant.

---

## Onglets

| Onglet | Statut | Rôle |
|---|---|---|
| `Coordonnées` | Existant, **obligatoire** | Identités de l'émetteur et des écoles, SIRET, IBAN, BIC, mentions |
| `Sessions` | Existant, **obligatoire** | Journal des séances facturées (ADR003) |
| `Cours` | Existant (suivi) | Volumes contractuels, heures dispensées, facturées, à facturer |
| `Tarifs` | Existant (suivi) | Taux horaires par école, module et programme |
| `Écoles` | Existant (suivi) | Liste des écoles et code de facturation |
| `Devis` | Existant (suivi) | Devis et bons de commande reçus |
| `Règles_facturation` | Prévu (F-14) | Règles propres à chaque école |
| `Correspondances` | Prévu (F-12, F-13) | Reconnaissance des créneaux d'agenda et heures facturables |
| `Factures` | Prévu (F-25) | Registre des factures : échéance, envoi, paiement |
| `Conventions_horaires` | Obsolète | Remplacé par `Correspondances` (ADR016) |

`facturation profil verifier` signale une erreur si un onglet obligatoire manque, et un point d'attention si un onglet de suivi manque.

### `Coordonnées` — existant, obligatoire

La seule source des identités. Première ligne = en-têtes.

| Colonne | A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|---|
| | Type | Code | Nom / Raison sociale | Adresse L1 | Adresse L2 | Adresse L3 | Email | Téléphone |

Trois valeurs de `Type`, écrites exactement ainsi :

- **`Émetteur`** (avec l'accent) — une ligne, `Code` = l'entité (`ROSARI`). `Nom` = nom et prénom suivis de « EI » ou « Entrepreneur individuel » (ADR030, valable pour ROSARI comme pour KAUR) ; sans cette mention, l'outil refuse d'écrire. Alimente l'encadré émetteur.
- **`Destinataire`** — une ligne par école, `Code` = code de l'école. `Nom` = **dénomination juridique de la société qui paie**, telle qu'écrite au contrat (ADR030), pas le nom commercial. Une ligne sans nom ou sans adresse bloque la facture (étape 4 du PCF).
- **`Info`** — `Code` = l'entité, `Nom` porte le libellé et `Adresse L1` la valeur. `SIRET`, `IBAN` et `BIC` sont obligatoires : leur absence lève une erreur au chargement. `TVA` porte la mention de TVA. `NDA` porte le texte complet de la mention de déclaration d'activité, recopié tel quel sur la facture ; il est obligatoire depuis F-24 (sans lui, l'outil refuse d'écrire). Les lignes `Info` ne concernent que l'émetteur : on n'y range pas le SIRET d'une école.

`Adresse L3` sert au pays. L'encadré destinataire n'offre que 4 lignes : ville et pays sont réunis sur une seule (`00000 VILLE-EXEMPLE, France`).

### `Sessions` — existant, obligatoire

| A | B | C | D | E | F | G | H | I | J | K | L | M |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Date_log | École | Période_facture | N°_Facture | Date_séance | Début | Fin | Durée_h | Promotion | Module / UE | Taux | Total HT | Summary calendrier verbatim |

Journal d'audit : il justifie le contenu d'une facture après coup. Ajout seul, avec une exception (ADR003 amendé) : une facture régénérée remplace ses propres lignes, repérées par le `N°_Facture`, au lieu d'en ajouter. Aucune autre ligne n'est modifiée ni supprimée.

### `Cours` — existant (suivi)

| A | B | C | D | E | F | G | H | I | J |
|---|---|---|---|---|---|---|---|---|---|
| École | Module / Matière | Classe / Groupe | Volume contractuel (h) | Heures dispensées (h) | Heures facturées (h) | Heures à facturer (h) | N° Facture liée | N° Devis lié | Statut |

Invariant vérifié par `scripts/check_coherence.py` : `Heures dispensées = Heures facturées + Heures à facturer`. Le script signale aussi des heures facturées au-delà des heures dispensées.

### `Tarifs` — existant (suivi)

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| École | Module / Matière | Classe / Programme | Taux HT/h | Notes | Réf. contrat *(prévu)* |

Le **contrat signé fait foi** pour un taux (ADR026) : la colonne `Réf. contrat` garde sa référence. Toute autre source de taux s'aligne sur cet onglet.

### `Écoles` — existant (suivi)

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| École | Code facturation | Entité légale | Contact | Email contact | Statut mission |

Le `Code facturation` relie une école à sa ligne `Destinataire` dans `Coordonnées`. Code d'une nouvelle école : 3 lettres, unique dans l'entité (ADR029).

### `Devis` — existant (suivi)

| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| École | Module | N° Devis | Date émission | Montant HT | Statut | PO reçu | N° Facture générée |

### `Règles_facturation` — prévu (F-14) · `💡 [proposé — à valider]`

Une ligne par école (ou par contrat quand une école en a plusieurs).

| Colonne | Contenu |
|---|---|
| École | Code de l'école |
| Statut lot | `Lot 1`, `Plus tard` ou `Hors facturation` : le périmètre d'un lot est lu ici |
| Rythme | `mensuel` (défaut), `fin de module`, ou `cycle` |
| Début de cycle | Jour du mois où commence le cycle de l'école, si `cycle` |
| Date limite d'envoi | Jour du mois et heure limites fixés par l'école, s'il y en a |
| Délai de paiement | Délai du contrat ou de l'école (ex. `30 jours`, `45 jours fin de mois`) ; vide = date d'émission + 30 jours (ADR030) |
| N° de contrat | Référence affichée sur la facture |
| Une facture par contrat | `oui` / `non` |
| Bon de commande requis | `oui` / `non` ; une facture par ligne de bon de commande si `oui` |
| Colonne Jour | `affichée` / `masquée` |
| Regroupement | `séance` (défaut) ou `jour` (une ligne par jour et par cours) |
| Canal d'envoi | Mail, portail fournisseur, plateforme de l'école, courrier |
| Pièces jointes | Ex. `PDF`, `PDF + récapitulatif` |
| Notes | Mentions particulières demandées par l'école |

### `Correspondances` — prévu (F-12, F-13) · `💡 [proposé — à valider]`

| Colonne | Contenu |
|---|---|
| Clé | Code d'événement (`#ABC-M1-STATS`) ou intitulé d'événement |
| Type de clé | `code`, `exact` ou `préfixe` (ordre de recherche, ADR016) |
| École | Code de l'école |
| Module | Module ou matière |
| Promotion | Classe ou groupe |
| Clé tarif | Ligne de `Tarifs` à appliquer |
| Mode heures | `agenda` (défaut), `coefficient` ou `forfait` |
| Valeur | Coefficient (ex. `0.875` pour 7/8) ou nombre d'heures du forfait |
| Gardée le | Date à laquelle Adrian a gardé cette règle (ADR028) |
| Notes | Justification |

### `Factures` — prévu (F-25) · `💡 [proposé — à valider]`

Le registre : une ligne par facture émise.

| Colonne | Contenu |
|---|---|
| N° facture | Ex. `R_ABC008` |
| École | Code de l'école |
| Contrat | N° de contrat, s'il existe |
| Période | Ex. `Nov 2026` |
| Date d'émission | Le 1er du mois suivant (ADR027) |
| Échéance | Calculée selon `Règles_facturation` (ADR030) |
| Total HT | Montant de la facture |
| Date d'envoi | Saisie par Adrian (le mail envoyé fait foi) |
| Date de paiement | Saisie par Adrian (ADR031) |
| Statut | `émise`, `envoyée`, `payée`, `échue` |
| Note | Ex. raison d'un numéro manquant (ADR029) |

---

## Classeurs de factures et numérotation

Un classeur par école, un **onglet par facture**, nommé `<N° facture> <Période>` (ex. `R_ABC009 Juil 2026`). Le préfixe `PAID` ajouté au nom d'un onglet réglé est un usage de lecture, sans valeur de preuve (ADR031).

Numérotation : `<préfixe d'entité>_<CODE ÉCOLE><NNN>` (ADR002, ADR029), par exemple `R_ABC009`.
- Le numéro suivant est le plus grand numéro de la série plus un ; il se lit aujourd'hui dans `Sessions` (colonne `N°_Facture`), demain dans `Factures`.
- Un numéro manquant n'est jamais réutilisé ; sa raison est notée.
- Un numéro n'est attribué qu'à l'émission.

---

## Créer un MASTER de zéro

1. Classeur avec les onglets existants ci-dessus (`Coordonnées`, `Sessions`, `Cours`, `Tarifs`, `Écoles`, `Devis`), en-têtes en ligne 1, dans cet ordre de colonnes.
2. Remplir `Coordonnées` : une ligne `Émetteur` (nom suivi de « EI »), une ligne `Info` par valeur (`SIRET`, `IBAN`, `BIC`, `TVA`, `NDA`), une ligne `Destinataire` par école.
3. Vérifier : `python scripts/check_coherence.py --entite <ENTITE>`.
4. Valider l'installation : `facturation profil verifier`, qui signale un émetteur, un SIRET, un IBAN, un BIC ou un destinataire incomplet.

Un exemple complet, avec une entité et une école fictives : [RECETTE_DOC.md](RECETTE_DOC.md), étape 3.
