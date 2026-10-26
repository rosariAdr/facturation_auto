# MASTER_SCHEMA — Structure du classeur de données

**Dernière mise à jour** : 2026-09-24

> Le module ne contient aucune donnée : tout vient du MASTER de l'entité, lu au
> runtime. Ce document décrit ce que le classeur doit contenir pour que le
> module fonctionne.

---

## Emplacement

```
<racine de données>/
├── ROSARI/
│   ├── ROSARI_MASTER.xlsx          ← source de vérité de l'entité
│   ├── EPITA/ROSARI_EPITA.xlsx     ← un classeur par école, 1 onglet par facture
│   └── DOR/ROSARI_DOR.xlsx
└── KAUR/
    └── KAUR_MASTER.xlsx
```

La racine est résolue dans cet ordre : `set_data_root()` → `$PCF_DATA_ROOT` →
dossier courant. Le chemin du MASTER est donc toujours
`<racine>/<ENTITE>/<ENTITE>_MASTER.xlsx`.

---

## Onglets

Deux onglets sont **obligatoires** pour générer une facture : `Coordonnées` et
`Sessions`. Les autres servent au suivi et aux contrôles de cohérence.

### `Coordonnées` — obligatoire

La seule source des identités. Première ligne = en-têtes.

| Colonne | A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|---|
| | Type | Code | Nom / Raison sociale | Adresse L1 | Adresse L2 | Adresse L3 | Email | Téléphone |

Trois valeurs possibles pour `Type` :

- **`Émetteur`** — une ligne, `Code` = l'entité (`ROSARI`). Alimente l'encadré
  émetteur.
- **`Destinataire`** — une ligne par école, `Code` = code court de l'école
  (`EPITA`, `DOR`…). Une ligne dont tous les champs sont vides est traitée comme
  « destinataire inconnu » : la génération s'arrête (gate étape 4 du PCF).
- **`Info`** — `Code` = l'entité, `Nom` porte le label (`SIRET`, `IBAN`, `BIC`,
  `TVA`) et `Adresse L1` la valeur. `SIRET`, `IBAN` et `BIC` sont obligatoires ;
  leur absence lève une erreur explicite au chargement.

`Adresse L3` sert au pays. Comme l'encadré destinataire n'offre que 4 lignes,
ville et pays sont réunis sur une seule (`94270 LE KREMLIN BICETRE, France`).

### `Sessions` — journal des séances facturées

| A | B | C | D | E | F | G | H | I | J | K | L | M |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Date_log | École | Période_facture | N°_Facture | Date_séance | Début | Fin | Durée_h | Promotion | Module / UE | Taux | Total HT | Summary calendrier verbatim |

Journal d'audit : c'est lui qui justifie le contenu d'une facture après coup.
Le PCF le veut en **append-only** (cf. `ADR.md`) ; un script de génération
ré-exécuté peut cependant mettre ses propres lignes à jour sur place, à condition
de le faire en s'appuyant sur `N°_Facture` — sinon chaque régénération duplique
le journal et fausse les totaux.

### `Cours` — suivi contractuel

| A | B | C | D | E | F | G | H | I | J |
|---|---|---|---|---|---|---|---|---|---|
| École | Module / Matière | Classe / Groupe | Volume contractuel (h) | Heures dispensées (h) | Heures facturées (h) | Heures à facturer (h) | N° Facture liée | N° Devis lié | Statut |

Invariant vérifié par `scripts/check_coherence.py` :
`Heures dispensées = Heures facturées + Heures à facturer`.

### `Tarifs`

| A | B | C | D | E |
|---|---|---|---|---|
| École | Module / Matière | Classe / Programme | Taux HT/h | Notes |

### `Écoles`

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| École | Code facturation | Entité légale | Contact | Email contact | Statut mission |

Le `Code facturation` est la clé qui relie une école à sa ligne `Destinataire`
dans `Coordonnées`.

### `Conventions_horaires` — optionnel

| A | B | C | D |
|---|---|---|---|
| École | Module / Matière | Coefficient | Notes |

Coefficient appliqué aux heures du calendrier pour obtenir les heures facturables
(`H_facturable = H_calendrier × coefficient`), quand une école ne facture pas au
temps réel de présence. L'étape 3.2 du PCF s'arrête si la convention manque pour
un couple école/module. Onglet à créer au premier besoin.

### `Devis`

| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| École | Module | N° Devis | Date émission | Montant HT | Statut | PO reçu | N° Facture générée |

---

## Classeurs de factures

Un classeur par école, un **onglet par facture**, nommé
`<N° facture> <Période>` (ex. `R_EPITA009 Juil 2026`). Les factures réglées sont
renommées avec un préfixe (`PAID …`) — convention de lecture, sans effet sur le code.

La numérotation suit `<Initiale entité>_<CODE ÉCOLE><NNN>` : `R_EPITA009`,
`R_DOR004`. Le numéro suivant se lit dans `MASTER.Sessions`, colonne `N°_Facture`.

---

## Créer un MASTER de zéro

1. Classeur avec les six onglets ci-dessus, en-têtes en ligne 1, dans cet ordre
   de colonnes.
2. Remplir `Coordonnées` : une ligne `Émetteur`, une ligne `Info` par valeur
   (SIRET, IBAN, BIC, TVA), une ligne `Destinataire` par école.
3. Vérifier : `python scripts/check_coherence.py --entite ROSARI`.
4. Valider le format : `python scripts/test_pcf_format.py` — le test 2 échoue si
   `Coordonnées` est incomplet.
