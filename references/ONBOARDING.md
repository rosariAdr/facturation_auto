# ONBOARDING — Ajouter une école

**Dernière mise à jour** : 2026-11-23

> La checklist pour intégrer une nouvelle école dans une entité (ADR011). À suivre **dans l'ordre**. Toutes les informations recueillies sont des données réelles : elles vont dans le MASTER de l'entité (ADR026), jamais dans le dépôt. Les exemples ci-dessous sont fictifs.

---

## Vue d'ensemble

| Phase | Acteur | Résultat |
|---|---|---|
| 1. Informations légales et financières | École + Adrian ou Parv | Identité, contacts, conditions de paiement, références |
| 2. Informations pédagogiques | École + Adrian ou Parv | Missions, heures facturables, code école |
| 3. Saisie dans le MASTER | Adrian ou Parv, avec un agent | Lignes du MASTER complètes, aperçu validé avant chaque écriture |
| 4. Facture pilote | Adrian ou Parv | Une facture vérifiée point par point, école déclarée active |

---

## Phase 1 — Informations légales et financières

> **Phase bloquante** : sans ces données, aucune facture ne peut être émise correctement (ADR030).

### 1.1 — Identité du client
À demander **par écrit**, ou à lire dans le contrat :

```
□ Dénomination juridique exacte de la société qui paie, telle qu'écrite au contrat
  (ce n'est pas forcément le nom commercial de l'école)
□ N° SIRET (ou équivalent étranger) et forme juridique
□ Adresse de facturation complète (rue, code postal, ville, pays)
□ Adresse de prestation, si elle diffère
□ N° de TVA du client : vérifier s'il doit figurer sur la facture (pas requis dans les cas vérifiés jusqu'ici)
```

### 1.2 — Contacts
Distinguer deux contacts, qui peuvent être la même personne :

```
□ Contact pédagogique (responsable de programme) : nom, email
□ Contact facturation (à qui envoyer les factures) : nom ou service, email
□ Canal d'envoi : email, portail fournisseur, plateforme de l'école, courrier postal
□ Pièces à joindre : PDF seul, ou PDF + récapitulatif demandé par l'école
```

### 1.3 — Conditions financières et références

```
□ Taux horaire HT, et la référence du contrat qui le fixe (le contrat fait foi, ADR026)
□ Modalité : à l'heure, au forfait, ou mixte
□ Délai de paiement prévu par le contrat ou par l'école
  (vide = date d'émission + 30 jours, ADR030)
□ N° de contrat à afficher sur la facture ; plusieurs contrats → une facture par contrat ?
□ Bon de commande : l'école en émet-elle un avant la prestation ?
  ▸ si oui : comment l'obtenir, et une facture par ligne de bon de commande ?
□ Rythme : mensuel (par défaut), fin de module, ou cycle propre à l'école
  ▸ si cycle : jour de début, date et heure limites d'envoi
```

### 1.4 — Mentions particulières

```
□ Mentions demandées par l'école en plus de celles de l'ADR030
  (ex. n° de déclaration d'activité, code du module)
□ N° de fournisseur attribué par l'école, s'il y en a un
□ Contrat signé : oui / non, date
```

---

## Phase 2 — Informations pédagogiques

### 2.1 — Missions
Pour **chaque mission** :

```
□ Module / matière (ex. « Statistiques appliquées »)
□ Programme / promotion (ex. « M1 groupe A »)
□ Volume contractuel total (heures)
□ Période : date de début, date de fin (estimée si besoin)
□ Fréquence approximative (ex. « une séance de 3 h par semaine »)
```

### 2.2 — Heures facturables
Comment passer des heures de l'agenda aux heures facturées. La règle va dans l'onglet `Correspondances` (colonnes « Mode heures » et « Valeur », ADR016).

```
□ Mode « agenda » (par défaut) : heures facturées = heures de l'agenda
□ Mode « coefficient » : ex. 9 h d'agenda = 7 h facturées (pause non facturée), coefficient 7/9
□ Mode « forfait » : ex. un jury de 7 h d'agenda facturé 7,5 h
□ Justification de la règle (contrat, échange écrit avec l'école)
```

Sans règle, l'outil facture les heures de l'agenda. Un créneau qu'il ne reconnaît pas va dans la file à classer et n'est jamais facturé par défaut (ADR028).

### 2.3 — Code école

```
□ Code de 3 lettres en majuscules pour toute nouvelle école (ADR029), ex. ABC
□ Vérifier qu'il est libre dans l'entité
□ Le préfixe d'entité (R_ ou K_) s'ajoute à la numérotation : R_ABC001
```

---

## Phase 3 — Saisie dans le MASTER

> **Pré-requis** : phases 1 et 2 complètes et validées par Adrian ou Parv. Chaque écriture dans le MASTER est précédée d'un aperçu et d'une validation explicite.

Schéma de chaque onglet : [MASTER_SCHEMA.md](MASTER_SCHEMA.md). Un onglet marqué « prévu » dans le schéma se remplit dès qu'il existe.

| Onglet | Ce qu'on ajoute |
|---|---|
| `Écoles` | Une ligne : école, code, entité légale, contact, email, statut de mission |
| `Coordonnées` | Une ligne `Destinataire` : code, dénomination juridique, adresse, email. Aucune ligne `Info` pour l'école : ces lignes sont réservées à l'émetteur |
| `Tarifs` | Une ligne par module et programme : taux HT, notes, référence du contrat |
| `Règles_facturation` | Statut dans les lots, rythme ou cycle, date limite, délai de paiement, n° de contrat, bon de commande, colonne Jour, regroupement, canal, pièces jointes |
| `Correspondances` | Une ligne par intitulé ou code d'événement : école, module, promotion, clé tarif, mode heures, valeur |
| `Cours` | Une ligne par mission : volume contractuel, heures dispensées = 0, facturées = 0, à facturer = 0, statut « À planifier » |

Le dossier `<ENTITE>/<CODE>/` et le classeur `<ENTITE>_<CODE>.xlsx` sont créés à la première facture : rien à préparer à la main.

Contrôle : `facturation profil verifier` puis `python scripts/check_coherence.py --entite <ENTITE>`.

### Brief à donner à un agent

```
Onboarding d'une école (references/ONBOARDING.md, phase 3)

ENTITÉ : [ROSARI ou KAUR]
ÉCOLE : [nom] — CODE : [3 lettres]

Montre un aperçu avant chaque écriture dans le MASTER et attends ma validation.
N'invente aucune valeur : demande-moi ce qui manque.

Données légales et financières : [phase 1]
Données pédagogiques : [phase 2]

À faire : compléter Écoles, Coordonnées, Tarifs, Règles_facturation,
Correspondances et Cours, puis lancer les deux contrôles et me donner le résultat.
```

---

## Phase 4 — Facture pilote

> **Objectif** : vérifier la chaîne de bout en bout avant de déclarer l'école active.

### 4.1 — Émettre la facture pilote
- **Cas A, des heures ont déjà été faites** : émettre la première vraie facture en suivant [PROC00_pcf.md](PROC00_pcf.md).
- **Cas B, aucune heure faite** : produire une facture fictive dans une racine de données de test (voir [RECETTE_DOC.md](RECETTE_DOC.md)), jamais dans le MASTER réel.

### 4.2 — Vérifications obligatoires (ADR030)

```
□ Émetteur : nom suivi de « EI », adresse, email, téléphone, SIRET ; IBAN et BIC corrects
□ Client : dénomination juridique exacte, adresse de facturation complète
□ N° de facture au bon format (R_ABC001), n° de contrat et n° de bon de commande sur des lignes distinctes
□ Date d'émission au 1er du mois suivant ; date d'échéance explicite
□ Tableau des séances : date, durée, taux, montant ; total HT juste
□ « Total HT », mention « TVA non applicable, article 293 B du Code général des impôts », « Net à payer »
□ Pénalités de retard, indemnité forfaitaire de 40 €, « Pas d'escompte »
□ N° de déclaration d'activité
□ Mentions particulières demandées par l'école
□ Une seule page ; PDF généré sans erreur
```

### 4.3 — Essai d'envoi
Avant la première vraie facture, préparer un brouillon de mail adressé à soi-même : objet clair, PDF lisible en pièce jointe, signature. Aujourd'hui à la main ; par l'outil au lot 2 (F-32), qui ne joindra que le PDF.

### 4.4 — Validation finale

```
□ Adrian ou Parv valide visuellement la facture pilote
□ Si c'est bon : statut de l'école → active, et statut dans les lots renseigné
□ Sinon : retour à la phase 1, 2 ou 3 selon le problème ; rien n'est envoyé à l'école
```

---

## Cas particuliers

### École déjà active, jamais onboardée
1. Faire la phase 1 après coup (recueillir ce qui manque, notamment la dénomination juridique et le délai de paiement).
2. Compléter le MASTER.
3. Sauter la phase 4 si des factures ont déjà été émises et vérifiées.

### Mission ponctuelle (jury, intervention courte)
- Phase 1 : le minimum (dénomination, SIRET, contact facturation, taux ou forfait, délai).
- Phase 2 : seulement la règle d'heures si elle diffère de l'agenda.
- Phase 3 : statut de mission « Ponctuelle ».
- Phase 4 : la facture unique sert de facture pilote.

### École présente dans les deux entités
**Étanchéité absolue** (ADR001) : faire l'onboarding deux fois, une fois dans chaque MASTER. Aucune donnée partagée.

---

## Documents à fournir à l'école (idée)

Un formulaire qui reprend la phase 1, pour que l'école le remplisse. Pas encore créé.
