# ONBOARDING — Processus nouveau client

**Dernière mise à jour** : 2026-05-04

> Checklist exhaustive A→Z pour intégrer une nouvelle école/client dans le système. À suivre **dans l'ordre** sans sauter d'étape.

---

## Vue d'ensemble

| Phase | Acteur | Durée |
|---|---|---|
| 1. Collecte info légale | Client + Adrian/Parv | Variable |
| 2. Collecte info pédagogique | Client + Adrian/Parv | Variable |
| 3. Création technique | Claude Code | ~10 min |
| 4. Validation par facture pilote | Adrian/Parv | ~15 min |

---

## Phase 1 — Collecte des informations légales

> **Cette phase est blocante** : sans ces données, aucune facture ne peut être émise légalement.

### 1.1 — Identité légale du client

À demander **par écrit** au contact opérationnel :

```
□ Raison sociale exacte (telle qu'elle apparaîtra sur la facture)
□ N° SIRET (ou équivalent étranger)
□ Forme juridique (SAS, SARL, association...)
□ Adresse de facturation complète (L1, L2, L3, code postal, ville)
□ Adresse de prestation (si différente de facturation)
□ Pays
```

### 1.2 — Contacts

Distinguer **deux contacts** (peuvent être la même personne) :

```
□ CONTACT OPÉRATIONNEL (RH / responsable pédagogique)
   - Nom, prénom
   - Email
   - Téléphone (optionnel)

□ CONTACT COMPTABILITÉ (à qui envoyer les factures)
   - Nom, prénom (ou "Service Comptabilité" / "AP — Accounts Payable")
   - Email à utiliser pour envoi
   - Process particulier ?
     □ Portail fournisseur dédié (URL, login)
     □ Email seulement
     □ Courrier postal aussi
```

### 1.3 — Conditions financières

```
□ Taux horaire négocié HT (€/h)
□ Modalité de facturation
   □ Horaire pur (heures × taux)
   □ Forfait
   □ Mixte (préciser)
□ Délai de paiement
   □ 30 jours net
   □ Fin de mois
   □ 45 jours fin de mois
   □ Autre (préciser)
□ Mode de paiement
   □ Virement (par défaut)
   □ Autre
□ Bon de commande (PO) requis avant facturation ?
   □ NON
   □ OUI — processus :
      ▸ Comment obtenir le PO ?
      ▸ N° de PO à mentionner sur la facture (oui/non)
      ▸ Devis préalable obligatoire (oui/non)
```

### 1.4 — Annexes éventuelles

```
□ Mentions spécifiques à apposer sur la facture (ex. "Bon de commande n° XXX")
□ Numéro de fournisseur attribué par le client (si le client en attribue)
□ Conditions générales de vente échangées (oui/non)
□ Contrat de prestation signé (oui/non, si oui : date)
```

---

## Phase 2 — Collecte des informations pédagogiques

### 2.1 — Mission(s)

Pour **chaque mission** confiée :

```
□ Module / Matière (ex. "Mathematics for Data Science")
□ Programme / Promotion (ex. "Master 2 DIA", "BTS SIO 1ère année")
□ Volume contractuel total (heures)
□ Période contractuelle
   ▸ Date de début
   ▸ Date de fin (estimée si pas figée)
□ Fréquence approximative (ex. "1 séance/semaine de 3h")
```

### 2.2 — Convention horaire

**Critique pour le calcul** : le coefficient horaire de l'école est porté par l'onglet `Conventions_horaires` du MASTER (cf. [MASTER_SCHEMA.md](MASTER_SCHEMA.md)). Sans convention pour le couple école/module, le PCF s'arrête à l'étape 3.2.

```
□ Heures planifiées dans le calendrier par séance type
□ Heures réellement facturables par séance type
□ Différence éventuelle (pauses repas, pauses techniques)
□ Coefficient à appliquer (calculé)
□ Justification (à documenter)
```

Exemples de conventions actuelles :
- DOR M2 RMD : 8h cal = 8h facturable (1:1)
- DOR M2 DIA : 9h cal = 7h facturable (pause non facturée)
- CDP Masters : 9h cal = 8h facturable (1h pause)

### 2.3 — Code école interne

À définir avec Adrian (ou Parv selon entité) :

```
□ Code court (3-6 caractères, MAJUSCULES)
   ex. DOR, EPITA, ISG, EMN, CDP, ASC, IDRAC, ESCA, IFAG
□ Vérifier l'unicité au sein de l'entité (pas de collision)
□ Préfixe entité (R_ ou K_) sera ajouté automatiquement à la numérotation
```

---

## Phase 3 — Création technique (Claude Code)

> **Acteur** : Claude Code (App, onglet Code).
>
> **Pré-requis** : Phase 1 et Phase 2 complètes et validées par Adrian/Parv.

### Prompt à utiliser

Copier-coller dans Claude Code après avoir rempli les variables :

```
Onboarding nouveau client

ENTITÉ : [ROSARI ou KAUR]
ÉCOLE : [Nom complet]
CODE : [Code court]

Données légales :
- Raison sociale : [...]
- SIRET : [...]
- Adresse facturation : [...]
- Contact opérationnel : [Nom, email]
- Contact comptabilité : [Nom, email]
- Délai paiement : [...]
- PO requis : [oui/non, modalités]

Données pédagogiques (mission 1, dupliquer si plusieurs) :
- Module : [...]
- Programme : [...]
- Volume contractuel : [X]h
- Période : [date début → date fin]
- Convention horaire : [Hcal]/[Hfact] (coef [X.XXX])

Actions à exécuter (dans l'ordre, en montrant aperçu avant chaque écriture) :

1. Ajouter ligne dans [ENTITE]_MASTER.xlsx > Écoles
   Code | Nom | Entité légale | Contact opérationnel | Email contact | Statut mission = Active

2. Ajouter ligne dans [ENTITE]_MASTER.xlsx > Tarifs
   Code | Module | Programme | Taux HT/h | Notes (PO si applicable)

3. Ajouter ligne(s) dans [ENTITE]_MASTER.xlsx > Coordonnées
   Type=Destinataire | Code | Nom complet | Adresse L1/L2/L3 | Email | Téléphone
   + ligne Info pour SIRET du client

4. Ajouter ligne(s) dans [ENTITE]_MASTER.xlsx > Conventions_horaires
   Code | Module/Programme | H calendrier | H facturable | Coefficient | Justification

5. Ajouter ligne dans [ENTITE]_MASTER.xlsx > Cours
   Code | Module | Programme | Volume contractuel | H dispensées=0 | H facturées=0
   | H à facturer=0 | N° Devis lié=vide | N° Facture liée=vide | Statut=À planifier

6. Créer dossier [ENTITE]/[CODE_OU_NOM]/ s'il n'existe pas

7. Créer fichier [ENTITE]/[CODE_OU_NOM]/[ENTITE]_[CODE].xlsx
   - Onglet "Modèle" basé sur _templates/PCF_facture_canonical.xlsx
   - Aucun onglet de facture (sera créé au premier PCF)

8. Mettre à jour CONTEXT.md
   - Ajouter l'école dans la section "Écoles actives" de la bonne entité
   - Si nouveau code, ajouter règle de classification calendrier
   - Mettre à jour la date "Dernière mise à jour"

9. Mettre à jour TASKS.md
   - Cocher la case "Onboarding [Nom école] terminé" si déjà tracée
   - Sinon, ajouter ligne ✅ Déjà fait

Fournir à la fin un récap des modifications effectuées + un lien vers la facture pilote
à émettre en Phase 4.
```

---

## Phase 4 — Validation par facture pilote

> **Objectif** : valider le bout-en-bout avant de considérer l'école comme opérationnelle.

### 4.1 — Émission d'une facture pilote

Soit :
- **Cas A — Données réelles disponibles** : émettre la première vraie facture pour les heures déjà dispensées (suivre [PROC00_pcf.md](PROC00_pcf.md)).
- **Cas B — Pas encore de cours dispensés** : émettre une facture **pilote fictive** (1h, montant symbolique) pour valider la chaîne, à supprimer ensuite.

### 4.2 — Vérifications obligatoires sur la facture pilote

```
□ Émetteur (ROSARI ou KAUR) : nom, adresse, email, téléphone, SIRET, IBAN/BIC corrects
□ Destinataire : raison sociale exacte, adresse complète, contact comptabilité
□ N° de facture au bon format (R_<CODE>NNN ou K_<CODE>NNN)
□ Mention TVA non applicable — art. 293B CGI
□ Date d'émission cohérente
□ Délai de paiement mentionné
□ Mentions légales (pénalités) présentes
□ Tableau des séances : date, durée, taux, montant
□ Total HT correctement calculé
□ Si PO requis : N° de PO mentionné
□ Format visuel conforme au template canonique
□ Export PDF généré sans erreur
```

### 4.3 — Test envoi (si possible)

Avant la première vraie facture, **tester l'envoi** sur sa propre adresse :

```
□ Brouillon Gmail créé
□ Pièce jointe PDF lisible
□ Pièce jointe XLSX non corrompue
□ Objet du mail clair
□ Corps du mail professionnel
□ Signature présente
```

### 4.4 — Validation finale

```
□ Adrian/Parv valide visuellement la facture pilote
□ Si OK : statut école → Active, prêt pour production
□ Si KO : retour Phase 1/2/3 selon le problème, ne pas envoyer au client
```

---

## Checklist récapitulative

À cocher une fois l'onboarding complet :

```
PHASE 1 — Légal
  □ Identité légale collectée
  □ Contacts opérationnel et comptabilité distincts identifiés
  □ Conditions financières actées (taux, délai, PO)

PHASE 2 — Pédagogique
  □ Mission(s) détaillée(s)
  □ Convention horaire documentée
  □ Code école interne défini et unique

PHASE 3 — Technique
  □ MASTER mis à jour : Écoles, Tarifs, Coordonnées, Conventions_horaires, Cours
  □ Dossier et fichier école créés
  □ CONTEXT.md mis à jour
  □ TASKS.md mis à jour

PHASE 4 — Validation
  □ Facture pilote émise et validée visuellement
  □ Test envoi effectué (sur sa propre adresse)
  □ Statut école → Active
```

---

## Cas particuliers

### Migration d'une école existante non-onboardée formellement

Pour les écoles déjà actives mais sans onboarding propre (ex. EMN au statut actuel) :

1. Faire la **Phase 1** rétroactivement (collecter les infos manquantes)
2. Compléter le MASTER (Coordonnées, Conventions_horaires si manquantes)
3. Sauter la Phase 4 si des factures ont déjà été émises

### Mission ponctuelle (jury, intervention courte)

Process allégé :
- Phase 1 : minimum vital (raison sociale, SIRET, contact, taux/forfait)
- Phase 2 : skip
- Phase 3 : statut mission = `Ponctuelle` au lieu d'`Active`
- Phase 4 : facture unique = la facture pilote

### École sur deux entités (ex. CDP International présente sur ROSARI ET KAUR)

**Étanchéité absolue** (cf. [ADR001](ADR.md)) : faire l'onboarding **deux fois**, une dans chaque MASTER. Pas de mutualisation des données.

---

## Documents à fournir au client

Pour faciliter sa Phase 1, on peut lui transmettre un formulaire récapitulatif. À créer (cf. ticket à venir) :
- `_templates/onboarding_form.docx` — formulaire pré-rempli pour le client
- `_templates/onboarding_form.pdf` — version PDF imprimable
