# ADR — Architecture Decision Records

**Dernière mise à jour** : 2026-05-04

> Toutes les décisions structurantes du projet, dans l'ordre chronologique. Chaque ADR est immuable une fois acté — si une décision change, créer un nouvel ADR qui supersède l'ancien.

Format de chaque ADR : **Contexte** → **Décision** → **Conséquences** → **Alternatives écartées**.

---

## ADR001 — Étanchéité totale ROSARI / KAUR

**Date** : 2026-04-22
**Statut** : Acté

### Contexte
Deux auto-entrepreneurs distincts (ROSARI et KAUR) avec deux SIRETs et deux comptes bancaires séparés. Obligations comptables strictes : pas de mélange entre les deux entités.

### Décision
**MASTERs séparés** (`ROSARI_MASTER.xlsx` + `KAUR_MASTER.xlsx`), **dossiers séparés**, **numérotations séparées**. Aucun croisement automatique.

### Conséquences
- Duplication de structure entre les deux entités (acceptable)
- Code Python doit prendre l'entité en paramètre explicitement
- La consolidation ne se fait que **côté dashboard** en lecture seule

### Alternatives écartées
- MASTER unique avec colonne `Entité` : risque de croisement involontaire, fragilité comptable
- Dashboards séparés : perte de la vision foyer

---

## ADR002 — Convention de numérotation `R_/K_<ECOLE><NNN>`

**Date** : 2026-04-26
**Statut** : Acté (à partir d'avril 2026)

### Contexte
Avant : numérotation hétérogène (`EPITA001`, `DOR002`, `1`, `2`…). Manque de lisibilité multi-entités.

### Décision
Format : `<PRÉFIXE_ENTITÉ>_<CODE_ÉCOLE><NNN>`
- `R_` pour ROSARI, `K_` pour KAUR
- Code école court (DOR, EPITA, ASC, CDP, IDRAC…)
- Séquence 3 chiffres **par école**, indépendante par entité

Exemples : `R_DOR001`, `K_ASC001`.

### Conséquences
- Lecture immédiate de l'entité et de l'école dans le N°
- Séquences indépendantes par école, simple à incrémenter
- Compatibilité avec systèmes externes (Tiime, Shine) qui acceptent ce format

### Alternatives écartées
- Séquence globale par entité : perte de l'information école, gestion plus complexe
- Format avec année (`R_DOR_2026_001`) : trop verbeux

### Note
Les factures émises **avant avril 2026** conservent leur ancienne numérotation. Pas de renommage rétroactif.

---

## ADR003 — Onglet `Sessions` append-only

**Date** : 2026-04-26
**Statut** : Acté

### Contexte
Besoin d'un audit trail des séances rattachées à chaque facture, pour litige client ou contrôle fiscal.

### Décision
Onglet `Sessions` dans chaque MASTER, mode **append-only** strict. Schéma :
`Date_log | École | Période_facture | N°_Facture | Date_séance | Début | Fin | Durée_h | Promotion | Module/UE | Taux | Total HT | Summary calendrier verbatim`

### Conséquences
- Croissance linéaire du fichier (acceptable pour 5–10 ans d'activité)
- Possibilité de retracer rétroactivement n'importe quelle facture
- Robustesse en cas de litige

### Alternatives écartées
- Onglet par facture : multiplie les onglets, ingérable
- Stockage hors MASTER : casse le principe "MASTER = source de vérité unique"

---

## ADR004 — Excel + openpyxl, pas HTML→PDF

**Date** : 2026-05-04
**Statut** : Acté

### Contexte
Génération de factures sur des layouts complexes (bandeau, encadrés, tableaux). Deux approches possibles :
1. Excel + openpyxl + export PDF
2. HTML + CSS + WeasyPrint (HTML→PDF)

### Décision
**Excel + openpyxl + LibreOffice CLI** pour l'export PDF.

### Conséquences
- Garde le contrôle sur les fichiers XLSX (édition manuelle possible si besoin)
- Compatibilité directe avec les attentes des clients (XLSX = norme freelance)
- Plus de friction sur la mise en forme via openpyxl (acceptée)
- Nécessite installation LibreOffice pour l'export PDF en CLI

### Alternatives écartées
- HTML→PDF : perte du XLSX éditable, dépendance à WeasyPrint, courbe d'apprentissage CSS
- Export Excel natif via macro VBA : non portable, dépendance Microsoft Office

---

## ADR005 — Source de vérité = MASTER, pas calendrier

**Date** : 2026-04-26
**Statut** : Acté

### Contexte
Confusion possible entre "ce qui est planifié dans le calendrier" et "ce qui doit être facturé". Un cours peut être annulé, déplacé, ou ajouté hors calendrier.

### Décision
Le **MASTER** de chaque entité est la **source de vérité unique**. Le calendrier sert uniquement à :
- Détecter les nouvelles séances non encore reportées dans le MASTER
- Documenter le `Summary calendrier verbatim` dans l'onglet `Sessions`

### Conséquences
- Toute génération de facture doit d'abord consulter le MASTER
- Les écarts calendrier ↔ MASTER sont normaux et documentés (séances annulées, ajoutées, etc.)
- Discipline humaine requise pour reporter les changements dans le MASTER

### Alternatives écartées
- Calendrier comme source : fragile (annulations non répercutées, etc.)
- Double source : risque de divergence sans arbitrage clair

### À évoluer
Plus tard, possibilité d'ajouter un document tiers reprenant infos calendrier + infos contractuelles + mails de planification. Décision repoussée.

---

## ADR006 — Migration partielle Tiime/Shine envisagée (post v1.0)

**Date** : 2026-05-04
**Statut** : Acté (intention long terme)

### Contexte
Outil maison performant pour le portfolio mais représente une charge de maintenance non négligeable. Tiime ou Shine offrent : conformité fiscale assurée, déclarations URSSAF auto, ~12–50 €/mois.

### Décision
**Approche hybride post-v1.0** :
- L'outil maison reste actif (portfolio LinkedIn + missions complexes)
- Tiime ou Shine peut être adopté **à la carte** pour certaines factures (nouveaux clients, paiements complexes)
- Pas d'imposition d'un outil tiers — l'utilisateur choisit au cas par cas

### Conséquences
- Réconciliation annuelle obligatoire entre les deux côtés (export CSV Tiime → MASTER)
- Risque de divergence à surveiller
- Liberté maximale pour l'utilisateur

### Alternatives écartées
- Migration totale vers Tiime : perte du projet portfolio
- Refus de toute migration : friction maintenance à long terme

---

## ADR007 — Repo Git privé géré par Claude Code

**Date** : 2026-05-04
**Statut** : Acté

### Contexte
Versionner le projet via Git pour : traçabilité, mise en valeur sur LinkedIn (post-v1.0), backup, collaboration future.

### Décision
- Repo **privé** d'abord, public à v1.0 (avec données anonymisées)
- Initialisation et configuration **par Claude Code** (pas par Claude Chat)
- Convention : `main` (stable) / `develop` (en cours) / `feature/*` (nouveautés)
- Conventional commits : `feat:`, `fix:`, `docs:`, `refactor:`, `test:`
- `.gitignore` strict : exclure tous les `.xlsx` réels, garder uniquement les templates et exemples anonymisés

### Conséquences
- Historique propre du projet pour LinkedIn
- Rollback possible en cas d'erreur
- Friction minimale (Claude Code intégré au workflow)

### Alternatives écartées
- Pas de versioning : perte historique, pas de rollback
- Repo public dès le départ : exposition prématurée des données réelles

---

## ADR008 — Multi-utilisateur KAUR gelé pour stabilisation

**Date** : 2026-05-04
**Statut** : Acté (décision provisoire)

### Contexte
Question initiale : faut-il une interface dédiée pour Parv (KAUR) afin qu'elle puisse lancer ses propres facturations ?

### Décision
**Gelé pour stabilisation.** Adrian opère pour les deux entités à court ET long terme. Pas d'interface utilisateur séparée prévue.

Réévaluation possible **après v1.0** une fois le système stable.

### Conséquences
- Simplicité maximale (pas de surface UI à maintenir)
- Adrian reste opérateur unique
- Parv n'a pas à apprendre l'outil

### Alternatives écartées
- Scénario B (Streamlit comme interface) : ajout de complexité avant stabilité
- Scénario C (Claude Code chez Parv) : prématuré, courbe d'apprentissage

---

## ADR009 — Conventions horaires formalisées en table

**Date** : 2026-05-04
**Statut** : Acté

### Contexte
Connaissance métier critique aujourd'hui éparpillée :
- DOR M2 RMD = 8h cal / 8h facturable (1:1)
- DOR M2 DIA = 9h cal / 7h facturable (pause)
- CDP Masters = 9h cal / 8h facturable (1h pause)

Risque : oubli, divergence entre sessions Claude, calculs incorrects.

### Décision
Créer un onglet `Conventions_horaires` dans chaque MASTER avec :
`École | Module/Classe | H calendrier | H facturable | Coefficient | Justification`

Le PCF (étape 3) doit lire cette table et appliquer le coefficient automatiquement.

### Conséquences
- Source de vérité unique pour les conventions
- Calcul automatique sans intervention humaine
- Audit possible (justification documentée)

### Alternatives écartées
- Documentation prose dans CONTEXT.md : non exploitable par script
- Hard-code dans `pcf_format.py` : pas évolutif

---

## ADR010 — Convention de nomenclature .md alignée Loop Breaker

**Date** : 2026-05-04
**Statut** : Acté

### Contexte
Cohérence entre les projets perso (jeu roguelite, facturation, etc.). Pattern éprouvé : `CONTEXT.md` (permanent), `TASKS.md` (backlog), tickets atomiques `<SCOPE><NN>_<slug>.md` pour les gros chantiers, `ADR.md` pour les décisions.

### Décision
Migration des fichiers existants :
- `CLAUDE.md` actuel → scindé en `CLAUDE.md` (entrée minimale) + `CONTEXT.md` (détaillé)
- `TODO.md` → renommé `TASKS.md`
- PCF inline dans CLAUDE.md → migré vers `PROC00_pcf.md`

Tickets atomiques uniquement pour les **gros chantiers** (Option D, dashboards, scripts) — le reste reste dans `TASKS.md` monolithique.

ADR centralisés dans un seul fichier `ADR.md` (pas de dossier `DECISIONS/`).

### Conséquences
- Cohérence visuelle entre projets perso
- Granularité contrôlée (pas trop de fichiers)
- Continuité Claude Code optimisée (CLAUDE.md → CONTEXT.md → ticket en cours)

### Alternatives écartées
- TASKS.md monolithique pur : perte de focus sur les gros chantiers
- Tickets atomiques pour TOUT : explosion du nombre de fichiers, friction

---

## ADR011 — Onboarding nouveau client formalisé en checklist

**Date** : 2026-05-04
**Statut** : Acté

### Contexte
Aujourd'hui, ajouter une école = 7 étapes manuelles dispersées (Tarifs, Écoles, Coordonnées, dossier, fichier, classification CLAUDE.md, Conventions_horaires). Risque d'oubli systémique.

### Décision
Créer `ONBOARDING.md` à la racine = checklist exhaustive A→Z. Phase 1 (collecte info légale) → Phase 2 (collecte info pédagogique) → Phase 3 (création technique automatisée par Claude Code) → Phase 4 (validation par facture pilote).

### Conséquences
- Onboarding reproductible et fiable
- Réduction du risque d'oubli
- Support à fournir au client : liste claire des informations à transmettre

### Alternatives écartées
- Pas de processus formel : cause de bugs (coordonnées manquantes, taux non documentés, etc.)
- Process automatisé sans checklist : la checklist humaine reste utile pour la collecte d'info côté client

---

## Index des ADR

| # | Sujet | Statut |
|---|---|---|
| ADR001 | Étanchéité totale ROSARI / KAUR | Acté |
| ADR002 | Convention numérotation `R_/K_<ECOLE><NNN>` | Acté |
| ADR003 | Onglet Sessions append-only | Acté |
| ADR004 | Excel + openpyxl, pas HTML→PDF | Acté |
| ADR005 | Source de vérité = MASTER, pas calendrier | Acté |
| ADR006 | Migration partielle Tiime/Shine post-v1.0 | Acté (intention) |
| ADR007 | Repo Git privé géré par Claude Code | Acté |
| ADR008 | Multi-utilisateur KAUR gelé | Acté (provisoire) |
| ADR009 | Conventions horaires formalisées | Acté |
| ADR010 | Nomenclature .md alignée Loop Breaker | Acté |
| ADR011 | Onboarding nouveau client formalisé | Acté |
