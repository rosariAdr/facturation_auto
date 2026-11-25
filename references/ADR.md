# ADR — Architecture Decision Records

**Dernière mise à jour** : 2026-11-23

> **Ce fichier fait foi pour toutes les décisions du projet** (ADR025). Il réunit les ADR001 à 011 (v1), les ADR012 à 024 (v2, rapatriés du journal de décisions Notion le 23/11/2026) et les décisions suivantes. Notion n'en garde que l'index.
>
> Une décision actée ne se réécrit pas : si elle change, un nouvel ADR la remplace et le statut de l'ancien est mis à jour. Format : **Contexte** → **Décision** → **Conséquences** → **Alternatives écartées**.
>
> Aucune donnée réelle ici (règle de `AGENTS.md` §7) : les exemples sont fictifs. Les règles propres à chaque école vivent dans le MASTER (ADR026).

---

## Index

| # | Sujet | Statut |
|---|---|---|
| ADR001 | Étanchéité totale ROSARI / KAUR | Acté |
| ADR002 | Numérotation `R_/K_<ECOLE><NNN>` | Acté, complété par ADR029 |
| ADR003 | Onglet Sessions en ajout seul | Acté, amendé le 26/10/2026 |
| ADR004 | Excel + openpyxl, pas HTML→PDF | Acté |
| ADR005 | Source de vérité = MASTER, pas l'agenda | Acté |
| ADR006 | Migration partielle Tiime/Shine après la v1.0 | Acté (intention), précisé par ADR033 |
| ADR007 | Dépôt Git privé géré par Claude Code | Acté ; modèle de branches remplacé par `AGENTS.md` §1 |
| ADR008 | Multi-utilisateur KAUR gelé | Remplacé par ADR014 |
| ADR009 | Conventions horaires formalisées en table | Repris par ADR016 |
| ADR010 | Nomenclature des fichiers .md | Acté pour le dépôt ; `CONTEXT.md` abandonné (ADR026) |
| ADR011 | Onboarding d'une école en checklist | Acté |
| ADR012 | Exécution par le code, pas par MCP | Acté |
| ADR013 | Fichiers plats : ni base de données ni serveur | Acté |
| ADR014 | Multi-utilisateur par configuration | Acté (remplace ADR008) |
| ADR015 | Gabarit unique par entité + règles par école | Acté |
| ADR016 | Rapprochement agenda : code, sinon table Correspondances | Acté, précisé par ADR028 |
| ADR017 | Lancement en fin de mois | Remplacé par ADR024 |
| ADR018 | Ordre de livraison : moteur, appli, Parv | Acté |
| ADR019 | Lecture de l'agenda par adresse iCal secrète | Acté |
| ADR020 | Appli du lot 2 : appli web locale empaquetée en .exe | Acté, précisé par ADR032 |
| ADR021 | PDF par LibreOffice sans interface | Acté |
| ADR022 | Documentation : Notion pour piloter, le dépôt pour la référence technique | Remplacé par ADR025 |
| ADR023 | Tests : pytest, Playwright et banque de tests documentée | Acté |
| ADR024 | Lancement le 25 sur le mois entier, émission au 1er | Remplacé par ADR027 |
| ADR025 | La documentation fait foi dans le dépôt | Acté (23/11/2026) |
| ADR026 | Les règles réelles des écoles vivent uniquement dans le MASTER | Acté (23/11/2026) |
| ADR027 | Calendrier mensuel : préparation le 20, rappels, émission au 1er | Acté (23/11/2026) |
| ADR028 | Classement des créneaux : file à classer, règles gardées au choix | Acté (23/11/2026) |
| ADR029 | Numérotation : codes conservés, aucun numéro réutilisé | Acté (23/11/2026) |
| ADR030 | Contenu légal de la facture | Acté (23/11/2026) |
| ADR031 | Émission : validation par facture, reprise, suivi d'envoi et de paiement | Acté (23/11/2026) |
| ADR032 | Lot 2 : le design d'abord, deux postes, un seul écrivain | Acté (23/11/2026) |
| ADR033 | Facturation électronique : Tiime pour la réception | Acté (23/11/2026) |

---

# Décisions de la v1 (ADR001 à 011)

## ADR001 — Étanchéité totale ROSARI / KAUR

**Date** : 2026-04-22
**Statut** : Acté

### Contexte
Deux auto-entrepreneurs distincts (ROSARI et KAUR) avec deux SIRET et deux comptes bancaires séparés. Obligations comptables strictes : pas de mélange entre les deux entités.

### Décision
**MASTER séparés** (`ROSARI_MASTER.xlsx` + `KAUR_MASTER.xlsx`), **dossiers séparés**, **numérotations séparées**. Aucun croisement automatique.

### Conséquences
- Duplication de structure entre les deux entités (acceptable)
- Le code Python prend l'entité en paramètre explicitement
- La consolidation ne se fait que **côté tableau de bord**, en lecture seule

### Alternatives écartées
- MASTER unique avec colonne `Entité` : risque de croisement involontaire, fragilité comptable
- Tableaux de bord séparés : perte de la vision foyer

---

## ADR002 — Convention de numérotation `R_/K_<ECOLE><NNN>`

**Date** : 2026-04-26
**Statut** : Acté (à partir d'avril 2026) ; complété par ADR029 (codes, numéros manquants)

### Contexte
Avant : numérotation hétérogène (`ABC001`, `ECO002`, `1`, `2`…). Manque de lisibilité multi-entités.

### Décision
Format : `<PRÉFIXE_ENTITÉ>_<CODE_ÉCOLE><NNN>`
- `R_` pour ROSARI, `K_` pour KAUR
- Code école court, en majuscules
- Séquence de 3 chiffres **par école**, indépendante par entité

Exemples fictifs : `R_ABC001`, `K_ECO001`.

### Conséquences
- Lecture immédiate de l'entité et de l'école dans le numéro
- Séquences indépendantes par école, simples à incrémenter
- Compatibilité avec des services externes (Tiime, Shine) qui acceptent ce format

### Alternatives écartées
- Séquence globale par entité : perte de l'information école, gestion plus complexe
- Format avec année (`R_ABC_2026_001`) : trop verbeux

### Note
Les factures émises **avant avril 2026** conservent leur ancienne numérotation. Pas de renommage rétroactif.

---

## ADR003 — Onglet `Sessions` en ajout seul

**Date** : 2026-04-26
**Statut** : Acté ; amendé le 26/10/2026 (voir « Amendement »)

### Contexte
Besoin d'une piste d'audit des séances rattachées à chaque facture, pour un litige client ou un contrôle fiscal.

### Décision
Onglet `Sessions` dans chaque MASTER, en **ajout seul**. Schéma :
`Date_log | École | Période_facture | N°_Facture | Date_séance | Début | Fin | Durée_h | Promotion | Module/UE | Taux | Total HT | Summary calendrier verbatim`

### Amendement (26/10/2026, journal de décisions v2)
La mise à jour sur place est admise quand elle s'appuie sur le `N°_Facture` : une facture régénérée remplace ses propres lignes au lieu d'en ajouter de nouvelles, sinon chaque régénération duplique le journal et fausse les totaux. Hors de ce cas, aucune ligne n'est modifiée ni supprimée.

### Conséquences
- Croissance linéaire du fichier (acceptable pour 5 à 10 ans d'activité)
- Possibilité de retracer n'importe quelle facture
- Robustesse en cas de litige

### Alternatives écartées
- Un onglet par facture : multiplie les onglets, ingérable
- Stockage hors MASTER : casse le principe « MASTER = source de vérité »

---

## ADR004 — Excel + openpyxl, pas HTML→PDF

**Date** : 2026-05-04
**Statut** : Acté

### Contexte
Génération de factures sur des mises en page complexes (bandeau, encadrés, tableaux). Deux approches possibles :
1. Excel + openpyxl + export PDF
2. HTML + CSS + WeasyPrint (HTML→PDF)

### Décision
**Excel + openpyxl + LibreOffice en ligne de commande** pour l'export PDF.

### Conséquences
- Garde le contrôle sur les fichiers XLSX (édition manuelle possible si besoin)
- Compatibilité directe avec les attentes des clients
- Plus de friction sur la mise en forme via openpyxl (acceptée)
- Nécessite LibreOffice pour l'export PDF

### Alternatives écartées
- HTML→PDF : perte du XLSX éditable, dépendance à WeasyPrint
- Export Excel natif par macro VBA : non portable, dépendance à Microsoft Office

---

## ADR005 — Source de vérité = MASTER, pas l'agenda

**Date** : 2026-04-26
**Statut** : Acté

### Contexte
Confusion possible entre « ce qui est planifié dans l'agenda » et « ce qui doit être facturé ». Un cours peut être annulé, déplacé ou ajouté hors agenda.

### Décision
Le **MASTER** de chaque entité est la **source de vérité unique**. L'agenda sert à :
- détecter les séances pas encore reportées dans le MASTER ;
- documenter le `Summary calendrier verbatim` dans l'onglet `Sessions`.

### Conséquences
- Toute génération de facture consulte d'abord le MASTER
- Les écarts agenda ↔ MASTER sont normaux et documentés (séances annulées, ajoutées…)
- Les changements sont reportés dans le MASTER

### Alternatives écartées
- Agenda comme source : fragile (annulations non répercutées…)
- Double source : risque de divergence sans arbitrage clair

---

## ADR006 — Migration partielle Tiime/Shine envisagée (après la v1.0)

**Date** : 2026-05-04
**Statut** : Acté (intention long terme) ; précisé par ADR033 (Tiime retenue pour la réception des factures électroniques)

### Contexte
Outil maison utile, mais avec une charge de maintenance. Tiime ou Shine offrent la conformité fiscale et des déclarations simplifiées.

### Décision
**Approche hybride après la v1.0** :
- l'outil maison reste actif ;
- Tiime ou Shine peut être adopté **à la carte** pour certaines factures ;
- pas d'imposition d'un outil tiers : choix au cas par cas.

### Conséquences
- Réconciliation annuelle entre les deux côtés (export CSV → MASTER), F-52
- Risque de divergence à surveiller

### Alternatives écartées
- Migration totale vers un service tiers : perte du projet
- Refus de toute migration : friction de maintenance à long terme

---

## ADR007 — Dépôt Git privé géré par Claude Code

**Date** : 2026-05-04
**Statut** : Acté. Le modèle de branches ci-dessous est remplacé par celui de `AGENTS.md` §1 (`main` / `dev` / `feat|fix|chore/*`) depuis le 17/11/2026.

### Contexte
Versionner le projet pour la traçabilité, la sauvegarde et une collaboration future.

### Décision
- Dépôt **privé** d'abord, public à la v1.0 (avec données anonymisées)
- Convention d'origine : `main` (stable) / `develop` (en cours) / `feature/*` — remplacée, voir le statut
- Conventional Commits
- `.gitignore` strict : aucun `.xlsx` réel, seulement le gabarit et des exemples fictifs

### Conséquences
- Historique propre, retour arrière possible
- Avant le passage en public, vérifier que l'historique ne contient aucune donnée réelle

### Alternatives écartées
- Pas de versionnement : pas d'historique ni de retour arrière
- Dépôt public dès le départ : exposition prématurée

---

## ADR008 — Multi-utilisateur KAUR gelé pour stabilisation

**Date** : 2026-05-04
**Statut** : **Remplacé par ADR014** (26/10/2026)

### Contexte
Faut-il une interface dédiée pour Parv (KAUR) afin qu'elle lance ses propres facturations ?

### Décision (historique)
Gelé : Adrian opère pour les deux entités, pas d'interface séparée.

### Alternatives écartées à l'époque
- Streamlit comme interface : complexité avant stabilité
- Claude Code chez Parv : prématuré

---

## ADR009 — Conventions horaires formalisées en table

**Date** : 2026-05-04
**Statut** : **Repris par ADR016** (26/10/2026) : les conventions deviennent les colonnes « Mode heures » et « Valeur » de l'onglet Correspondances.

### Contexte
Connaissance métier critique éparpillée. Exemples (rendus fictifs le 23/11/2026) :
- École A, module X : 8 h d'agenda = 8 h facturées (1:1)
- École A, module Y : 9 h d'agenda = 7 h facturées (pause non facturée)
- École B : 9 h d'agenda = 8 h facturées (1 h de pause)

Risque : oubli, divergence entre sessions de travail, calculs faux.

### Décision
Une table des conventions dans chaque MASTER, lue par le PCF, qui applique le coefficient automatiquement.

### Conséquences
- Source unique pour les conventions, calcul sans intervention humaine, justification documentée

### Alternatives écartées
- Prose dans `CONTEXT.md` : non exploitable par un script
- Valeurs codées en dur dans le module : pas évolutif

---

## ADR010 — Nomenclature des fichiers .md

**Date** : 2026-05-04
**Statut** : Acté pour le dépôt. La partie `CONTEXT.md` est abandonnée (ADR026) ; `TASKS.md` et `CLAUDE.md` restent des fichiers locaux sans règle propre (ADR025).

### Contexte
Cohérence entre projets personnels : un fichier permanent, un backlog, des tickets pour les gros chantiers, un fichier d'ADR.

### Décision
- ADR centralisés dans un seul fichier `ADR.md` (pas de dossier `DECISIONS/`)
- Tickets atomiques seulement pour les gros chantiers

### Conséquences
- Granularité contrôlée, pas d'explosion du nombre de fichiers

---

## ADR011 — Onboarding d'une école formalisé en checklist

**Date** : 2026-05-04
**Statut** : Acté

### Contexte
Ajouter une école demandait des étapes manuelles dispersées. Risque d'oubli systémique.

### Décision
`references/ONBOARDING.md` = checklist de A à Z : collecte légale → collecte pédagogique → création technique → facture pilote.

### Conséquences
- Onboarding reproductible ; liste claire des informations à demander à l'école

### Alternatives écartées
- Pas de processus formel : coordonnées manquantes, taux non documentés
- Process automatisé sans checklist : la collecte côté école reste humaine

---

# Décisions de la v2 (ADR012 à 024, rapatriées de Notion le 23/11/2026)

## ADR012 — Exécution par le code, pas par MCP

**Date** : 2026-10-26
**Statut** : Acté

### Contexte
La v1 s'appuyait sur Claude et ses connecteurs pour lire l'agenda et rapprocher les séances.

### Décision
Au moment de facturer, seul du code déterministe s'exécute : aucun appel MCP, aucun modèle d'IA. Claude sert à développer, tester et documenter.

### Conséquences
- Chaque règle métier devient une donnée (MASTER) ou du code testé
- Le résultat est reproductible et vérifiable
- Toute lecture d'un service extérieur (agenda, Gmail, banque) passe par du code ou par un fichier exporté

---

## ADR013 — Fichiers plats : ni base de données ni serveur

**Date** : 2026-10-26
**Statut** : Acté

### Décision
Données dans des classeurs Excel, configuration en TOML, état d'exécution en JSON Lines, tout sur le PC de l'utilisateur.

### Conséquences
- Rien à héberger ni à payer ; données lisibles et modifiables à la main
- Écriture concurrente évitée par un seul écrivain (l'outil) et une sauvegarde avant chaque écriture (voir ADR032 pour plusieurs postes)

---

## ADR014 — Multi-utilisateur par configuration (remplace ADR008)

**Date** : 2026-10-26
**Statut** : Acté

### Contexte
Adrian veut partager l'outil avec Parv (entité KAUR, son propre agenda), et peut-être d'autres formateurs plus tard.

### Décision
Chaque utilisateur a son profil local, son MASTER et son agenda ; aucune valeur d'entité en dur dans le code. Parv arrive au lot 3, les autres formateurs au lot 4.

### Conséquences
- Configuration par utilisateur construite dès le lot 0
- Étanchéité de l'ADR001 vérifiée par un test (T-CFG-01)

---

## ADR015 — Gabarit unique par entité + règles par école

**Date** : 2026-10-26
**Statut** : Acté

### Décision
Une seule mise en page, seules les informations de l'émetteur changent ; chaque école a ses règles (onglet Règles_facturation) ; un gabarit attitré n'est créé que si une école ne rentre pas dans le gabarit unique.

### Conséquences
- Une correction de format profite à toutes les écoles
- Les dérogations sont visibles dans les règles
- Depuis la v0.2, le gabarit commun est `facturation/assets/PCF_facture_standard.xlsx` ; un gabarit propre à une entité est prioritaire s'il existe (`references/FORMAT.md`)

---

## ADR016 — Rapprochement agenda : code, sinon table Correspondances (reprend ADR009)

**Date** : 2026-10-26
**Statut** : Acté (format exact du code : proposé) ; précisé par ADR028

### Décision
Un code écrit dans l'événement fait foi ; à défaut, l'intitulé est cherché dans l'onglet Correspondances (exact, puis par préfixe) ; sinon l'événement va dans la file à classer et n'est jamais facturé par défaut. Les conventions horaires de l'ADR009 deviennent les colonnes « Mode heures » et « Valeur » de Correspondances.

Format proposé du code : `#ÉCOLE-PROMO-MODULE` dans la description de l'événement, par exemple `#ABC-M1-STATS` (fictif).

### Conséquences
- L'agenda n'a pas besoin d'être modifié ; ajouter un code supprime toute ambiguïté

---

## ADR017 — Lancement en fin de mois, émission au 1er du mois suivant

**Date** : 2026-10-26
**Statut** : **Remplacé par ADR024** (28/10/2026)

### Décision (historique)
L'outil tourne le dernier jour du mois sur le mois en cours ; les factures sont datées du 1er du mois suivant.

---

## ADR018 — Ordre de livraison : moteur, appli, Parv

**Date** : 2026-10-26
**Statut** : Acté

### Décision
Lot 1 : moteur, audit et planification pour toutes les écoles d'Adrian. Lot 2 : appli .exe. Lot 3 : Parv. La configuration multi-utilisateur est prévue dès le lot 0 pour ne rien réécrire. Détail des lots : `references/ROADMAP.md`.

---

## ADR019 — Lecture de l'agenda par adresse iCal secrète

**Date** : proposé le 2026-10-26, acté le 2026-10-29
**Statut** : Acté

### Contexte
Deux voies : l'adresse iCal secrète de l'agenda (aucun projet Google Cloud, aucune connexion OAuth, un simple fichier .ics à lire), ou l'API Google Agenda (temps réel, mais projet Google Cloud et application publiée en production pour éviter l'expiration des jetons à 7 jours).

### Décision
iCal pour le lot 1, derrière une interface qui permet de brancher l'API sans toucher au reste. L'adresse est rangée dans `agenda.secret`, à côté du profil, jamais affichée.

### Conséquences
- À vérifier au lot 1 : le délai de mise à jour du flux iCal après une modification d'agenda (test T-AGD-01)

---

## ADR020 — Appli du lot 2 : appli web locale empaquetée en .exe

**Date** : proposé le 2026-10-26, acté le 2026-10-29
**Statut** : Acté ; précisé par ADR032 (phase de design d'abord)

### Décision
Une interface servie en local par Python, empaquetée en .exe avec PyInstaller. Le choix entre NiceGUI et FastAPI + HTMX se fait au démarrage du lot 2, après la phase de design (ADR032).

### Conséquences
- Testable par Playwright comme une page web, une seule langue pour tout le projet, aucun serveur distant

---

## ADR021 — PDF par LibreOffice en mode sans interface

**Date** : proposé le 2026-10-26, acté le 2026-10-29
**Statut** : Acté

### Décision
LibreOffice en mode sans interface : gratuit, scriptable, même rendu chez Parv. Si le rendu diverge trop de celui d'Excel sur le PC d'Adrian, export par Excel en option (Q7, à juger lors du premier essai complet sur le PC habituel).

---

## ADR022 — Documentation : Notion pour piloter, le dépôt pour la référence technique

**Date** : 2026-10-26
**Statut** : **Remplacé par ADR025** (23/11/2026)

### Décision (historique)
Notion portait cadrage, planification, décisions, banque de tests et journal de bord ; le dépôt gardait la référence technique, et renvoyait à Notion pour les décisions à partir de l'ADR012.

---

## ADR023 — Tests : pytest, Playwright et banque de tests documentée

**Date** : 2026-10-26
**Statut** : Acté

### Décision
Tests unitaires, d'intégration et de non-régression avec pytest ; tests d'utilisation de l'appli avec Playwright ; chaque test inscrit dans la banque de tests avec sa nature et son dernier résultat, pour être rejoué à chaque évolution.

---

## ADR024 — Lancement le 25 sur le mois entier, émission au 1er du mois suivant (remplace ADR017)

**Date** : 2026-10-28
**Statut** : **Remplacé par ADR027** (23/11/2026). Restent valables : préparation du mois entier, date d'émission au 1er du mois suivant, aucun envoi par l'outil.

### Décision (historique)
L'outil préparait les factures le 25 de chaque mois, pour le mois en cours tout entier : séances déjà tenues et séances encore prévues. Factures datées du 1er du mois suivant. Du 25 au 1er, Adrian vérifiait et programmait l'envoi lui-même. Heure du lancement : 16 h (Q6, 29/10/2026).

### Conséquences reprises par ADR027
- L'aperçu distingue les séances tenues des séances à venir et peut être régénéré jusqu'à la validation
- Après émission, une séance facturée puis annulée est signalée par l'audit et se corrige par un avoir, jamais en modifiant une facture émise (F-23)

---

# Décisions du cadrage du 23/11/2026 (ADR025 à 033)

Source : réponses d'Adrian au questionnaire de cadrage du 23/11/2026 et audit d'une facture émise, mené le 21/11/2026. Les points marqués `💡 [proposé — à valider]` sont des précisions de mise en œuvre qui restent à confirmer.

## ADR025 — La documentation fait foi dans le dépôt (remplace ADR022)

**Date** : 2026-11-23
**Statut** : Acté

### Contexte
Le dépôt (`AGENTS.md`) et le hub Notion se disaient chacun « source de vérité ». Les ADR012 à 024 n'existaient que dans Notion : un agent ou une personne qui découvrait le dépôt ne les voyait pas, et le hub n'était plus à jour.

### Décision
- Le **dépôt** fait foi pour : les conventions de travail (`AGENTS.md`), toutes les décisions (ce fichier), le comportement de l'outil (`references/`), la feuille de route et les critères d'acceptation (`references/ROADMAP.md`), le vocabulaire (`references/GLOSSAIRE.md`).
- **Notion** porte l'avancement : statut des fonctionnalités, banque de tests et derniers résultats, journal de bord, to-do. Il renvoie au dépôt pour les règles et ne recopie pas le texte des ADR.
- Le **MASTER** porte les données réelles (ADR026).
- Les fichiers locaux (`CLAUDE.md`, `memory/`, `TASKS.md`) ne portent aucune règle qui n'existe pas aussi dans le dépôt.

### Conséquences
- ADR012 à 024 rapatriés ici le 23/11/2026
- Une règle vit dans un seul fichier ; les autres y renvoient par un lien
- Un nouvel agent comprend le projet en lisant le dépôt seul ; `references/RECETTE_DOC.md` le vérifie

### Alternatives écartées
- Notion comme référence : non versionné, invisible sans connecteur, déjà en retard sur le dépôt
- Double tenue : divergence constatée le 22/11/2026

---

## ADR026 — Les règles réelles des écoles vivent uniquement dans le MASTER

**Date** : 2026-11-23
**Statut** : Acté

### Contexte
Taux, rythmes, heures facturables, mentions et canaux d'envoi propres à chaque école étaient éparpillés : MASTER, archives de conversations, mémoire de l'assistant, page Notion, `CONTEXT.md` prévu. Le dépôt n'accepte aucune donnée réelle.

### Décision
- Toute règle propre à une école est une donnée du MASTER de l'entité : `Tarifs`, `Règles_facturation`, `Correspondances`, `Coordonnées` (`references/MASTER_SCHEMA.md`).
- Le **contrat signé fait foi pour un taux** ; le MASTER en garde la référence (colonne « Réf. contrat » de `Tarifs`).
- La liste des écoles d'un lot (par exemple celles du lot 1) est une donnée du MASTER, pas du dépôt.
- `CONTEXT.md` n'est plus utilisé.

### Conséquences
- Le protocole, l'onboarding et `SKILL.md` ne renvoient plus à `CONTEXT.md`
- Toute autre source de taux (par exemple un tableau de bord du chiffre d'affaires tenu hors du dépôt) s'aligne sur le MASTER

### Alternatives écartées
- `CONTEXT.md` à côté du MASTER : de la prose, non exploitable par le code (ADR009)
- Règles dans le dépôt : interdit par `AGENTS.md` §7

---

## ADR027 — Calendrier mensuel : préparation le 20, rappels, émission au 1er (remplace ADR024)

**Date** : 2026-11-23
**Statut** : Acté

### Contexte
Préparer le 25 laissait peu de marge, et des échéances d'école ont été manquées en octobre et novembre 2026. Adrian veut produire les factures en avance et être relancé tant qu'il n'a pas validé.

### Décision
- **Le 20** : l'outil prépare les factures du mois entier (séances tenues et séances encore prévues) et demande la validation par une notification Windows.
- **Le 25** : nouvelle notification de validation.
- **Du 28 au dernier jour du mois** : une notification par jour tant qu'une facture reste à valider.
- **Date d'émission** : le 1er du mois suivant (inchangée).
- Une école qui a son propre cycle ou une date limite plus tôt est préparée avant cette date, selon sa règle dans `Règles_facturation`.
- Si le PC est éteint, le lancement a lieu au démarrage de session suivant. L'heure reste un réglage du profil (16 h par défaut, Q6).
- L'outil n'envoie jamais rien : Adrian programme l'envoi lui-même.

### Conséquences
- Le profil passera du jour 25 au jour 20 par défaut avec F-20 ; le code accepte aujourd'hui un jour de 1 à 28 et vaut 25 par défaut
- Entre le 20 et la validation, l'aperçu signale les séances ajoutées, déplacées ou annulées dans l'agenda
- Une séance modifiée après émission est signalée, puis corrigée par un avoir (F-23)

### Alternatives écartées
- Le 25 (ADR024) : trop peu de marge
- La fin de mois (ADR017) : aucune marge

---

## ADR028 — Classement des créneaux : file à classer, règles gardées au choix (précise ADR016)

**Date** : 2026-11-23
**Statut** : Acté

### Décision
- Tout créneau de l'agenda des cours est candidat.
- Reconnu (code d'événement, puis `Correspondances`) : il entre dans l'aperçu.
- Non reconnu : il va dans la file à classer et n'est jamais facturé par défaut.
- Quand Adrian classe ou corrige une séance, il choisit si la correction devient une règle (une ligne de `Correspondances`) ou reste ponctuelle.

### Conséquences
- Le critère de F-12 (« classé une fois l'est automatiquement ensuite ») devient « classé et gardé comme règle » ; idem pour F-31

### Alternatives écartées
- Facturer d'office les créneaux non reconnus : risque de facturer une vacation ou une réunion
- Mémoriser toujours la correction : une exception ponctuelle deviendrait une règle

---

## ADR029 — Numérotation : codes conservés, aucun numéro réutilisé (complète ADR002)

**Date** : 2026-11-23
**Statut** : Acté

### Décision
- Format inchangé : `<préfixe d'entité>_<code école><NNN>` (ADR002).
- Les codes école existants sont conservés, quelle que soit leur longueur. Une nouvelle école reçoit un code de **3 lettres**, unique dans l'entité.
- Le numéro suivant est le plus grand numéro de la série plus un.
- Un numéro manquant n'est jamais réutilisé : sa raison est notée et la série continue après le plus grand numéro.
- Un numéro n'est attribué qu'à l'émission : un brouillon abandonné n'en consomme pas (F-15).
- Les séries antérieures à avril 2026 ne sont pas renommées.

### Conséquences
- `💡 [proposé — à valider]` la raison d'un numéro manquant est notée dans le registre des factures (onglet `Factures`, `references/MASTER_SCHEMA.md`)

### Alternatives écartées
- Tout passer à 3 lettres, anciennes écoles comprises : rupture de lecture des séries existantes
- Combler un trou après coup : casse l'ordre chronologique de la série

---

## ADR030 — Contenu légal de la facture

**Date** : 2026-11-23
**Statut** : Acté

### Contexte
L'audit, le 21/11/2026, d'une facture émise a relevé des mentions absentes ou imprécises : statut d'entrepreneur individuel, indemnité forfaitaire de recouvrement, échéance datée, libellé de TVA, références. Le questionnaire du 23/11/2026 a tranché la présentation.

### Décision
Toute facture porte :
1. **Émetteur** : nom et prénom suivis de « EI » (ou « Entrepreneur individuel »), adresse, SIRET, email, téléphone.
2. **Client** : dénomination juridique de la société qui paie, telle qu'écrite au contrat (pas le nom commercial), et son adresse de facturation. Le n° de TVA du client n'est pas affiché quand il n'est pas requis (vérifié le 23/11/2026 pour l'école pilote ; à vérifier à l'onboarding de chaque école).
3. **Références**, sur trois lignes distinctes : n° de facture ; n° de contrat quand il existe ; n° de bon de commande quand l'école en a émis un. Une facture par contrat quand l'école en a plusieurs (règle d'école).
4. **Dates** : date d'émission (ADR027), dates des séances ou période, et une **date d'échéance explicite** : le délai du contrat ou de l'école ; à défaut, la date d'émission plus 30 jours.
5. **Totaux** : « Total HT », la mention de franchise, « Net à payer ». Pas de ligne « TVA : 0 € ». Les libellés « À payer TTC » et « Montant non assujetti à la TVA » sont retirés.
6. **TVA** : « TVA non applicable, article 293 B du Code général des impôts. » La TVA passe au Code des impositions sur les biens et services (CIBS) le 01/01/2027 ; les factures peuvent continuer à citer le CGI jusqu'au 30/06/2028 (BOI-RES-TVA-000253, vérifié le 23/11/2026). Le libellé est à revoir avant cette date. Aucune mention d'exonération formation (art. 261-4-4° a du CGI) tant que l'attestation n'est pas accordée (demande pas encore déposée au 23/11/2026).
7. **Paiement** : pénalités de retard à trois fois le taux d'intérêt légal, exigibles dès le lendemain de l'échéance, sans rappel préalable ; indemnité forfaitaire de 40 € pour frais de recouvrement, hors du net à payer ; « Pas d'escompte en cas de paiement anticipé ».
8. **N° de déclaration d'activité** (organisme de formation) sur toutes les factures, lu dans le MASTER.
9. **Mise en page** : la facture pilote tient sur une page. La règle pour les factures plus longues sera décidée après le pilote.
10. **Blocage** : si la valeur d'une mention obligatoire manque, l'émission de cette facture est bloquée (F-17).

Sources à consulter pour la recette : Code de commerce, articles L441-9 et L441-10 ; article 293 B du CGI ; fiches Service-Public Entreprendre F31808 (mentions obligatoires) et F23211 (délais et pénalités).

### Conséquences
- Gabarit et module : F-24 (`references/FORMAT.md`) ; données : `Coordonnées` et `Règles_facturation` (`references/MASTER_SCHEMA.md`) ; collecte : `references/ONBOARDING.md`
- À appliquer dès les factures de novembre 2026, datées du 01/12/2026 : par le gabarit si F-24 est livré à temps, sinon par retouche avant l'envoi
- La facture historique auditée reste un constat d'audit ; la rectifier est une décision métier distincte
- Si la franchise en base cesse (dépassement de seuil) ou si l'attestation d'exonération formation est accordée, les mentions de TVA changent : un nouvel ADR le décidera

### Alternatives écartées
- Afficher « TVA : 0 € » : non retenu, la mention de franchise suffit
- Garder « payable fin de mois » : échéance ambiguë

---

## ADR031 — Émission : validation par facture, reprise, suivi d'envoi et de paiement

**Date** : 2026-11-23
**Statut** : Acté

### Décision
- **Validation** : chaque facture est validée seule. Une facture incomplète est bloquée sans retenir les autres.
- **Reprise** : l'émission enchaîne des étapes (onglet du classeur, journal `Sessions`, suivi `Cours`, registre des factures, PDF, puis brouillon de mail au lot 2). L'état de chaque étape est consigné (journal d'exécution en JSON Lines, ADR013). Après un arrêt, l'outil montre où il en est et ne refait que les étapes restantes : une commande au lot 1, un bouton au lot 2.
- **Envoyée** : le mail envoyé depuis Gmail fait foi.
- **Payée** : Adrian reporte lui-même la date de paiement dans le MASTER.
- **Tableau de bord de suivi** : au lot 2 (F-35).

### Conséquences
- `💡 [proposé — à valider]` tant que l'outil ne lit pas Gmail, Adrian reporte aussi la date d'envoi dans le MASTER
- Le préfixe « PAID » ajouté aux onglets réglés reste un usage de lecture, sans valeur de preuve
- Une facture échue et non payée devient visible dans le registre des factures (F-25)

---

## ADR032 — Lot 2 : le design d'abord, deux postes, un seul écrivain (précise ADR020)

**Date** : 2026-11-23
**Statut** : Acté

### Décision
- Après le lot 1 et **avant tout développement d'écran**, une phase de design : parcours écran par écran, maquettes cliquables, charte et couleurs, essai avec Adrian sur un vrai mois.
- L'outil de design est libre (ChatGPT est envisagé). Aucune donnée réelle n'y est partagée : seulement des données fictives. Les livrables reviennent dans le dépôt sous une forme lisible par un agent (Markdown, images, HTML) ; `💡 [proposé — à valider]` dans `references/design/`.
- Le choix NiceGUI ou FastAPI + HTMX (ADR020) se fait après cette phase.
- Première recette de l'appli sur deux postes Windows : celui d'Adrian et celui de Parv, avec des données fictives tant que le lot 3 n'a pas commencé.
- Un seul poste écrit dans un MASTER donné à la fois.
- Brouillons Gmail : la connexion de l'outil à Gmail n'est pas encore acceptée (23/11/2026). Quand elle le sera, un brouillon par facture, avec le PDF seul en pièce jointe.

### Conséquences
- Les critères de sortie du lot 2 comprennent : usage rapide, comparaison des factures pour valider vite, mails écrits avec leur pièce jointe, tableau de bord de suivi (`references/ROADMAP.md`)

---

## ADR033 — Facturation électronique : Tiime pour la réception (précise ADR006)

**Date** : 2026-11-23
**Statut** : Acté

### Décision
- Tiime est la plateforme agréée retenue.
- Réception des factures électroniques : à tester (pas encore fait au 23/11/2026).
- Émission : plateforme et premier parcours à décider avant l'échéance d'émission de septembre 2027 (F-51).

### Conséquences
- Le pont avec l'outil (F-52) et la conformité de l'émission (F-51) restent au lot 4
