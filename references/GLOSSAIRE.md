# GLOSSAIRE — Vocabulaire du projet

**Dernière mise à jour** : 2026-11-23

> Les termes employés dans le dépôt, dans le code et dans les échanges. Exemples fictifs uniquement : les noms réels des écoles et leurs codes vivent dans le MASTER (onglet `Écoles`).

| Terme | Sens | Où en savoir plus |
|---|---|---|
| **Agenda des cours** | L'agenda Google où figurent tous les cours de la personne qui facture ; l'outil le lit par son adresse iCal secrète. | ADR019 |
| **Aperçu** | Ce que l'outil montre avant toute écriture : séances, heures, taux, totaux, mentions, informations manquantes. Rien n'est écrit sans validation. | `PROC00_pcf.md` |
| **Audit** | Comparaison entre l'agenda et ce qui a été facturé : séances tenues non facturées, séances modifiées après émission. | F-19, F-23 |
| **Avoir** | Document qui corrige une facture émise (par exemple une séance annulée après coup). Une facture émise n'est jamais modifiée. | ADR024, F-23 |
| **Bon de commande (BC, PO)** | Document émis par certaines écoles avant la prestation ; son numéro figure alors sur la facture, sur sa propre ligne. | ADR030 |
| **Code d'événement** | Code écrit dans la description d'un événement d'agenda, format proposé `#ÉCOLE-PROMO-MODULE` (ex. `#ABC-M1-STATS`) ; il fait foi sur l'intitulé. | ADR016 |
| **Code école** | Identifiant court d'une école dans une entité (ex. `ABC`) ; 3 lettres pour toute nouvelle école. | ADR029 |
| **Contrôle après écriture** (post-mortem) | Relecture des cellules clés d'une facture écrite pour vérifier qu'elles correspondent à l'aperçu validé. | `PROC00_pcf.md` |
| **Correspondances** | Onglet du MASTER (prévu) qui relie un code ou un intitulé d'événement à une école, un module, une clé de tarif et un mode de calcul des heures. | ADR016, `MASTER_SCHEMA.md` |
| **Créneau, séance** | Un événement de l'agenda des cours ; une fois reconnu et rattaché à une école, c'est une séance à facturer. | ADR028 |
| **Échéance** | Date à laquelle le paiement est dû, écrite sur la facture : délai du contrat ou de l'école, sinon date d'émission plus 30 jours. | ADR030 |
| **EI** | Entrepreneur individuel : la mention suit le nom de l'émetteur sur la facture. | ADR030 |
| **Émission** | Écriture de la facture validée (classeur, journal, suivi, PDF). | ADR031 |
| **Entité** | La structure qui facture (ex. `ROSARI`, `KAUR`) ; chaque entité a son MASTER, ses dossiers, sa numérotation. | ADR001 |
| **Facture pilote** | Première facture complète d'une école, ou du premier parcours complet de l'outil, vérifiée point par point. | `ONBOARDING.md` |
| **File à classer** | Créneaux que l'outil ne reconnaît pas ; jamais facturés par défaut. | ADR028 |
| **Franchise en base** | Régime de TVA de l'émetteur : pas de TVA facturée, mention de l'article 293 B du CGI. | ADR030 |
| **Gabarit** | Classeur de mise en page sans aucune donnée (`facturation/assets/PCF_facture_standard.xlsx`, onglet `MODELE`). | `FORMAT.md` |
| **Heures facturables** | Heures retenues sur la facture : celles de l'agenda, ou un calcul (coefficient, forfait) défini dans Correspondances. | F-13 |
| **Journal Sessions** | Onglet du MASTER qui garde une ligne par séance facturée ; piste d'audit. | ADR003 |
| **Lot** | Étape de livraison du projet (0 socle, 1 moteur, 2 appli, 3 Parv, 4 plus tard). | `ROADMAP.md` |
| **MASTER** | Classeur `<ENTITE>_MASTER.xlsx` de chaque entité : source de vérité des données réelles. | ADR005, ADR026 |
| **NDA** | Numéro de déclaration d'activité d'organisme de formation, affiché sur les factures. | ADR030 |
| **Numéro de facture** | `<préfixe d'entité>_<code école><NNN>`, ex. `R_ABC008`. | ADR002, ADR029 |
| **PAID** | Préfixe ajouté à la main au nom de l'onglet d'une facture réglée ; usage de lecture, sans valeur de preuve. | ADR031 |
| **PCF** | Protocole de création de factures : les étapes et les validations humaines. | `PROC00_pcf.md` |
| **Porte** | La commande qui doit être verte avant toute fusion : `python -m pytest`. | `AGENTS.md` §5 et §7 |
| **Préparation du 20** | Construction mensuelle des factures le 20, suivie de rappels jusqu'à validation. | ADR027 |
| **Profil** | Fichier TOML de la personne qui facture, hors du dépôt (entité, agenda, racine de données, lancement, PDF). | `README.md` |
| **Racine de données** | Dossier hors dépôt qui contient `<ENTITE>/<ENTITE>_MASTER.xlsx` et les classeurs de factures. | `MASTER_SCHEMA.md` |
| **Recette** | Essai complet d'un lot ou d'un document contre des critères écrits à l'avance. | `ROADMAP.md`, `RECETTE_DOC.md` |
| **Registre des factures** | Onglet `Factures` du MASTER (prévu) : une ligne par facture, avec échéance, date d'envoi, date de paiement. | ADR031, F-25 |
| **Règles_facturation** | Onglet du MASTER (prévu) : statut de l'école dans les lots, rythme ou cycle, date limite, délai de paiement, n° de contrat, bon de commande, regroupement. | ADR015, ADR026 |
| **Spec** | Fichier JSON qui décrit une facture pour `scripts/nouvelle_facture.py` (outil actuel). | `examples/spec_exemple.json` |
| **F-xx / T-XXX-NN / ADRxxx / Qx** | Fonctionnalité / test de la banque / décision / question ouverte. | `ROADMAP.md`, `ADR.md` |
