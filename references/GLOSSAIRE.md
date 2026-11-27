# GLOSSAIRE — Vocabulaire du projet

**Dernière mise à jour** : 2026-11-27

> Les termes employés dans le dépôt, dans le code et dans les échanges. Exemples fictifs uniquement : les noms réels des écoles et leurs codes vivent dans le MASTER (onglet `Écoles`).

| Terme | Sens | Où en savoir plus |
|---|---|---|
| **À FAIRE** | Section en tête de l'aperçu global d'une préparation : ce qui attend une action d'Adrian (valider, débloquer, classer, revalider, préparer à la main). | F-16a, `PROC00_pcf.md` |
| **Agenda des cours** | L'agenda Google où figurent tous les cours de la personne qui facture ; l'outil le lit par son adresse iCal secrète. | ADR019 |
| **Aperçu** | Ce que l'outil montre avant toute écriture : séances, heures, taux, totaux, mentions, informations manquantes. Rien n'est écrit sans validation. Depuis F-16a, la préparation écrit un aperçu global (`apercu.txt`) et un aperçu par facture (`apercus/<identifiant>.txt`), en texte 💡 [proposé — à valider]. | `PROC00_pcf.md` |
| **Audit** | Comparaison entre l'agenda et le journal `Sessions` (`facturation audit`) : séances tenues non facturées, lignes sans séance, doublons, écarts d'heures. Les séances modifiées après émission relèvent de F-23. | F-19, F-23 |
| **Avoir** | Document qui corrige une facture émise (par exemple une séance annulée après coup). Une facture émise n'est jamais modifiée. | ADR024, F-23 |
| **Blocage** | Donnée obligatoire manquante ou incohérente qui empêche d'émettre une facture préparée (destinataire incomplet, taux introuvable, NDA absent, bon de commande non saisi…) ; la facture reste visible et les autres continuent. | ADR030, ADR031 |
| **Bon de commande (BC, PO)** | Document émis par certaines écoles avant la prestation ; son numéro figure alors sur la facture, sur sa propre ligne. | ADR030 |
| **Brouillon** | Facture préparée par l'outil, pas encore émise : école (ou contrat), période, lignes, totaux, blocages. Identifiant `ABC-2026-12`, ou `ABC-2026-12-<contrat>` quand l'école facture par contrat (référence réduite à ses lettres et chiffres) ; pas de numéro de facture avant l'émission. | F-16a |
| **Code d'événement** | Code écrit dans la description d'un événement d'agenda, format proposé `#ÉCOLE-PROMO-MODULE` (ex. `#ABC-M1-STATS`) ; il fait foi sur l'intitulé. | ADR016 |
| **Code école** | Identifiant court d'une école dans une entité (ex. `ABC`) ; 3 lettres pour toute nouvelle école. | ADR029 |
| **Contrôle après écriture** (post-mortem) | Relecture des cellules clés d'une facture écrite pour vérifier qu'elles correspondent à l'aperçu validé. | `PROC00_pcf.md` |
| **Correspondances** | Onglet du MASTER qui relie un code ou un intitulé d'événement à une école, un module, une clé de tarif et un mode de calcul des heures. | ADR016, `MASTER_SCHEMA.md` |
| **Créneau, séance** | Un événement de l'agenda des cours ; une fois reconnu et rattaché à une école, c'est une séance à facturer. | ADR028 |
| **Décision** | Choix ponctuel d'Adrian sur un mois préparé (classer, exclure, inclure, bon de commande, annuler), consigné dans le journal du mois ; il ne vaut que pour ce mois, à la différence d'une règle gardée dans `Correspondances`. | ADR028, F-16a, F-16b |
| **Échéance** | Date à laquelle le paiement est dû, écrite sur la facture : délai du contrat ou de l'école, sinon date d'émission plus 30 jours. | ADR030 |
| **EI** | Entrepreneur individuel : la mention suit le nom de l'émetteur sur la facture. | ADR030 |
| **Empreinte** | Signature (sha256) de tout ce qui sera écrit sur une facture préparée ; une validation porte sur une empreinte, et un contenu qui change demande de revalider. | F-16a, F-17 |
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
| **Numéro court** | `F1`, `F2`… pour les factures, `S1`, `S2`… pour les séances d'une préparation, pour les désigner dans une commande (`facturation apercu --facture F1`). | F-16a |
| **Numéro de facture** | `<préfixe d'entité>_<code école><NNN>`, ex. `R_ABC008`. | ADR002, ADR029 |
| **Numéro probable** | Numéro que recevrait une facture préparée si elle était émise maintenant, affiché à titre indicatif dans l'aperçu ; le numéro n'est attribué qu'à l'émission. | ADR029, F-16a |
| **PAID** | Préfixe ajouté à la main au nom de l'onglet d'une facture réglée ; usage de lecture, sans valeur de preuve. | ADR031 |
| **PCF** | Protocole de création de factures : les étapes et les validations humaines. | `PROC00_pcf.md` |
| **Pièce d'émission** | Fichier écrit par l'émission (`_sorties/emissions/<N°>.json`) qui fige le brouillon validé et son numéro ; un brouillon qui a sa pièce n'est plus jamais régénéré par la préparation. | F-17 (prévu), F-16a |
| **Porte** | La commande qui doit être verte avant toute fusion : `python -m pytest`. | `AGENTS.md` §5 et §7 |
| **Préparation** | Construction, par `facturation preparer`, des brouillons d'un mois pour les écoles du lot, avec leurs aperçus, sous `_sorties/<AAAA-MM>/` ; rien n'est écrit dans le MASTER. | F-16a, `PROC00_pcf.md` |
| **Préparation du 20** | Construction mensuelle des factures le 20, suivie de rappels jusqu'à validation. | ADR027 |
| **Profil** | Fichier TOML de la personne qui facture, hors du dépôt (entité, agenda, racine de données, lancement, PDF). | `README.md` |
| **Racine de données** | Dossier hors dépôt qui contient `<ENTITE>/<ENTITE>_MASTER.xlsx` et les classeurs de factures. | `MASTER_SCHEMA.md` |
| **Recette** | Essai complet d'un lot ou d'un document contre des critères écrits à l'avance. | `ROADMAP.md`, `RECETTE_DOC.md` |
| **Registre des factures** | Onglet `Factures` du MASTER : une ligne par facture (et par numéro manquant), avec échéance, date d'envoi, date de paiement et statut ; tenu par `facturation registre`. | ADR031, F-25, `MASTER_SCHEMA.md` |
| **Règles_facturation** | Onglet du MASTER, une ligne par contrat : statut dans les lots, rythme ou cycle, date limite, délai de paiement, n° de contrat, bon de commande, regroupement. | ADR015, ADR026, `MASTER_SCHEMA.md` |
| **Séance tenue, à venir** | Une séance est tenue quand elle est terminée au moment de la préparation ; les séances à venir du mois sont préparées aussi, puis recontrôlées avant l'envoi. | ADR027, F-23 |
| **Sorties** | Dossier `<racine>/<ENTITE>/_sorties/`, hors dépôt, où l'outil écrit ce qui ne va pas dans le MASTER : préparations et aperçus du mois, rapports d'audit, journaux. | `PROC00_pcf.md` |
| **Spec** | Fichier JSON qui décrit une facture pour `scripts/nouvelle_facture.py` (outil actuel). | `examples/spec_exemple.json` |
| **Verrou** | Fichier `<ENTITE>_MASTER.verrou` posé le temps d'une écriture : un seul écrivain à la fois par entité. | ADR032 |
| **F-xx / T-XXX-NN / ADRxxx / Qx** | Fonctionnalité / test de la banque / décision / question ouverte. | `ROADMAP.md`, `ADR.md` |
