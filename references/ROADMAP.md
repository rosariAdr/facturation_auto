# ROADMAP — Lots, critères de sortie et fonctionnalités

**Dernière mise à jour** : 2026-11-23

> Ce fichier fait foi pour le découpage en lots, les critères de sortie et les critères d'acceptation des fonctionnalités (ADR025). Le **statut** de chaque fonctionnalité (à faire, en cours, bloqué, fait) et les derniers résultats de tests vivent dans Notion (base « Fonctionnalités », « Banque de tests »).
>
> Aucune école n'est nommée ici : la liste des écoles d'un lot est une donnée du MASTER (onglet `Règles_facturation`, ADR026). « L'école pilote » désigne l'école choisie pour le premier parcours complet.

---

## Où en est le projet (23/11/2026)

| Élément | État |
|---|---|
| Lot 0 — socle | Livré le 29/10/2026 : paquet, profil, vérification de l'installation, tests, CI. Correctifs du 21/11/2026 fusionnés sur `dev` (voir `CHANGELOG.md`). |
| Lot 1 — moteur | Pas commencé côté code : aucune lecture d'agenda, aucune préparation automatique, aucun export PDF par l'outil. |
| Factures de novembre 2026 | Faites à la main avec l'outil actuel (spec JSON, `scripts/nouvelle_facture.py`), comme en octobre, avec les mentions de l'ADR030. |
| Aperçus pour toutes les écoles du lot 1 | Dès que la lecture de l'agenda (F-10) est prête, sans date fixe. |

Ce qui est disponible aujourd'hui, et ce qui ne l'est pas : `README.md`, section « Ce que l'outil fait aujourd'hui ».

---

## Lots

| Lot | Objectif | Critère de sortie | Jalon |
|---|---|---|---|
| **0 — Socle** | Un dépôt propre, configurable et testé | Dépôt en place, `pytest` vert en CI, documentation à jour | Atteint le 29/10/2026 |
| **1 — Moteur** | Agenda → aperçus → factures pour les écoles marquées « Lot 1 » dans le MASTER ; audit ; préparation planifiée | 1. Parcours complet sur l'école pilote : lecture de l'agenda, aperçu, validation, classeur XLSX, MASTER à jour, PDF. 2. Aperçus pour toutes les écoles du lot 1, les cas particuliers d'abord traités à la main puis intégrés. 3. Mentions de l'ADR030 sur toutes les factures. La recette septembre–octobre (F-22) ne bloque pas la sortie. | Aperçus globaux dès que F-10 est prête ; pas de date fixe |
| **2 — Appli** | Tout piloter depuis un .exe, sans ligne de commande | **Condition d'entrée** : phase de design terminée (ADR032). **Sortie** : parcours Playwright au vert ; un mois entier validé depuis l'appli ; usage rapide ; comparaison des factures pour valider vite ; brouillons de mail avec PDF (quand Gmail est accepté) ; tableau de bord de suivi ; installation sur deux postes (celui d'Adrian et celui de Parv). | À fixer après le lot 1 |
| **3 — Parv** | Parv facture seule ses écoles (entité KAUR) | Premier mois préparé ensemble ; le mois suivant, Parv émet un mois complet sans aide | À fixer |
| **4 — Plus tard** | Facturation électronique, pont avec un service tiers, ouverture | Réception testée sur Tiime ; émission électronique conforme avant l'échéance de septembre 2027 ; ouverture à d'autres formateurs après un cycle réussi pour les deux entités, avec un testeur de plus | Septembre 2027 pour l'émission électronique |

### Actions de la première version de l'appli (lot 2)
Choisir le mois et les écoles ; relire les aperçus et les différences ; valider et suivre l'émission ; ouvrir les classeurs, les PDF et les résultats du contrôle.

### Phase de design (entrée du lot 2)
Parcours écran par écran, maquettes cliquables, charte et couleurs, essai avec Adrian sur un vrai mois. Outil libre, sans donnée réelle ; livrables déposés dans le dépôt (ADR032).

---

## Fonctionnalités et critères d'acceptation

Priorités : **Indispensable**, **Important**, **Souhaitable**. « Mis à jour le 23/11 » signale un critère modifié par le cadrage du 23/11/2026.

### Lot 0 — Socle

| Réf | Fonctionnalité | Priorité | Critère d'acceptation |
|---|---|---|---|
| F-01 | Dépôt GitHub privé et conventions | Indispensable | Dépôt privé, branches `main` et `dev`, clonable sur le PC d'Adrian, aucune donnée réelle dans l'historique. |
| F-02 | Configuration par utilisateur | Indispensable | Deux profils fictifs différents produisent chacun leurs sorties dans leur propre dossier ; un profil incomplet donne un message qui nomme le champ manquant. |
| F-03 | Banque de tests et intégration continue | Indispensable | `pytest` rejoue toute la banque en une commande ; la CI est verte sur `dev` ; chaque test figure dans la banque de tests. |
| F-04 | Documentation autoportante (mis à jour le 23/11) | Important | Une personne ou un agent qui n'a pas suivi le projet reprend le travail en lisant le dépôt seul, sans reposer de question déjà tranchée (`references/RECETTE_DOC.md`). |

### Lot 1 — Moteur

| Réf | Fonctionnalité | Priorité | Critère d'acceptation |
|---|---|---|---|
| F-10 | Lecture de l'agenda | Indispensable | Sur un .ics fictif avec récurrences, exceptions (séances déplacées ou annulées) et changement d'heure, la liste des séances (date, début, fin, durée, intitulé, description) est exacte. |
| F-11 | Lecture et écriture sûres du MASTER | Indispensable | Un MASTER mal formé produit une erreur qui nomme l'onglet et la colonne ; aucune écriture sans sauvegarde préalable. |
| F-12 | Classement des séances (mis à jour le 23/11) | Indispensable | 100 % des séances de la période classées ou mises dans la file à classer, aucune facturée par erreur ; quand Adrian classe ou corrige une séance, il choisit si la règle est gardée ; une règle gardée s'applique au lancement suivant (ADR028). |
| F-13 | Heures facturables | Indispensable | Un forfait de jury : 7 h d'agenda donnent 7,5 h facturées ; une journée 9 h–17 h avec un coefficient 7/8 donne 7 h ; sans règle, les heures facturées sont celles de l'agenda. |
| F-14 | Règles de facturation par école (mis à jour le 23/11) | Indispensable | Une école sans rythme renseigné est facturée chaque mois ; une école à bons de commande produit une facture par ligne, colonne Jour masquée ; une école sans bon de commande reçu est préparée puis bloquée ; une école en fin de module n'est facturée qu'à la fin ; une école à cycle propre est préparée avant sa date limite ; une facture par contrat quand la règle le demande. |
| F-15 | Numérotation automatique (mis à jour le 23/11) | Indispensable | Après `R_ABC007`, la facture suivante de l'école est `R_ABC008` ; un brouillon abandonné ne consomme pas de numéro ; aucun doublon ; un numéro manquant n'est jamais réutilisé et sa raison est notée (ADR029). |
| F-16 | Préparation multi-écoles des brouillons (mis à jour le 23/11) | Indispensable | Un mois préparé le 20 : un brouillon par école du lot 1 (ou par contrat) couvrant tout le mois ; séances tenues et à venir distinguées ; les totaux de l'aperçu égalent la somme heures × taux de chaque ligne ; une relance après modification de l'agenda remplace les brouillons sans doublon et montre les différences. |
| F-17 | Émission après validation (mis à jour le 23/11) | Indispensable | Chaque facture est validée seule ; une facture incomplète est bloquée sans retenir les autres ; rejouer l'émission ne duplique rien ; après un arrêt, l'état est affiché et seules les étapes restantes sont refaites ; le contrôle après écriture (émetteur, destinataire, SIRET, IBAN, total, mentions) est vert pour chaque facture (ADR031). |
| F-18 | Fiabilisation du gabarit (mis à jour le 23/11) | Indispensable | T-FMT-07 au vert ; rendu PDF de 1, 7, 17 et 40 séances conforme aux images de référence ; la facture pilote tient sur une page ; mentions toujours présentes sur la dernière page ; une dénomination ou une mention longue n'est pas coupée. |
| F-19 | Audit agenda ↔ facturé | Indispensable | Sur une période de référence, le rapport liste les séances tenues non facturées, et aucun écart pour l'école pilote. |
| F-20 | Préparation planifiée et rappels (mis à jour le 23/11) | Indispensable | Préparation le 20 avec notification Windows ; rappel le 25 ; rappel quotidien du 28 à la fin du mois tant qu'une facture reste à valider ; tâche installée en une commande, déclenchée au premier démarrage de session si le PC était éteint, sans session Claude, avec un journal d'exécution même en cas d'échec (ADR027). |
| F-21 | Compléter le MASTER pour toutes les écoles | Indispensable | Chaque école du lot 1 a des coordonnées complètes (dénomination juridique), un tarif avec sa référence de contrat, une règle et au moins une correspondance. |
| F-22 | Recette septembre–octobre 2026 (mis à jour le 23/11) | Important | Non bloquante : l'école pilote est identique à ses trois factures de septembre–octobre (lignes, heures, taux, totaux) ; chaque autre école : facture juste ou blocage motivé. Pertinence à réévaluer après les premiers aperçus. |
| F-23 | Séances modifiées après émission et avoirs (mis à jour le 23/11) | Important | D'abord : une séance supprimée ou modifiée après émission est signalée avec la facture et la séance concernées. Ensuite : elle donne un avoir correct ; la facture d'origine reste intacte. |
| F-24 | Mentions légales et échéance (nouveau, 23/11) | Indispensable | Sur une facture fictive : émetteur suivi de « EI » ; client sous sa dénomination juridique ; n° de contrat et n° de bon de commande sur des lignes distinctes du n° de facture ; échéance datée ; « Total HT », mention 293 B, « Net à payer » ; pénalités exigibles dès le lendemain de l'échéance ; indemnité de 40 € ; « Pas d'escompte » ; n° de déclaration d'activité ; ni « Montant non assujetti à la TVA » ni « À payer TTC » ; une valeur manquante bloque l'émission (ADR030). |
| F-25 | Suivi d'envoi et de paiement (nouveau, 23/11) | Important | Le registre des factures (onglet `Factures`) porte, pour chaque facture, son échéance, sa date d'envoi et sa date de paiement ; une facture échue et non payée est signalée (ADR031). |

### Lot 2 — Appli

| Réf | Fonctionnalité | Priorité | Critère d'acceptation |
|---|---|---|---|
| F-30 | Application locale .exe (mis à jour le 23/11) | Important | Adrian prépare, valide et émet un mois entier sans ligne de commande : choix du mois et des écoles, relecture des aperçus et des différences, validation et suivi de l'émission, accès aux classeurs, PDF et résultats du contrôle. |
| F-31 | Classement assisté dans l'appli (mis à jour le 23/11) | Important | Une séance classée dans l'appli, et gardée comme règle par Adrian, est reconnue au lancement suivant. |
| F-32 | Brouillons de mail (mis à jour le 23/11) | Important | Après l'accord d'Adrian, un brouillon Gmail par facture, avec le PDF en pièce jointe ; aucun n'est envoyé. Condition : la connexion Gmail est acceptée (pas encore au 23/11/2026). |
| F-33 | Tests d'utilisation Playwright | Important | `pytest -m e2e` vert avant chaque version de l'appli. |
| F-34 | Installeur et guide d'installation (mis à jour le 23/11) | Important | Une installation sur un PC vierge aboutit en suivant le guide seul ; première recette sur deux postes. |
| F-35 | Tableau de bord de suivi (nouveau, 23/11) | Important | Vue des factures émises, envoyées, payées et échues, par école et par mois, lue dans le registre des factures. |

### Lot 3 — Parv

| Réf | Fonctionnalité | Priorité | Critère d'acceptation |
|---|---|---|---|
| F-40 | Gabarit KAUR | Indispensable | Tests de format KAUR au vert ; Parv valide le rendu. |
| F-41 | Profil et données KAUR | Indispensable | Aucune donnée ROSARI visible ni écrite depuis le profil KAUR (ADR001). |
| F-42 | Prise en main par Parv (mis à jour le 23/11) | Important | Premier mois préparé ensemble ; le mois suivant, Parv émet un mois complet sans aide. |

### Lot 4 — Plus tard

| Réf | Fonctionnalité | Priorité | Critère d'acceptation |
|---|---|---|---|
| F-50 | Ouverture à d'autres formateurs (mis à jour le 23/11) | Souhaitable | Condition d'entrée : un cycle réussi pour les deux entités, avec un testeur de plus. Critère : un formateur extérieur installe et facture un mois avec la seule documentation. |
| F-51 | Conformité facturation électronique (mis à jour le 23/11) | Important | La réception est testée sur Tiime ; une facture émise par l'outil est acceptée par la plateforme retenue (ADR033). |
| F-52 | Pont Tiime / Shine | Souhaitable | Réconciliation annuelle sans écart entre l'outil et le service tiers. |

---

## Questions encore ouvertes

| Sujet | Ce qui reste à décider | Qui |
|---|---|---|
| Factures longues | La règle au-delà d'une page, après le pilote | Adrian |
| Facture historique auditée | La rectifier ou non (décision métier distincte, ADR030) | Adrian |
| Mention de franchise | Le libellé après le passage de la TVA au CIBS ; la citation du CGI reste admise jusqu'au 30/06/2028 (ADR030) | Avant le 30/06/2028 |
| Mentions de TVA | Si la franchise cesse ou si l'attestation d'exonération formation est accordée | Adrian, avec un expert-comptable |
| Rendu PDF (Q7) | LibreOffice suffit-il, ou faut-il Excel ? | Après le premier essai complet sur le PC habituel |
| Lot 3 | École et mois de test, facture de référence de Parv (Q10), rapprochement des anciennes factures | À vérifier |
| Lot 4 | Premier parcours d'émission électronique ; plateforme imposée par une école ou non | À vérifier |
