# PROC00 — Protocole de Création de Factures (PCF)

**Statut** : ✅ Acté — v2 (23/11/2026)
**Priorité** : P0 (cœur du système)
**Créé** : 2026-04-26
**Dernière mise à jour** : 2026-11-27

> Le protocole à suivre pour **toute** facture ou tout devis. Ce fichier décrit **comment** on facture ; les raisons des choix sont dans [ADR.md](ADR.md), les critères d'acceptation dans [ROADMAP.md](ROADMAP.md), la mise en page dans [FORMAT.md](FORMAT.md), les données dans [MASTER_SCHEMA.md](MASTER_SCHEMA.md).
>
> Deux parties : **ce que l'on fait aujourd'hui** avec l'outil actuel, puis **la cible du lot 1**. Une étape marquée « prévu » n'existe pas encore dans l'outil.

---

## Principe directeur

Le PCF est un protocole hybride : des étapes automatiques, et des validations humaines qu'on n'automatise jamais. Elles évitent les erreurs irrécupérables : mauvais destinataire, mauvais montant, mauvaise période.

Règles constantes :
- **Aperçu avant toute écriture**, puis validation explicite.
- **Une validation par facture** ; une facture incomplète ne retient pas les autres (ADR031).
- **Aucun envoi par l'outil** : brouillons seulement, l'envoi est programmé par la personne qui facture.
- **Le MASTER fait foi** (ADR005) ; les règles propres à chaque école y vivent (ADR026).
- **Les entités ne se mélangent jamais** (ADR001).
- **Ne jamais inventer** un taux, une coordonnée ou un volume horaire : demander.

---

## Partie 1 — Aujourd'hui : l'outil actuel (v1)

C'est ainsi que sont produites les factures de novembre 2026.

| Étape | Qui | Comment, aujourd'hui |
|---|---|---|
| 1. Périmètre | Humain | Choisir l'entité, l'école, la période et le contrat. Numéro : plus grand numéro de la série + 1 (ADR029). Date d'émission : le 1er du mois suivant (ADR027). |
| 2. Relevé de l'agenda | Script + humain | `facturation agenda lire --mois AAAA-MM` liste les séances du mois (F-10) ; le rattachement aux écoles et le report dans la spec restent manuels jusqu'à F-12 et F-16. |
| 3. Contrôle du MASTER | Script + humain | `python scripts/check_coherence.py --sessions` : état de `Cours` et de `Sessions`, en lecture seule ; il additionne les lignes d'une facture **sans repérer les doublons**. Les heures facturables suivent les règles de l'école. |
| 4. Coordonnées | Script + humain | `facturation profil verifier` signale un émetteur ou un destinataire incomplet ; `nouvelle_facture.py` refuse un destinataire absent ou sans adresse, un NDA manquant et un émetteur sans « EI », dès l'aperçu. |
| 5. Format | Script | `scripts/nouvelle_facture.py` applique le gabarit et le module (`FORMAT.md`). |
| 6. Aperçu | Humain | `--dry-run` affiche émetteur, destinataire, lignes, totaux, sans rien écrire. Validation : oui, non ou corriger. |
| 7. Écriture et contrôle | Script | Écrit un nouvel onglet dans le classeur de l'école (refuse un onglet ou un numéro déjà présent), relit les cellules clés et sort en erreur en cas d'écart. |
| Après | Humain | **Hors outil pour l'instant** : journal `Sessions` et suivi `Cours` dans le MASTER, PDF, brouillon de mail. Les mentions de l'ADR030 sont écrites par l'outil depuis F-24. |

```bash
python scripts/nouvelle_facture.py ma_spec.json --data-root "<racine>" --dry-run   # aperçu
python scripts/nouvelle_facture.py ma_spec.json --data-root "<racine>"             # écriture
```

Format de la spec : [examples/spec_exemple.json](../examples/spec_exemple.json). Plusieurs factures d'un coup (par exemple une par ligne de bon de commande) : un script dédié sur le modèle de [examples/factures_par_bon_de_commande.py](../examples/factures_par_bon_de_commande.py).

---

## Partie 2 — Cible du lot 1 (en partie outillée)

### Ce que l'outil fait déjà (27/11/2026)

Les sorties s'écrivent sous `<racine>/<ENTITE>/_sorties/`, hors dépôt. Aucune de ces commandes n'envoie quoi que ce soit, et seules les actions de `registre` écrivent dans le MASTER.

| Étape | Commande | Ce qu'elle produit | Où |
|---|---|---|---|
| A. Préparation, lancée à la main | `facturation preparer [--mois AAAA-MM] [--ecoles ABC,ECO] [--fichier AGENDA.ics]` | Un brouillon par école du lot, ou par contrat : séances tenues et à venir, heures, taux, total, numéro probable, blocages, avertissements, différences depuis la préparation précédente ; la section « À FAIRE » ; la file à classer. Code 3 tant qu'il reste une action « À FAIRE ». | `_sorties/<AAAA-MM>/` : `preparation.json` (la précédente rangée dans `historique/`), `apercu.txt`, `apercus/<identifiant>.txt`, `a_classer.txt`, `journal.jsonl` |
| A. Relecture de l'aperçu | `facturation apercu [--mois AAAA-MM] [--facture <identifiant ou F<n>>] [--ouvrir]` | L'aperçu global, ou celui d'une facture, sans relire l'agenda ni rien écrire ; `--ouvrir` ouvre le dossier du mois. Aperçus en texte seulement 💡 [proposé — à valider]. | Lit `_sorties/<AAAA-MM>/` |
| D. Audit | `facturation audit` avec `--mois AAAA-MM`, ou `--du JJ/MM/AAAA --au JJ/MM/AAAA`, et `[--ecoles ABC,ECO] [--fichier AGENDA.ics]` | Écarts entre l'agenda et `Sessions` : séance tenue non facturée, ligne absente de l'agenda, doublon, écart d'heures, créneau tenu non reconnu, écart de volume (école à bon de commande), même jour à d'autres heures 💡 [proposé — à valider]. Période par défaut : du 1er du mois M−3 à hier 💡 [proposé — à valider]. Code 3 s'il y a un écart, ou une ligne de `Sessions` illisible ou sans date 💡 [proposé — à valider]. | `_sorties/audits/audit_<du>_<au>_<AAAAMMJJ-HHMMSS>.txt` et `.csv` (Excel) ; rien dans le MASTER |
| D. Suivi : registre | `facturation registre [--ecole X] [--mois AAAA-MM] [--echues] [--trous]` | Tableau des factures, échues en tête, avec leurs anomalies ; `--trous` : numéros manquants sans raison notée. Code 3 s'il y a une facture échue ou une anomalie. | Lit l'onglet `Factures` du MASTER |
| D. Suivi : saisies | `facturation registre initialiser [--reprendre-historique]`, `registre envoyee N° --le DATE`, `registre payee N° --le DATE`, `registre note N° "…"`, `registre manquant N° --raison "…"` | Montre l'avant et l'après, puis écrit sur « oui » (ou `--oui`). `--remplacer` pour changer une date ou une note déjà saisie. | Onglet `Factures` du MASTER ; sauvegarde dans `sauvegardes/` ; événement dans `_sorties/journal.jsonl` |

La préparation n'écrit que sous `_sorties/<AAAA-MM>/`, sous le verrou de l'entité : une préparation qui échoue (agenda indisponible, MASTER illisible) laisse la précédente intacte. Une facture bloquée ne retient pas les autres (ADR031). Une séance déjà présente dans `Sessions` est écartée ; une séance qui recoupe, à d'autres heures, une ligne de `Sessions` du même jour est seulement signalée (« possible doublon ») 💡 [proposé — à valider]. Une école au rythme « fin de module » n'est pas préparée : elle est listée « à préparer à la main » 💡 [proposé — à valider].

**Reste manuel, ou à venir :**

- **Planification et rappels (F-20)** : `facturation preparer` se lance à la main ; ni tâche planifiée ni notification.
- **Étape B (F-16b, F-17)** : les décisions d'Adrian (classer, exclure, inclure, bon de commande) se lisent déjà dans le journal du mois, mais aucune commande ne les écrit encore. L'aperçu cite `facturation classer`, `facturation corriger` et `facturation valider`, qui n'existent pas encore : en attendant, une règle se garde à la main dans `Correspondances`, puis on relance `facturation preparer` ; la validation reste la relecture de l'aperçu. Une école à bon de commande reste bloquée tant que sa référence ne peut pas être saisie : sa facture se fait comme en partie 1.
- **Étape C (F-17, F-17b)** : l'émission n'est pas outillée. La facture s'écrit avec `scripts/nouvelle_facture.py` (partie 1) ; le journal `Sessions` et le suivi `Cours` se tiennent à la main ; le PDF se fait à part. Les factures faites à la main pendant la bascule sont reportées dans `Sessions` et dans `Factures` 💡 [proposé — à valider] : `facturation registre initialiser --reprendre-historique` ajoute au registre les N° de `Sessions` qui y manquent, avec l'école, la période et le total ; la date d'émission et l'échéance restent vides, à compléter dans Excel, car l'outil ne les invente jamais.
- **Séance modifiée après émission (F-23)** : pas encore de signalement automatique ; l'aperçu cite `facturation recontroler`, qui n'existe pas encore. L'audit montre les écarts entre l'agenda et `Sessions`.

### Le calendrier du mois (ADR027)

| Quand | Ce qui se passe |
|---|---|
| Le 20 | Préparation des factures du mois entier, séances tenues et à venir ; notification Windows de demande de validation. |
| Le 25 | Nouvelle notification de validation. |
| Du 28 au dernier jour | Une notification par jour tant qu'une facture reste à valider. |
| Le 1er du mois suivant | Date d'émission des factures. Adrian programme l'envoi lui-même. |
| Avant sa date limite | Une école à cycle propre est préparée selon sa règle (`Règles_facturation`). |

Si le PC est éteint, le lancement a lieu au démarrage de session suivant.

### Pré-requis
- Profil valide (`facturation profil verifier` sans erreur), adresse iCal enregistrée (`facturation profil agenda`).
- MASTER v2 : `Coordonnées`, `Tarifs` avec référence de contrat, `Règles_facturation`, `Correspondances`, `Sessions`, `Cours`, `Factures` (`MASTER_SCHEMA.md`).
- LibreOffice installé pour le PDF (ADR021).

### Étape A — Préparation (auto, F-10 à F-16, F-20)

Outillée par `facturation preparer` et `facturation apercu` (F-16a), sauf le lancement planifié (F-20).

1. Lire l'agenda des cours par l'adresse iCal (ADR019) : récurrences, exceptions, changements d'heure.
2. Classer chaque créneau (ADR016, ADR028) : code d'événement, puis `Correspondances` ; sinon **file à classer**, jamais facturée par défaut.
3. Calculer les heures facturables : heures de l'agenda, ou coefficient, ou forfait, selon `Correspondances` (F-13).
4. Appliquer les règles de l'école (`Règles_facturation`, F-14) : rythme ou cycle, une facture par contrat, bon de commande, regroupement.
5. Produire un **aperçu par facture** : séances tenues et à venir, heures, taux, total par école, informations manquantes, mentions de l'ADR030, différences depuis la préparation précédente.

### Étape B — Validation (humain)

Pas encore outillée (F-16b pour le classement et les corrections, F-17 pour la validation).
- Une validation par facture : **oui**, **non** ou **corriger**.
- Pour une séance de la file à classer, Adrian la classe et choisit si la correction devient une règle (ADR028).
- Une facture à laquelle il manque une donnée obligatoire est bloquée ; les autres continuent.

### Étape C — Émission (auto, F-15, F-17, F-24)

Pas encore outillée : la facture s'écrit comme en partie 1. Le registre sait déjà recevoir la ligne d'une facture émise (F-25).

Pour chaque facture validée, dans l'ordre :
1. Attribuer le numéro (ADR029).
2. Écrire l'onglet du classeur de l'école (gabarit + module, mentions de l'ADR030).
3. Écrire le journal `Sessions` (règle de l'ADR003 amendé, voir plus bas).
4. Mettre à jour `Cours` (heures facturées, heures à facturer, n° de facture lié, statut) et le registre `Factures`.
5. Exporter le PDF par LibreOffice, à côté du classeur.
6. Contrôle après écriture (voir plus bas).

L'état de chaque étape est consigné dans un journal d'exécution (JSON Lines, ADR013). Après un arrêt, l'outil montre où il en est et ne refait que les étapes restantes (ADR031). Le brouillon de mail vient au lot 2 (F-32).

### Étape D — Après l'émission (F-19, F-23, F-25)
- **Audit** agenda ↔ facturé : séances tenues non facturées, écarts. Outillé : `facturation audit` (F-19).
- **Séance modifiée après émission** : signalement avec la facture et la séance concernées ; correction ensuite par un avoir, jamais en modifiant la facture émise 💡 [proposé — à valider] (confirmation de l'expert-comptable attendue). Pas encore outillé (F-23).
- **Suivi** : « envoyée » prouvé par le mail envoyé depuis Gmail ; « payée » : Adrian reporte la date de paiement dans le MASTER (ADR031). Outillé : `facturation registre envoyee` et `registre payee` (F-25), ou la saisie directe dans Excel. Le statut « abandonnée » ne s'écrira que sur une facture restée « en émission », sans date d'envoi ni de paiement (abandon prévu avec F-17), et son numéro n'est jamais réutilisé 💡 [proposé — à valider] ; un numéro manquant est noté avec sa raison par `facturation registre manquant` (ADR029).

---

## Le journal `Sessions` (ADR003 amendé)

Une ligne par séance facturée :

```
Date_log | École | Période_facture | N°_Facture | Date_séance | Début | Fin | Durée_h | Promotion | Module/UE | Taux | Total HT | Summary calendrier verbatim
```

Ajout seul. Seule exception : une facture régénérée remplace ses propres lignes, repérées par le `N°_Facture`, au lieu d'en ajouter de nouvelles. Aucune autre ligne n'est modifiée ni supprimée.

---

## Le contrôle après écriture

Cellules relues par `post_mortem` et comparées à l'aperçu validé (`verifier_facture`) : titre, nom de l'émetteur, nom du destinataire, n° de facture, SIRET, école, premier libellé, formule du total, IBAN, BIC.

- `destinataire_nom`, `siret` et `iban` ne peuvent venir que des appels `apply_*` : une valeur vide signale un appel oublié, donc une facture inexploitable.
- Depuis F-24 : le contrôle porte aussi sur l'échéance, le NDA et la ligne de contrat.

Si une cellule diffère : la facture n'est pas finalisée, l'outil affiche l'écart (« attendu … trouvé … ») et rend la main.

---

## Critères de fin (Definition of Done) d'une facture

| Critère | Aujourd'hui | Cible du lot 1 |
|---|---|---|
| Aperçu validé par la personne qui facture | Oui | Oui |
| Onglet écrit dans le bon classeur, contrôle après écriture sans écart | Oui (outil) | Oui (outil) |
| Mentions de l'ADR030 présentes | Oui (outil, F-24) | Oui (outil) |
| Journal `Sessions` et suivi `Cours` à jour | À la main | Outil |
| Registre `Factures` (échéance, envoi, paiement) | `facturation registre` (F-25) : ligne reprise de `Sessions`, dates saisies par Adrian | Outil, dates d'envoi et de paiement saisies par Adrian |
| PDF généré | À part | Outil (LibreOffice) |

---

## Tests et validation

- La banque de tests (`python -m pytest`) couvre le format, le profil, les scripts et les garde-fous, sur des données fictives. Le défaut du gabarit suivi par T-FMT-07 (formats perdus, mentions en page 2) est corrigé depuis F-18.
- Prévu (F-18) : rendus PDF de 1, 7, 17 et 40 séances comparés à des images de référence.
- Prévu (F-17) : rejouer deux fois la même émission ne duplique rien dans `Sessions`.
- Déjà couvert : une école sans destinataire ou sans adresse bloque l'écriture.

---

## Notes

- **Pourquoi des validations humaines** : elles protègent contre les erreurs irrécupérables ; les étapes automatiques ne sont jamais sautées.
- **Pourquoi un journal `Sessions`** : piste d'audit en cas de litige ou de contrôle fiscal (ADR003).
- **Pourquoi `nouvelle_facture.py` ne couvre pas tout** : il applique le format, l'aperçu et l'écriture d'une facture ; le relevé de l'agenda, le journal, le suivi et le PDF restent hors de l'outil jusqu'au lot 1.
