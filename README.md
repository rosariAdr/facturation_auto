# facturation_auto

Factures mensuelles d'un formateur indépendant, à partir de son agenda de cours, des tarifs de chaque école et d'un gabarit de facture XLSX. Un module Python force les invariants du gabarit, et un protocole garde les validations humaines (PCF).

> **État au 27/11/2026** : v0.2 en développement (refonte v2). Le lot 0 (socle) est livré. Le lot 1 est en cours : l'outil lit l'agenda, prépare les brouillons et les aperçus du mois, audite l'agenda contre le facturé et tient le registre des factures ; la validation et l'émission ne sont pas encore outillées, les factures s'écrivent toujours avec `scripts/nouvelle_facture.py`. Lots et critères : [references/ROADMAP.md](references/ROADMAP.md). Décisions : [references/ADR.md](references/ADR.md).

## Ce que l'outil fait aujourd'hui

| Capacité | Statut | Comment |
|---|---|---|
| Créer, situer et vérifier un profil | Disponible | `facturation profil init`, `chemin`, `verifier` |
| Enregistrer l'adresse iCal secrète de l'agenda | Disponible | `facturation profil agenda` ; l'adresse est lue par `agenda lire`, `preparer` et `audit`, jamais affichée |
| Aperçu d'une facture décrite par une spec JSON | Disponible | `python scripts/nouvelle_facture.py <spec> --dry-run` ; montants en décimal, spec mal formée refusée dès l'aperçu |
| Écrire la facture dans le classeur de l'école, avec contrôle après écriture | Disponible | `python scripts/nouvelle_facture.py <spec>` (refuse d'écraser une facture existante) |
| Contrôler le MASTER (`Cours`, `Sessions`) en lecture seule | Disponible | `python scripts/check_coherence.py` |
| Lire l'agenda des cours (récurrences, exceptions, changement d'heure) | Disponible (F-10) | `facturation agenda lire --mois AAAA-MM` (ou `--fichier agenda.ics`) ; l'adresse iCal n'est jamais affichée |
| Préparer les brouillons du mois pour les écoles du lot (une facture par école ou par contrat), avec blocages, avertissements et file à classer | Disponible (F-16a) | `facturation preparer [--mois AAAA-MM] [--ecoles ABC,ECO] [--fichier agenda.ics]` ; écrit sous `_sorties/<AAAA-MM>/`, jamais dans le MASTER ; code 3 tant qu'il reste une action « À FAIRE » |
| Relire l'aperçu du mois, ou d'une facture | Disponible (F-16a) | `facturation apercu [--facture F1] [--ouvrir]` : relit la préparation sans relire l'agenda ni rien écrire ; aperçus en texte seulement 💡 [proposé — à valider] |
| Lire le MASTER avec contrôle des onglets et des colonnes, écrire avec sauvegarde et verrou | Disponible (F-11) | Utilisé par `preparer`, `audit` et `registre` ; seul `registre` écrit dans le MASTER, et seulement l'onglet `Factures` |
| Proposer le numéro de facture suivant | Disponible (F-15) | Indicatif dans l'aperçu (« numéro probable ») ; le numéro sera attribué à l'émission (F-17) |
| Facture d'une page jusqu'à 18 séances, sauts de page maîtrisés au-delà, mentions lisibles | Disponible (F-18) | Automatique dans `nouvelle_facture.py` |
| Classer les créneaux de l'agenda (code d'événement, intitulé exact, préfixe ; sinon file à classer) | Disponible (F-12) | Appliqué par `preparer` et `audit` ; un créneau non reconnu va dans `a_classer.txt` et n'est jamais facturé |
| Calculer les heures facturables (agenda, coefficient, forfait) et lire le taux d'une ligne de Tarifs | Disponible (F-13) | Appliqué par `preparer` et `audit` ; le taux se lit aussi en texte (« 60 €/h »), jamais inventé |
| Lire les règles de facturation par école (rythme, cycle, délai de paiement, contrat, bon de commande) | Disponible (F-14) | Appliquées par `preparer` (période, échéance, contrat, blocages) et `audit` ; une ligne par contrat |
| Auditer l'agenda contre le facturé (`Sessions`) | Disponible (F-19) | `facturation audit` avec `--mois AAAA-MM`, ou `--du JJ/MM/AAAA --au JJ/MM/AAAA` (du 1er du mois M−3 à hier par défaut 💡 [proposé — à valider]), `--ecoles` ; rapport `.txt` et `.csv` sous `_sorties/audits/`, rien écrit dans le MASTER |
| Tenir le registre `Factures` : échéances, retards, dates d'envoi et de paiement, numéros manquants | Disponible (F-25) | `facturation registre` (factures échues en tête, anomalies, `--trous`) ; `registre initialiser`, `envoyee`, `payee`, `note`, `manquant` montrent l'avant et l'après, puis écrivent sur « oui » (ou `--oui`), avec sauvegarde ; la ligne d'une facture émise sera créée par l'émission (F-17) |
| Classer un créneau, corriger une préparation (exclure, inclure, bon de commande), garder une règle | Prévu, lot 1 (F-16b) | Les décisions se lisent déjà mais ne s'écrivent pas encore : `facturation classer` et `facturation corriger`, cités par l'aperçu, n'existent pas encore ; une règle se garde à la main dans `Correspondances` |
| Valider chaque facture, puis l'émettre (numéro, onglet, journal `Sessions`, registre, PDF), avec reprise après un arrêt | Prévu, lot 1 (F-17, ADR021) | `facturation valider`, cité par l'aperçu, n'existe pas encore ; aujourd'hui : `nouvelle_facture.py`, puis `Sessions` à la main ; LibreOffice est seulement détecté par `profil verifier` |
| Tenir le suivi `Cours` | Prévu, lot 1 (F-17b) | À la main aujourd'hui |
| Signaler une séance modifiée après émission | Prévu, lot 1 (F-23) | Correction par un avoir, jamais en modifiant la facture émise 💡 [proposé — à valider] |
| Préparer le 20, rappeler le 25 et du 28 à la fin du mois | Prévu, lot 1 (F-20, ADR027) | `facturation preparer` se lance à la main ; le profil porte encore le 25 par défaut |
| Mentions légales de l'ADR030 (échéance, contrat, NDA, pénalités, 40 €…) | Disponible (F-24) | Écrites par `nouvelle_facture.py` ; refus si le NDA ou « EI » manque ; dans la préparation, leur absence bloque le brouillon |
| Appli .exe, brouillons Gmail, tableau de bord de suivi | Prévu, lot 2 | Après une phase de design (ADR032) |
| Formulaire d'onboarding à envoyer aux écoles | Idée | — |

## Où vit quoi (ADR025)

| Quoi | Où |
|---|---|
| Conventions de travail (branches, commits, tests, règles des agents) | [AGENTS.md](AGENTS.md) |
| Décisions et leurs raisons | [references/ADR.md](references/ADR.md) |
| Lots, critères de sortie, critères d'acceptation des fonctionnalités | [references/ROADMAP.md](references/ROADMAP.md) |
| Comment facturer, pas à pas | [references/PROC00_pcf.md](references/PROC00_pcf.md) |
| Données réelles : écoles, taux, règles, coordonnées | Le MASTER de l'entité, dans la racine de données, hors dépôt (ADR026) |
| Secrets : profil et adresse iCal | `%APPDATA%\facturation\`, hors dépôt |
| Avancement : statuts, banque de tests, journal de bord | Notion (second cerveau du projet) |

Le dépôt ne contient **aucune donnée réelle** : ni nom d'école, ni adresse, ni SIRET, ni IBAN, ni montant. Le gabarit est vide de toute coordonnée ; le module les lit dans `MASTER.Coordonnées` au moment de l'exécution.

## Le dépôt

```
facturation_auto/
├── pyproject.toml                paquet `facturation-auto`, commande `facturation`
├── facturation/
│   ├── pcf_format.py             module de format (gabarit, encadrés, séances, contrôles)
│   ├── facture.py                une facture décrite par une spec : aperçu, refus, écriture, contrôle
│   ├── config.py                 profil utilisateur (TOML hors dépôt)
│   ├── verification.py           contrôles de l'installation
│   ├── agenda.py, master.py, sessions.py, classement.py, heures.py, regles.py, numerotation.py
│   │                             lecture de l'agenda et du MASTER, calculs du lot 1
│   ├── preparation.py, decisions.py, audit.py, registre.py
│   │                             préparation du mois, audit, registre des factures
│   ├── fichiers.py, chemins.py, journal.py   écritures sûres, chemins des sorties, journaux
│   ├── cli.py                    commande `facturation`
│   ├── commandes/                une sous-commande par module (profil, agenda, preparer, audit, registre)
│   └── assets/
│       └── PCF_facture_standard.xlsx   gabarit commun, sans données
├── scripts/
│   ├── nouvelle_facture.py       spec JSON → facture
│   ├── check_coherence.py        contrôles du MASTER (lecture seule)
│   └── build_template.py         reconstruction du gabarit (maintenance)
├── tests/                        pytest : unit, integration, regression (données fictives)
├── examples/
│   ├── profil_exemple.toml
│   ├── spec_exemple.json
│   └── factures_par_bon_de_commande.py   plusieurs factures depuis des bons de commande (fictif)
├── references/                   décisions, feuille de route, protocole, format, MASTER, onboarding, glossaire, recette
├── AGENTS.md                     conventions de travail (Claude Code et Codex)
└── SKILL.md                      point d'entrée skill Claude Code
```

## Installation

Python 3.11 ou plus.

```bash
git clone git@github.com:rosariAdr/facturation_auto.git
cd facturation_auto
python -m pip install -e ".[dev]"
```

LibreOffice sera nécessaire pour l'export PDF (lot 1).

## Profil utilisateur

Chaque personne qui facture a son profil : entité, agenda, racine de données, jour et heure de préparation, moteur PDF. Il est rangé hors du dépôt, par défaut dans `%APPDATA%\facturation\profil.toml` sous Windows (`~/.config/facturation/profil.toml` ailleurs, ou le chemin de `FACTURATION_PROFIL`).

```bash
facturation profil init --nom "Prénom Nom" --entite ROSARI --racine "G:/Mon Drive/Factures"
facturation profil agenda      # colle l'adresse iCal secrète, sans l'afficher
facturation profil verifier
```

L'adresse iCal secrète se trouve dans Google Agenda (navigateur) : survoler l'agenda de cours dans la colonne de gauche, cliquer sur les trois points, « Paramètres et partage », puis copier « Adresse secrète au format iCal ». `profil agenda` la range dans `agenda.secret`, à côté du profil, jamais dans le dépôt.

`verifier` contrôle sans rien écrire : dossier de données et MASTER, onglets, coordonnées de l'émetteur, destinataires incomplets, gabarit, adresse iCal secrète (jamais affichée), LibreOffice. Il signale aussi un onglet de suivi que l'outil ignore faute de colonne, et ce qu'un enregistrement du MASTER par l'outil perdrait (contrôle fait sur une copie). Code de retour 1 en cas d'erreur bloquante. Exemple commenté : [examples/profil_exemple.toml](examples/profil_exemple.toml).

Valeurs par défaut du profil aujourd'hui : préparation le 25 du mois à 16 h, factures datées du 1er du mois suivant, agenda lu par adresse iCal, PDF par LibreOffice. Décision du 23/11/2026 (ADR027) : préparation le 20, avec rappels ; le défaut changera avec F-20.

## Racine de données

Elle contient `<ENTITE>/<ENTITE>_MASTER.xlsx` : voir [references/MASTER_SCHEMA.md](references/MASTER_SCHEMA.md). Les scripts la résolvent dans cet ordre : `--data-root` > `--profil` > variable `PCF_DATA_ROOT` > profil par défaut s'il existe > dossier courant.

## Usage aujourd'hui

Aperçu, puis écriture, d'une facture décrite par une spec JSON ([format](examples/spec_exemple.json)) :

```bash
python scripts/nouvelle_facture.py ma_spec.json --dry-run
python scripts/nouvelle_facture.py ma_spec.json
```

Contrôler le MASTER avant d'émettre :

```bash
python scripts/check_coherence.py --entite ROSARI --sessions
```

Préparer le mois, relire les aperçus, auditer et suivre les factures (lot 1, en cours) :

```bash
facturation preparer --mois 2026-12          # brouillons et aperçus sous <racine>/<ENTITE>/_sorties/2026-12/
facturation apercu --mois 2026-12 --facture F1
facturation audit                            # du 1er du mois M−3 à hier 💡 [proposé — à valider]
facturation registre --echues
facturation registre payee R_ABC008 --le 15/01/2027   # montre l'avant et l'après, puis demande « oui »
```

Aucune de ces commandes n'envoie quoi que ce soit. `preparer`, `apercu` et `audit` n'écrivent rien dans le MASTER ; les actions de `registre` n'y écrivent que l'onglet `Factures`, après validation. Le protocole complet, et ce qui reste manuel : [references/PROC00_pcf.md](references/PROC00_pcf.md).

En Python (les codes, montants et références sont fictifs) :

```python
from datetime import date
from openpyxl import Workbook
from facturation import pcf_format as fmt

fmt.set_data_root("/chemin/vers/Factures")
wb = Workbook(); wb.remove(wb.active)
ws = fmt.clone_template(wb, "R_ABC010 Sept 2026", entite="ROSARI")
fmt.apply_emetteur(ws, entite="ROSARI")
fmt.apply_destinataire(ws, dest=fmt.get_destinataire("ABC", entite="ROSARI"), entite="ROSARI")
fmt.apply_metadata(ws, n_facture="R_ABC010", date_emission=date(2026, 10, 1),
                   entite="ROSARI", ref_bon_commande="BC-EXEMPLE-0001")
fmt.place_sessions(ws, "ABC", groups=[
    ("M1", "Statistiques — groupe A", [(date(2026, 9, 15), 3.5, 50)]),
], entite="ROSARI")
fmt.apply_mentions(ws, entite="ROSARI", date_emission=date(2026, 10, 1))  # échéance : +30 jours
fmt.apply_rib(ws, entite="ROSARI")
fmt.apply_print_setup(ws, entite="ROSARI")
wb.save("facture.xlsx")
print(fmt.post_mortem(ws, entite="ROSARI"))
```

L'entité est toujours un argument nommé (`entite=…`) : il n'y a pas d'entité par défaut. `apply_mentions` exige une ligne `Info` / `NDA` dans `MASTER.Coordonnées`.

## Gabarit

Une mise en page unique, `facturation/assets/PCF_facture_standard.xlsx`, pour toutes les entités : seules les informations de l'émetteur changent. Une entité peut avoir son propre gabarit `PCF_facture_<ENTITE>.xlsx`, prioritaire s'il existe. Plan des cellules, mentions, pièges : [references/FORMAT.md](references/FORMAT.md).

## Tests

```bash
python -m pytest                  # toute la banque (la porte)
python -m pytest -m unit          # fonctions pures : profil, agenda, MASTER, règles, clonage du gabarit
python -m pytest -m integration   # scripts et commande facturation
python -m pytest -m regression    # format et rendu des factures
python -m pytest -m recette       # facture complète, vérifiée sur le fichier relu
```

Tous les tests tournent sur des données fictives créées à la volée. Un défaut connu peut être marqué `xfail` strict : il reste visible dans le rapport et fait échouer la suite le jour où il est corrigé, pour rappeler de retirer le marqueur (les défauts T-FMT-07 du gabarit ont été corrigés par F-18). GitHub Actions rejoue la banque à chaque pull request et à chaque push sur `main`, `dev` et les branches de travail (`feat/**`, `fix/**`, `chore/**`), sous Windows et Ubuntu, avec Python 3.11 et 3.12. Le rendu PDF est sauté si LibreOffice est absent.

## Documentation

| Fichier | Contenu |
|---|---|
| [references/ADR.md](references/ADR.md) | Toutes les décisions, ADR001 à 033 |
| [references/ROADMAP.md](references/ROADMAP.md) | Lots, critères de sortie, fonctionnalités et critères d'acceptation |
| [references/PROC00_pcf.md](references/PROC00_pcf.md) | Le protocole : aujourd'hui, puis la cible du lot 1 |
| [references/FORMAT.md](references/FORMAT.md) | Gabarit, mentions, invariants, API du module |
| [references/MASTER_SCHEMA.md](references/MASTER_SCHEMA.md) | Structure du MASTER, onglets existants et prévus |
| [references/ONBOARDING.md](references/ONBOARDING.md) | Ajouter une école |
| [references/GLOSSAIRE.md](references/GLOSSAIRE.md) | Vocabulaire du projet |
| [references/RECETTE_DOC.md](references/RECETTE_DOC.md) | Vérifier que la documentation suffit, sur un exemple fictif |
| [CHANGELOG.md](CHANGELOG.md) | Historique des versions |

## Conventions

Les conventions de travail (branches `main` / `dev` / `feat|fix|chore/*`, Conventional Commits, porte de tests, règles pour les agents) vivent dans [AGENTS.md](AGENTS.md), lu par Claude Code et Codex.
