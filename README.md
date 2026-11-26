# facturation_auto

Factures mensuelles d'un formateur indépendant, à partir de son agenda de cours, des tarifs de chaque école et d'un gabarit de facture XLSX. Un module Python force les invariants du gabarit, et un protocole garde les validations humaines (PCF).

> **État au 23/11/2026** : v0.2 en développement (refonte v2). Le lot 0 (socle) est livré ; des correctifs du 21/11 sont sur `dev`. Le lot 1 (lecture de l'agenda, préparation automatique) n'est pas encore commencé dans le code : les factures se font avec `scripts/nouvelle_facture.py`. Lots et critères : [references/ROADMAP.md](references/ROADMAP.md). Décisions : [references/ADR.md](references/ADR.md).

## Ce que l'outil fait aujourd'hui

| Capacité | Statut | Comment |
|---|---|---|
| Créer, situer et vérifier un profil | Disponible | `facturation profil init`, `chemin`, `verifier` |
| Enregistrer l'adresse iCal secrète de l'agenda | Disponible | `facturation profil agenda` (l'adresse est rangée, pas encore lue) |
| Aperçu d'une facture décrite par une spec JSON | Disponible | `python scripts/nouvelle_facture.py <spec> --dry-run` |
| Écrire la facture dans le classeur de l'école, avec contrôle après écriture | Disponible | `python scripts/nouvelle_facture.py <spec>` (refuse d'écraser une facture existante) |
| Contrôler le MASTER (`Cours`, `Sessions`) en lecture seule | Disponible | `python scripts/check_coherence.py` |
| Lire l'agenda automatiquement | Prévu, lot 1 (F-10) | — |
| Classer les séances, calculer les heures facturables | Prévu, lot 1 (F-12, F-13) | — |
| Préparer le 20, rappeler le 25 et du 28 à la fin du mois | Prévu, lot 1 (F-20, ADR027) | Le profil porte encore le 25 par défaut |
| Tenir le journal `Sessions`, le suivi `Cours` et le registre `Factures` | Prévu, lot 1 (F-17, F-25) | À la main aujourd'hui |
| Exporter le PDF | Prévu, lot 1 (ADR021) | LibreOffice est seulement détecté par `profil verifier` |
| Mentions légales de l'ADR030 (échéance, contrat, NDA, pénalités, 40 €…) | Disponible (F-24) | Écrites par `nouvelle_facture.py` ; refus si le NDA ou « EI » manque |
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
│   ├── config.py                 profil utilisateur (TOML hors dépôt)
│   ├── verification.py           contrôles de l'installation
│   ├── cli.py                    commande `facturation`
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

`verifier` contrôle sans rien écrire : dossier de données et MASTER, onglets, coordonnées de l'émetteur, destinataires incomplets, gabarit, adresse iCal secrète (jamais affichée), LibreOffice. Code de retour 1 en cas d'erreur bloquante. Exemple commenté : [examples/profil_exemple.toml](examples/profil_exemple.toml).

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
python -m pytest                 # toute la banque (la porte)
python -m pytest -m regression   # format et rendu des factures
python -m pytest -m unit         # profil et fonctions pures
```

Tous les tests tournent sur des données fictives créées à la volée. Les défauts connus du gabarit (T-FMT-07, à corriger par F-18) sont marqués `xfail` strict : ils restent visibles dans le rapport et feront échouer la suite le jour où ils seront corrigés, pour rappeler de retirer le marqueur. GitHub Actions rejoue la banque à chaque pull request et à chaque push sur `main`, `dev` et les branches de travail (`feat/**`, `fix/**`, `chore/**`), sous Windows et Ubuntu, avec Python 3.11 et 3.12. Le rendu PDF est sauté si LibreOffice est absent.

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
