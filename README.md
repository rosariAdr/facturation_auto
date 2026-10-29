# facturation_auto

Factures mensuelles d'un formateur indépendant, à partir de son agenda de cours,
des tarifs de chaque école et d'un gabarit de facture XLSX. Un module Python force
les invariants du gabarit, et un protocole en 7 étapes garde les validations
humaines (PCF).

> **État :** v0.2 en développement (refonte v2). Le lot 0 livre le socle : paquet
> installable, profil par utilisateur, vérification de l'installation, tests et
> intégration continue. La lecture de l'agenda et la préparation automatique des
> factures arrivent au lot 1. Le pilotage (cadrage, fonctionnalités, décisions,
> banque de tests, journal) vit dans le second cerveau Notion du projet.

## Ce que le dépôt contient, et ce qu'il ne contient pas

Le code et le format, jamais les données. Aucun nom d'école, adresse, SIRET, IBAN
ni montant n'est versionné : le gabarit est vide de toute coordonnée, et le module
les lit au runtime dans `MASTER.Coordonnées`. Les données métier restent dans un
dossier séparé, la *racine de données* (chez l'utilisateur principal, un dossier
Google Drive pour ordinateur).

```
facturation_auto/
├── pyproject.toml                paquet `facturation-auto`, commande `facturation`
├── facturation/
│   ├── pcf_format.py             module de format (API de 10 fonctions)
│   ├── config.py                 profil utilisateur (TOML hors dépôt)
│   ├── verification.py           contrôles de l'installation
│   ├── cli.py                    commande `facturation`
│   └── assets/
│       └── PCF_facture_standard.xlsx   gabarit MODELE commun, sans données
├── scripts/
│   ├── nouvelle_facture.py       spec JSON → facture
│   ├── check_coherence.py        contrôles du MASTER (lecture seule)
│   └── build_template.py         reconstruction du gabarit (maintenance)
├── tests/                        pytest : unit, integration, regression (données fictives)
├── examples/
│   ├── profil_exemple.toml
│   ├── spec_exemple.json
│   └── generate_epita_pos.py     cas réel : 7 factures depuis 2 bons de commande
├── references/                   protocole PCF, format, schéma du MASTER, onboarding, ADR
└── SKILL.md                      point d'entrée skill Claude Code
```

## Installation

Python 3.11 ou plus.

```bash
git clone git@github.com:rosariAdr/facturation_auto.git
cd facturation_auto
python -m pip install -e ".[dev]"
```

Pour l'export PDF (lots suivants), installer aussi LibreOffice.

## Profil utilisateur

Chaque personne qui facture a son profil : entité, agenda, racine de données,
jour de lancement, moteur PDF. Il est rangé hors du dépôt, par défaut dans
`%APPDATA%\facturation\profil.toml` sous Windows (`~/.config/facturation/profil.toml`
ailleurs, ou le chemin de `FACTURATION_PROFIL`).

```bash
facturation profil init --nom "Prénom Nom" --entite ROSARI --racine "G:/Mon Drive/Factures"
facturation profil agenda      # colle l'adresse iCal secrète, sans l'afficher
facturation profil verifier
```

L'adresse iCal secrète se trouve dans Google Agenda (navigateur) : survoler l'agenda
de cours dans la colonne de gauche, cliquer sur les trois points, « Paramètres et
partage », puis copier « Adresse secrète au format iCal ». `profil agenda` la range
dans `agenda.secret`, à côté du profil, jamais dans le dépôt.

`verifier` contrôle sans rien écrire : présence du dossier de données et du MASTER,
onglets, coordonnées de l'émetteur, destinataires incomplets, gabarit, adresse iCal
secrète (jamais affichée), LibreOffice. Code de retour 1 en cas d'erreur bloquante.
Exemple commenté : [examples/profil_exemple.toml](examples/profil_exemple.toml).

Valeurs par défaut issues du cadrage : lancement le 25 du mois à 16 h sur le mois en cours,
factures datées du 1er du mois suivant, agenda lu par adresse iCal, PDF par LibreOffice.

## Racine de données

Elle contient `<ENTITE>/<ENTITE>_MASTER.xlsx` — voir
[references/MASTER_SCHEMA.md](references/MASTER_SCHEMA.md). Les scripts la résolvent
dans cet ordre : `--data-root` > `--profil` > variable `PCF_DATA_ROOT` > profil par
défaut s'il existe > dossier courant.

## Usage (v1, toujours disponible)

Aperçu puis émission d'une facture décrite par une spec JSON :

```bash
python scripts/nouvelle_facture.py examples/spec_exemple.json --dry-run
python scripts/nouvelle_facture.py examples/spec_exemple.json
```

Contrôler le MASTER avant d'émettre :

```bash
python scripts/check_coherence.py --ecole EPITA --sessions
```

En Python :

```python
from datetime import date
from openpyxl import Workbook
from facturation import pcf_format as fmt

fmt.set_data_root("/chemin/vers/Factures")
wb = Workbook(); wb.remove(wb.active)
ws = fmt.clone_template(wb, "R_EPITA010 Sept 2026", entite="ROSARI")
fmt.apply_emetteur(ws, entite="ROSARI")
fmt.apply_destinataire(ws, dest=fmt.get_destinataire("EPITA"), entite="ROSARI")
fmt.apply_metadata(ws, n_facture="R_EPITA010", date_emission=date(2026, 9, 30),
                   ref_bon_commande="PO-XX-0001")
fmt.place_sessions(ws, "EPITA", groups=[
    ("AIS S1", "[2026][PI][AIS S1][UE][98H]", [(date(2026, 9, 15), 21, 90)]),
], entite="ROSARI", masquer_jour=True)
fmt.apply_rib(ws, entite="ROSARI")
fmt.apply_print_setup(ws, entite="ROSARI")
wb.save("facture.xlsx")
print(fmt.post_mortem(ws, entite="ROSARI"))
```

## Gabarit

Une mise en page unique, `facturation/assets/PCF_facture_standard.xlsx`, pour toutes
les entités : seules les informations de l'émetteur changent. Une entité peut avoir
son propre gabarit `PCF_facture_<ENTITE>.xlsx`, prioritaire s'il existe. Voir
[references/FORMAT.md](references/FORMAT.md).

## Tests

```bash
python -m pytest                 # toute la banque
python -m pytest -m regression   # format et rendu des factures
python -m pytest -m unit         # profil et fonctions pures
```

Tous les tests tournent sur des données fictives créées à la volée. Les défauts
connus du gabarit (T-FMT-07, à corriger par F-18) sont marqués `xfail` strict : ils
restent visibles dans le rapport et feront échouer la suite le jour où ils seront
corrigés, pour rappeler de retirer le marqueur. GitHub Actions rejoue la banque à
chaque push sous Windows et Ubuntu. Le rendu PDF est sauté si LibreOffice est absent.

## Conventions

- Branches : `main` (stable), `develop` (en cours), `feature/*`.
- Commits conventionnels : `feat:`, `fix:`, `docs:`, `test:`, `refactor:`, `chore:`.
- Aucune donnée réelle dans le dépôt, y compris dans les tests et les exemples.
- Décisions : ADR001 à ADR011 dans [references/ADR.md](references/ADR.md), les
  suivantes dans le journal de décisions du second cerveau Notion.
