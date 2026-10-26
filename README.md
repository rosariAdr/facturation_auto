# facturation_module

Génération de factures XLSX pour auto-entrepreneur formateur : un gabarit Excel
canonique, un module Python qui en force les invariants, et un protocole en 7
étapes avec validations humaines.

Le paquet s'utilise **soit** comme un dépôt d'outils en ligne de commande, **soit**
comme un skill Claude Code (`SKILL.md` à la racine).

## Ce qu'il contient — et ce qu'il ne contient pas

Le code et le format, jamais les données. Aucun nom, adresse, SIRET, IBAN ni
montant n'est versionné : le gabarit livré est vide de toute coordonnée, et le
module les lit au runtime dans `MASTER.Coordonnées`. Les données métier restent
dans un dossier séparé, la *racine de données*.

```
facturation_module/
├── SKILL.md                      point d'entrée skill
├── references/
│   ├── PROC00_pcf.md             protocole de création de factures (7 étapes)
│   ├── FORMAT.md                 plan du gabarit, invariants, pièges
│   ├── MASTER_SCHEMA.md          structure attendue du classeur de données
│   ├── ONBOARDING.md             ajouter une école
│   └── ADR.md                    décisions d'architecture
├── scripts/
│   ├── pcf_format.py             module canonique (API de 10 fonctions)
│   ├── nouvelle_facture.py       CLI : spec JSON → facture
│   ├── check_coherence.py        contrôles MASTER (lecture seule)
│   ├── build_template.py         reconstruction du gabarit (maintenance)
│   └── test_pcf_format.py        tests de non-régression
├── assets/
│   └── PCF_facture_ROSARI.xlsx   gabarit MODELE, sans données
└── examples/
    ├── spec_exemple.json
    └── generate_epita_pos.py     cas réel : 7 factures depuis 2 bons de commande
```

## Installation

```bash
git clone <ce-dépôt> facturation_module
cd facturation_module
pip install -r requirements.txt
```

Python ≥ 3.9. Seule dépendance : `openpyxl`.

### Comme skill Claude Code

Cloner (ou lier) le dépôt dans le dossier des skills, de sorte que `SKILL.md` se
trouve à la racine du dossier du skill :

```bash
git clone <ce-dépôt> ~/.claude/skills/facturation-pcf
```

Sous Windows : `%USERPROFILE%\.claude\skills\facturation-pcf`. Pour un skill limité
à un projet, utiliser `.claude/skills/` dans ce projet.

## Racine de données

Le module ne sait rien de l'emplacement des données tant qu'on ne le lui dit pas.
Elle est résolue dans cet ordre :

1. `pcf_format.set_data_root("<chemin>")` en Python, ou `--data-root` en ligne de commande
2. la variable d'environnement `PCF_DATA_ROOT`
3. à défaut, le dossier courant

Elle doit contenir `<ENTITE>/<ENTITE>_MASTER.xlsx` — voir
[references/MASTER_SCHEMA.md](references/MASTER_SCHEMA.md).

```bash
export PCF_DATA_ROOT="/chemin/vers/Factures"     # bash
$env:PCF_DATA_ROOT = "C:\chemin\vers\Factures"   # PowerShell
```

## Usage

Aperçu puis émission d'une facture décrite par une spec JSON :

```bash
python scripts/nouvelle_facture.py examples/spec_exemple.json --dry-run
python scripts/nouvelle_facture.py examples/spec_exemple.json
```

Contrôler le MASTER avant d'émettre :

```bash
python scripts/check_coherence.py --entite ROSARI --ecole EPITA --sessions
```

En Python, pour un cas qui ne rentre pas dans une spec :

```python
import sys; sys.path.insert(0, "scripts")
from datetime import date
from openpyxl import Workbook
import pcf_format as fmt

fmt.set_data_root("/chemin/vers/Factures")

wb = Workbook(); wb.remove(wb.active)
ws = fmt.clone_template(wb, "R_EPITA010 Sept 2026", entite="ROSARI")
fmt.apply_emetteur(ws, entite="ROSARI")
fmt.apply_destinataire(ws, dest=fmt.get_destinataire("EPITA"), entite="ROSARI")
fmt.apply_metadata(ws, n_facture="R_EPITA010", date_emission=date(2026, 9, 30),
                   ref_bon_commande="PO-26-ASSOTA-02001")
fmt.place_sessions(ws, "EPITA", groups=[
    ("AIS S1", "[2026][PI][AIS S1][Algorithmic Workshop][98H][S8_25-26][26.06]",
     [(date(2026, 9, 15), 21, 90)]),
], entite="ROSARI", masquer_jour=True)
fmt.apply_rib(ws, entite="ROSARI")
fmt.apply_print_setup(ws, entite="ROSARI")
wb.save("facture.xlsx")
print(fmt.post_mortem(ws, entite="ROSARI"))
```

## Tests

```bash
python scripts/test_pcf_format.py --data-root "/chemin/vers/Factures"
```

Six tests : structure du gabarit, injection des coordonnées depuis le MASTER,
adresse destinataire complète (non-régression de la fusion `M7:P8`), destinataire
vide, session unique, bon de commande et masquage de la colonne Jour. Les tests
ont besoin d'un MASTER accessible et écrivent dans un dossier temporaire.

## Entité KAUR

Non implémentée : `clone_template(entite="KAUR")` lève `NotImplementedError` tant
que `assets/PCF_facture_KAUR.xlsx` n'existe pas. Voir la fin de
[references/FORMAT.md](references/FORMAT.md).
