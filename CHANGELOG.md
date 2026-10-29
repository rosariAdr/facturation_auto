# Changelog

Format : [Keep a Changelog](https://keepachangelog.com/fr/1.1.0/). Versions : `0.x` tant que la v2 n'est pas en service.

## [0.2.0.dev0] — lot 0, socle (en cours)

### Ajouté
- Paquet Python installable `facturation-auto` (`pyproject.toml`) et commande `facturation`.
- Profil par utilisateur en TOML hors dépôt (`facturation/config.py`) : entité, agenda, racine de données, lancement (25 du mois par défaut), moteur PDF ; validation champ par champ.
- `facturation profil init`, `facturation profil chemin`, `facturation profil verifier` (contrôles de l'installation, sans écriture).
- Banque de tests pytest sur données fictives : format (T-FMT-01 à 07), profil (T-CFG-01…), scripts et vérification (T-CLI-01/02, T-COH-01, T-VER-01/02).
- Intégration continue GitHub Actions (Windows et Ubuntu).
- `examples/profil_exemple.toml`, `CHANGELOG.md`.

### Modifié
- `pcf_format.py` déplacé dans le paquet `facturation/` ; entités non limitées à ROSARI/KAUR ; gabarit standard commun `PCF_facture_standard.xlsx` (ADR015), gabarit propre à une entité prioritaire s'il existe.
- Scripts `nouvelle_facture.py` et `check_coherence.py` : option `--profil`, racine résolue depuis le profil.

### Retiré
- `scripts/test_pcf_format.py`, remplacé par `tests/regression/test_format.py`.
- `requirements.txt`, remplacé par `pyproject.toml`.

### Connu
- T-FMT-07 (xfail strict) : formats perdus après la 7e séance, mentions légales en page 2 — correction prévue en F-18 (lot 1).

## [0.1.0] — 24/09/2026

Import de la v1 : module de format, CLI spec JSON, contrôles du MASTER, gabarit ROSARI, protocole PCF, ADR001 à ADR011.
