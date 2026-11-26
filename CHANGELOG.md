# Changelog

Format : [Keep a Changelog](https://keepachangelog.com/fr/1.1.0/). Versions : `0.x` tant que la v2 n'est pas en service.

## [Unreleased]

### Ajouté (F-24, mentions légales de l'ADR030)
- Bas de facture écrit par le module (`apply_mentions`) : « Total HT », « Net à payer », mention 293 B, déclaration d'activité lue dans `Info / NDA`, échéance datée (émission + 30 jours par défaut), « Pas d'escompte », pénalités exigibles dès le lendemain de l'échéance, indemnité forfaitaire de 40 €.
- Ligne « N° de contrat » (`ref_contrat`) et champ `echeance` dans la spec de `nouvelle_facture.py` ; aperçu complété.
- Refus d'écrire, dès l'aperçu, si le NDA manque ou si l'émetteur ne porte pas « EI » ; mêmes contrôles en ERREUR dans `profil verifier`.
- Contrôle après écriture étendu à l'échéance, au NDA et au contrat.

### Corrigé (F-24)
- Au-delà de 60 séances, le net à payer pointait vers une mauvaise cellule et la fin de facture sortait de la zone d'impression.

### Corrigé (fusionné sur `dev` le 21/11/2026)
- `nouvelle_facture.py` refuse d'écraser une facture existante (onglet ou numéro déjà présent).
- Entité exigée partout ; préfixe ou dossier d'une autre entité refusés.
- Facture bloquée si le destinataire n'a pas d'adresse.
- Contrôle après écriture vérifié : sortie en erreur en cas d'écart.
- Aperçu : émetteur et libellé complet affichés.
- `check_coherence.py` signale des heures facturées au-delà des heures dispensées.
- Profil : une adresse iCal collée dans le profil est refusée, sans être recopiée.
- Scripts et tests : le paquet de ce dépôt est toujours celui importé.

### Modifié
- CI déclenchée sur `dev` et les branches `feat/fix/chore`, avec Python 3.11 ; `AGENTS.md` ajouté.

### Documentation (23/11/2026, cadrage)
- `references/ADR.md` : toutes les décisions dans le dépôt (ADR012–024 rapatriés de Notion, ADR025–033 ajoutés, statuts à jour).
- Nouveaux : `references/ROADMAP.md`, `references/GLOSSAIRE.md`, `references/RECETTE_DOC.md`.
- Réécrits : `PROC00_pcf.md` (aujourd'hui / cible du lot 1), `MASTER_SCHEMA.md` (v2), `FORMAT.md` (mentions de l'ADR030), `ONBOARDING.md` (MASTER v2, plus de `CONTEXT.md`), `README.md` (ce qui est disponible, prévu ou à l'état d'idée ; où vit quoi ; exemple Python corrigé), `SKILL.md` (aiguillage).
- `AGENTS.md` §7 : où vit quoi, liste des documents de référence, durée de la porte sur le PC d'Adrian.

### Retiré
- Données réelles des documents de référence et de l'exemple par bon de commande, renommé `examples/factures_par_bon_de_commande.py` (données fictives, comportement inchangé). Restent à traiter : `examples/spec_exemple.json`, les tests et deux docstrings, liés entre eux (ticket à part).

## [0.2.0.dev0] — lot 0, socle (en cours)

### Ajouté
- Paquet Python installable `facturation-auto` (`pyproject.toml`) et commande `facturation`.
- Profil par utilisateur en TOML hors dépôt (`facturation/config.py`) : entité, agenda, racine de données, lancement (25 du mois par défaut), moteur PDF ; validation champ par champ.
- `facturation profil init`, `facturation profil chemin`, `facturation profil verifier` (contrôles de l'installation, sans écriture).
- `facturation profil agenda` : enregistre l'adresse iCal secrète dans `agenda.secret`, à côté du profil, saisie sans affichage (ADR019).
- Heure de lancement par défaut : 16 h le 25 du mois (Q6).
- Banque de tests pytest sur données fictives : format (T-FMT-01 à 08), profil (T-CFG-01…), scripts et vérification (T-CLI-01/02, T-COH-01, T-VER-01/02, T-CFG-02/03).
- Intégration continue GitHub Actions (Windows et Ubuntu).
- `examples/profil_exemple.toml`, `CHANGELOG.md`.

### Modifié
- Clonage du gabarit 8 fois plus rapide : chacun des ~60 styles distincts n'est recopié qu'une fois au lieu d'être recopié pour chacune des ~26 800 cellules stylées ; résultat identique au fichier près, garanti par le nouveau test T-FMT-08. Écriture d'une facture : 5,4 s → 1,3 s sur la machine de référence ; banque de tests deux fois plus rapide.
- `pcf_format.py` déplacé dans le paquet `facturation/` ; entités non limitées à ROSARI/KAUR ; gabarit standard commun `PCF_facture_standard.xlsx` (ADR015), gabarit propre à une entité prioritaire s'il existe.
- Scripts `nouvelle_facture.py` et `check_coherence.py` : option `--profil`, racine résolue depuis le profil.

### Retiré
- `scripts/test_pcf_format.py`, remplacé par `tests/regression/test_format.py`.
- `requirements.txt`, remplacé par `pyproject.toml`.

### Connu
- T-FMT-07 (xfail strict) : formats perdus après la 7e séance, mentions légales en page 2 — correction prévue en F-18 (lot 1).

## [0.1.0] — 24/09/2026

Import de la v1 : module de format, CLI spec JSON, contrôles du MASTER, gabarit ROSARI, protocole PCF, ADR001 à ADR011.
