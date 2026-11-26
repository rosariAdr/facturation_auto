---
name: facturation-pcf
description: Génère et contrôle les factures et devis de formateurs indépendants (entités ROSARI et KAUR) au format XLSX du projet facturation_auto, en suivant le protocole PCF. Utiliser dès qu'il s'agit d'émettre, régénérer, corriger ou contrôler une facture ou un devis pour une école, de tenir le journal des séances dans le MASTER, de vérifier la cohérence heures dispensées / facturées, de facturer un bon de commande ou de faire évoluer le gabarit de facture.
---

# Facturation PCF

Factures XLSX pour des formateurs indépendants qui facturent des écoles à l'heure de cours. Deux entités aujourd'hui, **ROSARI** et **KAUR**, strictement séparées.

Ce fichier est un aiguillage : les règles vivent dans `AGENTS.md` et dans `references/` (ADR025).

## Règles absolues

1. **Aperçu avant toute écriture** (classeur, MASTER), puis validation explicite. Une validation par facture.
2. **Étanchéité des entités** : une commande porte sur une entité et une seule (ADR001).
3. **Jamais d'email envoyé** : brouillons seulement.
4. **Le MASTER de l'entité fait foi** (ADR005), et c'est là que vivent les règles de chaque école (ADR026). L'agenda sert à repérer les séances, pas à les valider.
5. **Ne jamais inventer** un taux, une coordonnée ou un volume horaire : demander.
6. **Aucune donnée réelle dans le dépôt**.

## Avant de commencer

- Localiser la racine des données (le dossier qui contient `ROSARI/`, `KAUR/`) et la passer aux scripts : `--data-root`, `--profil`, ou la variable `PCF_DATA_ROOT`. Sans elle, le dossier courant est utilisé et le MASTER sera introuvable.
- Lire les règles de l'école dans le MASTER (`Tarifs`, `Règles_facturation`, `Correspondances`, `Coordonnées`) : voir [references/MASTER_SCHEMA.md](references/MASTER_SCHEMA.md).

## Protocole

Toute facture suit [references/PROC00_pcf.md](references/PROC00_pcf.md) : la partie 1 décrit ce que l'on fait aujourd'hui, la partie 2 la cible du lot 1. Les validations humaines ne s'automatisent pas.

Contenu légal obligatoire de chaque facture : ADR030 dans [references/ADR.md](references/ADR.md). Tant que le gabarit ne porte pas toutes ces mentions (F-24), les vérifier à la main avant l'envoi.

## Générer une facture (outil actuel)

```bash
python scripts/nouvelle_facture.py ma_spec.json --data-root "<racine>" --dry-run   # aperçu
python scripts/nouvelle_facture.py ma_spec.json --data-root "<racine>"             # écriture
```

Format de la spec : [examples/spec_exemple.json](examples/spec_exemple.json). Le script fait l'aperçu, écrit l'onglet de la facture (il refuse d'écraser une facture existante) et contrôle les cellules clés. **Il ne met à jour ni le journal `Sessions` ni le suivi `Cours`, et ne produit pas de PDF** : ces étapes restent à faire à part jusqu'au lot 1.

Plusieurs factures d'un coup : un script dédié sur le modèle de [examples/factures_par_bon_de_commande.py](examples/factures_par_bon_de_commande.py), qui émet une facture par ligne de bon de commande et tient le journal des séances.

## Contrôler

```bash
python scripts/check_coherence.py --entite ROSARI --sessions
```

Signale les écarts `heures dispensées ≠ facturées + à facturer` et des heures facturées au-delà des heures dispensées, puis récapitule les totaux par facture. Il additionne les lignes du journal sans repérer les doublons.

## Facturer un bon de commande

Quand l'école émet des bons de commande (règle `Bon de commande requis` de l'école) :

- **Une facture par ligne de bon de commande.** Ne pas regrouper : le rapprochement comptable se fait ligne à ligne.
- **Recopier le code produit à l'identique** dans le libellé, y compris sa notation à crochets.
- **Renseigner `ref_bon_commande`** : la référence va sur sa propre ligne, pas dans le libellé. Si plusieurs factures partagent un bon, préciser la ligne (`… (ligne 3)`).
- **Masquer la colonne Jour** (`masquer_jour: true`) : ces factures portent un volume horaire, pas des séances datées.
- **Vérifier que le total facturé égale le total du bon** avant de valider.

## Modifier le format

Le format a deux sources : le gabarit `facturation/assets/` et le module `facturation/pcf_format.py`. Lire [references/FORMAT.md](references/FORMAT.md) avant d'y toucher : plan des cellules, mentions, invariants (ne jamais `delete_rows`) et pièges vérifiés. Après toute modification : `python -m pytest` (données fictives, aucune racine requise).

## Autres références

- [references/ROADMAP.md](references/ROADMAP.md) — lots et critères d'acceptation
- [references/ONBOARDING.md](references/ONBOARDING.md) — ajouter une école
- [references/GLOSSAIRE.md](references/GLOSSAIRE.md) — vocabulaire
- [AGENTS.md](AGENTS.md) — conventions de travail
