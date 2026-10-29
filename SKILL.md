---
name: facturation-pcf
description: Génère les factures et devis d'auto-entrepreneur formateur (entités ROSARI/KAUR) au format XLSX canonique, en appliquant le Protocole de Création de Factures en 7 étapes. Utiliser dès qu'il s'agit d'émettre, régénérer, corriger ou contrôler une facture ou un devis pour une école (EPITA, DOR, ISG, EMN, CDP…), de journaliser des séances dans le MASTER, de vérifier la cohérence heures dispensées / facturées, ou de faire évoluer le gabarit de facture. Utiliser aussi quand l'utilisateur fournit un bon de commande (PO) à facturer.
---

# Facturation PCF

Génération de factures XLSX pour deux auto-entrepreneurs formateurs distincts,
**ROSARI** et **KAUR**, facturant des écoles à l'heure de cours.

## Règles absolues

1. **Aperçu obligatoire avant toute écriture** (`.xlsx` comme `.md`). Afficher ce
   qui va être écrit, attendre la validation.
2. **Étanchéité ROSARI / KAUR** : aucun croisement, jamais. Une commande porte sur
   une entité et une seule.
3. **Jamais d'email envoyé automatiquement** — brouillon uniquement.
4. **Source de vérité = le MASTER de l'entité.** Le calendrier sert à détecter les
   séances, pas à les valider.
5. **Demander confirmation si une information manque** (taux, contact, volume
   horaire). Ne jamais inventer un montant ni une coordonnée.

## Avant de commencer

Localiser la racine des données (le dossier contenant `ROSARI/`, `KAUR/`) et la
passer au module — `--data-root`, ou la variable d'environnement `PCF_DATA_ROOT`.
Sans elle, le dossier courant est utilisé et le MASTER sera introuvable.

Si un `CONTEXT.md` existe à la racine des données, le lire : il porte les écoles,
les tarifs et les conventions de classement du calendrier, qui ne sont
volontairement pas dans ce paquet.

## Protocole

Toute génération suit les 7 étapes de [references/PROC00_pcf.md](references/PROC00_pcf.md).
Les étapes 1, 4 et 6 sont des validations humaines : ne pas les automatiser, elles
protègent contre les erreurs irrécupérables (mauvais destinataire, mauvais montant,
mauvaise période).

En résumé : périmètre → extraction calendrier → croisement MASTER et journalisation
→ coordonnées → format → **aperçu validé** → écriture et post-mortem.

## Générer une facture

Cas courant — écrire une spec JSON puis lancer le CLI :

```bash
python scripts/nouvelle_facture.py ma_spec.json --data-root "<racine>" --dry-run   # aperçu (étape 6)
python scripts/nouvelle_facture.py ma_spec.json --data-root "<racine>"             # écriture (étape 7)
```

Le format de la spec est décrit dans [examples/spec_exemple.json](examples/spec_exemple.json).
Le CLI couvre les étapes 5 et 7 ; la journalisation dans `MASTER.Sessions`
(étape 3) reste à faire — voir le script d'exemple qui la traite.

Cas complexe (plusieurs factures d'un coup, logique métier particulière) : écrire
un script dédié sur le modèle de
[examples/generate_epita_pos.py](examples/generate_epita_pos.py), qui émet sept
factures à partir de deux bons de commande et met à jour le journal des sessions
de façon idempotente.

## Contrôler

```bash
python scripts/check_coherence.py --entite ROSARI --sessions
```

Signale les écarts `heures dispensées ≠ facturées + à facturer` et récapitule les
totaux par facture — à lancer avant toute nouvelle émission (étape 3.1).

## Facturer un bon de commande

Quand l'utilisateur fournit un PO (EPITA passe par Basware) :

- **Une facture par ligne de PO.** Ne pas regrouper : le rapprochement comptable
  se fait ligne à ligne.
- **Recopier le code produit du PO à l'identique** dans le libellé, y compris sa
  notation à crochets, sans le reformater ni l'harmoniser entre POs.
- **Renseigner `ref_bon_commande`** : la référence du PO va sur sa propre ligne de
  la facture, pas dans le libellé. Quand plusieurs factures partagent un PO,
  préciser la ligne (`PO-26-… (ligne 3)`).
- **Masquer la colonne Jour** (`masquer_jour: true`) : ces factures portent un
  volume horaire, pas des séances datées.
- **Vérifier que le total facturé égale le total du PO** avant de valider.

## Modifier le format

Le format a deux sources et deux seulement : le gabarit `facturation/assets/` et le module
`facturation/pcf_format.py`. Lire [references/FORMAT.md](references/FORMAT.md) avant
d'y toucher — il documente le plan des cellules, les invariants (ne jamais
`delete_rows`) et les pièges vérifiés, dont la fusion `M7:P8` qui fait disparaître
silencieusement la première ligne d'adresse.

Après toute modification : `python -m pytest` (données fictives, aucune racine requise).

## Autres références

- [references/MASTER_SCHEMA.md](references/MASTER_SCHEMA.md) — structure attendue du classeur MASTER
- [references/ONBOARDING.md](references/ONBOARDING.md) — ajouter une nouvelle école
- [references/ADR.md](references/ADR.md) — pourquoi le système est fait ainsi
