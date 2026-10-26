# PROC00 — Protocole de Création de Factures (PCF)

**Statut** : ✅ Acté — version courante
**Priorité** : P0 (cœur du système)
**Version** : v0.3 → v0.4 (intégration Conventions_horaires + module Python)
**Créé** : 2026-04-26
**Dernière mise à jour** : 2026-05-04

> Protocole obligatoire à appliquer pour **toute** génération de facture ou devis. Aucune génération ne peut se faire sans dérouler ces 7 étapes.

---

## Principe directeur

Le PCF est un **protocole hybride** :
- **Étapes 2, 3, 5, 7** : automatisables — Claude Code / scripts Python les exécutent
- **Étapes 1, 4, 6** : validation humaine obligatoire — gates non automatisables

Ne pas tenter d'automatiser les gates humains : ils existent pour éviter les erreurs irrécupérables (mauvais destinataire, mauvais montant, mauvaise période).

---

## Pré-requis avant lancement

```
□ CONTEXT.md de la racine de données lu (écoles, tarifs, conventions)
□ Entité identifiée (ROSARI ou KAUR)
□ Racine de données connue (--data-root ou $PCF_DATA_ROOT)
□ MASTER de l'entité accessible (cf. MASTER_SCHEMA.md)
□ Calendrier connecté (gcal access)
□ Module scripts/pcf_format.py et gabarit assets/ disponibles (cf. FORMAT.md)
```

---

## Les 7 étapes

### Étape 1 — Période & portée [HUMAIN]

**Objectif** : énoncer clairement le périmètre de la facture/devis avant tout calcul.

**Données à fournir** :
- Entité : ROSARI ou KAUR
- École : code court (DOR, EPITA, CDP, etc.)
- Période couverte : dates précises (ex. "Mars 2026" = 1er → 31 mars 2026)
- Type : Facture ou Devis
- N° prévu : déterminé en lisant la dernière séquence dans MASTER + 1
- Date d'émission : aujourd'hui par défaut

**Sortie** :
```
ENTITÉ : ROSARI
ÉCOLE : DOR
PÉRIODE : Mars 2026 (01/03/2026 → 31/03/2026)
TYPE : Facture
N° : R_DOR003 (dernière = R_DOR002)
DATE ÉMISSION : 04/05/2026
```

---

### Étape 2 — Extraction calendrier [AUTO]

**Objectif** : récupérer toutes les sessions du calendrier sur la période, avec tous les détails utiles.

**Procédure** :
- Appel `gcal_list_events` avec `condenseEventDetails: true`
- Plage temporelle = période définie en étape 1
- Filtrage par école selon les règles de classification du `CONTEXT.md` de la racine de données

**Format de sortie obligatoire** :
```
date | début–fin | durée_décimale | summary verbatim
2026-03-12 | 09:00–12:00 | 3.0 | [NEX] Maths BTS SIO A1 — Algèbre linéaire
2026-03-19 | 09:00–12:00 | 3.0 | [NEX] Maths BTS SIO A1 — Suites
...
```

**Le calendrier est la source du contenu réel des cours dispensés** — pas le MASTER (qui est la source de vérité agrégée).

---

### Étape 3 — Croisement MASTER + journalisation [AUTO]

**Objectif** : vérifier la cohérence et alimenter le journal Sessions.

#### 3.1 — Vérification

- Lire `MASTER.Cours` pour la ligne école/module/programme concerné
- Comparer `H dispensées` avec la somme calendrier de l'étape 2
- **Signaler tout écart avant de continuer** (ne pas masquer)

#### 3.2 — Application coefficient horaire

- Lire `MASTER.Conventions_horaires` pour le couple école/module
- Appliquer le coefficient à chaque session :
  ```
  H_facturable = H_calendrier × coefficient
  ```
- Si la convention manque → **STOP**, créer la ligne en `Conventions_horaires` avant de continuer

#### 3.3 — Journalisation Sessions (append-only)

Pour chaque session retenue, ajouter une ligne dans `MASTER.Sessions` :

```
Date_log | École | Période_facture | N°_Facture | Date_séance | Début | Fin | Durée_h_facturable | Promotion | Module/UE | Taux | Total HT | Summary calendrier verbatim
```

**Mode append-only strict** : jamais d'écrasement. Si la facture est régénérée, ajouter de nouvelles lignes (avec un nouveau `Date_log`) — l'historique reste tracé.

---

### Étape 4 — Coordonnées [HUMAIN]

**Objectif** : valider que les coordonnées sont complètes des deux côtés.

#### 4.1 — Émetteur (selon entité)

Lecture depuis `MASTER.Coordonnées` filtré sur `Type=Emetteur` :
- Nom, adresse complète, email, téléphone
- SIRET, IBAN, BIC, mention TVA (depuis ligne `Info`)

#### 4.2 — Destinataire (selon école)

Lecture depuis `MASTER.Coordonnées` filtré sur `Type=Destinataire` et le code école :
- Raison sociale, adresse complète, email contact comptabilité

#### 4.3 — Gate humain : vérification

```
SI destinataire vide ou incomplet :
   → STOP
   → Demander à l'utilisateur de compléter via ONBOARDING.md
   → Ne pas continuer tant que l'info manque
```

---

### Étape 5 — Format [AUTO]

**Objectif** : appliquer le format canonique sans déviation.

#### 5.1 — Source unique du format

- Gabarit : `assets/PCF_facture_<ENTITE>.xlsx > MODELE`
- Module : `scripts/pcf_format.py`

**Pas de spec visuelle ailleurs.** Toute évolution du format passe par modification du gabarit + adaptation du module. Détail du plan des cellules et des pièges : [FORMAT.md](FORMAT.md).

#### 5.2 — Procédure

Cas courant : écrire une spec JSON et lancer `scripts/nouvelle_facture.py`
(`--dry-run` produit l'aperçu de l'étape 6). Cas particulier, en Python :

```python
import sys; sys.path.insert(0, "scripts")
import pcf_format as fmt

fmt.set_data_root(racine)                      # ou $PCF_DATA_ROOT

ws = fmt.clone_template(dst_wb, sheet_name=f"{numero} {periode}", entite="ROSARI")
fmt.apply_emetteur(ws, entite="ROSARI")        # ou "KAUR"
fmt.apply_destinataire(ws, dest=fmt.get_destinataire(code_ecole), entite="ROSARI")
fmt.apply_metadata(ws, n_facture=numero, date_emission=date_emission,
                   entite="ROSARI", ref_bon_commande=ref_po)   # ref_po optionnel
fmt.place_sessions(ws, ecole_name=ecole, groups=sessions_groupees,
                   entite="ROSARI", masquer_jour=False)
fmt.apply_rib(ws, entite="ROSARI")
fmt.apply_print_setup(ws, entite="ROSARI")
```

#### 5.3 — Règles invariantes

- Bandeau violet B2:S3 (préservé du gabarit, non modifiable)
- **Émetteur (ROSARI)** : col C rows 7-11, **encadré bordures fines noires (C7:E11)**
  - Row 7 : nom (gras 12pt)
  - Rows 8-9 : adresse postale
  - Row 10 : email
  - Row 11 : téléphone
- **Destinataire (ROSARI)** : col M rows 7-10, **encadré bordures fines noires (M7:P10)**
  - Row 7 : nom (gras 12pt)
  - Rows 8-9 : adresse postale
  - Row 10 : email
- **Format KAUR à venir** (cf. FMT01 itération 2 — template KAUR à créer)
- Auto-resize : `row_dimensions[r].height = None` pour toutes lignes avec texte
- Si nb_sessions < 60 : masquer rows excédentaires (`hidden=True`), vider contenu
- **Ne JAMAIS `delete_rows`** — casse les plages fusionnées en-dessous (IBAN)
- `insert_rows()` autorisé uniquement pour extension > 60 sessions, avec démerge préalable
- Formule TOTAL : `=SUM(Q20:R79)` — fonctionne avec rows masquées (somme = 0 pour rows vides)

---

### Étape 6 — Aperçu avant écriture [HUMAIN]

**Objectif** : validation finale avant écriture sur disque + envoi.

#### 6.1 — Aperçu à afficher

```
====== APERÇU FACTURE ======
Type : FACTURE / DEVIS
N° : R_DOR003
Date émission : 04/05/2026
Émetteur : Adrian Rosari (Sannois)
Destinataire : NEXA Digital School — [adresse complète]

DÉTAIL DES SÉANCES :
| Date | Module | Promo | Durée | Taux | Total |
| 12/03 | Algèbre lin | A1 | 3h | 60€ | 180€ |
| 19/03 | Suites | A1 | 3h | 60€ | 180€ |
| ...

TOTAL HEURES : 14h
TOTAL HT : 840 €
DÉLAI PAIEMENT : Fin de mois
============================

VALIDER ? [OUI / NON / CORRIGER]
```

#### 6.2 — Gate humain

```
SI utilisateur répond OUI :
   → Passer à l'étape 7

SI utilisateur répond NON ou CORRIGER :
   → Retour à l'étape 1, 2, 3 ou 4 selon le problème
   → Ne pas écrire le fichier
```

---

### Étape 7 — Écriture & vérification post-mortem [AUTO]

**Objectif** : écrire le fichier et confirmer que rien n'a été corrompu.

#### 7.1 — Écriture

- Sauvegarder le XLSX dans `[ENTITE]/[ECOLE]/[ENTITE]_[ECOLE].xlsx` (nouvel onglet)
- Mettre à jour `MASTER.Cours` :
  - `H facturées` += heures de cette facture
  - `H à facturer` -= heures de cette facture
  - `N° Facture liée` = numéro émis
  - `Statut` → `Facturé`
- Si applicable : mettre à jour `MASTER.Devis` (si la facture résulte d'un devis accepté)

#### 7.2 — Post-mortem (lecture des cellules clés)

```python
checks = fmt.post_mortem(ws, entite="ROSARI")
# titre, emetteur_nom, destinataire_nom, n_facture, siret,
# ecole, premier_libelle, total_formule, iban, bic
```

Contrôler en particulier que `destinataire_nom`, `siret` et `iban` ne sont pas
`None` : le gabarit ne les contient pas, ils ne peuvent venir que des appels
`apply_*`. Un `None` signale un appel oublié, donc une facture inexploitable.

#### 7.3 — Confirmation

```
SI toutes cellules clés sont conformes à l'aperçu de l'étape 6 :
   → "OK — facture R_DOR003 émise et MASTER mis à jour"

SI au moins une cellule diffère :
   → "ERREUR sur cellule X : attendu Y, trouvé Z"
   → Ne pas marquer comme finalisée
   → Demander instructions à l'utilisateur
```

#### 7.4 — Export PDF (cf. FMT02)

Après validation de l'écriture XLSX :
- Générer le PDF via LibreOffice CLI
- Stocker à côté du XLSX : `R_DOR003_DOR_Mars2026.pdf`

---

## Critères de fin (Definition of Done)

Une facture est considérée comme **finalisée** quand :

```
□ Étapes 1-7 toutes exécutées
□ XLSX généré dans le bon dossier
□ PDF généré
□ MASTER.Cours mis à jour (H facturées, statut, N° lié)
□ MASTER.Sessions alimenté (toutes les séances loguées)
□ Post-mortem 6 cellules → OK
□ Aperçu validé par l'utilisateur en étape 6
```

---

## Tests / Validation

### Test de non-régression du format

À chaque évolution du module `pcf_format.py` :
- Régénérer une facture pilote (ex. `R_DOR_TEST001`)
- Comparer pixel-à-pixel avec la référence canonique
- Vérifier les 6 cellules clés
- Détruire la facture pilote après validation

### Test de l'append-only Sessions

Régénérer 2 fois la même facture → vérifier que `MASTER.Sessions` contient 2× les lignes (pas d'écrasement).

### Test du gate destinataire vide

Lancer le PCF sur une école sans destinataire dans `Coordonnées` → vérifier que le STOP étape 4 fonctionne.

---

## Notes / Décisions

- **Pourquoi 7 étapes et pas moins** : chaque étape a un rôle distinct, les gates humains (1, 4, 6) protègent contre les erreurs irrécupérables, les étapes auto (2, 3, 5, 7) ne sont jamais skippées.
- **Pourquoi append-only Sessions** : audit trail légal en cas de litige client ou contrôle fiscal — voir [ADR003](ADR.md).
- **Pourquoi `nouvelle_facture.py` ne couvre pas tout le cycle** : le script applique les étapes automatisables (5 et 7) et affiche l'aperçu de l'étape 6, mais ne remplace pas le protocole. Les gates humains et la journalisation Sessions (étape 3) restent explicites.
