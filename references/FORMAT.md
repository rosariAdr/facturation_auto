# FORMAT — Gabarit canonique et invariants de mise en forme

**Dernière mise à jour** : 2026-11-23

> Source unique du format : `facturation/assets/PCF_facture_standard.xlsx > MODELE` (ou `PCF_facture_<ENTITE>.xlsx` si l'entité a son propre gabarit) + `facturation/pcf_format.py`. Toute spec visuelle écrite ailleurs est obsolète. Faire évoluer le format = modifier le gabarit **et** adapter le module, jamais l'un sans l'autre.
>
> Le contenu légal exigé sur chaque facture est décidé par l'ADR030 ; ce fichier dit où il se place. Une zone marquée « prévu » n'est pas encore dans le gabarit.

---

## Principe

Le gabarit porte **la mise en forme uniquement** : bandeau, encadrés, bordures, largeurs, hauteurs, formules, mentions fixes. Il ne contient **aucune donnée personnelle** : ni nom, ni adresse, ni SIRET, ni IBAN. Ces valeurs sont injectées depuis `MASTER.Coordonnées` par `apply_emetteur`, `apply_destinataire`, `apply_metadata` et `apply_rib`.

Conséquence : le gabarit peut être versionné et partagé sans rien exposer, et changer de RIB ne demande pas de régénérer le gabarit.

---

## Plan du gabarit standard (onglet `MODELE`)

| Zone | Cellules | Rempli par |
|---|---|---|
| Bandeau titre | `B2:S3` | `apply_metadata` (titre + date d'émission) |
| Encadré émetteur | `C7:E11` | `apply_emetteur` (5 lignes) |
| Encadré destinataire | `M7:P10` | `apply_destinataire` (4 lignes) |
| Bon de commande | `C14` + `D14:H14` | `apply_metadata(ref_bon_commande=…)` |
| N° SIRET | `C16` + `D16` | `apply_metadata` |
| N° de facture | `C17` + `D17` | `apply_metadata` |
| En-tête du tableau | ligne 19 | gabarit (fixe) |
| Lignes de cours | lignes 20-79 (60) | `place_sessions` |
| TOTAL | ligne 80 | gabarit : `=SUM(Q20:R79)` |
| À payer, TVA | lignes 83-85 | gabarit (texte fixe, voir « Mentions et totaux ») |
| RIB | `D88` / `D89` | `apply_rib` |
| Mentions de paiement | lignes 91-93 | gabarit (texte fixe, voir « Mentions et totaux ») |

### Colonnes du tableau de cours

| Colonne | Fusion | Contenu |
|---|---|---|
| C | C:D | ÉCOLE (fusionnée verticalement sur les lignes utilisées) |
| E | E:F | Promotion |
| G | G:J | Libellé de cours / groupe |
| K | K:L | Jour — masquable via `place_sessions(masquer_jour=True)` |
| M | M:N | Nb heures |
| O | O:P | Taux horaire HT |
| Q | Q:R | Total — formule `=O{row}*M{row}` |

---

## Mentions et totaux : aujourd'hui et cible (ADR030, F-24)

| Élément | Aujourd'hui dans le gabarit | Cible (prévu, F-24) |
|---|---|---|
| Nom de l'émetteur | Tel qu'écrit dans `Coordonnées` | Nom et prénom suivis de « EI » : c'est une donnée du MASTER, rien à changer dans le gabarit |
| Nom du client | Tel qu'écrit dans `Coordonnées` | Dénomination juridique de la société qui paie : donnée du MASTER |
| Références | Bon de commande (`C14`), n° de facture (`C17`) | Trois lignes distinctes : n° de facture, n° de contrat (nouvelle ligne), n° de bon de commande si l'école en émet un |
| Total | `TOTAL` (ligne 80) | « Total HT » |
| Montant dû | « À payer TTC* » (`C83`) | « Net à payer » |
| Astérisque | « *Montant non assujetti a la TVA » (`C84`) | Retirée |
| Mention de TVA | « TVA non applicable-art. 293 B du CGI- franchise en base » (`C85`) | « TVA non applicable, article 293 B du Code général des impôts. » (citation du CGI admise jusqu'au 30/06/2028, ADR030) |
| Déclaration d'activité | Absente ; ajoutée à la main en `C86` depuis le 30/10/2026 | Une ligne remplie depuis `Coordonnées` (`Info` / `NDA`) |
| Échéance | « Facture Payable Fin de Mois » (`C91`) | « Échéance : JJ/MM/AAAA », calculée : délai du contrat ou de l'école, sinon date d'émission + 30 jours |
| Escompte | « Pas d'escompte en cas de paiement anticipe » (`C92`) | Conservée (accents corrigés) |
| Pénalités | Trois fois le taux d'intérêt légal (`C93`) | Conservée, complétée : « exigibles dès le lendemain de l'échéance, sans rappel préalable » |
| Frais de recouvrement | Absente | « Indemnité forfaitaire pour frais de recouvrement : 40 € » |

Une donnée obligatoire manquante (n° de contrat attendu, délai, NDA…) bloque l'émission au lieu de laisser une zone vide (F-17). Pas de ligne « TVA : 0 € » (ADR030).

---

## Longueur de la facture

- La **facture pilote tient sur une page** (ADR030). Aujourd'hui, au-delà d'environ 17 séances, les mentions passent en page 2 (T-FMT-07, à corriger par F-18).
- Quand une école l'accepte, le regroupement « une ligne par jour et par cours » (règle `Regroupement` de `Règles_facturation`) réduit la longueur en gardant les dates visibles.
- La règle pour les factures de plus d'une page sera décidée après le pilote.

---

## Invariants

- **Ne jamais utiliser `delete_rows()`** : casse les plages fusionnées en dessous (RIB, mentions). Pour retirer une ligne inutilisée : `row_dimensions[r].hidden = True` et vider le contenu.
- `insert_rows()` uniquement pour étendre au-delà de 60 sessions, avec démerge préalable.
- La formule TOTAL reste `=SUM(Q20:R79)` quelles que soient les lignes masquées (une ligne masquée vide compte pour 0).
- Hauteur des lignes de cours : 46.5 pt, `wrap_text` activé (les libellés de bon de commande sont longs).
- Toute ligne portant du texte en 18 pt doit avoir `row_dimensions[r].height = None` (hauteur automatique) : le gabarit hérite de hauteurs de 14.25 pt qui tronquent.

---

## Pièges connus (vérifiés)

**La fusion `M7:P8` avale la première ligne d'adresse.** Le nom du destinataire occupe deux lignes dans le gabarit. Écrire naïvement les lignes d'adresse en `M7`, `M8`, `M9`, `M10` fait tomber la rue dans une `MergedCell`, où l'écriture est ignorée **sans erreur** : la facture part sans son adresse postale. `apply_destinataire` démerge donc la zone et refusionne une plage par ligne.

**Les bordures sont assignables sur une `MergedCell`, pas les valeurs.** Pour encadrer une zone fusionnée, il faut parcourir toutes ses cellules : ne styler que la cellule d'ancrage laisse le cadre droit manquant. C'est ce que fait `_frame_box`.

**Le gabarit porte des bordures résiduelles au milieu des encadrés**, héritées de l'époque où les zones faisaient 3 lignes. `_frame_box` les nettoie à chaque génération : cadre extérieur seul, aucun trait horizontal interne.

**Sans `apply_print_setup`, l'export PDF coupe les colonnes Taux et Total** sur une seconde page : 18 colonnes ne tiennent pas en portrait par défaut. La fonction force portrait + ajustement à une page de large.

**Au-delà de la 7e ligne de séances, les colonnes Heures, Taux et Total perdent leur format nombre** (T-FMT-07). Contourné à la main en octobre 2026 ; correction prévue par F-18.

**Une dénomination longue est coupée dans l'encadré destinataire (`M7`)**, et un libellé de bon de commande long est tronqué en `C14` + `D14:H14`. Contournés à la main en octobre 2026 ; à corriger avec F-18 et F-24.

---

## API du module

Toutes les fonctions qui touchent une facture exigent l'entité en argument nommé (`entite=…`) : il n'y a pas d'entité par défaut.

| Fonction | Rôle |
|---|---|
| `valider_entite(entite)` | Contrôle le format d'un code d'entité |
| `gabarit_pour(entite)` | Gabarit à utiliser : propre à l'entité s'il existe, sinon standard |
| `set_data_root(chemin)` / `get_data_root()` | Racine des données métier |
| `get_emetteur(*, entite)` | Émetteur, lu depuis `MASTER.Coordonnées` |
| `get_destinataire(code_ecole, *, entite)` | Destinataire depuis `MASTER.Coordonnées`, ou `None` |
| `champs_destinataire_manquants(dest)` | Colonnes obligatoires vides pour un destinataire |
| `clone_template(dst_wb, sheet_name, *, entite)` | Clone `MODELE` dans un classeur cible |
| `apply_emetteur(ws, *, entite)` | Encadré émetteur |
| `apply_destinataire(ws, dest, *, entite)` | Encadré destinataire (`None` = encadré vide) |
| `apply_metadata(ws, n_facture, date_emission, *, entite, ref_bon_commande=None)` | Titre, bon de commande, SIRET, n° de facture |
| `place_sessions(ws, ecole_name, groups, *, entite, masquer_jour=False)` | Lignes de cours groupées par (promotion, libellé) |
| `apply_rib(ws, *, entite)` | IBAN + BIC |
| `apply_print_setup(ws, *, entite)` | Portrait, ajusté à une page de large, zone d'impression |
| `post_mortem(ws, *, entite)` | Dictionnaire des cellules clés |
| `verifier_facture(ws, *, entite, n_facture, date_emission, dest, ecole_name, groups)` | Liste des écarts entre la facture relue et les valeurs attendues |

Ordre d'appel : `clone_template` → `apply_emetteur` → `apply_destinataire` → `apply_metadata` → `place_sessions` → `apply_rib` → `apply_print_setup` → `post_mortem` / `verifier_facture`.

---

## Reconstruire le gabarit

Rarement nécessaire. Le gabarit dérive d'une facture réelle servant de référence visuelle :

```bash
python scripts/build_template.py --source "<classeur de référence>.xlsx" --onglet "<onglet de référence>"
```

Le script étend la zone de cours à 60 lignes, trace les encadrés, vide toutes les données et neutralise les coordonnées avant d'écrire `facturation/assets/PCF_facture_standard.xlsx`.

---

## Autres entités (KAUR…)

Depuis la v0.2 (ADR015), toute entité utilise le gabarit standard `PCF_facture_standard.xlsx` : seules les informations de l'émetteur changent, et elles viennent de son MASTER. Si une entité veut une mise en page différente, on dépose `facturation/assets/PCF_facture_<ENTITE>.xlsx` (construit depuis une facture de référence avec `build_template.py`) et, si les cellules bougent, on déclare son plan dans `LAYOUTS` du module. Le gabarit propre est alors prioritaire.
