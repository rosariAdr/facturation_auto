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
| N° de contrat | `C15` + `D15:H15` | `apply_metadata(ref_contrat=…)` ; ligne vide sans contrat |
| N° SIRET | `C16` + `D16` | `apply_metadata` |
| N° de facture | `C17` + `D17` | `apply_metadata` |
| En-tête du tableau | ligne 19 | gabarit (fixe) |
| Lignes de cours | lignes 20-79 (60) | `place_sessions` |
| Total HT | ligne 80 | gabarit : `=SUM(Q20:R79)` ; libellé écrit par `apply_mentions` |
| Net à payer, TVA, déclaration d'activité | lignes 83-86 | `apply_mentions` (voir « Mentions et totaux ») |
| RIB | `D88` / `D89` | `apply_rib` |
| Échéance et mentions de paiement | lignes 91-94 | `apply_mentions` (voir « Mentions et totaux ») |

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

## Mentions et totaux (ADR030, F-24)

Depuis F-24, les textes du bas de facture sont écrits par le module (`apply_mentions`, appelée après `place_sessions`) : les anciens textes du gabarit sont remplacés à chaque génération. Les cellules suivent le décalage des lignes au-delà de 60 séances.

| Élément | Cellule | Contenu |
|---|---|---|
| Nom de l'émetteur | `C7` | Donnée du MASTER : nom suivi de « EI » ; sans cette mention, l'écriture est refusée |
| Nom du client | `M7` | Donnée du MASTER : dénomination juridique de la société qui paie |
| N° de contrat | `C15` + `D15:H15` | « N° de contrat : » et la valeur de la spec (`ref_contrat`) ; ligne vide sans contrat |
| Bon de commande | `C14` + `D14:H14` | Inchangé (`ref_bon_commande`) |
| Total | `O80` | « Total HT » |
| Montant dû | `C83` + `E83` | « Net à payer : » et une formule vers le total (décalée au-delà de 60 séances) |
| Ancienne astérisque | `C84` | Vidée |
| Mention de TVA | `C85:K85` | « TVA non applicable, article 293 B du Code général des impôts. » (fusion élargie pour ne pas couper la phrase ; citation du CGI admise jusqu'au 30/06/2028) |
| Déclaration d'activité | `C86` | Texte de `Coordonnées` / `Info` / `NDA`, recopié tel quel, sans fusion ; sans cette ligne, l'écriture est refusée |
| Échéance | `C91` | « Échéance : JJ/MM/AAAA » : la date de la spec (`echeance`), sinon date d'émission + 30 jours |
| Escompte | `C92` | « Pas d'escompte en cas de paiement anticipé. » |
| Pénalités | `C93:K93` | « Pénalités de retard : trois fois le taux d'intérêt légal, exigibles dès le lendemain de l'échéance, sans rappel préalable. » |
| Frais de recouvrement | `C94:K94` | « Indemnité forfaitaire pour frais de recouvrement en cas de retard de paiement : 40 €. » |

La zone d'impression va jusqu'à la ligne 94 (plus le décalage). Pas de ligne « TVA : 0 € » (ADR030). Une échéance antérieure à la date d'émission est refusée.

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

**Une dénomination longue est coupée dans l'encadré destinataire (`M7`)**, et le libellé « Bon de commande : » est tronqué en `C14` + `D14:H14`. Contournés à la main en octobre 2026 ; à corriger avec F-18.

**Au-delà de 60 séances, seules les mentions suivent le décalage des lignes.** `apply_rib` écrit toujours en `D88` / `D89`, et `post_mortem` / `verifier_facture` lisent le total et le RIB à leur place fixe : une telle facture est rejetée par le contrôle après écriture, ce qui bloque l'envoi au lieu de laisser passer une erreur. À corriger avec F-18.

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
| `apply_metadata(ws, n_facture, date_emission, *, entite, ref_bon_commande=None, ref_contrat=None)` | Titre, bon de commande, n° de contrat, SIRET, n° de facture |
| `place_sessions(ws, ecole_name, groups, *, entite, masquer_jour=False)` | Lignes de cours groupées par (promotion, libellé) |
| `apply_mentions(ws, *, entite, date_emission, echeance=None)` | Total HT, net à payer, TVA, déclaration d'activité, échéance, escompte, pénalités, indemnité de 40 € (ADR030) |
| `apply_rib(ws, *, entite)` | IBAN + BIC |
| `apply_print_setup(ws, *, entite)` | Portrait, ajusté à une page de large, zone d'impression |
| `post_mortem(ws, *, entite)` | Dictionnaire des cellules clés, dont échéance, NDA et contrat |
| `verifier_facture(ws, *, entite, n_facture, date_emission, dest, ecole_name, groups, echeance=None, ref_contrat=None)` | Liste des écarts entre la facture relue et les valeurs attendues |
| `calculer_echeance(date_emission, echeance=None)` | Échéance donnée, sinon émission + 30 jours ; refuse une échéance antérieure |
| `emetteur_porte_ei(nom)` | Le nom de l'émetteur porte-t-il « EI » ou « Entrepreneur individuel » ? |
| `get_nda(*, entite)` | Texte de `Coordonnées` / `Info` / `NDA` ; erreur s'il manque |

Ordre d'appel : `clone_template` → `apply_emetteur` → `apply_destinataire` → `apply_metadata` → `place_sessions` → `apply_mentions` → `apply_rib` → `apply_print_setup` → `post_mortem` / `verifier_facture`.

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
