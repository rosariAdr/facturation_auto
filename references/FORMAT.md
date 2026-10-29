# FORMAT — Gabarit canonique et invariants de mise en forme

**Dernière mise à jour** : 2026-09-24

> Source unique du format : `facturation/assets/PCF_facture_standard.xlsx > MODELE` (ou
> `PCF_facture_<ENTITE>.xlsx` si l'entité a son propre gabarit) + `facturation/pcf_format.py`.
> Toute spec visuelle écrite ailleurs est obsolète. Faire évoluer le format = modifier
> le gabarit **et** adapter le module, jamais l'un sans l'autre.

---

## Principe

Le gabarit porte **la mise en forme uniquement** : bandeau, encadrés, bordures,
largeurs, hauteurs, formules. Il ne contient **aucune donnée personnelle** — ni
nom, ni adresse, ni SIRET, ni IBAN. Ces valeurs sont injectées au runtime depuis
`MASTER.Coordonnées` par `apply_emetteur`, `apply_metadata` et `apply_rib`.

Conséquence directe : le gabarit peut être versionné et partagé sans exposer quoi
que ce soit, et changer de RIB ne demande pas de régénérer le gabarit.

---

## Plan du gabarit ROSARI (onglet `MODELE`)

| Zone | Cellules | Rempli par |
|---|---|---|
| Bandeau titre | `B2:S3` | `apply_metadata` (titre + date d'émission) |
| Encadré émetteur | `C7:E11` | `apply_emetteur` (5 lignes) |
| Encadré destinataire | `M7:P10` | `apply_destinataire` (4 lignes) |
| Bon de commande | `C14` + `D14:H14` | `apply_metadata(ref_bon_commande=…)` |
| N° SIRET | `C16` + `D16` | `apply_metadata` |
| N° de Facture | `C17` + `D17` | `apply_metadata` |
| En-tête tableau | ligne 19 | gabarit (fixe) |
| Lignes de cours | lignes 20-79 (60) | `place_sessions` |
| TOTAL | ligne 80 | gabarit : `=SUM(Q20:R79)` |
| À payer / TVA | lignes 83-85 | gabarit |
| RIB | `D88` / `D89` | `apply_rib` |
| Mentions légales | lignes 91-93 | gabarit |

### Colonnes du tableau de cours

| Colonne | Fusion | Contenu |
|---|---|---|
| C | C:D | ECOLE (fusionnée verticalement sur les lignes utilisées) |
| E | E:F | Promotion |
| G | G:J | Libellé de cours / groupe |
| K | K:L | Jour — masquable via `place_sessions(masquer_jour=True)` |
| M | M:N | Nb heures |
| O | O:P | Taux horaire HT |
| Q | Q:R | Total — formule `=O{row}*M{row}` |

---

## Invariants

- **Ne jamais utiliser `delete_rows()`** — casse les plages fusionnées en dessous
  (RIB, mentions). Pour retirer une ligne inutilisée : `row_dimensions[r].hidden = True`
  et vider le contenu.
- `insert_rows()` uniquement pour étendre au-delà de 60 sessions, avec démerge préalable.
- La formule TOTAL reste `=SUM(Q20:R79)` quelles que soient les lignes masquées
  (une ligne masquée vide compte pour 0).
- Hauteur des lignes de cours : 46.5 pt, `wrap_text` activé (les libellés de bon de
  commande sont longs).
- Toute ligne portant du texte en 18 pt doit avoir `row_dimensions[r].height = None`
  (hauteur automatique) : le gabarit hérite de hauteurs de 14.25 pt qui tronquent.

---

## Pièges connus (vérifiés)

**La fusion `M7:P8` avale la première ligne d'adresse.** Le nom du destinataire
occupe deux lignes dans le gabarit. Écrire naïvement les lignes d'adresse en
`M7`, `M8`, `M9`, `M10` fait tomber la rue dans une `MergedCell`, où l'écriture
est ignorée **sans erreur** — la facture part sans son adresse postale.
`apply_destinataire` démerge donc la zone et refusionne une plage par ligne.

**Les bordures sont assignables sur une `MergedCell`, pas les valeurs.** Pour
encadrer une zone fusionnée, il faut parcourir toutes ses cellules : ne styler
que la cellule d'ancrage laisse le cadre droit manquant. C'est ce que fait
`_frame_box`.

**Le gabarit porte des bordures résiduelles au milieu des encadrés**, hérité de
l'époque où les zones faisaient 3 lignes. `_frame_box` les nettoie à chaque
génération : cadre extérieur seul, aucun trait horizontal interne.

**Sans `apply_print_setup`, l'export PDF coupe les colonnes Taux et Total** sur
une seconde page : 18 colonnes ne tiennent pas en portrait par défaut. La
fonction force portrait + ajustement à une page de large.

---

## API du module

| Fonction | Rôle |
|---|---|
| `set_data_root(chemin)` / `get_data_root()` | Racine des données métier |
| `clone_template(dst_wb, sheet_name, entite)` | Clone `MODELE` dans un classeur cible |
| `apply_emetteur(ws, entite)` | Émetteur, lu depuis `MASTER.Coordonnées` |
| `apply_destinataire(ws, dest, entite)` | Destinataire (`None` = encadré vide) |
| `apply_metadata(ws, n_facture, date_emission, entite, ref_bon_commande=None)` | Titre, bon de commande, SIRET, n° facture |
| `place_sessions(ws, ecole_name, groups, entite, masquer_jour=False)` | Lignes de cours groupées par (promotion, libellé) |
| `apply_rib(ws, entite)` | IBAN + BIC |
| `apply_print_setup(ws, entite)` | Portrait, ajusté à 1 page de large, zone d'impression |
| `post_mortem(ws, entite)` | Dict des cellules clés, pour le contrôle d'étape 7 |
| `get_destinataire(code_ecole, entite)` | Destinataire depuis `MASTER.Coordonnées` |

Ordre d'appel : `clone_template` → `apply_emetteur` → `apply_destinataire` →
`apply_metadata` → `place_sessions` → `apply_rib` → `apply_print_setup` → `post_mortem`.

---

## Reconstruire le gabarit

Rarement nécessaire. Le gabarit dérive d'une facture réelle servant de référence
visuelle :

```bash
python scripts/build_template.py --source "<...>/ROSARI_DOR.xlsx" --onglet "DOR001 012026"
```

Le script étend la zone de cours à 60 lignes, trace les encadrés, vide toutes les
données et neutralise les coordonnées avant d'écrire `facturation/assets/PCF_facture_standard.xlsx`.

---

## Autres entités (KAUR…)

Depuis la v0.2 (ADR015), toute entité utilise le gabarit standard
`PCF_facture_standard.xlsx` : seules les informations de l'émetteur changent, et
elles viennent de son MASTER. Si une entité veut une mise en page différente, on
dépose `facturation/assets/PCF_facture_<ENTITE>.xlsx` (construit depuis une
facture de référence avec `build_template.py`) et, si les cellules bougent, on
déclare son plan dans `LAYOUTS` du module. Le gabarit propre est alors prioritaire.
