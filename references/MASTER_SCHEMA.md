# MASTER_SCHEMA — Structure du classeur de données

**Dernière mise à jour** : 2026-11-27

> Le module ne contient aucune donnée : tout vient du MASTER de l'entité, lu au moment de l'exécution. Ce document décrit ce que le classeur doit contenir. Le MASTER est la source de vérité (ADR005) et le seul endroit où vivent les règles réelles des écoles (ADR026).
>
> Chaque onglet est marqué **obligatoire**, **de suivi** ou **lu s'il existe** (onglets ajoutés pour le lot 1). Les colonnes des onglets ajoutés pour le lot 1 (`Règles_facturation`, `Correspondances`, `Factures`) sont celles que l'outil lit ; elles restent `💡 [proposé — à valider]` tant qu'Adrian ne les a pas validées. Les exemples sont fictifs.

---

## Emplacement

```
<racine de données>/
├── ROSARI/
│   ├── ROSARI_MASTER.xlsx          ← source de vérité de l'entité
│   ├── ABC/ROSARI_ABC.xlsx         ← un classeur par école, un onglet par facture
│   └── ECO/ROSARI_ECO.xlsx
└── KAUR/
    └── KAUR_MASTER.xlsx
```

Le chemin du MASTER est toujours `<racine>/<ENTITE>/<ENTITE>_MASTER.xlsx`.

**Résolution de la racine**
- Par les scripts et la commande `facturation` : `--data-root` > `--profil` > variable `PCF_DATA_ROOT` > profil par défaut s'il existe > dossier courant.
- Par le module seul (`facturation.pcf_format`) : `set_data_root()` > `PCF_DATA_ROOT` > dossier courant.

---

## Onglets

| Onglet | Statut | Rôle |
|---|---|---|
| `Coordonnées` | Existant, **obligatoire** | Identités de l'émetteur et des écoles, SIRET, IBAN, BIC, mentions |
| `Sessions` | Existant, **obligatoire** | Journal des séances facturées (ADR003) ; lu par la préparation, l'audit et le registre |
| `Cours` | Existant (suivi) | Volumes contractuels, heures dispensées, facturées, à facturer |
| `Tarifs` | Existant (suivi) | Taux horaires par école, module et programme ; lu par la préparation et l'audit |
| `Écoles` | Existant (suivi) | Liste des écoles et code de facturation ; traduit un libellé d'école en code |
| `Devis` | Existant (suivi) | Devis et bons de commande reçus ; pas lu par l'outil |
| `Règles_facturation` | Lu s'il existe (F-14) | Règles propres à chaque école, une ligne par contrat |
| `Correspondances` | Lu s'il existe (F-12, F-13) | Reconnaissance des créneaux d'agenda et heures facturables ; écrit par F-16b (prévu) |
| `Factures` | Lu s'il existe ; écrit par `facturation registre` (F-25) | Registre des factures : échéance, envoi, paiement |
| `Conventions_horaires` | Obsolète | Remplacé par `Correspondances` (ADR016) |

`facturation profil verifier` signale une erreur si un onglet obligatoire manque, et un point d'attention si un onglet de suivi manque.

**Lecture par l'outil** (`facturation/master.py`) : les colonnes se lisent par leur nom (ligne 1), quel que soit leur ordre ; une colonne en plus est permise.
- Un onglet obligatoire absent, ou une colonne absente d'un onglet obligatoire ou d'un onglet lu s'il existe, fait refuser le MASTER : l'erreur nomme l'onglet et la colonne, et `profil verifier` la signale en ERREUR.
- Un onglet de suivi absent, ou auquel il manque une colonne exigée (ou qui la porte deux fois), est ignoré sans bloquer : `profil verifier` le signale en point d'attention. Colonnes exigées : `Tarifs` : `École`, `Module / Matière`, `Classe / Programme`, `Taux HT/h` ; `Écoles` : `École`, `Code facturation` ; `Cours` : ses dix colonnes.
- `profil verifier` signale aussi ce qu'un enregistrement du MASTER par l'outil perdrait (graphiques, images, validations, formules dont la valeur se lirait vide…), contrôle fait sur une copie temporaire.

### Ce que l'outil écrit dans le MASTER (27/11/2026)

| Commande | Écrit dans le MASTER |
|---|---|
| `facturation profil verifier`, `facturation agenda lire`, `scripts/check_coherence.py` | Rien |
| `facturation preparer`, `facturation apercu` | Rien : brouillons et aperçus sous `<racine>/<ENTITE>/_sorties/<AAAA-MM>/` |
| `facturation audit` | Rien : rapport sous `<racine>/<ENTITE>/_sorties/audits/` |
| `facturation registre` (sans action) | Rien : lecture de `Factures` et de `Sessions` |
| `facturation registre initialiser`, `envoyee`, `payee`, `note`, `manquant` | L'onglet `Factures`, et lui seul |
| `scripts/nouvelle_facture.py` | Rien : il écrit l'onglet de la facture dans le classeur de l'école |

Une écriture du registre suit toujours ce chemin : l'avant et l'après sont montrés, puis validés par « oui » (ou `--oui`) ; refus si le classeur est ouvert dans Excel ou LibreOffice ; verrou `<ENTITE>_MASTER.verrou` ; copie intacte dans `sauvegardes/<nom>_AAAAMMJJ-HHMMSS.xlsx` ; refus si l'enregistrement perdrait quelque chose (même contrôle que `profil verifier`), ou si le MASTER a changé depuis l'aperçu ; enregistrement atomique, puis relecture : seules les cellules annoncées ont changé ; événement `registre` dans `_sorties/journal.jsonl`. L'outil écrit des valeurs, jamais de formule, et n'écrase jamais une cellule qui en contient une.

Prévu : la ligne gardée de `Correspondances` (F-16b) ; les lignes de `Sessions` et la ligne `Factures` d'une facture émise (F-17) ; le suivi `Cours` (F-17b).

### `Coordonnées` — existant, obligatoire

La seule source des identités. Première ligne = en-têtes.

| Colonne | A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|---|
| | Type | Code | Nom / Raison sociale | Adresse L1 | Adresse L2 | Adresse L3 | Email | Téléphone |

Trois valeurs de `Type`, écrites exactement ainsi :

- **`Émetteur`** (avec l'accent) — une ligne, `Code` = l'entité (`ROSARI`). `Nom` = nom et prénom suivis de « EI » ou « Entrepreneur individuel » (ADR030, valable pour ROSARI comme pour KAUR) ; sans cette mention, l'outil refuse d'écrire. Alimente l'encadré émetteur.
- **`Destinataire`** — une ligne par école, `Code` = code de l'école. `Nom` = **dénomination juridique de la société qui paie**, telle qu'écrite au contrat (ADR030), pas le nom commercial. Une ligne sans nom ou sans adresse bloque la facture (étape 4 du PCF).
- **`Info`** — `Code` = l'entité, `Nom` porte le libellé et `Adresse L1` la valeur. `SIRET`, `IBAN` et `BIC` sont obligatoires : leur absence lève une erreur au chargement. `TVA` porte la mention de TVA. `NDA` porte le texte complet de la mention de déclaration d'activité, recopié tel quel sur la facture ; il est obligatoire depuis F-24 (sans lui, l'outil refuse d'écrire). Les lignes `Info` ne concernent que l'émetteur : on n'y range pas le SIRET d'une école.

`Adresse L3` sert au pays. L'encadré destinataire n'offre que 4 lignes : ville et pays sont réunis sur une seule (`00000 VILLE-EXEMPLE, France`).

### `Sessions` — existant, obligatoire

| A | B | C | D | E | F | G | H | I | J | K | L | M |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Date_log | École | Période_facture | N°_Facture | Date_séance | Début | Fin | Durée_h | Promotion | Module / UE | Taux | Total HT | Summary calendrier verbatim |

Journal d'audit : il justifie le contenu d'une facture après coup. Ajout seul, avec une exception (ADR003 amendé) : une facture régénérée remplace ses propres lignes, repérées par le `N°_Facture`, au lieu d'en ajouter. Aucune autre ligne n'est modifiée ni supprimée.

**Lecture par l'outil** (`facturation/sessions.py`, lecture seule) : la préparation y écarte les séances déjà facturées, l'audit y rapproche l'agenda, le registre y contrôle les totaux. Les formes historiques sont tolérées :
- `École` : le code, ou le libellé de la colonne `École` de l'onglet `Écoles` ;
- `Date_séance` : une date, ou un texte JJ/MM/AAAA ou AAAA-MM-JJ ;
- `Début` et `Fin` : une heure, ou un texte « 09:00 », « 9h », « 9h30 », « 09h30 » ; vides, la ligne est « sans heure ». Une ligne à une seule heure (début ou fin vide) compte comme « même jour sans heure » dans la recherche des doublons, jamais comme une correspondance exacte 💡 [proposé — à valider] ;
- `Durée_h`, `Taux`, `Total HT` : un nombre, ou un texte « 3,5 » ;
- `Période_facture` : texte libre lu tel quel (intervalles, « Mois-Mois AAAA », parenthèses) ; l'outil ne s'y fie pas pour repérer une période déjà facturée, ce que font les dates de séance et le `N°_Facture`.

Une cellule illisible ne fait jamais planter : la ligne est signalée avec son rang Excel, la colonne et la valeur. La préparation en fait un avertissement « ligne de Sessions à corriger », l'audit la rapporte (code 3) 💡 [proposé — à valider], et la reprise de l'historique du registre refuse de continuer.

Les lignes d'un N° au statut « abandonnée » dans `Factures` restent dans `Sessions` (piste d'audit), mais l'anti-doublon et l'audit les ignorent. Le statut se reconnaît sans tenir compte de la casse ni de l'accent (« abandonnee ») 💡 [proposé — à valider].

### `Cours` — existant (suivi)

| A | B | C | D | E | F | G | H | I | J |
|---|---|---|---|---|---|---|---|---|---|
| École | Module / Matière | Classe / Groupe | Volume contractuel (h) | Heures dispensées (h) | Heures facturées (h) | Heures à facturer (h) | N° Facture liée | N° Devis lié | Statut |

Invariant vérifié par `scripts/check_coherence.py` : `Heures dispensées = Heures facturées + Heures à facturer`. Le script signale aussi des heures facturées au-delà des heures dispensées.

### `Tarifs` — existant (suivi)

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| École | Module / Matière | Classe / Programme | Taux HT/h | Notes *(facultative)* | Réf. contrat *(facultative)* |

Le **contrat signé fait foi** pour un taux (ADR026) : la colonne `Réf. contrat` garde sa référence. Toute autre source de taux s'aligne sur cet onglet.

**Lecture par l'outil** (`facturation/heures.py`) :
- une ligne se désigne par sa clé tarif `ÉCOLE|Module|Classe` (colonne `Clé tarif` de `Correspondances`), chaque segment comparé sans tenir compte de la casse ni des espaces autour, sans repli sur une ligne voisine ;
- `École` porte le code, ou le libellé de l'onglet `Écoles` ;
- `Taux HT/h` se lit en nombre, ou en texte de la forme « 60 », « 60 € », « 60 €/h », « 60,5 €/h » ou « 60.50 » (espaces, même insécables, tolérés). Une cellule vide, un taux nul ou négatif, ou tout autre texte donnent une erreur qui nomme la clé : le taux n'est jamais inventé, et la facture préparée est bloquée ;
- plusieurs lignes de même clé ne sont acceptées que si elles portent le même taux et la même `Réf. contrat` ;
- `Réf. contrat` relie la ligne de tarif à un contrat de `Règles_facturation` ; le n° imprimé sur la facture vient de la ligne de `Règles_facturation`.

### `Écoles` — existant (suivi)

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| École | Code facturation | Entité légale | Contact | Email contact | Statut mission |

Le `Code facturation` relie une école à sa ligne `Destinataire` dans `Coordonnées`. Code d'une nouvelle école : 3 lettres, unique dans l'entité (ADR029).

L'outil lit `École` et `Code facturation` pour traduire un libellé d'école en code, dans `Sessions`, `Tarifs`, `Correspondances` et `Factures` ; un libellé inconnu de `Écoles` (qui n'a pas la forme d'un code école), ou porté par deux codes, donne une erreur qui le cite. Le libellé `École` est aussi celui qui s'imprime sur la facture préparée ; sans lui, le code est imprimé, avec un avertissement.

### `Devis` — existant (suivi)

| A | B | C | D | E | F | G | H |
|---|---|---|---|---|---|---|---|
| École | Module | N° Devis | Date émission | Montant HT | Statut | PO reçu | N° Facture générée |

### `Règles_facturation` — lu s'il existe (F-14, F-16a, F-19) · `💡 [proposé — à valider]`

**Une ligne par contrat.** La clé d'une ligne est le couple (`École`, `N° de contrat`) : une école peut avoir plusieurs lignes, à condition que chacune porte un `N° de contrat` distinct, comparé sans tenir compte de la casse. Deux lignes de la même école sans n°, une ligne sans n° à côté d'une ligne avec n°, ou le même n° deux fois sont refusés. Le n° de contrat imprimé sur la facture vient de cette ligne ; `Tarifs.Réf. contrat` y relie chaque ligne de tarif.

| Colonne | Contenu |
|---|---|
| École | Code de l'école |
| Statut lot | `Lot 1`, `Plus tard` ou `Hors facturation`, porté par chaque ligne (chaque contrat) : le périmètre d'un lot est lu ici ; vide, la ligne n'est dans aucun lot |
| Rythme | `mensuel` (défaut), `fin de module`, ou `cycle`. Une école en `fin de module` n'est pas préparée par l'outil : elle est listée « à préparer à la main » 💡 [proposé — à valider] |
| Début de cycle | Jour du mois où commence le cycle de l'école, de 2 à 28, exigé si `cycle` |
| Date limite d'envoi | Jour du mois et heure limites fixés par l'école, s'il y en a (texte libre, ex. « 15 à 12 h ») |
| Délai de paiement | Délai du contrat ou de l'école : vide = date d'émission + 30 jours (ADR030) ; `N jours` ; `fin de mois` ; `N jours fin de mois` = date d'émission + N jours, puis fin de ce mois-là 💡 [proposé — à valider] ; N au plus 365 |
| N° de contrat | Référence affichée sur la facture ; distincte pour chaque ligne d'une même école |
| Une facture par contrat | `oui` / `non` (vide = non) |
| Bon de commande requis | `oui` / `non` (vide = non). Avec `oui`, la facture préparée est bloquée tant que la référence du bon de commande n'est pas saisie (décision `bon_commande`, F-16b) ; une facture par ligne de bon de commande se fait par un script dédié, sur le modèle de `examples/factures_par_bon_de_commande.py` 💡 [proposé — à valider] |
| Colonne Jour | `affichée` (défaut) / `masquée` |
| Regroupement | `séance` (défaut) ou `jour` (une ligne par jour et par cours) |
| Canal d'envoi | Mail, portail fournisseur, plateforme de l'école, courrier |
| Pièces jointes | Ex. `PDF`, `PDF + récapitulatif` |
| Notes | Mentions particulières demandées par l'école |

**Lecture par l'outil** (`facturation/regles.py`) : textes lus sans les espaces autour, `Statut lot`, `Rythme`, `Colonne Jour` et `Regroupement` sans tenir compte de la casse ; `oui` / `non` ou un booléen Excel. Une valeur illisible donne une erreur qui nomme l'école, la colonne et la valeur, et toutes les erreurs de l'onglet sont rapportées ensemble : une règle n'est jamais devinée. Seule exception : la ligne illisible d'une école `Plus tard` ou `Hors facturation` est écartée avec un avertissement, sans retenir les autres ; l'audit garde alors cette école hors facturation, en lisant son `Statut lot` sur la ligne telle quelle 💡 [proposé — à valider].

**Préparation** (`facturation preparer`) : une école à plusieurs lignes est facturée par contrat, et bloquée si l'une de ses lignes ne dit pas `Une facture par contrat = oui` ; une `Réf. contrat` de `Tarifs` sans ligne de règles est un blocage ; sans `Une facture par contrat`, une `Réf. contrat` différente du `N° de contrat` de la règle aussi. La date d'émission est le 1er du mois suivant pour toutes les écoles, y compris celles à cycle, qui reçoivent un avertissement (ADR027).

### `Correspondances` — lu s'il existe (F-12, F-13) · `💡 [proposé — à valider]`

| Colonne | Contenu |
|---|---|
| Clé | Code d'événement (`#ABC-M1-STATS`) ou intitulé d'événement |
| Type de clé | `code`, `exact` ou `préfixe` (ordre de recherche, ADR016) |
| École | Code de l'école |
| Module | Module ou matière |
| Promotion | Classe ou groupe |
| Clé tarif | Ligne de `Tarifs` à appliquer |
| Mode heures | `agenda` (défaut), `coefficient` ou `forfait` |
| Valeur | Coefficient (ex. `0.875` pour 7/8) ou nombre d'heures du forfait |
| Gardée le | Date à laquelle Adrian a gardé cette règle (ADR028) |
| Notes | Justification |

La colonne `École` porte le code, ou le libellé de l'onglet `Écoles`. En attendant F-16b, une règle se garde en ajoutant la ligne à la main, puis en relançant `facturation preparer`.

### `Factures` — lu par l'outil, écrit par `facturation registre` (F-25) · `💡 [proposé — à valider]`

Le registre : une ligne par facture, et une ligne par numéro manquant. L'outil crée l'onglet s'il manque (`facturation registre initialiser`), en fin de classeur, avec ces 11 en-têtes en ligne 1 et les volets figés en A2. Les colonnes se lisent par leur nom (`facturation/registre.py`).

| Colonne | Contenu | Lecture et écriture par l'outil |
|---|---|---|
| N° facture | Ex. `R_ABC008` | Obligatoire sur une ligne remplie, une seule ligne par N°. Deux N° qui ne diffèrent que par la casse, des espaces, un caractère invisible ou le tiret bas avant le rang (`R_ABC_008`) désignent la même facture : un double est refusé |
| École | Code de l'école | Le code, ou le libellé de l'onglet `Écoles` |
| Contrat | N° de contrat, s'il existe | Texte |
| Période | Ex. `Déc 2026` ou `Décembre 2026` | Texte. La préparation y repère une période déjà facturée hors outil (même école, même période, même contrat), écrite en abrégé, en toutes lettres, ou saisie en date (Excel convertit « Déc 2026 » en 01/12/2026) |
| Date d'émission | Le 1er du mois suivant (ADR027) | Date Excel, ou texte JJ/MM/AAAA ou AAAA-MM-JJ ; écrite en date, au format JJ/MM/AAAA |
| Échéance | Calculée selon `Règles_facturation` (ADR030) | Comme la date d'émission. Le jour de l'échéance n'est pas échu : la facture l'est le lendemain |
| Total HT | Montant de la facture | Un nombre, ou un texte « 1 234,56 » ; « 1 234,56 € », avec le signe, est refusé 💡 [proposé — à valider]. Écrit en nombre, arrondi au centime |
| Date d'envoi | Saisie par Adrian (le mail envoyé fait foi) | Date ; `facturation registre envoyee N° --le DATE` |
| Date de paiement | Saisie par Adrian (ADR031) | Date ; `facturation registre payee N° --le DATE` |
| Statut | `en émission`, `émise`, `envoyée`, `payée`, `abandonnée` | Lu avec ou sans accents ; « échue », toléré, se lit comme un statut vide ; tout autre texte est refusé |
| Note | Ex. raison d'un numéro manquant (ADR029) | Texte ; `facturation registre note N° "…"` |

Une cellule illisible (date qui n'en est pas une, montant, statut inconnu) arrête la lecture du registre avec un message qui nomme la ligne et la colonne, par exemple « onglet « Factures », ligne 7, colonne « Date de paiement » : « fin nov » n'est pas une date (JJ/MM/AAAA) ».

**Statuts** 💡 [proposé — à valider] :
- **saisis** (écrits dans la colonne) : `en émission` (posé par l'émission, F-17, prévu), `émise`, `envoyée`, `payée`, `abandonnée`. `abandonnée` ne s'écrit que sur une facture restée en émission, sans date d'envoi ni de paiement (abandon prévu avec F-17 ; une facture émise se corrige par un avoir), ou sur la ligne d'un numéro manquant (`facturation registre manquant`) ; son numéro n'est jamais réutilisé ;
- **affichés** (tableau de `facturation registre`), dans cet ordre : `abandonnée` ; date de paiement → `payée` ; `en émission` (jamais échue) ; échéance dépassée → `échue` ; date d'envoi → `envoyée` ; date d'émission ou statut `émise` → `émise` ; sinon `à compléter`. `échue` et `à compléter` sont calculés, jamais écrits.

**Ce que chaque action écrit** (après l'aperçu et « oui ») :
- `initialiser` : l'onglet et ses en-têtes. Avec `--reprendre-historique`, une ligne par N° de `Sessions` absent du registre : école, période (`Période_facture`), total (somme des `Total HT`, arrondie une fois), statut `émise`, note « repris de Sessions le JJ/MM/AAAA : dates à compléter ». La date d'émission et l'échéance restent vides : l'outil ne les invente jamais ;
- `envoyee`, `payee` : la date ; une date d'envoi fait passer `émise` à `envoyée`, une date de paiement fait passer à `payée`. Refusé sur une facture `en émission` ou `abandonnée`, ou pour une date antérieure à l'émission ; une date déjà saisie ne se remplace qu'avec `--remplacer` ;
- `note` : la note (seule colonne qui change encore sur une facture abandonnée) ; une note déjà saisie ne se remplace qu'avec `--remplacer` ;
- `manquant` : une ligne avec le N°, l'école déduite du N°, le statut `abandonnée` et la raison en note, sans date ni montant (ADR029). Seulement pour un numéro qui manque à la série (un trou, ou le numéro qui suit le plus grand), jamais pour un N° présent dans `Sessions`.

**Anomalies** signalées par `facturation registre` : N° de `Sessions` absent du registre ; date d'envoi, de paiement ou échéance antérieure à l'émission ; paiement sans envoi ; `Total HT` différent de la somme des `Total HT` de `Sessions` pour ce N° (au centime) ; facture non abandonnée sans ligne dans `Sessions` ; statut incohérent avec les dates, dont une facture `en émission` ou `abandonnée` qui porte une date d'envoi ou de paiement 💡 [proposé — à valider].

---

## Classeurs de factures et numérotation

Un classeur par école, un **onglet par facture**, nommé `<N° facture> <Période>` (ex. `R_ABC009 Juil 2026`). Le préfixe `PAID` ajouté au nom d'un onglet réglé est un usage de lecture, sans valeur de preuve (ADR031).

Numérotation : `<préfixe d'entité>_<CODE ÉCOLE><NNN>` (ADR002, ADR029), par exemple `R_ABC009`.
- Le numéro suivant est le plus grand numéro de la série plus un ; il se lit dans `Sessions` (colonne `N°_Facture`) et dans `Factures` (colonne `N° facture`), s'il existe.
- Un numéro manquant n'est jamais réutilisé ; sa raison est notée (`facturation registre manquant`, qui l'inscrit dans `Factures`).
- Un numéro n'est attribué qu'à l'émission.

---

## Créer un MASTER de zéro

1. Classeur avec les onglets existants ci-dessus (`Coordonnées`, `Sessions`, `Cours`, `Tarifs`, `Écoles`, `Devis`), en-têtes en ligne 1, dans cet ordre de colonnes.
2. Remplir `Coordonnées` : une ligne `Émetteur` (nom suivi de « EI »), une ligne `Info` par valeur (`SIRET`, `IBAN`, `BIC`, `TVA`, `NDA`), une ligne `Destinataire` par école.
3. Pour préparer les factures avec l'outil (lot 1) : ajouter `Règles_facturation` et `Correspondances`, avec les colonnes ci-dessus ; créer `Factures` avec `facturation registre initialiser`.
4. Vérifier : `python scripts/check_coherence.py --entite <ENTITE>`.
5. Valider l'installation : `facturation profil verifier`, qui signale un émetteur, un SIRET, un IBAN, un BIC ou un destinataire incomplet, un onglet de suivi ignoré et ce qu'un enregistrement par l'outil perdrait.

Un exemple complet, avec une entité et une école fictives : [RECETTE_DOC.md](RECETTE_DOC.md), étape 3.
