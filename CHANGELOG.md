# Changelog

Format : [Keep a Changelog](https://keepachangelog.com/fr/1.1.0/). Versions : `0.x` tant que la v2 n'est pas en service.

## [Unreleased]

### Ajouté (vague 2 du lot 1)
- F-12 : classement des créneaux de l'agenda (`facturation/classement.py`) : code d'événement `#ÉCOLE-PROMO-MODULE`, puis intitulé exact, puis plus long préfixe ; sinon file à classer (« non reconnue », « ambiguë », « code inconnu »), jamais facturée par défaut ; ligne de Correspondances à garder au choix d'Adrian (ADR028).
- F-13 : heures facturables (`facturation/heures.py`) : agenda, coefficient ou forfait ; lecture du taux d'une ligne de Tarifs sans jamais l'inventer.
- F-14 : règles de facturation par école (`facturation/regles.py`) : rythme, période à facturer (mois ou cycle), échéance selon le délai de paiement, écoles d'un lot, blocages (bon de commande requis).

### Ajouté (vague 1 du lot 1)
- F-10 : lecture de l'agenda des cours par son adresse iCal (`facturation/agenda.py`) : récurrences, exceptions, occurrences déplacées, événements annulés et journées entières exclus, durée exacte au changement d'heure ; commande `facturation agenda lire --mois AAAA-MM [--fichier]` ; l'adresse n'apparaît jamais dans un message. Dépendances : `icalendar`, `recurring-ical-events`.
- F-11 : lecture contrôlée du MASTER (onglets et colonnes nommés dans l'erreur) et écriture sûre (`facturation/master.py`) : sauvegarde horodatée avant écriture, verrou par poste (refus si Excel ou LibreOffice a le classeur ouvert, ou si une autre instance écrit), verrou retiré même en cas d'erreur.
- F-15 : numérotation (`facturation/numerotation.py`) : numéro suivant, numéros manquants, codes école (3 lettres pour une nouvelle école), lecture des numéros du MASTER.

### Modifié
- `profil verifier` lit `agenda.secret` avec ou sans BOM et signale en ERREUR une adresse mal formée ou un fichier illisible.

### Corrigé (F-18, fiabilisation du gabarit)
- Formats de nombre sur toutes les lignes de cours ; facture d'une page jusqu'à 18 séances, sauts de page maîtrisés au-delà, bloc du bas jamais coupé ; bas de facture juste au-delà de 60 séances ; libellés et dénomination du client entiers ; traits parasites et bandeau corrigés ; mentions lisibles en Calibri ; contrôle après écriture étendu aux textes fixes et à la formule du net à payer.

### Corrigé (clonage du gabarit)
- Clonage du gabarit complet : protection des cellules, préfixe texte, lignes et colonnes masquées ou groupées (largeur reprise sur chaque colonne d'un groupe), mise en page, marges, options d'impression, en-têtes et pieds de page, protection de la feuille, affichage, volets figés, zone et titres d'impression sont désormais repris. La sélection d'onglet ne l'est pas, pour qu'Excel ne groupe pas la facture avec l'onglet actif de l'école. Fidélité vérifiée cellule par cellule et propriété par propriété, en mémoire (T-FMT-08) et après enregistrement (T-REC-01).
- Une fusion du gabarit qui échoue, ou qu'openpyxl ignore sans erreur, arrête la génération avec un message qui nomme la plage, au lieu d'être passée sous silence. Ce que le clonage ne sait pas recopier (mise en forme conditionnelle, validations, images, tableaux, commentaires, styles nommés autres que « Normal ») provoque aussi une erreur explicite.
- Le gabarit n'est lu sur le disque qu'une fois par processus tant qu'il ne change pas : les factures suivantes d'un même lancement sont clonées 3 à 5 fois plus vite que la première.
- openpyxl épinglé en 3.1.x : la voie rapide du clonage repose sur une structure interne, dont le contrat est vérifié par T-CLO-01. Si elle change malgré tout, la copie repasse d'elle-même par la voie publique, plus lente mais exacte (T-CLO-04).
- Tests : clonage du gabarit (T-CLO-01 à 11) et recette d'une facture complète relue depuis le disque (T-REC-01 à 03, marqueur `recette`).

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

### Ajouté (P-0b, socle du lot 1)
- `facturation/fichiers.py` : écritures atomiques (texte, JSON, classeur), journaux JSON Lines, empreintes, masquage des adresses ; `facturation/chemins.py` : mois, noms d'onglet (Janv… Déc), arborescence des sorties, racine refusée dans un dépôt ; `facturation/sessions.py` : lecture typée de `Sessions` (heures « 9h », école par code ou libellé) ; `facturation/journal.py` : schéma figé des journaux, état de validation, pièces d'émission.
- `master.py` : onglets de suivi lus sans bloquer, rangs Excel, verrou seul, classeur d'école ouvert, contrôle de fidélité d'un enregistrement (`fidelite_aller_retour`) ; `heures.py` : taux en texte (« 60 €/h »), `ligne_tarif` et `Réf. contrat`, montants exacts ; `regles.py` : une ligne par contrat (Q-B1) ; `profil verifier` signale les onglets de suivi illisibles et les pertes d'un enregistrement.
- Dépendance `Pillow`, pour que les images (logo) d'un classeur survivent à l'enregistrement.

### Corrigé (revue de P-0b)
- Une ligne de `Sessions` à une seule heure n'échappe plus à la recherche des doublons ; un n° abandonné saisi en minuscules ou au statut « abandonnee » est reconnu.
- Un ajout au journal après une écriture interrompue ne le rend plus illisible ; un journal avec BOM se lit ; une sauvegarde n'apparaît sous son nom qu'une fois complète ; le classeur est mis sur disque avant de remplacer l'original.
- Une racine écrite avec « .. », ou atteinte par un lien, est jugée sur son vrai chemin ; un mois en chiffres non ASCII est refusé.
- Une pièce d'émission abîmée ne conseille plus de la supprimer ; une clé tarif peut nommer l'école par son libellé ; `arrondi_centime` refuse un `float` ; deux n° de contrat qui ne diffèrent que par la casse sont refusés ; un MASTER refusé par l'outil est une ERREUR de `profil verifier`.

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
- Clonage du gabarit plus rapide : chacun des ~60 styles distincts n'est posé qu'une fois au lieu de l'être pour chacune des ~26 800 cellules stylées. Écriture d'une facture : 5,4 s → 1,3 s sur la machine de référence.
- `pcf_format.py` déplacé dans le paquet `facturation/` ; entités non limitées à ROSARI/KAUR ; gabarit standard commun `PCF_facture_standard.xlsx` (ADR015), gabarit propre à une entité prioritaire s'il existe.
- Scripts `nouvelle_facture.py` et `check_coherence.py` : option `--profil`, racine résolue depuis le profil.

### Retiré
- `scripts/test_pcf_format.py`, remplacé par `tests/regression/test_format.py`.
- `requirements.txt`, remplacé par `pyproject.toml`.

### Connu
- T-FMT-07 (xfail strict) : formats perdus après la 7e séance, mentions légales en page 2 — correction prévue en F-18 (lot 1).

## [0.1.0] — 24/09/2026

Import de la v1 : module de format, CLI spec JSON, contrôles du MASTER, gabarit ROSARI, protocole PCF, ADR001 à ADR011.
