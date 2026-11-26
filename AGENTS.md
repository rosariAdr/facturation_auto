# AGENTS.md — Conventions de travail (facturation_auto)

> **Source de vérité unique des conventions de ce dépôt.** Lue nativement par **Codex**, et par **Claude Code** via `CLAUDE.md`, qui l'importe (`@AGENTS.md`).
> **Socle commun** — les §0 à §6 sont le socle partagé par tous les projets : même texte partout, à quelques valeurs près (nom de la branche stable, contenu de la commande `porte`, outils de test), toutes rappelées en §7. La §7 est propre à ce dépôt.
> En cas de conflit entre ce fichier et le code : **s'arrêter et signaler**, ne jamais trancher en silence.

---

## 0. Règles pour les agents (Claude & Codex)

- **AGENTS.md fait foi.** Si une info manque (une règle, un choix d'archi, une valeur métier), **demander — ne jamais inventer**. Toute proposition de comblement est taguée `💡 [proposé — à valider]` dans la doc ou la PR.
- **Un agent = un ticket = une branche à la fois.** Deux agents ne travaillent jamais sur les mêmes fichiers en parallèle.
- **Revue croisée.** Faire relire le diff d'un agent par l'autre (ex. `/code-review` côté Claude) avant d'ouvrir la PR.
- **Commits : l'agent propose, le dev commite.** L'agent rédige le message de commit ; le dev le valide et commite lui-même (GitKraken). **Aucun footer `Co-Authored-By`.**
- **Jamais** de commit direct sur `main`/`dev`, **jamais** de `push`/`merge`/`tag` par un agent sans demande explicite.

### Plusieurs agents en parallèle

- **Le brief écrit le périmètre** : les fichiers que l'agent possède, et ceux auxquels il ne touche pas.
- **Un port de test par copie.** Chaque copie du dépôt (worktree) qui lance les tests fixe son port (`PORT_E2E=3101`, `3102`…). Sinon `reuseExistingServer` fait lire le serveur d'une autre copie, et la porte voit des tests verts sur un code qui n'est pas le sien.
- **Aucun relevé depuis une branche** (nombre de tests, couverture, compteurs) : il est périmé dès la fusion suivante. Les relevés se font une fois, sur `dev`, après la dernière fusion.
- **Fichiers partagés** (CHANGELOG, journal) : ajouter en fin de section ; un conflit se résout en gardant les deux côtés.
- **Numérotation partagée** (tickets, ADR) : le brief annonce le dernier numéro attribué, et on vérifie qu'un ID est libre avant de le poser. En cas de collision, celui qui est déjà fusionné garde son numéro.

---

## 1. Workflow Git

```
main  ← stable, taggué (vX.Y.Z) — PR depuis dev uniquement
  └── dev  ← intégration continue — PR depuis les branches de travail
        ├── feat/<TICKET_ID>      (ex. feat/F-18-gabarit)
        ├── fix/<description>     (ex. fix/arrondi-sessions)
        └── chore/<description>   (maintenance, docs, deps, outillage)
```

- Pas de commit direct sur `main` ni `dev` → **tout passe par PR**.
- Une PR = un ticket (ou un lot cohérent de petits tickets liés).
- Branches de travail → `dev` : **squash merge**. `dev → main` : **merge classique** (préserve l'historique). Tag à chaque release.
- **Resync obligatoire après chaque `dev → main`** :
  ```bash
  git checkout dev && git merge --ff-only main && git push origin dev
  ```
  Check qui doit passer : `git log --oneline main..dev` **vide**.
- `wip:` autorisé pour un commit de session interrompue sur une branche de travail (squashé à la PR).

---

## 2. Conventional Commits

Format : `type(scope): description courte`

| Type | Quand l'utiliser |
|---|---|
| `feat` | nouvelle fonctionnalité visible |
| `fix` | correction de bug |
| `test` | ajout/modif de tests sans changement de code prod |
| `refactor` | restructuration sans changement de comportement |
| `chore` | maintenance, deps, scripts, config |
| `docs` | documentation seule |
| `style` | formatage (rien sur la logique) |
| `perf` | optimisation de performance |

Les **scopes** et la **langue des messages** sont propres au projet → voir §7.

---

## 3. Definition of Done (par type de ticket)

Vérifier **tous** les critères du type avant de fermer un ticket.

- **Feature (logique)** — code au pattern existant · **test unitaire** (logique pure) · **test fonctionnel** (action d'état) · **scénario** si la feature entre dans une boucle · **porte verte** (§5) · doc d'état à jour.
- **Feature (UI)** — respecte le design system / tokens · **smoke test de rendu** · pas de régression sur les écrans voisins · accessibilité minimum (`alt`, contraste) · capture si pertinent.
- **Bug fix** — **test de régression écrit AVANT le fix (TDD strict)** : il échoue d'abord, c'est ce qui prouve qu'il garde (§5) · tests verts après · stack trace originale notée dans le commit si pertinent.
- **Refactor** — tests existants verts (zéro changement de comportement) · pas de nouvelle dépendance injustifiée · taille du livrable stable (±5 %).
- **Data / contenu** — **test de structure** (champs requis, IDs uniques, références valides) · nommage cohérent (§7).
- **Docs / process** — markdown lisible · liens internes valides · CHANGELOG mis à jour si impact workflow.

---

## 4. Process de Pull Request

**Branche de travail → `dev`**
1. **Porte verte** (§5) sur la branche.
2. Push la branche (`git push -u origin <branche>`).
3. PR titrée **comme un commit** (`type(scope): …`).
4. Description = lien vers le ticket + changements clés + **comment tester** (étapes) + captures si UI. Une branche apporte ses tests, ou la PR dit pourquoi elle n'en a pas besoin.
5. Revue croisée (`/code-review` côté Claude) + auto-review du diff → **squash merge**.

**`dev → main` (release)**
1. `dev` porte une suite de features cohérente (pas un mix incomplet).
2. **🚦 Gate de release obligatoire** : porte verte sur `dev` + recette E2E du projet (§7) — sine qua non.
3. CHANGELOG : `[Unreleased] → [vX.Y.Z] - AAAA-MM-JJ`.
4. PR `dev → main` → **merge classique** → `git tag vX.Y.Z && git push --tags` → **resync `dev`** (§1).

---

## 5. Tests

**Pyramide** (unitaire → fonctionnel → scénarios → UI) :
- **Unitaire** — logique pure, testable sans mock (à côté du source).
- **Fonctionnel** — actions d'état / API (fichiers dédiés par domaine).
- **Scénarios** — multi-actions qui simulent un parcours réel.
- **UI** — rendu + navigation (smoke test obligatoire sur les flux critiques).

**Couverture visée** : logique pure ~90 % · le reste ≥ 80 % (au moins un test par action ou fonction publique) · UI = smoke + flux critiques.

**La porte — une seule commande avant toute fusion.**
- `porte` enchaîne tout ce qui doit être vert (lint/format, typage s'il y en a, tests, build, e2e) et **s'arrête à la première erreur**. Commande exacte : §7.
- On fusionne **parce que la porte est verte**, pas parce que le travail est fini — même pour un petit changement.
- Si un test échoue, on corrige le code — **jamais le test**, sauf à écrire pourquoi le test avait tort.
- La même commande tourne en local, dans le hook et en CI (§6) : ce qui est vert ici est vert là-bas.

**Une garde ne vaut que cassée au moins une fois.**
- Avant d'affirmer qu'un test protège quelque chose, le faire **échouer exprès** (casser le code ou l'attente), constater le rouge, puis rétablir.
- Obligatoire pour les **assertions négatives** (« n'apparaît pas », `toHaveCount(0)`, `not.toBeVisible()`, `assert x not in …`) : sur un mauvais sélecteur, elles passent à vide.

**E2E — Playwright (standard dès qu'il y a une interface).**
- Les parcours critiques se valident en **Playwright**, en vrai navigateur, et font partie de la porte.
- Tests **déterministes** : attendre les éléments (`await expect(locator)`), jamais de `waitForTimeout` fixe ; sélecteurs par rôle / `data-testid`, pas par classe CSS.
- Port paramétrable (`PORT_E2E`, §0) et serveur forcé en IPv4 (`127.0.0.1`) : sous Windows, `localhost` peut résoudre en `::1`.
- Parcours couverts et état du portage : §7.

---

## 6. Garde-fous (enforcement)

Les règles ci-dessus sont du *contexte* qu'un agent peut ne pas suivre. Pour les rendre **imposées** :

- **Pre-commit** : formatage + lint des fichiers indexés (outil et état : §7).
- **`/code-review`** (Claude Code) avant chaque PR — il lit ce fichier et `CLAUDE.md`.
- `💡 [proposé]` **Hook `PreToolUse`** pour imposer les règles dures : bloquer `git push` sur `main`/`dev`, refuser un nom de branche hors convention, et **exiger la porte verte avant un commit/merge**. (`/hooks` ou `.claude/settings.json`.)
- **CI** : exécute la même commande `porte` à chaque PR (état : §7).

---

## 7. Spécifique à ce dépôt (facturation_auto)

> **La seule section à remplacer pour réutiliser ce fichier dans un autre projet.**

- **Objet** : factures mensuelles d'un formateur indépendant, à partir de son agenda de cours, des tarifs de chaque école et d'un gabarit XLSX. Le protocole **PCF** (7 étapes) garde des **validations humaines aux étapes 1, 4 et 6** — voir `references/PROC00_pcf.md`.
- **Stack** : Python ≥ 3.11, paquet `facturation-auto` (commande `facturation`), `openpyxl`, `icalendar` et `recurring-ical-events` (lecture de l'agenda, F-10) ; LibreOffice pour le PDF. Installation : `python -m pip install -e ".[dev]"`.
- **Branche stable** : `main`. Intégration sur `dev`, créée depuis `main` le 17 novembre 2026. Ce modèle remplace l'ancien `main` / `develop` / `feature/*` ; `origin/develop`, identique à `main`, ne sert plus.
- **IDs de ticket** : `F-xx` fonctionnalité (critères dans `references/ROADMAP.md`, statuts dans Notion), `T-XXX-NN` test, `ADRxxx` décision (`references/ADR.md`, dernier attribué : ADR033), `Qx` question ouverte. Branche : `feat/F-18-gabarit`.
- **Porte** : `python -m pytest` — toute la banque (marqueurs `unit`, `integration`, `regression`, `e2e`), environ une minute et demie en CI et sur un poste rapide (96 s mesurées le 18 novembre 2026), de 1 min 20 à 3 min 30 sur le PC d'Adrian selon l'état du poste (77 s et 214 s le 23 novembre 2026, sur le code de `dev` en 50f0c42). `xfail_strict` : un défaut connu marqué `xfail` qui se met à passer fait échouer la suite, pour rappeler de retirer le marqueur. `💡 [proposé]` y ajouter `ruff check` + `ruff format --check` derrière un `scripts/porte.py`.
- **E2E** : Playwright via `pytest-playwright`, marqueur `e2e` déjà déclaré — **à partir du lot 2** (l'appli). Rien d'installé avant.
- **CI** : `.github/workflows/tests.yml` rejoue la banque à chaque PR et à chaque push sur `main`, `dev`, `feat/**`, `fix/**` et `chore/**`, sous Windows et Ubuntu, avec Python 3.11 et 3.12 (rendu PDF sauté si LibreOffice est absent) — c'est déjà l'équivalent de la porte.
- **Pre-commit** : aucun. `💡 [proposé]` `pre-commit` + `ruff`.
- **Scopes de commit** (relevés dans l'historique et complétés) : `format`, `gabarit`, `profil`, `agenda`, `master`, `cli`, `pdf`, `tests`, `ci`, `git`, et `lotN` pour le livrable d'un lot. Messages **en français**.
- **Règles non négociables** :
  - **Aucune donnée réelle dans le dépôt** — ni nom d'école, adresse, SIRET, IBAN, montant ou coordonnée —, y compris dans les tests et les exemples (données fictives créées à la volée). Les données métier vivent hors dépôt, dans la *racine de données*.
  - **Secrets hors dépôt** : profil dans `%APPDATA%\facturation\profil.toml`, adresse iCal dans `agenda.secret` à côté — **jamais affichée**.
  - **Aperçu avant toute écriture** dans le MASTER ou un classeur d'école (`--dry-run`), puis **validation explicite**.
  - **Jamais d'email envoyé automatiquement** : brouillons seulement.
  - **Ne jamais inventer** un taux, une coordonnée ou un volume horaire : demander.
  - Du code testé plutôt que des appels MCP ; des fichiers plats, pas de base de données.
  - Les entités ne se mélangent jamais (préfixes de facture `R_` / `K_`).
  - Documentation et messages de l'outil **en français**.
- **Fichiers locaux, jamais commités** : `CLAUDE.md`, `CLAUDE.local.md`, `CONTEXT.md`, `TASKS.md`, `dashboard.html`, `memory/`, `_archives/`, `.claude/settings.local.json` (ignorés par `.gitignore`), `installer_poste.*` et `.claude/worktrees/` (exclusion locale `.git/info/exclude`). Ce fichier-ci, lui, est versionné : il ne contient aucune donnée.
- **Où vit quoi (ADR025)** : le dépôt fait foi pour les décisions, les règles de l'outil, la feuille de route et les critères d'acceptation ; Notion porte l'avancement (statuts, banque de tests, journal de bord) ; le MASTER porte les données réelles, dont les règles de chaque école (ADR026). Les fichiers locaux ne portent aucune règle propre.
- **Docs de référence** : `README.md` · `references/ADR.md` (toutes les décisions, ADR001–033) · `references/ROADMAP.md` (lots, critères F-xx) · `references/PROC00_pcf.md` · `references/FORMAT.md` · `references/MASTER_SCHEMA.md` · `references/ONBOARDING.md` · `references/GLOSSAIRE.md` · `references/RECETTE_DOC.md` · `CHANGELOG.md` · `SKILL.md`.
