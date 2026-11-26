"""Outils communs des commandes et répartition de la ligne de commande (P-0b ; banque de tests :
T-CLI-13 à T-CLI-19).

Tests d'acceptation écrits avant l'implémentation : chaque test importe `facturation.commandes`
(ou `facturation.cli`) lui-même, pour que l'absence du module fasse échouer les tests sans bloquer
la collecte des autres. Profils fictifs écrits dans `tmp_path`.

Contrat (contrats des vagues 3 à 5, § 1.6 et 3.7) :
- `cli.py` répartit les commandes sur `args.func` ; `_COMMANDES` contient un module par ligne,
  chacun exposant `ajouter(sous) -> None` ;
- `facturation/commandes/__init__.py` : `argument_mois`, `argument_date` (JJ/MM/AAAA ou
  AAAA-MM-JJ), `mois_par_defaut` (le mois en cours si `aujourd_hui.day >= profil.lancement_jour`,
  sinon le précédent), `charger_profil` (affiche l'erreur), `demander` (mot entier, aucun défaut ;
  fin d'entrée → None), `entete` (« Mois : décembre 2026 — entité ROSARI »).

Lectures retenues (`💡 [proposé — à valider]`) :
- `argument_mois` et `argument_date` servent de `type=` à argparse : un refus est une
  `argparse.ArgumentTypeError` qui cite la valeur, comme `--mois` de `agenda lire` ;
- `demander` repose la question tant que la réponse n'est pas un des mots proposés (espaces
  autour retirés) ; la question figure dans l'invite ;
- `charger_profil` rend None et affiche l'erreur du profil sur la sortie d'erreur ou la sortie
  standard ;
- `main` rend ce que rend `args.func(args)`.
"""

from __future__ import annotations

import argparse
import dataclasses
import importlib
import types
from datetime import date
from pathlib import Path

import pytest

from facturation import config

pytestmark = pytest.mark.unit


def _cmd():
    """Module testé, importé dans le test : son absence fait échouer le test, pas la collecte."""
    return importlib.import_module("facturation.commandes")


def _profil(tmp_path: Path, entite: str = "ROSARI") -> config.Profil:
    chemin = tmp_path / f"profil_{entite}.toml"
    chemin.write_text(config.rediger("Personne A", entite, str(tmp_path / "donnees")), encoding="utf-8")
    return config.charger(chemin)


def test_t_cli_13_argument_mois():
    """T-CLI-13 — `argument_mois` rend le mois « AAAA-MM » ; tout autre texte lève
    `ArgumentTypeError` qui le cite."""
    cmd = _cmd()
    assert cmd.argument_mois("2026-12") == "2026-12"
    for valeur in ("2026-13", "12/2026", "2026-1", "décembre"):
        with pytest.raises(argparse.ArgumentTypeError, match=valeur):
            cmd.argument_mois(valeur)


def test_t_cli_14_argument_date():
    """T-CLI-14 — `argument_date` lit JJ/MM/AAAA ou AAAA-MM-JJ ; une date impossible ou une autre forme
    lève `ArgumentTypeError` qui la cite."""
    cmd = _cmd()
    assert cmd.argument_date("01/12/2026") == date(2026, 12, 1)
    assert cmd.argument_date("2026-12-01") == date(2026, 12, 1)
    assert cmd.argument_date("29/02/2028") == date(2028, 2, 29)
    for valeur in ("31/02/2026", "2026-13-01", "demain", "1/12/26", "2026/12/01"):
        with pytest.raises(argparse.ArgumentTypeError, match=valeur):
            cmd.argument_date(valeur)


@pytest.mark.parametrize(
    ("jour_lancement", "aujourd_hui", "attendu"),
    [
        (20, date(2026, 12, 20), "2026-12"),
        (20, date(2026, 12, 31), "2026-12"),
        (20, date(2026, 12, 19), "2026-11"),
        (20, date(2027, 1, 5), "2026-12"),
        (1, date(2027, 1, 1), "2027-01"),
        (25, date(2026, 12, 24), "2026-11"),
    ],
    ids=["jour-du-lancement", "fin-du-mois", "veille", "debut-janvier", "lancement-le-1er", "lancement-le-25"],
)
def test_t_cli_15_mois_par_defaut(tmp_path, jour_lancement, aujourd_hui, attendu):
    """T-CLI-15 — mois par défaut : le mois en cours à partir du jour de lancement du profil, sinon le
    mois précédent (changement d'année compris)."""
    profil = dataclasses.replace(_profil(tmp_path), lancement_jour=jour_lancement)
    assert _cmd().mois_par_defaut(profil, aujourd_hui) == attendu


def test_t_cli_16_charger_profil(tmp_path, capsys):
    """T-CLI-16 — `charger_profil` rend le profil, ou None en affichant l'erreur du profil (fichier
    introuvable) au lieu de lever."""
    cmd = _cmd()
    profil = _profil(tmp_path)
    assert cmd.charger_profil(profil.chemin) == profil
    assert cmd.charger_profil(tmp_path / "absent.toml") is None
    sortie = capsys.readouterr()
    assert "introuvable" in sortie.out + sortie.err


def _entree(*reponses: str):
    """Fausse fonction `input` : rend les réponses dans l'ordre, puis lève `EOFError` ; garde les invites."""
    restantes = list(reponses)
    invites: list[str] = []

    def entree(invite: str = "") -> str:
        invites.append(invite)
        if not restantes:
            raise EOFError
        return restantes.pop(0)

    entree.invites = invites
    return entree


def test_t_cli_17_demander_un_mot_entier_sans_defaut():
    """T-CLI-17 — `demander` rend l'un des mots proposés, saisi en entier ; une réponse vide, un début
    de mot ou un autre mot font reposer la question (aucun défaut) ; la fin de l'entrée rend None."""
    cmd = _cmd()
    entree = _entree("oui")
    assert cmd.demander("Émettre les 2 factures ?", ("oui", "non"), entree=entree) == "oui"
    assert "Émettre les 2 factures ?" in entree.invites[0]

    entree = _entree("", "o", "peut-être", " non ")
    assert cmd.demander("Émettre ?", ("oui", "non"), entree=entree) == "non"
    assert len(entree.invites) == 4

    assert cmd.demander("Émettre ?", ("oui", "non"), entree=_entree()) is None
    assert cmd.demander("Émettre ?", ("oui", "non"), entree=_entree("", "o")) is None


def test_t_cli_18_entete(tmp_path):
    """T-CLI-18 — première ligne de chaque commande : « Mois : décembre 2026 — entité ROSARI »."""
    cmd = _cmd()
    assert cmd.entete(_profil(tmp_path), "2026-12") == "Mois : décembre 2026 — entité ROSARI"
    assert cmd.entete(_profil(tmp_path, "KAUR"), "2027-02") == "Mois : février 2027 — entité KAUR"


def test_t_cli_19_repartition_par_args_func(monkeypatch):
    """T-CLI-19 — `cli.py` construit ses sous-commandes depuis `_COMMANDES` (un module par ligne, qui
    expose `ajouter(sous)`) et répartit sur `args.func` : un module ajouté en fin de liste suffit à
    brancher une commande, dont `main` rend le code."""
    cli = importlib.import_module("facturation.cli")
    assert isinstance(cli._COMMANDES, (list, tuple))
    for module in cli._COMMANDES:
        assert callable(getattr(module, "ajouter", None)), module

    recus = []

    def ajouter(sous) -> None:
        commande = sous.add_parser("essai-fictif", help="Commande fictive des tests")
        commande.add_argument("--valeur", type=int, required=True)
        commande.set_defaults(func=lambda args: recus.append(args.valeur) or args.valeur)

    fictif = types.ModuleType("facturation.commandes.essai_fictif")
    fictif.ajouter = ajouter
    monkeypatch.setattr(cli, "_COMMANDES", [*cli._COMMANDES, fictif])
    assert cli.main(["essai-fictif", "--valeur", "3"]) == 3
    assert recus == [3]
