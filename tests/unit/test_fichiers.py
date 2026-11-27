"""Primitives de fichiers du lot 1 (P-0b, `facturation/fichiers.py` ; banque de tests : T-FIC-01 à T-FIC-19,
puis T-FIC-35, 37 et 39 à 41 pour la revue).

Tests d'acceptation écrits avant l'implémentation : chaque test importe `facturation.fichiers`
lui-même, pour que l'absence du module fasse échouer les tests sans bloquer la collecte des
autres. Fichiers fictifs créés dans `tmp_path`.

Contrat (contrats des vagues 3 à 5, § 1.3, 1.4, 1.6 et 3.1) :
- écritures atomiques : fichier temporaire `<nom>.tmp-<pid>` dans le même dossier, `fsync`,
  puis `os.replace` ; sur erreur, le temporaire est supprimé, l'original reste intact et
  l'exception remonte ; un classeur refusé par `os.replace` (`PermissionError`) devient
  `ClasseurVerrouille` ;
- JSON Lines : ajout seul, une ligne compacte par appel ; une dernière ligne tronquée est
  ignorée avec un avertissement, une ligne illisible ailleurs lève `FichierAbime` (fichier et
  numéro de ligne) ;
- empreinte : sha256 du JSON canonique (clés triées, séparateurs `(",", ":")`,
  `ensure_ascii=False`), après `en_json` ; aucun `float` n'y entre ;
- `masquer_secrets` remplace toute sous-chaîne `\\w+://\\S+` par `<adresse masquée>`.

Lectures retenues là où le contrat se tait (`💡 [proposé — à valider]`) : l'avertissement
d'une ligne tronquée est un `UserWarning` (catégorie par défaut de `warnings.warn`) ; le JSON
canonique est encodé en UTF-8 avant le hachage ; un `set` (non figé) est refusé par `en_json`
comme tout type non listé.
"""

from __future__ import annotations

import dataclasses
import hashlib
import importlib
import json
import math
import os
import re
import warnings
from datetime import date, datetime, time
from decimal import Decimal
from pathlib import Path
from zoneinfo import ZoneInfo

import pytest
from openpyxl import Workbook, load_workbook

pytestmark = pytest.mark.unit

PARIS = ZoneInfo("Europe/Paris")
ADRESSE_FICTIVE = "https://calendar.example.org/calendar/ical/personne%40example.org/private-0123abcd/basic.ics"


def _fic():
    """Module testé, importé dans le test : son absence fait échouer le test, pas la collecte."""
    return importlib.import_module("facturation.fichiers")


def _temporaires(dossier: Path) -> list[str]:
    """Fichiers temporaires laissés dans le dossier (`<nom>.tmp-<pid>` ou tout autre `.tmp`)."""
    return sorted(p.name for p in dossier.iterdir() if ".tmp" in p.name)


def _canonique(donnees) -> str:
    return json.dumps(donnees, sort_keys=True, separators=(",", ":"), ensure_ascii=False)


# ──────────────────────────────────────────────────────────────────────
# Classes d'erreur
# ──────────────────────────────────────────────────────────────────────


def test_t_fic_05_classes_d_erreur():
    """T-FIC-05 — `FichierAbime` est une `ValueError` ; `ClasseurVerrouille` une `RuntimeError`."""
    fic = _fic()
    assert issubclass(fic.FichierAbime, ValueError)
    assert issubclass(fic.ClasseurVerrouille, RuntimeError)


# ──────────────────────────────────────────────────────────────────────
# Écritures atomiques
# ──────────────────────────────────────────────────────────────────────


def test_t_fic_06_ecrire_texte_atomique(tmp_path):
    """T-FIC-06 — `ecrire_texte_atomique` crée les dossiers parents, écrit en UTF-8, remplace un
    fichier existant et ne laisse aucun fichier temporaire."""
    fic = _fic()
    cible = tmp_path / "a" / "b" / "apercu.txt"
    fic.ecrire_texte_atomique(cible, "Aperçu fictif : 1 800,00 €\n")
    assert cible.read_bytes().decode("utf-8") == "Aperçu fictif : 1 800,00 €\n"
    fic.ecrire_texte_atomique(cible, "Deuxième version")
    assert cible.read_text(encoding="utf-8") == "Deuxième version"
    assert _temporaires(cible.parent) == []


def test_t_fic_07_ecrire_json_atomique(tmp_path):
    """T-FIC-07 — `ecrire_json_atomique` : UTF-8 sans échappement des accents, clés triées,
    indentation de 2 ; relu par `lire_json` à l'identique ; aucun fichier temporaire."""
    fic = _fic()
    cible = tmp_path / "sorties" / "preparation.json"
    donnees = {"zèbre": 1, "école": "ÉCOLE FICTIVE A", "liste": [1, 2], "imbrique": {"b": 2, "a": 1}}
    fic.ecrire_json_atomique(cible, donnees)
    texte = cible.read_bytes().decode("utf-8")
    assert "ÉCOLE FICTIVE A" in texte and "\\u00c9" not in texte
    assert texte.index('"imbrique"') < texte.index('"liste"') < texte.index('"zèbre"')
    assert '\n  "école"' in texte  # indentation de 2
    assert json.loads(texte) == donnees
    assert fic.lire_json(cible) == donnees
    assert _temporaires(cible.parent) == []


@pytest.mark.parametrize("ecriture", ["texte", "json"])
def test_t_fic_01_os_replace_qui_echoue_laisse_l_original_intact(tmp_path, monkeypatch, ecriture):
    """T-FIC-01 — un `os.replace` qui échoue laisse l'original identique octet pour octet, supprime
    le fichier temporaire (écrit dans le même dossier) et laisse remonter l'exception."""
    fic = _fic()
    cible = tmp_path / "etat.json"
    original = b'{"version": "originale"}\n'
    cible.write_bytes(original)
    sources = []

    def refus(src, dst, *args, **kwargs):
        sources.append(Path(src))
        raise OSError(28, "No space left on device (simulé)")

    monkeypatch.setattr(os, "replace", refus)
    with pytest.raises(OSError):
        if ecriture == "texte":
            fic.ecrire_texte_atomique(cible, "nouvelle version")
        else:
            fic.ecrire_json_atomique(cible, {"version": "nouvelle"})
    monkeypatch.undo()

    assert sources, "os.replace n'a pas été appelé : l'écriture n'est pas atomique"
    assert all(src.parent == cible.parent for src in sources), sources
    assert cible.read_bytes() == original
    assert _temporaires(tmp_path) == []


def test_t_fic_08_creer_json_exclusif(tmp_path):
    """T-FIC-08 — `creer_json_exclusif` crée le fichier et rend True ; s'il existe, rend False sans
    rien écrire."""
    fic = _fic()
    cible = tmp_path / "emissions" / "R_ABC001.json"
    cible.parent.mkdir()
    assert fic.creer_json_exclusif(cible, {"n_facture": "R_ABC001"}) is True
    assert json.loads(cible.read_text(encoding="utf-8")) == {"n_facture": "R_ABC001"}
    avant = cible.read_bytes()
    assert fic.creer_json_exclusif(cible, {"n_facture": "AUTRE"}) is False
    assert cible.read_bytes() == avant


def test_t_fic_09_enregistrer_classeur_atomique(tmp_path, monkeypatch):
    """T-FIC-09 — `enregistrer_classeur_atomique` enregistre le classeur par un fichier temporaire du
    même dossier, puis `os.replace` ; aucun temporaire ne reste."""
    fic = _fic()
    cible = tmp_path / "ABC" / "ROSARI_ABC.xlsx"
    cible.parent.mkdir()
    wb = Workbook()
    wb.active.title = "Feuille fictive"
    wb.active["A1"] = "valeur fictive"
    sources = []
    vrai_replace = os.replace

    def espion(src, dst, *args, **kwargs):
        sources.append((Path(src), Path(dst)))
        return vrai_replace(src, dst, *args, **kwargs)

    monkeypatch.setattr(os, "replace", espion)
    fic.enregistrer_classeur_atomique(wb, cible)
    monkeypatch.undo()

    assert [dst for _, dst in sources] == [cible]
    assert sources[0][0].parent == cible.parent
    assert load_workbook(cible)["Feuille fictive"]["A1"].value == "valeur fictive"
    assert _temporaires(cible.parent) == []


def test_t_fic_10_classeur_ouvert_devient_classeur_verrouille(tmp_path, monkeypatch):
    """T-FIC-10 — une `PermissionError` sur `os.replace` (classeur ouvert dans Excel) devient
    `ClasseurVerrouille`, qui nomme le classeur et dit que rien n'a été modifié ; l'original est
    intact et aucun temporaire ne reste."""
    fic = _fic()
    cible = tmp_path / "ROSARI_ABC.xlsx"
    wb = Workbook()
    wb.active["A1"] = "original"
    wb.save(cible)
    original = cible.read_bytes()
    wb.active["A1"] = "modifié"

    def refus(src, dst, *args, **kwargs):
        raise PermissionError(13, "Permission denied (simulé)", str(dst))

    monkeypatch.setattr(os, "replace", refus)
    with pytest.raises(fic.ClasseurVerrouille) as erreur:
        fic.enregistrer_classeur_atomique(wb, cible)
    monkeypatch.undo()

    message = str(erreur.value)
    assert "ROSARI_ABC.xlsx" in message
    assert re.search(r"(?i)rien n'a été modifié", message), message
    assert cible.read_bytes() == original
    assert _temporaires(tmp_path) == []


def test_t_fic_11_sauvegarder_fichier(tmp_path):
    """T-FIC-11 — `sauvegarder_fichier` copie le fichier à l'identique dans `<dossier>/sauvegardes/`
    par défaut, ou dans le dossier donné ; deux sauvegardes de suite ne s'écrasent jamais."""
    fic = _fic()
    classeur = tmp_path / "ABC" / "ROSARI_ABC.xlsx"
    classeur.parent.mkdir()
    classeur.write_bytes(b"contenu fictif du classeur")

    premiere = fic.sauvegarder_fichier(classeur)
    seconde = fic.sauvegarder_fichier(classeur)
    assert premiere.parent == seconde.parent == classeur.parent / "sauvegardes"
    assert premiere != seconde
    for copie in (premiere, seconde):
        assert copie.name.startswith("ROSARI_ABC") and copie.suffix == ".xlsx", copie.name
        assert copie.read_bytes() == b"contenu fictif du classeur"

    ailleurs = fic.sauvegarder_fichier(classeur, tmp_path / "autre_dossier")
    assert ailleurs.parent == tmp_path / "autre_dossier"
    assert ailleurs.read_bytes() == classeur.read_bytes()
    assert classeur.read_bytes() == b"contenu fictif du classeur"


# ──────────────────────────────────────────────────────────────────────
# JSON Lines
# ──────────────────────────────────────────────────────────────────────


def test_t_fic_12_ajouter_puis_lire_jsonl(tmp_path):
    """T-FIC-12 — `ajouter_jsonl` ajoute une ligne compacte par appel (dossiers créés au besoin) ;
    `lire_jsonl` les relit dans l'ordre ; un fichier absent se lit comme une liste vide."""
    fic = _fic()
    journal = tmp_path / "_sorties" / "2026-12" / "journal.jsonl"
    assert fic.lire_jsonl(journal) == []
    fic.ajouter_jsonl(journal, {"type": "preparation", "école": "ABC", "n": 1})
    fic.ajouter_jsonl(journal, {"type": "decision", "n": 2})
    texte = journal.read_bytes().decode("utf-8")
    lignes = texte.split("\n")
    assert lignes[-1] == "" and len(lignes) == 3, texte  # une ligne par appel, chacune terminée
    assert all(": " not in ligne and ", " not in ligne for ligne in lignes), texte  # compacte
    assert fic.lire_jsonl(journal) == [{"type": "preparation", "école": "ABC", "n": 1}, {"type": "decision", "n": 2}]


@pytest.mark.parametrize(
    "enregistrement",
    [["pas", "un", "dictionnaire"], "texte", {"montant": math.nan}, {"montant": math.inf},
     {"imbrique": {"valeurs": [1.0, -math.inf]}}],
    ids=["liste", "texte", "nan", "inf", "inf-imbrique"],
)
def test_t_fic_13_ajouter_jsonl_refuse(tmp_path, enregistrement):
    """T-FIC-13 — `ajouter_jsonl` lève `TypeError` pour un enregistrement qui n'est pas un `Mapping`
    ou qui contient un `float` non fini, sans rien écrire."""
    fic = _fic()
    journal = tmp_path / "journal.jsonl"
    journal.write_text('{"deja":"là"}\n', encoding="utf-8")
    with pytest.raises(TypeError):
        fic.ajouter_jsonl(journal, enregistrement)
    assert journal.read_text(encoding="utf-8") == '{"deja":"là"}\n'


def test_t_fic_02_ligne_tronquee_et_ligne_abimee(tmp_path):
    """T-FIC-02 — une dernière ligne sans fin de ligne et illisible (écriture interrompue) est ignorée
    avec un avertissement ; une ligne illisible au milieu lève `FichierAbime`, qui nomme le fichier
    et le numéro de la ligne."""
    fic = _fic()
    tronque = tmp_path / "tronque.jsonl"
    tronque.write_text('{"n":1}\n{"n":2}\n{"n":3,"etat":"inter', encoding="utf-8")
    with pytest.warns(UserWarning):
        assert fic.lire_jsonl(tronque) == [{"n": 1}, {"n": 2}]

    abime = tmp_path / "abime.jsonl"
    abime.write_text('{"n":1}\n{"n":2\n{"n":3}\n', encoding="utf-8")
    with pytest.raises(fic.FichierAbime) as erreur:
        fic.lire_jsonl(abime)
    message = str(erreur.value)
    assert "abime.jsonl" in message
    assert re.search(r"ligne 2\b", message), message


def test_t_fic_14_derniere_ligne_complete_mais_illisible(tmp_path):
    """T-FIC-14 — une dernière ligne illisible mais terminée par une fin de ligne n'est pas une
    écriture interrompue : `FichierAbime` (ligne 2), pas un simple avertissement."""
    fic = _fic()
    abime = tmp_path / "journal.jsonl"
    abime.write_text('{"n":1}\n{"n":2,\n', encoding="utf-8")
    with pytest.raises(fic.FichierAbime, match=r"ligne 2\b"):
        fic.lire_jsonl(abime)


def test_t_fic_15_fichier_sain_sans_avertissement(tmp_path):
    """T-FIC-15 — un fichier sain se lit sans aucun avertissement (la garde de T-FIC-02 ne passe pas
    à vide)."""
    fic = _fic()
    sain = tmp_path / "sain.jsonl"
    sain.write_text('{"n":1}\n{"n":2}\n', encoding="utf-8")
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        assert fic.lire_jsonl(sain) == [{"n": 1}, {"n": 2}]


# ──────────────────────────────────────────────────────────────────────
# Conversions et empreintes
# ──────────────────────────────────────────────────────────────────────


@dataclasses.dataclass(frozen=True)
class _Fictif:
    nom: str
    quand: date
    montant: Decimal
    heures: tuple


def test_t_fic_16_en_json_convertit_les_types_du_contrat():
    """T-FIC-16 — `en_json` : Decimal → texte, date → AAAA-MM-JJ, datetime → ISO avec décalage,
    time → HH:MM, Path → texte POSIX, tuple → liste, frozenset → liste triée, dataclass → dict,
    récursivement ; un float fini reste un float."""
    fic = _fic()
    assert fic.en_json(Decimal("1800.00")) == "1800.00"
    assert fic.en_json(date(2026, 12, 1)) == "2026-12-01"
    assert fic.en_json(datetime(2026, 12, 1, 9, 30, tzinfo=PARIS)) == "2026-12-01T09:30:00+01:00"
    assert fic.en_json(time(9, 30)) == "09:30"
    assert fic.en_json(Path("ABC") / "ROSARI_ABC.xlsx") == "ABC/ROSARI_ABC.xlsx"
    assert fic.en_json((1, "a")) == [1, "a"]
    assert fic.en_json(frozenset({"ECO", "ABC"})) == ["ABC", "ECO"]
    assert fic.en_json(3.5) == 3.5 and isinstance(fic.en_json(3.5), float)
    assert fic.en_json(_Fictif("ABC", date(2026, 12, 1), Decimal("157.50"), (3.5, 2))) == {
        "nom": "ABC", "quand": "2026-12-01", "montant": "157.50", "heures": [3.5, 2],
    }
    assert fic.en_json({"a": [Decimal("1.10"), {"b": (date(2026, 1, 2),)}], "c": None, "d": True}) == {
        "a": ["1.10", {"b": ["2026-01-02"]}], "c": None, "d": True,
    }


@pytest.mark.parametrize("valeur", [object(), b"octets", {1, 2}], ids=["objet", "octets", "set"])
def test_t_fic_17_en_json_refuse_un_autre_type(valeur):
    """T-FIC-17 — `en_json` lève `TypeError` pour tout type hors du contrat, même imbriqué."""
    fic = _fic()
    with pytest.raises(TypeError):
        fic.en_json(valeur)
    with pytest.raises(TypeError):
        fic.en_json({"imbrique": [valeur]})


def test_t_fic_03_empreinte_stable_et_sans_float():
    """T-FIC-03 — l'empreinte est le sha256 du JSON canonique (après `en_json`) : stable quand l'ordre
    des clés change, différente quand une valeur change ; un `float` est refusé (`TypeError`),
    même imbriqué."""
    fic = _fic()
    a = {"ecole": "ABC", "lignes": [{"heures": "3.5", "taux": "45.0"}], "total": Decimal("157.50"),
         "mois": "2026-12", "emission": date(2026, 12, 1)}
    b = {"emission": date(2026, 12, 1), "mois": "2026-12", "total": Decimal("157.50"),
         "lignes": [{"taux": "45.0", "heures": "3.5"}], "ecole": "ABC"}
    empreinte = fic.empreinte(a)
    assert re.fullmatch(r"[0-9a-f]{64}", empreinte), empreinte
    assert fic.empreinte(b) == empreinte
    attendu = hashlib.sha256(_canonique(fic.en_json(a)).encode("utf-8")).hexdigest()
    assert empreinte == attendu
    assert fic.empreinte({**a, "total": Decimal("157.51")}) != empreinte
    for avec_float in ({**a, "taux": 45.0}, {"lignes": [{"heures": 3.5}]}, 3.5):
        with pytest.raises(TypeError):
            fic.empreinte(avec_float)


def test_t_fic_18_empreinte_fichier(tmp_path):
    """T-FIC-18 — `empreinte_fichier` est le sha256 des octets du fichier."""
    fic = _fic()
    fichier = tmp_path / "R_ABC001 Déc 2026.pdf"
    fichier.write_bytes(b"%PDF-1.4 contenu fictif")
    assert fic.empreinte_fichier(fichier) == hashlib.sha256(b"%PDF-1.4 contenu fictif").hexdigest()


def test_t_fic_04_masquer_secrets():
    """T-FIC-04 — `masquer_secrets` remplace toute adresse (`https://…`, `webcal://…`) par
    « <adresse masquée> » et garde le reste du texte ; un texte sans adresse est inchangé."""
    fic = _fic()
    texte = f"Agenda indisponible ({ADRESSE_FICTIVE}) : réessayez. Miroir webcal://agenda.example.org/x.ics fin"
    masque = fic.masquer_secrets(texte)
    assert "calendar.example.org" not in masque and "private-0123abcd" not in masque
    assert "agenda.example.org" not in masque
    assert masque == "Agenda indisponible (<adresse masquée> : réessayez. Miroir <adresse masquée> fin"
    assert fic.masquer_secrets("Rien à masquer : 1 800,00 €") == "Rien à masquer : 1 800,00 €"


def test_t_fic_19_lire_json_abime_nomme_le_fichier(tmp_path):
    """T-FIC-19 — `lire_json` d'un fichier abîmé lève une `ValueError` (`FichierAbime` attendu) qui
    nomme le fichier, pas une `JSONDecodeError` muette sur son emplacement."""
    fic = _fic()
    abime = tmp_path / "preparation.json"
    abime.write_text('{"mois": "2026-12",\n "brouillons": [', encoding="utf-8")
    with pytest.raises(ValueError) as erreur:
        fic.lire_json(abime)
    assert "preparation.json" in str(erreur.value), str(erreur.value)


# ──────────────────────────────────────────────────────────────────────
# Revue de P-0b (T-FIC-35, T-FIC-37, T-FIC-39 à T-FIC-41 ; T-FIC-36, 38 et 42 dans test_chemins.py)
# ──────────────────────────────────────────────────────────────────────
#
# Constats des deux relecteurs. Écrits rouges avant le correctif : T-FIC-35, 39, 40 et 41.
# T-FIC-37 couvre une garde déjà codée que les relecteurs ont trouvée sans test : il est vert, et a
# été vérifié en cassant la garde (AGENTS.md §5).
# - T-FIC-35 : `ajouter_jsonl` collait l'enregistrement à une dernière ligne sans « \n » ; après
#   une écriture interrompue (§ 3.1), le journal devenait illisible (`FichierAbime`) et l'événement
#   neuf perdu. Lecture retenue (`💡 [proposé — à valider]`) : une fin illisible sans « \n » est
#   coupée avant l'ajout, avec un avertissement (un seul, entre l'ajout et la lecture qui suit) ;
#   une fin lisible sans « \n » est terminée par « \n », sans rien perdre ;
# - T-FIC-37 : un enregistrement de classeur abîmé ou interrompu ne remplace jamais l'original et ne
#   laisse aucun temporaire (relecture avant `os.replace`) ;
# - T-FIC-39 : `lire_jsonl` refusait un journal enregistré « UTF-8 avec BOM » (Bloc-notes), alors que
#   `lire_json` l'accepte ;
# - T-FIC-40 (`💡`) : une sauvegarde n'apparaît sous son nom définitif qu'une fois sa copie complète
#   et mise sur disque (temporaire, `fsync`, puis renommage) : un arrêt brutal ne laisse jamais une
#   copie tronquée qui a l'air valide ;
# - T-FIC-41 (`💡`) : `enregistrer_classeur_atomique` met le temporaire sur disque (`fsync`) avant
#   `os.replace`, comme `ecrire_*_atomique` : une coupure de courant ne laisse pas un MASTER vide.


def test_t_fic_35_ajout_apres_une_ecriture_interrompue(tmp_path):
    """T-FIC-35 — un ajout qui suit une écriture interrompue (dernière ligne tronquée, sans « \\n »)
    ne rend pas le journal illisible : la ligne tronquée est écartée avec un avertissement, et
    l'événement neuf se relit."""
    fic = _fic()
    journal = tmp_path / "journal.jsonl"
    journal.write_bytes(b'{"n":1}\n{"n":2,"etat":"inter')
    with pytest.warns(UserWarning) as releves:
        fic.ajouter_jsonl(journal, {"n": 3})
        lus = fic.lire_jsonl(journal)
    assert lus == [{"n": 1}, {"n": 3}]
    assert len([r for r in releves if issubclass(r.category, UserWarning)]) == 1, [str(r.message) for r in releves]


def test_t_fic_35_ajout_apres_une_derniere_ligne_sans_fin_de_ligne(tmp_path):
    """T-FIC-35 — une dernière ligne lisible mais sans « \\n » (écrite à la main) n'est pas collée à
    l'enregistrement suivant : les trois enregistrements se relisent, sans `FichierAbime`."""
    fic = _fic()
    journal = tmp_path / "journal.jsonl"
    journal.write_bytes(b'{"a":1}\n{"b":2}')
    fic.ajouter_jsonl(journal, {"c": 3})
    assert fic.lire_jsonl(journal) == [{"a": 1}, {"b": 2}, {"c": 3}]


def _enregistrement_factice(comportement: str):
    """Remplace `Workbook.save` : écrit un fichier qui n'est pas un classeur, ou s'interrompt après
    un début d'écriture. Accepte un chemin ou un fichier ouvert."""
    def enregistrer(destination, *args, **kwargs):
        octets = b"pas un zip" if comportement == "abime" else b"PK\x03\x04debut"
        if hasattr(destination, "write"):
            destination.write(octets)
        else:
            Path(destination).write_bytes(octets)
        if comportement == "interrompu":
            raise OSError(28, "No space left on device (simulé)")
    return enregistrer


@pytest.mark.parametrize("comportement", ["abime", "interrompu"])
def test_t_fic_37_classeur_illisible_jamais_substitue(tmp_path, monkeypatch, comportement):
    """T-FIC-37 — un enregistrement qui produit un fichier illisible (refusé à la relecture), ou qui
    s'interrompt, lève une exception : l'original reste identique octet pour octet et aucun
    temporaire ne reste."""
    fic = _fic()
    cible = tmp_path / "ROSARI_ABC.xlsx"
    wb = Workbook()
    wb.active["A1"] = "original"
    wb.save(cible)
    original = cible.read_bytes()
    wb.active["A1"] = "modifié"
    monkeypatch.setattr(wb, "save", _enregistrement_factice(comportement))
    with pytest.raises(Exception):
        fic.enregistrer_classeur_atomique(wb, cible)
    assert cible.read_bytes() == original
    assert _temporaires(tmp_path) == []
    assert sorted(p.name for p in tmp_path.iterdir()) == ["ROSARI_ABC.xlsx"]


def test_t_fic_39_journal_avec_bom(tmp_path):
    """T-FIC-39 — un journal enregistré « UTF-8 avec BOM » (Bloc-notes) se lit comme `lire_json` lit
    un JSON avec BOM ; un ajout ultérieur se relit aussi."""
    fic = _fic()
    journal = tmp_path / "journal.jsonl"
    journal.write_bytes("﻿".encode("utf-8") + b'{"n":1}\n{"n":2}\n')
    assert fic.lire_jsonl(journal) == [{"n": 1}, {"n": 2}]
    fic.ajouter_jsonl(journal, {"n": 3})
    assert fic.lire_jsonl(journal) == [{"n": 1}, {"n": 2}, {"n": 3}]


def test_t_fic_40_sauvegarde_jamais_visible_incomplete(tmp_path, monkeypatch):
    """T-FIC-40 — tant que la copie n'est pas mise sur disque (`fsync`), aucun fichier ne porte le nom
    d'une sauvegarde (💡 temporaire puis renommage) ; à la fin, la sauvegarde est complète."""
    fic = _fic()
    classeur = tmp_path / "ROSARI_MASTER.xlsx"
    contenu = b"contenu fictif du MASTER " * 4096
    classeur.write_bytes(contenu)
    dossier = tmp_path / "sauvegardes"
    nom_de_sauvegarde = re.compile(r"ROSARI_MASTER_\d{8}-\d{6}(?:-\d+)?\.xlsx")
    visibles = []
    vrai_fsync = os.fsync

    def espion(descripteur):
        visibles.append(sorted(p.name for p in dossier.iterdir() if nom_de_sauvegarde.fullmatch(p.name)))
        return vrai_fsync(descripteur)

    monkeypatch.setattr(os, "fsync", espion)
    copie = fic.sauvegarder_fichier(classeur)
    monkeypatch.undo()

    assert visibles, "la sauvegarde n'est pas mise sur disque (fsync) avant d'être rendue"
    assert all(noms == [] for noms in visibles), visibles
    assert nom_de_sauvegarde.fullmatch(copie.name), copie.name
    assert copie.read_bytes() == contenu
    assert sorted(p.name for p in dossier.iterdir()) == [copie.name]


def test_t_fic_41_classeur_mis_sur_disque_avant_substitution(tmp_path, monkeypatch):
    """T-FIC-41 — `enregistrer_classeur_atomique` met le temporaire sur disque (`fsync` sur ce même
    fichier) avant `os.replace` (💡, comme `ecrire_*_atomique`)."""
    fic = _fic()
    cible = tmp_path / "ROSARI_ABC.xlsx"
    wb = Workbook()
    wb.active["A1"] = "valeur fictive"
    appels = []
    vrai_fsync, vrai_replace = os.fsync, os.replace

    def espion_fsync(descripteur):
        statut = os.fstat(descripteur)
        appels.append(("fsync", (statut.st_dev, statut.st_ino)))
        return vrai_fsync(descripteur)

    def espion_replace(src, dst, *args, **kwargs):
        statut = os.stat(src)
        appels.append(("replace", (statut.st_dev, statut.st_ino)))
        return vrai_replace(src, dst, *args, **kwargs)

    monkeypatch.setattr(os, "fsync", espion_fsync)
    monkeypatch.setattr(os, "replace", espion_replace)
    fic.enregistrer_classeur_atomique(wb, cible)
    monkeypatch.undo()

    remplacements = [i for i, (nature, _) in enumerate(appels) if nature == "replace"]
    assert len(remplacements) == 1, appels
    temporaire = appels[remplacements[0]][1]
    assert ("fsync", temporaire) in appels[:remplacements[0]], appels
    assert load_workbook(cible).active["A1"].value == "valeur fictive"
