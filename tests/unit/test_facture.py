"""Facture dans le paquet (P-0a ; banque de tests : T-FAC-01, T-FAC-03 à T-FAC-17) et contrôles de la facture
laissés hors de P-0a (fix-facture ; T-FAC-18 à T-FAC-25, puis T-FAC-18c, 18d, 25b et 26 à 28 pour les constats de
sa revue).

T-FAC-02 (aide du script, en sous-processus) est à la fin de `tests/integration/test_scripts.py`, avec le côté
script de T-FAC-18 à T-FAC-29.

`facturation/facture.py` reprend les fonctions de `scripts/nouvelle_facture.py` sans changer leur
comportement, et en ajoute deux : `ecrire_onglet` (la séquence d'écriture de `generer`, sans rien
enregistrer) et `verifier_onglet` (le contrôle après écriture, valeurs attendues déduites de la
spec). Le script n'est plus qu'une enveloppe : `from facturation.facture import *`, plus `main()`.

Tests d'acceptation écrits avant l'implémentation : chaque test importe `facturation.facture`
lui-même, pour que son absence fasse échouer les tests sans bloquer la collecte des autres.
Données fictives uniquement (`tests.conftest.racine_fictive` : entité ROSARI, école ABC).
"""

from __future__ import annotations

import copy
import hashlib
import importlib
import importlib.util
import inspect
import json
import os
import re
from datetime import date
from pathlib import Path

import pytest
from openpyxl import Workbook, load_workbook

from facturation import pcf_format as fmt
from tests.conftest import creer_master

pytestmark = pytest.mark.unit

DEPOT = Path(__file__).resolve().parents[2]
SCRIPT = DEPOT / "scripts" / "nouvelle_facture.py"

ENTITE = "ROSARI"
# IBAN fictif de `tests.conftest.creer_master` ; T-FAC-04 vérifie qu'une seule cellule le porte.
IBAN_FICTIF = "FR76 0000 0000 0000 0000 0000 000"

# Noms publics de `facturation/facture.py` d'après le contrat, hors `fmt` et `config`.
NOMS_PUBLICS = {
    "EcritureRefusee", "PostMortemEnEcart", "echeance_de_la_spec", "charger_spec", "controler_spec",
    "construire_groupes", "conflits_ecriture", "apercu", "generer", "ecrire_onglet", "verifier_onglet",
}


def _seance(jour: str, heures: float, taux: float) -> dict:
    return {"date": jour, "heures": heures, "taux": taux}


# Toutes les clés reconnues : école à libellé, bon de commande, contrat, échéance, colonne Jour masquée,
# deux promotions et un taux décimal.
SPEC_COMPLETE = {
    "entite": ENTITE,
    "ecole": "ABC",
    "ecole_label": "ÉCOLE FICTIVE A",
    "fichier_sortie": "ROSARI/ABC/ROSARI_ABC.xlsx",
    "onglet": "R_ABC001 Sept 2026",
    "n_facture": "R_ABC001",
    "date_emission": "2026-09-30",
    "ref_bon_commande": "BC-EXEMPLE-0001",
    "ref_contrat": "CT-EXEMPLE-0001",
    "echeance": "2026-11-15",
    "masquer_jour": True,
    "lignes": [
        {"promotion": "M1 S1", "libelle": "[2026][M1 S1][Module A][6H]",
         "seances": [_seance("2026-09-15", 3, 90), _seance("2026-09-22", 3, 90)]},
        {"promotion": "M2 S1", "libelle": "[2026][M2 S1][Module B][1H30]",
         "seances": [_seance("2026-09-23", 1.5, 60.5)]},
    ],
}

# Les seules clés obligatoires : ni libellé d'école, ni bon de commande, ni contrat, ni échéance.
SPEC_MINIMALE = {
    "entite": ENTITE,
    "ecole": "ABC",
    "fichier_sortie": "ROSARI/ABC/ROSARI_ABC.xlsx",
    "onglet": "R_ABC002 Oct 2026",
    "n_facture": "R_ABC002",
    "date_emission": "2026-10-31",
    "lignes": [{"promotion": "M1 S1", "libelle": "Module A", "seances": [_seance("2026-10-06", 3.5, 90)]}],
}

# Au-delà de 60 séances, des lignes sont insérées et le bas de facture descend d'autant (F-18).
# Clés obligatoires seules, elle aussi.
SPEC_62_SEANCES = {
    **SPEC_MINIMALE,
    "onglet": "R_ABC003 Sept 2026",
    "n_facture": "R_ABC003",
    "date_emission": "2026-09-30",
    "lignes": [{"promotion": "M1 S1", "libelle": "Module A",
                "seances": [_seance(f"2026-09-{1 + i % 28:02d}", 1, 50) for i in range(62)]}],
}

SPECS = {"complete": SPEC_COMPLETE, "minimale": SPEC_MINIMALE, "62-seances": SPEC_62_SEANCES}


def _spec(nom: str = "complete", **surcharges) -> dict:
    """Copie indépendante d'une spec fictive, complétée des surcharges."""
    spec = copy.deepcopy(SPECS[nom])
    spec.update(surcharges)
    return spec


def _module():
    """Module testé, importé dans le test : son absence fait échouer le test, pas la collecte."""
    return importlib.import_module("facturation.facture")


def _charger_script():
    """Importe `scripts/nouvelle_facture.py`, qui n'est pas un paquet, par son chemin."""
    spec = importlib.util.spec_from_file_location("nouvelle_facture", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _nouveau_classeur() -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    return wb


@pytest.fixture
def dest(racine_fictive: Path) -> dict:
    """Destinataire fictif « ABC », lu dans le MASTER fictif comme le fait `generer`."""
    return fmt.get_destinataire("ABC", entite=ENTITE)


def _ecrire_et_relire(facture, spec: dict, dest: dict, dossier: Path):
    """Écrit l'onglet par `ecrire_onglet`, enregistre le classeur, puis rend l'onglet relu du fichier."""
    wb = _nouveau_classeur()
    ws = facture.ecrire_onglet(wb, spec, dest, entite=ENTITE)
    chemin = dossier / "classeur_relu.xlsx"
    wb.save(chemin)
    return load_workbook(chemin)[ws.title]


def _cellules_valant(ws, valeur) -> list[str]:
    return [c.coordinate for ligne in ws.iter_rows() for c in ligne if c.value == valeur]


def _cles(ecarts: list[str]) -> list[str]:
    """Clé de chaque écart « ERREUR sur cellule <clé> … » ; un écart d'une autre forme fait échouer."""
    cles = []
    for ecart in ecarts:
        trouve = re.match(r"ERREUR sur cellule (\w+)", ecart)
        assert trouve, f"écart de forme inattendue : {ecart!r}"
        cles.append(trouve.group(1))
    return cles


def _instantane(dossier: Path) -> dict[str, tuple[str, int]]:
    """Empreinte et date de modification de chaque fichier sous `dossier`, pour constater que rien n'a été
    écrit : la date attrape aussi une réécriture à l'identique (openpyxl peut redonner les mêmes octets)."""
    return {
        p.relative_to(dossier).as_posix(): (hashlib.sha256(p.read_bytes()).hexdigest(), p.stat().st_mtime_ns)
        for p in sorted(dossier.rglob("*")) if p.is_file()
    }


# ──────────────────────────────────────────────────────────────────────
# Le script n'est plus qu'une enveloppe du paquet
# ──────────────────────────────────────────────────────────────────────


def test_t_fac_01_le_script_expose_generer_du_paquet():
    """T-FAC-01 — `facturation.facture.generer` est l'objet exposé par le script chargé par `importlib`."""
    facture = _module()
    nf = _charger_script()
    assert nf.generer is facture.generer


def test_t_fac_05a_noms_publics_du_module():
    """T-FAC-05 — `__all__` liste exactement les noms du contrat, plus `fmt` et `config` ; `fmt` et `config`
    sont les modules du paquet, et chaque nom public est défini dans `facturation.facture` lui-même."""
    facture = _module()
    assert len(facture.__all__) == len(set(facture.__all__)), facture.__all__
    assert set(facture.__all__) == NOMS_PUBLICS | {"fmt", "config"}
    assert facture.fmt is importlib.import_module("facturation.pcf_format")
    assert facture.config is importlib.import_module("facturation.config")
    ailleurs = {nom: getattr(facture, nom).__module__ for nom in sorted(NOMS_PUBLICS)
                if getattr(facture, nom).__module__ != "facturation.facture"}
    assert ailleurs == {}, f"noms publics définis hors de facturation.facture : {ailleurs}"


def test_t_fac_05b_le_script_reexporte_le_paquet_et_garde_main():
    """T-FAC-05 — le script expose chaque nom public du paquet (mêmes objets, classes d'exception comprises :
    `main()` attrape celles que lève `generer`) et garde son propre `main()`."""
    facture = _module()
    nf = _charger_script()
    differents = [nom for nom in sorted(NOMS_PUBLICS | {"fmt", "config"})
                  if getattr(nf, nom, None) is not getattr(facture, nom)]
    assert differents == [], f"noms du script qui ne sont pas ceux du paquet : {differents}"
    assert callable(nf.main) and nf.main.__module__ == nf.__name__


_P = inspect.Parameter.POSITIONAL_OR_KEYWORD
_N = inspect.Parameter.KEYWORD_ONLY
_SANS = inspect.Parameter.empty

# Signatures du contrat (nom, nature, valeur par défaut) : `entite`, `prefixe`, `racine`, `emetteur`… nommés.
SIGNATURES = {
    "echeance_de_la_spec": [("spec", _P, _SANS)],
    "charger_spec": [("chemin", _P, _SANS)],
    "controler_spec": [("spec", _P, _SANS), ("prefixe", _N, _SANS), ("racine", _N, _SANS)],
    "construire_groupes": [("lignes", _P, _SANS)],
    "conflits_ecriture": [("wb", _P, _SANS), ("onglet", _P, _SANS), ("n_facture", _P, _SANS), ("entite", _N, _SANS)],
    "apercu": [
        ("spec", _P, _SANS), ("groupes", _P, _SANS), ("dest", _P, _SANS), ("conflits", _P, None),
        ("emetteur", _N, _SANS), ("echeance", _N, None), ("nda", _N, None),
    ],
    "generer": [("spec", _P, _SANS), ("dry_run", _P, False), ("prefixe", _N, None)],
    "ecrire_onglet": [("wb", _P, _SANS), ("spec", _P, _SANS), ("dest", _P, _SANS), ("entite", _N, _SANS)],
    "verifier_onglet": [("ws", _P, _SANS), ("spec", _P, _SANS), ("dest", _P, _SANS), ("entite", _N, _SANS)],
}


def test_t_fac_13_signatures_du_contrat():
    """T-FAC-13 — chaque fonction publique a la signature du contrat : mêmes paramètres, dans le même ordre,
    de même nature (l'entité, le préfixe, la racine et l'émetteur se passent par leur nom) et mêmes défauts."""
    facture = _module()
    differentes = {}
    for nom, attendue in SIGNATURES.items():
        obtenue = [(p.name, p.kind, p.default) for p in inspect.signature(getattr(facture, nom)).parameters.values()]
        if obtenue != attendue:
            differentes[nom] = obtenue
    assert differentes == {}, f"signatures différentes du contrat : {differentes}"


def test_t_fac_11_un_patch_de_nf_fmt_atteint_ecrire_onglet(racine_fictive, dest, monkeypatch):
    """T-FAC-11 — `nf.fmt` est le module `pcf_format` lui-même : un patch de `nf.fmt.apply_rib` (comme le fait
    la banque existante) atteint `ecrire_onglet`, et `verifier_onglet` signale l'IBAN faux qu'il écrit."""
    facture = _module()
    nf = _charger_script()
    assert nf.fmt is facture.fmt is fmt
    vrai_apply_rib = nf.fmt.apply_rib
    appels = []

    def apply_rib_fautif(ws, **kwargs):
        vrai_apply_rib(ws, **kwargs)
        appels.append(ws.title)
        for coordonnee in _cellules_valant(ws, IBAN_FICTIF):
            ws[coordonnee] = "FR76 IBAN FAUX"

    monkeypatch.setattr(nf.fmt, "apply_rib", apply_rib_fautif)
    spec = _spec()
    ws = facture.ecrire_onglet(_nouveau_classeur(), spec, dest, entite=ENTITE)

    assert appels == [spec["onglet"]]
    ecarts = facture.verifier_onglet(ws, spec, dest, entite=ENTITE)
    assert _cles(ecarts) == ["iban"], ecarts


# ──────────────────────────────────────────────────────────────────────
# ecrire_onglet et verifier_onglet
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize("nom", list(SPECS))
def test_t_fac_03_onglet_ecrit_puis_relu_conforme(racine_fictive, dest, tmp_path, nom):
    """T-FAC-03 — sur une spec fictive, `ecrire_onglet` suivi de `verifier_onglet` sur l'onglet relu du fichier
    enregistré donne `[]` (spec complète, spec réduite aux clés obligatoires, 62 séances)."""
    facture = _module()
    spec = _spec(nom)
    ws_relu = _ecrire_et_relire(facture, spec, dest, tmp_path)
    assert facture.verifier_onglet(ws_relu, spec, dest, entite=ENTITE) == []


def test_t_fac_04_iban_altere_signale(racine_fictive, dest, tmp_path):
    """T-FAC-04 — un IBAN altéré dans l'onglet relu est signalé par `verifier_onglet` (garde cassée : le même
    onglet, conforme avant l'altération, donne un seul écart, sur l'IBAN, qui cite la valeur trouvée)."""
    facture = _module()
    spec = _spec()
    ws_relu = _ecrire_et_relire(facture, spec, dest, tmp_path)
    assert facture.verifier_onglet(ws_relu, spec, dest, entite=ENTITE) == []

    cellules = _cellules_valant(ws_relu, IBAN_FICTIF)
    assert len(cellules) == 1, cellules
    ws_relu[cellules[0]] = "FR76 9999 9999 9999 9999 9999 999"

    ecarts = facture.verifier_onglet(ws_relu, spec, dest, entite=ENTITE)
    assert _cles(ecarts) == ["iban"], ecarts
    assert "FR76 9999 9999 9999 9999 9999 999" in ecarts[0]


def _onglet_de_reference(wb: Workbook, spec: dict, dest: dict):
    """La séquence d'écriture de `generer` avant P-0a, recopiée appel par appel (f487cce)."""
    date_emission = date.fromisoformat(spec["date_emission"])
    echeance = fmt.calculer_echeance(
        date_emission, date.fromisoformat(spec["echeance"]) if spec.get("echeance") else None
    )
    groupes = [
        (ligne["promotion"], ligne["libelle"],
         [(date.fromisoformat(s["date"]), float(s["heures"]), float(s["taux"])) for s in ligne["seances"]])
        for ligne in spec["lignes"]
    ]
    ws = fmt.clone_template(wb, spec["onglet"], entite=ENTITE)
    fmt.apply_emetteur(ws, entite=ENTITE)
    fmt.apply_destinataire(ws, dest=dest, entite=ENTITE)
    fmt.apply_metadata(
        ws, n_facture=spec["n_facture"], date_emission=date_emission, entite=ENTITE,
        ref_bon_commande=spec.get("ref_bon_commande"), ref_contrat=spec.get("ref_contrat") or None,
    )
    fmt.place_sessions(
        ws, ecole_name=spec.get("ecole_label", spec["ecole"]), groups=groupes, entite=ENTITE,
        masquer_jour=bool(spec.get("masquer_jour", False)),
    )
    fmt.apply_mentions(ws, entite=ENTITE, date_emission=date_emission, echeance=echeance)
    fmt.apply_rib(ws, entite=ENTITE)
    fmt.apply_print_setup(ws, entite=ENTITE)
    return ws


def _photo(ws) -> dict:
    """Ce qui fait l'onglet : valeurs et formats, fusions, lignes et colonnes (taille, masquage), mise en page."""
    valeurs, formats = {}, {}
    for ligne in ws.iter_rows():
        for cellule in ligne:
            if cellule.value is not None:
                valeurs[cellule.coordinate] = cellule.value
                formats[cellule.coordinate] = cellule.number_format
    return {
        "titre": ws.title,
        "valeurs": valeurs,
        "formats": formats,
        "fusions": sorted(str(plage) for plage in ws.merged_cells.ranges),
        "colonnes": {k: (d.width, bool(d.hidden)) for k, d in sorted(ws.column_dimensions.items())},
        "lignes": {k: (d.height, bool(d.hidden)) for k, d in sorted(ws.row_dimensions.items())},
        "zone_impression": ws.print_area,
        "titres_imprimes": ws.print_title_rows,
        "sauts": [saut.id for saut in ws.row_breaks.brk],
        "mise_en_page": (
            ws.page_setup.orientation, ws.page_setup.paperSize, ws.page_setup.fitToWidth,
            ws.page_setup.fitToHeight, getattr(ws.sheet_properties.pageSetUpPr, "fitToPage", None),
        ),
    }


@pytest.mark.parametrize("nom", ["complete", "62-seances"])
def test_t_fac_06_ecrire_onglet_suit_la_sequence_de_generer(racine_fictive, dest, nom):
    """T-FAC-06 — `ecrire_onglet` donne le même onglet que la séquence actuelle de `generer` : émetteur,
    destinataire, métadonnées (bon de commande, contrat), séances (`ecole_label`, sinon le code de l'école ;
    colonne Jour masquée ou non), mentions et échéance de la spec, RIB, mise en page (sauts de page compris,
    au-delà de 60 séances)."""
    facture = _module()
    spec = _spec(nom)
    attendu = _photo(_onglet_de_reference(_nouveau_classeur(), spec, dest))
    obtenu = _photo(facture.ecrire_onglet(_nouveau_classeur(), spec, dest, entite=ENTITE))
    differences = [cle for cle in attendu if obtenu[cle] != attendu[cle]]
    assert differences == [], f"l'onglet diffère de la séquence de generer sur : {differences}"


# La séquence de `generer` avant P-0a, que le contrat demande à `ecrire_onglet` d'exécuter telle quelle.
SEQUENCE_DE_GENERER = [
    "clone_template", "apply_emetteur", "apply_destinataire", "apply_metadata",
    "place_sessions", "apply_mentions", "apply_rib", "apply_print_setup",
]


@pytest.mark.parametrize(("nom", "echeance_attendue"), [("complete", date(2026, 11, 15)), ("minimale", date(2026, 11, 30))])
def test_t_fac_06b_ordre_et_arguments_des_appels(racine_fictive, dest, monkeypatch, nom, echeance_attendue):
    """T-FAC-06 — `ecrire_onglet` appelle les fonctions de `pcf_format` par le module (un patch les atteint),
    une fois chacune, dans l'ordre de `generer`, toutes sur l'onglet cloné et avec l'entité reçue, et leur passe
    les valeurs de la spec : bon de commande et contrat (ou rien), libellé d'école (sinon le code), groupes de
    séances, colonne Jour masquée ou non, date d'émission et échéance (celle de la spec, sinon émission + 30 j)."""
    facture = _module()
    appels = []
    for fonction in SEQUENCE_DE_GENERER:
        _espionner(monkeypatch, fmt, fonction, appels)
    spec = _spec(nom)

    ws = facture.ecrire_onglet(_nouveau_classeur(), spec, dest, entite=ENTITE)

    assert [fonction for fonction, _, _ in appels] == SEQUENCE_DE_GENERER
    arguments = {fonction: args for fonction, args, _ in appels}
    assert appels[0][2] is ws and arguments["clone_template"]["sheet_name"] == spec["onglet"]
    assert [f for f, args, _ in appels[1:] if args["ws"] is not ws] == []
    assert [f for f, args, _ in appels if args["entite"] != ENTITE] == []
    date_emission = date.fromisoformat(spec["date_emission"])
    assert arguments["apply_destinataire"]["dest"] == dest
    attendus = {
        "apply_metadata": {
            "n_facture": spec["n_facture"], "date_emission": date_emission,
            "ref_bon_commande": spec.get("ref_bon_commande"), "ref_contrat": spec.get("ref_contrat"),
        },
        "place_sessions": {
            "ecole_name": spec.get("ecole_label", "ABC"),
            "groups": [
                (ligne["promotion"], ligne["libelle"],
                 [(date.fromisoformat(s["date"]), float(s["heures"]), float(s["taux"])) for s in ligne["seances"]])
                for ligne in spec["lignes"]
            ],
            "masquer_jour": bool(spec.get("masquer_jour")),
        },
        "apply_mentions": {"date_emission": date_emission, "echeance": echeance_attendue},
    }
    differents = {
        (fonction, cle): arguments[fonction][cle]
        for fonction, valeurs in attendus.items() for cle, valeur in valeurs.items()
        if arguments[fonction][cle] != valeur
    }
    assert differents == {}, f"arguments différents de la séquence de generer : {differents}"


def test_t_fac_07_ecrire_onglet_n_enregistre_rien(racine_fictive, dest, tmp_path, monkeypatch):
    """T-FAC-07 — `ecrire_onglet` n'enregistre rien (aucun `save`, aucun fichier créé ni modifié) : il ajoute
    l'onglet en fin du classeur reçu, sans toucher aux autres onglets ni à la spec reçue (base de l'empreinte
    validée, §1.3), et le rend."""
    facture = _module()
    spec = _spec()
    wb = Workbook()
    ancien = wb.active
    ancien.title = "R_ABC000 Août 2026"
    ancien["D17"] = "R_ABC000"
    avant = _instantane(tmp_path)

    def enregistrement_interdit(self, *args, **kwargs):
        raise AssertionError("ecrire_onglet a enregistré un classeur")

    monkeypatch.setattr(Workbook, "save", enregistrement_interdit)
    ws = facture.ecrire_onglet(wb, spec, dest, entite=ENTITE)

    assert ws.parent is wb and ws.title == spec["onglet"]
    assert wb.sheetnames == ["R_ABC000 Août 2026", spec["onglet"]]
    assert wb["R_ABC000 Août 2026"]["D17"].value == "R_ABC000"
    assert _instantane(tmp_path) == avant
    assert spec == _spec(), "ecrire_onglet a modifié la spec reçue"


def _retirer(cle: str):
    def modifier(spec: dict):
        del spec[cle]
    return modifier


def _poser(cle: str, valeur):
    def modifier(spec: dict):
        spec[cle] = valeur
    return modifier


def _changer_taux(spec: dict):
    spec["lignes"][1]["seances"][0]["taux"] = 61.5


def _changer_heures(spec: dict):
    spec["lignes"][0]["seances"][1]["heures"] = 4


def _changer_premier_libelle(spec: dict):
    spec["lignes"][0]["libelle"] = "[2026][M1 S1][Module fictif modifié]"


# (nom du cas, modification de la spec de contrôle, clé de l'écart attendu). L'onglet est écrit une fois.
CAS_SPEC_DIFFERENTE = [
    ("ecole-label", _poser("ecole_label", "AUTRE ÉCOLE FICTIVE"), "ecole"),
    ("sans-ecole-label", _retirer("ecole_label"), "ecole"),  # l'école attendue redevient le code « ABC »
    ("echeance", _poser("echeance", "2026-12-01"), "echeance"),
    ("sans-echeance", _retirer("echeance"), "echeance"),  # attendue par défaut : émission + 30 jours
    ("contrat", _poser("ref_contrat", "CT-EXEMPLE-0002"), "contrat"),
    ("sans-contrat", _retirer("ref_contrat"), "contrat"),  # sans contrat, la ligne de contrat doit être vide
    ("numero", _poser("n_facture", "R_ABC009"), "n_facture"),
    ("date-emission", _poser("date_emission", "2026-09-29"), "titre"),
    ("taux", _changer_taux, "taux"),
    ("heures", _changer_heures, "heures"),
    ("libelle", _changer_premier_libelle, "premier_libelle"),
]


def test_t_fac_08_valeurs_attendues_deduites_de_la_spec_et_du_destinataire(racine_fictive, dest):
    """T-FAC-08 — `verifier_onglet` déduit ses valeurs attendues de la spec et du destinataire qu'il reçoit : un
    onglet écrit pour une spec, conforme à cette spec, donne un écart sur le point où diffère la spec de contrôle
    (libellé d'école ou son absence, échéance ou son absence, contrat ou son absence, numéro, date d'émission,
    taux, heures, libellé), et un seul écart, sur le nom, avec un autre destinataire. Le contrôle ne modifie
    ni la spec ni le destinataire reçus."""
    facture = _module()
    spec = _spec()
    ws = facture.ecrire_onglet(_nouveau_classeur(), spec, dest, entite=ENTITE)
    dest_avant = copy.deepcopy(dest)
    assert facture.verifier_onglet(ws, spec, dest, entite=ENTITE) == []
    assert spec == _spec() and dest == dest_avant, "verifier_onglet a modifié la spec ou le destinataire reçus"

    non_signales = {}
    for nom, modifier, cle in CAS_SPEC_DIFFERENTE:
        autre = _spec()
        modifier(autre)
        ecarts = facture.verifier_onglet(ws, autre, dest, entite=ENTITE)
        if cle not in _cles(ecarts):
            non_signales[nom] = ecarts
    ecarts = facture.verifier_onglet(ws, spec, {**dest, "nom": "AUTRE DESTINATAIRE FICTIF"}, entite=ENTITE)
    if _cles(ecarts) != ["destinataire_nom"]:
        non_signales["destinataire"] = ecarts
    assert non_signales == {}, f"différences non signalées par verifier_onglet : {non_signales}"


def test_t_fac_09_echeance_invalide_refusee(racine_fictive, dest):
    """T-FAC-09 — échéance de la spec antérieure à l'émission (30/09/2026) ou hors format AAAA-MM-JJ :
    `ecrire_onglet` et `verifier_onglet` lèvent la `ValueError` d'`echeance_de_la_spec`, qui nomme la cause."""
    facture = _module()
    with pytest.raises(ValueError, match="antérieure"):
        facture.ecrire_onglet(_nouveau_classeur(), _spec(echeance="2026-09-29"), dest, entite=ENTITE)
    ws = facture.ecrire_onglet(_nouveau_classeur(), _spec(), dest, entite=ENTITE)
    for echeance, cause in (("2026-09-29", "antérieure"), ("30/10/2026", "30/10/2026")):
        with pytest.raises(ValueError, match=cause):
            facture.verifier_onglet(ws, _spec(echeance=echeance), dest, entite=ENTITE)


# ──────────────────────────────────────────────────────────────────────
# generer passe par ecrire_onglet et verifier_onglet
# ──────────────────────────────────────────────────────────────────────


def _espionner(monkeypatch, module, nom: str, appels: list) -> None:
    """Remplace `module.<nom>` par un espion qui note ses arguments (par leur nom, défauts compris : passer
    `masquer_jour=False` ou l'omettre revient au même) et son résultat."""
    vraie = getattr(module, nom)
    signature = inspect.signature(vraie)

    def espion(*args, **kwargs):
        liaison = signature.bind(*args, **kwargs)
        liaison.apply_defaults()
        arguments = dict(liaison.arguments)
        resultat = vraie(*args, **kwargs)
        appels.append((nom, arguments, resultat))
        return resultat

    monkeypatch.setattr(module, nom, espion)


def test_t_fac_10a_generer_ecrit_par_ecrire_onglet_et_controle_l_onglet_relu(racine_fictive, monkeypatch, capsys):
    """T-FAC-10 — `generer` écrit par `ecrire_onglet`, puis contrôle par `verifier_onglet` l'onglet relu du
    fichier enregistré (un autre objet que l'onglet en mémoire), avec la même spec, le destinataire du MASTER
    et l'entité de la spec ; il rend le post-mortem et annonce la facture conforme."""
    facture = _module()
    appels = []
    _espionner(monkeypatch, facture, "ecrire_onglet", appels)
    _espionner(monkeypatch, facture, "verifier_onglet", appels)
    spec = _spec()

    post_mortem = facture.generer(spec)

    assert [nom for nom, _, _ in appels] == ["ecrire_onglet", "verifier_onglet"]
    (_, ecrire, ws_ecrit), (_, verifier, ecarts) = appels
    dest_master = fmt.get_destinataire("ABC", entite=ENTITE)
    assert ecrire["spec"] == spec and verifier["spec"] == spec
    assert ecrire["dest"] == dest_master and verifier["dest"] == dest_master
    assert ecrire["entite"] == verifier["entite"] == ENTITE
    assert ecarts == []
    ws_verifie = verifier["ws"]
    assert ws_verifie.title == ws_ecrit.title == spec["onglet"]
    assert ws_verifie is not ws_ecrit and ws_verifie.parent is not ecrire["wb"], (
        "verifier_onglet doit porter sur l'onglet relu du fichier enregistré, pas sur l'onglet en mémoire"
    )
    assert (racine_fictive / spec["fichier_sortie"]).exists()
    assert post_mortem["n_facture"] == spec["n_facture"]
    assert "OK — facture R_ABC001 conforme" in capsys.readouterr().out


def test_t_fac_10b_ecarts_de_verifier_onglet_levent_post_mortem_en_ecart(racine_fictive, monkeypatch, capsys):
    """T-FAC-10 — les écarts rendus par `verifier_onglet` font lever `PostMortemEnEcart` par `generer`, avec ces
    écarts tels quels ; la facture reste écrite, mais n'est pas annoncée conforme (comportement inchangé)."""
    facture = _module()
    ecarts_simules = ["ERREUR sur cellule iban : attendu 'FR76 0000', trouvé 'FR76 9999'"]
    monkeypatch.setattr(facture, "verifier_onglet", lambda *args, **kwargs: list(ecarts_simules))
    spec = _spec()

    with pytest.raises(facture.PostMortemEnEcart) as erreur:
        facture.generer(spec)

    assert erreur.value.ecarts == ecarts_simules
    assert (racine_fictive / spec["fichier_sortie"]).exists()
    assert "OK — facture" not in capsys.readouterr().out


def test_t_fac_10c_dry_run_sans_ecriture(racine_fictive, monkeypatch, capsys):
    """T-FAC-10 — en `dry_run`, `generer` affiche l'aperçu, rend `{}` et n'appelle pas `ecrire_onglet`."""
    facture = _module()

    def ecriture_interdite(*args, **kwargs):
        raise AssertionError("dry-run : ecrire_onglet ne doit pas être appelé")

    monkeypatch.setattr(facture, "ecrire_onglet", ecriture_interdite)

    assert facture.generer(_spec(), dry_run=True) == {}
    sortie = capsys.readouterr().out
    assert "APERÇU FACTURE" in sortie and "[dry-run] Aucun fichier écrit." in sortie
    assert not (racine_fictive / "ROSARI" / "ABC").exists()


# ──────────────────────────────────────────────────────────────────────
# Les fonctions déplacées gardent leurs refus (type et message)
# ──────────────────────────────────────────────────────────────────────


def test_t_fac_12a_exceptions_du_module():
    """T-FAC-12 — `EcritureRefusee` et `PostMortemEnEcart` sont des `RuntimeError` ; `PostMortemEnEcart` garde
    ses écarts dans `.ecarts` et les donne pour message, un par ligne."""
    facture = _module()
    assert issubclass(facture.EcritureRefusee, RuntimeError)
    assert issubclass(facture.PostMortemEnEcart, RuntimeError)
    ecarts = ["ERREUR sur cellule iban : attendu 'A', trouvé 'B'", "ERREUR sur cellule bic : attendu 'C', trouvé 'D'"]
    erreur = facture.PostMortemEnEcart(ecarts)
    assert erreur.ecarts == ecarts
    assert str(erreur) == "\n".join(ecarts)


def test_t_fac_12b_charger_spec(tmp_path):
    """T-FAC-12 — `charger_spec` lit la spec ; des clés obligatoires manquantes ou aucune ligne de cours lèvent
    une `ValueError` qui nomme le fichier et les clés, ou l'absence de ligne."""
    facture = _module()
    complete = tmp_path / "complete.json"
    complete.write_text(json.dumps(SPEC_COMPLETE, ensure_ascii=False), encoding="utf-8")
    assert facture.charger_spec(complete) == SPEC_COMPLETE

    incomplete = tmp_path / "incomplete.json"
    incomplete.write_text(json.dumps({k: v for k, v in SPEC_COMPLETE.items() if k not in ("onglet", "lignes")}),
                          encoding="utf-8")
    with pytest.raises(ValueError, match=r"incomplete\.json") as erreur:
        facture.charger_spec(incomplete)
    assert "'lignes'" in str(erreur.value) and "'onglet'" in str(erreur.value), erreur.value

    sans_ligne = tmp_path / "sans_ligne.json"
    sans_ligne.write_text(json.dumps({**SPEC_COMPLETE, "lignes": []}), encoding="utf-8")
    with pytest.raises(ValueError, match="aucune ligne"):
        facture.charger_spec(sans_ligne)


def test_t_fac_12c_echeance_de_la_spec():
    """T-FAC-12 — `echeance_de_la_spec` : l'échéance de la spec, sinon émission + 30 jours ; une échéance
    antérieure à l'émission lève une `ValueError` qui le dit."""
    facture = _module()
    assert facture.echeance_de_la_spec({"date_emission": "2026-09-30"}) == date(2026, 10, 30)
    assert facture.echeance_de_la_spec({"date_emission": "2026-09-30", "echeance": "2026-11-15"}) == date(2026, 11, 15)
    with pytest.raises(ValueError, match="antérieure"):
        facture.echeance_de_la_spec({"date_emission": "2026-09-30", "echeance": "2026-09-29"})


def test_t_fac_12d_refus_de_controler_spec_et_de_generer(racine_fictive, capsys):
    """T-FAC-12 — `controler_spec` et `generer` refusent par `facturation.facture.EcritureRefusee`, avec un message
    qui nomme la cause (préfixe d'une autre entité, sortie hors du dossier de l'entité, onglet existant) ; le
    classeur existant reste intact."""
    facture = _module()
    assert facture.controler_spec(_spec(), prefixe=None, racine=racine_fictive) == (
        racine_fictive / "ROSARI" / "ABC" / "ROSARI_ABC.xlsx"
    )
    with pytest.raises(facture.EcritureRefusee, match="préfixe R_"):
        facture.controler_spec(_spec(n_facture="K_ABC001"), prefixe=None, racine=racine_fictive)
    with pytest.raises(facture.EcritureRefusee, match="fichier_sortie"):
        facture.controler_spec(_spec(fichier_sortie="KAUR/ABC/KAUR_ABC.xlsx"), prefixe=None, racine=racine_fictive)

    spec = _spec()
    classeur = racine_fictive / spec["fichier_sortie"]
    classeur.parent.mkdir(parents=True)
    wb = Workbook()
    wb.active.title = spec["onglet"]
    wb.active["D17"] = spec["n_facture"]
    wb.save(classeur)
    # Date de modification ramenée au 01/01/2020 : une réécriture à l'identique (mêmes octets) se voit aussi.
    os.utime(classeur, ns=(1_577_836_800 * 10**9, 1_577_836_800 * 10**9))
    avant = _instantane(classeur.parent)

    with pytest.raises(facture.EcritureRefusee, match="existe déjà"):
        facture.generer(spec)
    assert _instantane(classeur.parent) == avant, "le classeur existant a été réécrit malgré le refus"


# ──────────────────────────────────────────────────────────────────────
# Constats de revue : ecrire_onglet et verifier_onglet se gardent eux-mêmes
# ──────────────────────────────────────────────────────────────────────
#
# `generer` contrôle l'onglet, le numéro et l'entité avant d'appeler `ecrire_onglet`, et n'enregistre rien
# après une exception : aucun des cas ci-dessous ne l'atteint. Mais l'émission (F-17) appellera
# `ecrire_onglet` et `verifier_onglet` directement, avec `entite=…`, sur le classeur d'une école qui porte
# déjà d'autres factures : un contrôle oublié par l'appelant ne doit ni réutiliser un numéro, ni mêler deux
# entités, ni laisser dans le classeur une facture à moitié écrite qu'un enregistrement suivant garderait.

ONGLET_EXISTANT = "R_ABC000 Août 2026"

# Facture fictive de l'entité KAUR, sur le modèle de la spec minimale (même école fictive ABC).
SPEC_KAUR = {
    **SPEC_MINIMALE,
    "entite": "KAUR",
    "fichier_sortie": "KAUR/ABC/KAUR_ABC.xlsx",
    "onglet": "K_ABC001 Sept 2026",
    "n_facture": "K_ABC001",
    "date_emission": "2026-09-30",
}


def _classeur_avec_une_facture() -> Workbook:
    """Classeur d'école qui porte déjà une facture : onglet « R_ABC000 Août 2026 », N° R_ABC000 en D17."""
    wb = _nouveau_classeur()
    wb.create_sheet(ONGLET_EXISTANT)["D17"] = "R_ABC000"
    return wb


def _etat(wb: Workbook) -> tuple[list[str], list]:
    """Onglets du classeur, dans l'ordre, et la cellule N° de facture (D17) de chacun."""
    return list(wb.sheetnames), [wb[nom]["D17"].value for nom in wb.sheetnames]


# (titre de l'onglet déjà présent, N° de sa cellule D17) : chacun interdit d'écrire R_ABC001 (FA-02).
CAS_DEJA_PRESENTS = {
    "meme-onglet": ("R_ABC001 Sept 2026", "R_ABC001"),
    "onglet-a-la-casse-pres": ("r_abc001 sept 2026", None),
    "numero-dans-la-cellule": ("Facture septembre", "R_ABC001"),
    "numero-dans-le-titre": ("PAID R_ABC001 Sept 2026", None),
}


@pytest.mark.parametrize(("titre", "numero"), list(CAS_DEJA_PRESENTS.values()), ids=list(CAS_DEJA_PRESENTS))
def test_t_fac_14a_ecrire_onglet_refuse_un_onglet_ou_un_numero_deja_present(racine_fictive, dest, titre, numero):
    """T-FAC-14 — `ecrire_onglet` refuse par `EcritureRefusee`, sans toucher au classeur, d'écrire une facture
    dont l'onglet existe déjà (y compris à la casse près : openpyxl créerait en silence « … 20261 ») ou dont le
    numéro figure déjà dans un onglet (cellule N° ou titre) : les conflits de `conflits_ecriture`. Le message
    nomme l'onglet en conflit. Défense en profondeur : `generer` refuse déjà ces cas avant l'aperçu."""
    facture = _module()
    spec = _spec()
    wb = _classeur_avec_une_facture()
    feuille = wb.create_sheet(titre)
    if numero:
        feuille["D17"] = numero
    assert facture.conflits_ecriture(wb, spec["onglet"], spec["n_facture"], entite=ENTITE), "cas mal construit"
    avant = _etat(wb)

    with pytest.raises(facture.EcritureRefusee) as erreur:
        facture.ecrire_onglet(wb, spec, dest, entite=ENTITE)

    assert titre in str(erreur.value), erreur.value
    assert _etat(wb) == avant, "ecrire_onglet a modifié le classeur malgré le conflit"


def test_t_fac_14b_second_appel_sur_le_meme_classeur_refuse(racine_fictive, dest):
    """T-FAC-14 — `ecrire_onglet` appelé deux fois avec la même spec sur le même classeur (une reprise, plusieurs
    factures d'une même école) : le second appel lève `EcritureRefusee` qui nomme l'onglet, le classeur garde un
    seul onglet à ce numéro, et la facture déjà écrite reste conforme."""
    facture = _module()
    spec = _spec()
    wb = _classeur_avec_une_facture()
    ws = facture.ecrire_onglet(wb, spec, dest, entite=ENTITE)
    avant = _etat(wb)

    with pytest.raises(facture.EcritureRefusee, match=re.escape(spec["onglet"])):
        facture.ecrire_onglet(wb, spec, dest, entite=ENTITE)

    assert _etat(wb) == avant and wb.sheetnames == [ONGLET_EXISTANT, spec["onglet"]], wb.sheetnames
    assert facture.verifier_onglet(ws, spec, dest, entite=ENTITE) == []


def test_t_fac_15_verifier_onglet_controle_le_titre_de_l_onglet(racine_fictive, dest):
    """T-FAC-15 — `verifier_onglet` compare le titre de l'onglet à `spec["onglet"]` : l'onglet suffixé
    « R_ABC001 Sept 20261 » (ce qu'openpyxl crée quand le titre est déjà pris) donne un seul écart, de clé
    `onglet`, qui cite le titre attendu et le titre trouvé. Garde : le même onglet sous son titre est conforme."""
    facture = _module()
    spec = _spec()
    ws = facture.ecrire_onglet(_nouveau_classeur(), spec, dest, entite=ENTITE)
    assert facture.verifier_onglet(ws, spec, dest, entite=ENTITE) == []

    ws.title = spec["onglet"] + "1"
    ecarts = facture.verifier_onglet(ws, spec, dest, entite=ENTITE)

    assert _cles(ecarts) == ["onglet"], ecarts
    assert f"'{spec['onglet']}'" in ecarts[0] and f"'{spec['onglet']}1'" in ecarts[0], ecarts


@pytest.mark.parametrize(("entite_spec", "entite_appel"), [("ROSARI", "KAUR"), ("KAUR", "ROSARI")])
def test_t_fac_16_entite_de_l_appel_differente_de_celle_de_la_spec(racine_fictive, entite_spec, entite_appel):
    """T-FAC-16 — les entités ne se mélangent jamais (ADR001, contrats §1.1) : `ecrire_onglet` et
    `verifier_onglet` appelés avec une entité différente de `spec["entite"]` lèvent une `ValueError` qui nomme
    les deux entités. `ecrire_onglet` ne touche pas au classeur ; sans ce refus, il écrirait l'émetteur et le
    RIB de l'une sous le numéro de l'autre, et `verifier_onglet`, qui lit ses attendus dans le MASTER de
    l'entité reçue, ne verrait rien. Garde : chaque onglet est conforme sous sa propre entité."""
    facture = _module()
    creer_master(racine_fictive, "KAUR")
    fmt.set_data_root(racine_fictive)
    spec = _spec() if entite_spec == ENTITE else copy.deepcopy(SPEC_KAUR)
    dest = fmt.get_destinataire("ABC", entite=entite_spec)
    wb = _classeur_avec_une_facture()
    avant = _etat(wb)

    with pytest.raises(ValueError) as erreur:
        facture.ecrire_onglet(wb, spec, dest, entite=entite_appel)
    assert entite_spec in str(erreur.value) and entite_appel in str(erreur.value), erreur.value
    assert _etat(wb) == avant, "ecrire_onglet a écrit dans le classeur pour une autre entité"

    ws = facture.ecrire_onglet(_nouveau_classeur(), spec, dest, entite=entite_spec)
    assert facture.verifier_onglet(ws, spec, dest, entite=entite_spec) == []
    with pytest.raises(ValueError) as erreur:
        facture.verifier_onglet(ws, spec, dest, entite=entite_appel)
    assert entite_spec in str(erreur.value) and entite_appel in str(erreur.value), erreur.value


def _sans(*cles: str):
    def modifier(spec: dict):
        for cle in cles:
            del spec[cle]
    return modifier


def _rien(*args) -> None:
    return None


def _master_sans_nda(racine: Path, monkeypatch) -> None:
    creer_master(racine, nda=None)
    fmt.set_data_root(racine)  # vide le cache du MASTER


def _apply_rib_en_echec(racine: Path, monkeypatch) -> None:
    def apply_rib(ws, **kwargs):
        raise RuntimeError("IBAN absent de MASTER.Coordonnées (simulé)")

    monkeypatch.setattr(fmt, "apply_rib", apply_rib)


# (modification de la spec, préparation de la racine, erreurs admises, motif du message). Toutes, sauf la
# dernière, surviennent après `clone_template` dans le code relu en revue ; l'échéance antérieure, déjà lue
# avant la première écriture, sert de témoin.
CAS_ERREUR_EN_COURS_D_ECRITURE = {
    "spec-sans-n_facture": (_sans("n_facture"), _rien, (KeyError, ValueError), "n_facture"),
    "spec-sans-ecole-ni-libelle": (_sans("ecole", "ecole_label"), _rien, (KeyError, ValueError), "ecole"),
    "master-sans-nda": (_rien, _master_sans_nda, ValueError, "NDA"),
    "apply-rib-en-echec": (_rien, _apply_rib_en_echec, RuntimeError, "IBAN"),
    "echeance-anterieure": (_poser("echeance", "2026-09-29"), _rien, ValueError, "antérieure"),
}


@pytest.mark.parametrize(
    ("modifier_spec", "preparer", "erreurs", "motif"),
    list(CAS_ERREUR_EN_COURS_D_ECRITURE.values()),
    ids=list(CAS_ERREUR_EN_COURS_D_ECRITURE),
)
def test_t_fac_17_erreur_en_cours_d_ecriture_classeur_intact(
    racine_fictive, dest, monkeypatch, modifier_spec, preparer, erreurs, motif
):
    """T-FAC-17 — la docstring d'`ecrire_onglet` promet qu'une spec invalide lève son erreur sans toucher au
    classeur. Toute erreur, de la spec (`n_facture` absent, ni `ecole` ni `ecole_label`) ou du MASTER (ligne
    Info / NDA absente, RIB incomplet), laisse le classeur tel qu'il était : mêmes onglets, la facture déjà
    présente intacte, aucun onglet à moitié écrit qu'un appelant qui attrape l'erreur enregistrerait ensuite."""
    facture = _module()
    spec = _spec()
    modifier_spec(spec)
    preparer(racine_fictive, monkeypatch)
    wb = _classeur_avec_une_facture()
    avant = _etat(wb)

    with pytest.raises(erreurs, match=motif):
        facture.ecrire_onglet(wb, spec, dest, entite=ENTITE)

    assert _etat(wb) == avant, "ecrire_onglet a laissé un onglet à moitié écrit dans le classeur"


# ──────────────────────────────────────────────────────────────────────
# fix-facture — contrôles de la facture laissés hors de P-0a (T-FAC-18 à T-FAC-25)
# ──────────────────────────────────────────────────────────────────────
#
# La revue de P-0a a relevé des défauts antérieurs que son contrat (comportement gelé) interdisait de
# corriger : aperçu arrondi sur le float binaire, ligne sans séance écrite (ou déclarée non conforme par un
# faux écart), séances mal formées converties en silence, spec avec BOM refusée, titre d'onglet trop long
# écrit quand même. Tests d'acceptation du correctif (contrats §18), rouges avant lui (TDD strict, bug fix).
#
# Les refus de structure vivent dans `charger_spec` (ValueError) ou dans `controler_spec` (EcritureRefusee) :
# les tests passent donc la spec par le chemin du script (JSON lu par `charger_spec`, puis `generer`) et
# admettent l'un ou l'autre type. Le message nomme la cause, le rang de la ligne (« ligne 2 », compté à partir
# de 1) et, pour une séance, son rang dans sa ligne (« séance 3 ») et la valeur fautive. Le côté script (code 1,
# « [ERREUR] », sans trace Python) est à la fin de `tests/integration/test_scripts.py`.

DEST_FICTIF = {"nom": "ÉCOLE FICTIVE A", "adresse_l1": "10 avenue Test", "adresse_l2": "00000 TESTVILLE"}
EMETTEUR_FICTIF = {
    "nom": "FORMATEUR FICTIF ROSARI EI", "adresse_l1": "1 rue de l'Exemple,", "adresse_l2": "00000 Villefictive",
}


def _apercu_des_groupes(facture, spec: dict) -> str:
    """Aperçu de la spec, groupes construits comme dans `generer` (heures et taux en float)."""
    return facture.apercu(
        spec, facture.construire_groupes(spec["lignes"]), DEST_FICTIF, emetteur=EMETTEUR_FICTIF, nda="NDA fictif",
    )


def _apercu_d_une_ligne(facture, seances: list[tuple]) -> str:
    """Aperçu d'une facture d'une ligne, avec une séance par couple (heures, taux)."""
    lignes = [{"promotion": "M1 S1", "libelle": "Module A",
               "seances": [_seance(f"2026-09-{1 + i:02d}", h, t) for i, (h, t) in enumerate(seances)]}]
    return _apercu_des_groupes(facture, _spec(lignes=lignes))


def _montants_affiches(texte: str) -> list[str]:
    """Dernière colonne de chaque ligne de séance de l'aperçu, telle qu'affichée (« |     50.00 € »)."""
    return [ligne[ligne.rindex("|"):] for ligne in texte.splitlines() if ligne.startswith("  | ")]


def _colonne_montant(montant: str) -> str:
    """Colonne montant au format d'affichage inchangé : 9 caractères alignés à droite, puis « € »."""
    return f"| {montant:>9} €"


def test_t_fac_18a_apercu_arrondi_au_centime_sur_la_valeur_decimale():
    """T-FAC-18 — l'aperçu calcule en décimal à partir de la valeur écrite (jamais du float binaire) et arrondit au
    centime, la moitié vers le haut : 1,5 h × 33,33 € = 49,995 € s'affiche 50.00 € sur la ligne et au TOTAL HT, au
    format inchangé. Avant le correctif, le float binaire (49,99499…) donnait 49.99 €."""
    facture = _module()
    texte = _apercu_d_une_ligne(facture, [(1.5, 33.33)])
    assert _montants_affiches(texte) == [_colonne_montant("50.00")], texte
    assert "TOTAL HT     : 50.00 €" in texte.splitlines(), texte


def test_t_fac_18b_total_ht_arrondi_une_seule_fois():
    """T-FAC-18 — TOTAL HT = arrondi(Σ heures × taux), une seule fois, comme le `=SUM` du gabarit qui additionne des
    produits non arrondis : trois séances de 1 h à 0,005 € donnent 0.01 € par ligne et 0.02 € au total (ni 0.03,
    somme des lignes arrondies, ni 0.01, arrondi du float binaire avant le correctif). Garde : les montants de la
    spec complète ne bougent pas (270.00, 270.00 et 90.75 ; total 630.75)."""
    facture = _module()
    complete = _apercu_des_groupes(facture, _spec())
    assert _montants_affiches(complete) == [_colonne_montant(m) for m in ("270.00", "270.00", "90.75")], complete
    assert "TOTAL HT     : 630.75 €" in complete.splitlines(), complete

    texte = _apercu_d_une_ligne(facture, [(1, 0.005)] * 3)
    assert _montants_affiches(texte) == [_colonne_montant("0.01")] * 3, texte
    assert "TOTAL HT     : 0.02 €" in texte.splitlines(), texte


# Deux lignes, la seconde à trois séances : une séance fautive y est la « séance 3 » de la « ligne 2 » (rangs
# comptés à partir de 1 ; à partir de 0, le message dirait « ligne 1 » et « séance 2 »).
SPEC_SEANCES = {
    **SPEC_MINIMALE,
    "onglet": "R_ABC001 Sept 2026",
    "n_facture": "R_ABC001",
    "date_emission": "2026-09-30",
    "lignes": [
        {"promotion": "M1 S1", "libelle": "[2026][M1 S1][Module A][3H]", "seances": [_seance("2026-09-15", 3, 90)]},
        {"promotion": "M2 S1", "libelle": "[2026][M2 S1][Module B][6H]",
         "seances": [_seance("2026-09-16", 2, 60), _seance("2026-09-17", 2, 60), _seance("2026-09-18", 2, 60)]},
    ],
}

# Classeur de l'école fictive ABC sous la racine de données, celui de toutes les specs fictives de ce fichier.
CLASSEUR_ABC = Path("ROSARI", "ABC", "ROSARI_ABC.xlsx")


def _spec_seances(numero: int = 1) -> dict:
    """Copie indépendante de SPEC_SEANCES, sous le numéro R_ABC<numero> (onglet « R_ABC<numero> Sept 2026 »)."""
    spec = copy.deepcopy(SPEC_SEANCES)
    spec["n_facture"] = f"R_ABC{numero:03d}"
    spec["onglet"] = f"R_ABC{numero:03d} Sept 2026"
    return spec


def _classeur_existant(racine: Path) -> Path:
    """Classeur de l'école ABC qui porte déjà une facture (« R_ABC000 Août 2026 », N° R_ABC000 en D17), créé s'il
    n'existe pas, daté du 01/01/2020 : une réécriture à l'identique (mêmes octets) se voit aussi."""
    classeur = racine / CLASSEUR_ABC
    if not classeur.exists():
        classeur.parent.mkdir(parents=True, exist_ok=True)
        _classeur_avec_une_facture().save(classeur)
    os.utime(classeur, ns=(1_577_836_800 * 10**9, 1_577_836_800 * 10**9))
    return classeur


def _refus_avant_ecriture(facture, spec: dict, racine: Path, dossier: Path) -> str:
    """Passe la spec par le chemin du script — JSON écrit dans `dossier`, lu par `charger_spec`, puis donné à
    `generer` — et attend un refus avant toute écriture : `ValueError` (refus de `charger_spec`) ou `EcritureRefusee`
    (refus de `controler_spec`). Le classeur de l'école, qui porte déjà une facture, garde ses octets et sa date
    (aucun onglet ajouté), et aucun fichier n'apparaît sous la racine. Rend le message du refus."""
    _classeur_existant(racine)
    avant = _instantane(racine)
    chemin = dossier / "spec_refusee.json"
    chemin.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
    with pytest.raises((ValueError, facture.EcritureRefusee)) as erreur:
        facture.generer(facture.charger_spec(chemin))
    assert _instantane(racine) == avant, "la spec refusée a modifié la racine de données (onglet ou fichier écrit)"
    return str(erreur.value)


def _verifier_rang(message: str, *, ligne: int, seance: int | None = None) -> None:
    """Le message nomme le rang de la ligne (« ligne 2 ») et, s'il y a lieu, celui de la séance dans sa ligne
    (« séance 3 »), comptés à partir de 1 ; « n° » est admis devant le nombre."""
    assert re.search(rf"\bligne (n° ?)?{ligne}\b", message, re.IGNORECASE), f"« ligne {ligne} » absent : {message!r}"
    if seance is not None:
        assert re.search(rf"\bs[ée]ance (n° ?)?{seance}\b", message, re.IGNORECASE), (
            f"« séance {seance} » absent : {message!r}"
        )


def _formes_de(valeur) -> set[str]:
    """Écritures admises de la valeur fautive dans un message : Python (repr, str), JSON, et virgule décimale.

    Une chaîne se cite entre guillemets (Python, JSON ou « … ») : nue, elle se confondrait avec un nombre (« 18 »)
    ou avec le texte fixe du message, et l'assertion passerait à vide (revue de fix-facture : « attendu true ou
    false » contient déjà false)."""
    if isinstance(valeur, str):
        return {repr(valeur), json.dumps(valeur, ensure_ascii=False), f"« {valeur} »", f"«{valeur}»"}
    formes = {repr(valeur), str(valeur), json.dumps(valeur, ensure_ascii=False)}
    if isinstance(valeur, float):
        formes.add(str(valeur).replace(".", ","))
    return formes


def _verifier_valeur(message: str, valeur) -> None:
    """Le message cite la valeur fautive, comme un mot entier (« 0 » ne se trouve pas dans « R_ABC001 »)."""
    cites = [f for f in _formes_de(valeur) if re.search(rf"(?<![\w.]){re.escape(f)}(?!\w)", message, re.IGNORECASE)]
    assert cites, f"valeur fautive {valeur!r} absente du message : {message!r}"


def _vider_les_seances(ligne: dict) -> None:
    ligne["seances"] = []


def _oter_les_seances(ligne: dict) -> None:
    del ligne["seances"]


@pytest.mark.parametrize("rang", [1, 2], ids=["premiere-ligne", "derniere-ligne"])
@pytest.mark.parametrize("retirer", [_vider_les_seances, _oter_les_seances], ids=["seances-vide", "seances-absente"])
def test_t_fac_19_ligne_sans_seance_refusee(racine_fictive, tmp_path, retirer, rang):
    """T-FAC-19 — une ligne dont `seances` est vide ou absente est refusée avant toute écriture, par un message qui
    nomme son rang et son libellé : aucun onglet ajouté au classeur de l'école, aucun fichier écrit. Avant le
    correctif, `seances` vide était écrite (la ligne disparaissait de la facture) et, en première ligne, déclarée
    non conforme par un faux écart sur le premier libellé ; `seances` absente levait une KeyError brute."""
    facture = _module()
    spec = _spec_seances()
    ligne = spec["lignes"][rang - 1]
    retirer(ligne)

    message = _refus_avant_ecriture(facture, spec, racine_fictive, tmp_path)

    _verifier_rang(message, ligne=rang)
    assert ligne["libelle"] in message, message


# (champ de la séance, valeur fautive) : rien de tout cela n'est un nombre fini.
CAS_HEURES_OU_TAUX_NON_NUMERIQUES = {
    "heures-chaine-virgule": ("heures", "1,5"),
    "heures-chaine-chiffres": ("heures", "18"),  # jamais de conversion silencieuse d'une chaîne en nombre
    "heures-booleen": ("heures", True),
    "heures-null": ("heures", None),
    "heures-nan": ("heures", float("nan")),
    "heures-infinies": ("heures", float("inf")),
    "taux-chaine-chiffres": ("taux", "90"),
    "taux-booleen": ("taux", False),
    "taux-null": ("taux", None),
    "taux-nan": ("taux", float("nan")),
    "taux-moins-infini": ("taux", float("-inf")),
}


@pytest.mark.parametrize(
    ("champ", "valeur"), list(CAS_HEURES_OU_TAUX_NON_NUMERIQUES.values()), ids=list(CAS_HEURES_OU_TAUX_NON_NUMERIQUES),
)
def test_t_fac_20_heures_ou_taux_qui_ne_sont_pas_des_nombres_refuses(racine_fictive, tmp_path, champ, valeur):
    """T-FAC-20 — heures ou taux qui ne sont pas des nombres finis (chaîne, même « 18 », booléen, null, NaN, infini)
    sont refusés avant toute écriture, sans jamais convertir une chaîne en nombre. Le message nomme le champ, le rang
    de la ligne, celui de la séance dans sa ligne et la valeur fautive. Avant le correctif : « 18 » et les booléens
    étaient convertis en silence, « 1,5 » levait la ValueError de `float()` sans rang, null une TypeError, et NaN ou
    l'infini étaient écrits puis déclarés non conformes."""
    facture = _module()
    spec = _spec_seances()
    spec["lignes"][1]["seances"][2][champ] = valeur

    message = _refus_avant_ecriture(facture, spec, racine_fictive, tmp_path)

    assert champ in message, message
    _verifier_rang(message, ligne=2, seance=3)
    _verifier_valeur(message, valeur)


@pytest.mark.parametrize(
    ("champ", "refusee", "admise"),
    [("heures", 0, 0.25), ("heures", -1.5, 0.25), ("taux", -0.01, 0)],
    ids=["heures-nulles", "heures-negatives", "taux-negatif"],
)
def test_t_fac_21_heures_nulles_ou_negatives_et_taux_negatif_refuses(racine_fictive, tmp_path, champ, refusee, admise):
    """T-FAC-21 — heures ≤ 0 ou taux < 0 : refus avant toute écriture, par un message qui nomme le champ, le rang de
    la ligne, celui de la séance et la valeur. Garde : la borne admise s'écrit et reste conforme (heures de 0,25 ; un
    taux à 0, accepté avant le correctif, le reste : ne rien inventer). Avant le correctif, les trois s'écrivaient."""
    facture = _module()
    _classeur_existant(racine_fictive)
    garde = _spec_seances(1)
    garde["lignes"][1]["seances"][2][champ] = admise
    facture.generer(garde)  # lève PostMortemEnEcart si la facture écrite n'est pas conforme
    assert garde["onglet"] in load_workbook(racine_fictive / CLASSEUR_ABC).sheetnames

    spec = _spec_seances(2)
    spec["lignes"][1]["seances"][2][champ] = refusee
    message = _refus_avant_ecriture(facture, spec, racine_fictive, tmp_path)

    assert champ in message, message
    _verifier_rang(message, ligne=2, seance=3)
    _verifier_valeur(message, refusee)


@pytest.mark.parametrize("valeur", ["false", "oui", 1], ids=["chaine-false", "chaine-oui", "entier-1"])
def test_t_fac_22_masquer_jour_non_booleen_refuse(racine_fictive, tmp_path, valeur):
    """T-FAC-22 — `masquer_jour` qui n'est pas un booléen (true ou false) est refusé avant toute écriture, par un
    message qui nomme la clé et la valeur. Avant le correctif, `bool()` le convertissait en silence : la chaîne
    « false » masquait la colonne Jour. Une chaîne se cite entre guillemets (`_formes_de`) : sans eux, le cas
    « false » passait à vide sur le texte fixe « attendu true ou false » (revue de fix-facture)."""
    facture = _module()
    spec = _spec_seances() | {"masquer_jour": valeur}

    message = _refus_avant_ecriture(facture, spec, racine_fictive, tmp_path)

    assert "masquer_jour" in message, message
    _verifier_valeur(message, valeur)


def test_t_fac_23_spec_en_utf8_avec_bom_acceptee(tmp_path):
    """T-FAC-23 — une spec JSON en UTF-8 avec BOM (Bloc-notes, PowerShell 5), fins de ligne Windows comprises, est lue
    par `charger_spec` comme la même spec sans BOM. Avant le correctif : JSONDecodeError « Unexpected UTF-8 BOM »."""
    facture = _module()
    texte = json.dumps(SPEC_COMPLETE, ensure_ascii=False, indent=2)
    sans_bom = tmp_path / "sans_bom.json"
    sans_bom.write_bytes(texte.encode("utf-8"))
    avec_bom = tmp_path / "avec_bom.json"
    avec_bom.write_bytes(b"\xef\xbb\xbf" + texte.replace("\n", "\r\n").encode("utf-8"))

    assert facture.charger_spec(sans_bom) == SPEC_COMPLETE
    assert facture.charger_spec(avec_bom) == SPEC_COMPLETE


# 31 caractères, mais 33 octets en UTF-8 (« û » et « É ») : la limite d'Excel compte des caractères. Puis 32.
TITRE_31 = "R_ABC001 Août 2026 École fictiv"
TITRE_32 = "R_ABC002 Août 2026 École fictive"


def test_t_fac_25_titre_d_onglet_de_plus_de_31_caracteres_refuse(racine_fictive, tmp_path):
    """T-FAC-25 — un titre d'onglet de plus de 31 caractères (limite d'Excel) est refusé avant toute écriture, par un
    message qui donne le titre et sa longueur. Garde : 31 caractères, accents compris, s'écrivent et restent
    conformes. Avant le correctif, openpyxl n'émettait qu'un avertissement et l'onglet de 32 caractères était écrit."""
    facture = _module()
    assert (len(TITRE_31), len(TITRE_32)) == (31, 32)
    _classeur_existant(racine_fictive)
    garde = _spec_seances(1) | {"onglet": TITRE_31}
    facture.generer(garde)  # lève PostMortemEnEcart si la facture écrite n'est pas conforme
    assert TITRE_31 in load_workbook(racine_fictive / CLASSEUR_ABC).sheetnames

    message = _refus_avant_ecriture(facture, _spec_seances(2) | {"onglet": TITRE_32}, racine_fictive, tmp_path)

    assert TITRE_32 in message, message
    assert re.search(r"(?<![\w.])32(?!\w)", message), f"longueur 32 absente du message : {message!r}"


# ──────────────────────────────────────────────────────────────────────
# Constats de la revue de fix-facture (T-FAC-18c, 18d, 25b, 26 à 28)
# ──────────────────────────────────────────────────────────────────────
#
# Même style de refus que T-FAC-19 à 25 : `ValueError` ou `EcritureRefusee`, avant toute écriture, message qui nomme
# la cause, le rang de la ligne et de la séance, et la valeur fautive. Le côté script (code 1, « [ERREUR] », sans
# trace Python) est à la fin de `tests/integration/test_scripts.py`.


def _colonne_taux(texte: str) -> list[str]:
    """Colonne taux (l'avant-dernière) de chaque ligne de séance de l'aperçu, telle qu'affichée (« |  45.13 € | »)."""
    return [ligne.split("|")[-2] for ligne in texte.splitlines() if ligne.startswith("  | ")]


def _taux_au_format(taux: str) -> str:
    """Colonne taux au format d'affichage inchangé : 6 caractères alignés à droite, puis « € »."""
    return f" {taux:>6} € "


@pytest.mark.parametrize(("taux", "affiche"), [(45.125, "45.13"), (60.555, "60.56")], ids=["45,125", "60,555"])
def test_t_fac_18c_taux_affiche_arrondi_comme_le_montant(taux, affiche):
    """T-FAC-18c — sur une ligne de l'aperçu (étape 6, validation humaine), la colonne taux s'arrondit comme la colonne
    montant : en décimal à partir de la valeur écrite, la moitié vers le haut, au format inchangé. 1 h à 45,125 €
    s'affiche « 45.13 € | 45.13 € », comme dans la facture Excel. Avant : le taux se formatait sur le float binaire
    (`{taux:>6.2f}` : 45,125 arrondi au pair, 60,555 qui vaut 60,55499… en binaire) et la ligne se contredisait
    (« 45.12 € | 45.13 € »). Garde : les taux à deux décimales de la spec complète ne bougent pas."""
    facture = _module()
    complete = _apercu_des_groupes(facture, _spec())
    assert _colonne_taux(complete) == [_taux_au_format(t) for t in ("90.00", "90.00", "60.50")], complete

    texte = _apercu_d_une_ligne(facture, [(1, taux)])

    assert _montants_affiches(texte) == [_colonne_montant(affiche)], texte
    assert _colonne_taux(texte) == [_taux_au_format(affiche)], texte


@pytest.mark.parametrize(("heures", "total"), [(0.7, "2.1"), (1.1, "3.3"), (0.1, "0.3")])
def test_t_fac_18d_total_heures_somme_decimale(heures, total):
    """T-FAC-18d — le TOTAL HEURES de l'aperçu se somme en décimal à partir de la valeur écrite, comme les euros
    (T-FAC-18) : trois séances de 0,7 h donnent « TOTAL HEURES : 2.1h ». Avant : somme en float binaire,
    « 2.0999999999999996h » (« 3.3000000000000003h », « 0.30000000000000004h »). Garde : le format ne change pas
    (spec complète : « 7.5h »)."""
    facture = _module()
    assert "TOTAL HEURES : 7.5h" in _apercu_des_groupes(facture, _spec()).splitlines()

    texte = _apercu_d_une_ligne(facture, [(heures, 90)] * 3)

    assert f"TOTAL HEURES : {total}h" in texte.splitlines(), texte


# Titres qu'Excel refuse, au-delà de la longueur (T-FAC-25) : vide, null (converti en « None » par `str()`), un
# nombre (jamais de conversion silencieuse), un caractère interdit (`/ \ ? * [ ] :`).
TITRES_INVALIDES = {
    "vide": "",
    "null": None,
    "nombre": 2026,
    "caractere-interdit": "R_ABC001 09/2026",
}


def _exiger(message: str, motif: str) -> None:
    """Le message contient le motif (expression régulière, sans tenir compte de la casse)."""
    assert re.search(motif, message, re.IGNORECASE), f"« {motif} » absent du message : {message!r}"


def _verifier_titre_cite(message: str, titre) -> None:
    """Le refus d'un titre d'onglet nomme l'onglet et, sauf titre vide (rien à citer), la valeur fautive."""
    _exiger(message, "onglet")
    if titre != "":
        _verifier_valeur(message, titre)


@pytest.mark.parametrize("titre", list(TITRES_INVALIDES.values()), ids=list(TITRES_INVALIDES))
def test_t_fac_25b_titre_d_onglet_invalide_pour_excel_refuse(racine_fictive, tmp_path, titre):
    """T-FAC-25b — un titre d'onglet qu'Excel refuse (vide, null, nombre, caractère interdit) est refusé avant toute
    écriture, comme un titre trop long, par un message qui nomme l'onglet et la valeur. Avant : null et "" étaient
    écrits dans un onglet « Sheet », puis déclarés NON CONFORMES par un faux écart (et la relance avec le bon titre
    était refusée : « le numéro R_ABC001 figure déjà dans l'onglet 'Sheet' ») ; le nombre levait une TypeError et le
    caractère interdit la ValueError d'openpyxl, sans le titre, après l'aperçu."""
    facture = _module()

    message = _refus_avant_ecriture(facture, _spec_seances() | {"onglet": titre}, racine_fictive, tmp_path)

    _verifier_titre_cite(message, titre)


def _modifier_seance_3_ligne_2(champ: str, valeur):
    def modifier(spec: dict):
        spec["lignes"][1]["seances"][2][champ] = valeur
    return modifier


def _vider_les_seances_de_la_ligne_1(spec: dict) -> None:
    spec["lignes"][0]["seances"] = []


# (modification de `_spec_seances()`, vérification du message) : la garde `_refuser_spec_mal_formee` sur le chemin
# d'`ecrire_onglet` (émission F-17) et de `generer` appelé avec une spec en dict, sans passer par `charger_spec`.
CAS_SPEC_MAL_FORMEE_SANS_CHARGER_SPEC = {
    "seances-vide": (
        _vider_les_seances_de_la_ligne_1,
        lambda m: (_verifier_rang(m, ligne=1), _exiger(m, re.escape(SPEC_SEANCES["lignes"][0]["libelle"]))),
    ),
    "heures-chaine": (
        _modifier_seance_3_ligne_2("heures", "18"),
        lambda m: (_exiger(m, "heures"), _verifier_rang(m, ligne=2, seance=3), _verifier_valeur(m, "18")),
    ),
    "taux-nan": (
        _modifier_seance_3_ligne_2("taux", float("nan")),
        lambda m: (_exiger(m, "taux"), _verifier_rang(m, ligne=2, seance=3), _verifier_valeur(m, float("nan"))),
    ),
    "masquer-jour-chaine": (
        _poser("masquer_jour", "false"),
        lambda m: (_exiger(m, "masquer_jour"), _verifier_valeur(m, "false")),
    ),
    "titre-32-caracteres": (
        _poser("onglet", TITRE_32),
        lambda m: (_exiger(m, re.escape(TITRE_32)), _verifier_valeur(m, 32)),
    ),
    **{
        f"titre-{nom}": (_poser("onglet", titre), lambda m, titre=titre: _verifier_titre_cite(m, titre))
        for nom, titre in TITRES_INVALIDES.items()
    },
}


@pytest.mark.parametrize(
    ("modifier", "verifier"),
    list(CAS_SPEC_MAL_FORMEE_SANS_CHARGER_SPEC.values()),
    ids=list(CAS_SPEC_MAL_FORMEE_SANS_CHARGER_SPEC),
)
def test_t_fac_26_ecrire_onglet_se_garde_d_une_spec_mal_formee(racine_fictive, dest, modifier, verifier):
    """T-FAC-26 — `ecrire_onglet`, appelé directement (chemin de l'émission F-17, qui ne passe ni par `charger_spec` ni
    par `generer`), refuse une spec mal formée avant d'écrire : ligne sans séance, heures « 18 », taux NaN,
    `masquer_jour` « false », titre de 32 caractères ou invalide pour Excel. Le message nomme le rang et la valeur,
    et le classeur, qui porte déjà une facture, garde ses onglets. Sans cette garde, l'onglet s'écrit : la séance
    disparaît, « 18 » est converti, « false » masque la colonne Jour (revue de fix-facture : retirer l'appel laissait
    les 75 tests T-FAC verts). Les cas du titre vide, null, nombre ou à caractère interdit sont aussi rouges sur le
    code d'avant ce constat."""
    facture = _module()
    spec = _spec_seances()
    modifier(spec)
    wb = _classeur_avec_une_facture()
    avant = _etat(wb)

    with pytest.raises((ValueError, facture.EcritureRefusee)) as erreur:
        facture.ecrire_onglet(wb, spec, dest, entite=ENTITE)

    assert _etat(wb) == avant, "ecrire_onglet a écrit dans le classeur une spec mal formée"
    verifier(str(erreur.value))


@pytest.mark.parametrize(
    ("modifier", "verifier"),
    list(CAS_SPEC_MAL_FORMEE_SANS_CHARGER_SPEC.values()),
    ids=list(CAS_SPEC_MAL_FORMEE_SANS_CHARGER_SPEC),
)
def test_t_fac_26b_generer_refuse_une_spec_mal_formee_avant_l_apercu(racine_fictive, capsys, modifier, verifier):
    """T-FAC-26b — `generer` reçoit une spec en dict (sans `charger_spec`) mal formée : même en `dry_run`, il la refuse
    avant l'aperçu, comme le dit sa docstring, par un message qui nomme le rang et la valeur, et n'écrit rien. Sans
    cette garde, le `--dry-run` afficherait « 18 » converti en silence, et l'écriture suivrait."""
    facture = _module()
    spec = _spec_seances()
    modifier(spec)

    with pytest.raises((ValueError, facture.EcritureRefusee)) as erreur:
        facture.generer(spec, dry_run=True)

    sortie = capsys.readouterr().out
    assert "APERÇU FACTURE" not in sortie, f"aperçu affiché avant le refus : {sortie}"
    assert not (racine_fictive / "ROSARI" / "ABC").exists()
    verifier(str(erreur.value))


def _lignes_en(transformer):
    def modifier(spec: dict):
        spec["lignes"] = transformer(spec["lignes"])
    return modifier


def _ligne_2_en(transformer):
    def modifier(spec: dict):
        spec["lignes"][1] = transformer(spec["lignes"][1])
    return modifier


def _seances_de_la_ligne_2_en(transformer):
    def modifier(spec: dict):
        spec["lignes"][1]["seances"] = transformer(spec["lignes"][1]["seances"])
    return modifier


def _seance_3_de_la_ligne_2_en(valeur):
    def modifier(spec: dict):
        spec["lignes"][1]["seances"][2] = valeur
    return modifier


# Un refus de structure dit ce qui était attendu : une liste (ou un « tableau », en JSON).
LISTE = r"liste|tableau"

# (modification de `_spec_seances()`, vérification du message) : types inattendus que la garde doit refuser sans
# trace Python (AttributeError, TypeError, OverflowError avant ce constat).
CAS_TYPES_INATTENDUS = {
    "lignes-en-objet": (
        _lignes_en(lambda lignes: lignes[0]),  # une ligne seule, sans la liste autour
        lambda m: (_exiger(m, r"\blignes\b"), _exiger(m, LISTE)),
    ),
    "lignes-en-chaine": (
        _lignes_en(lambda lignes: lignes[0]["libelle"]),
        lambda m: (_exiger(m, r"\blignes\b"), _exiger(m, LISTE)),
    ),
    "ligne-null": (_ligne_2_en(lambda ligne: None), lambda m: _verifier_rang(m, ligne=2)),
    "ligne-en-liste": (_ligne_2_en(lambda ligne: [ligne]), lambda m: _verifier_rang(m, ligne=2)),
    "seances-en-objet": (
        # Une seule séance écrite en objet au lieu d'une liste d'un élément : la clé « heures » y est bien présente.
        _seances_de_la_ligne_2_en(lambda seances: seances[0]),
        lambda m: (_verifier_rang(m, ligne=2), _exiger(m, r"s[ée]ances"), _exiger(m, LISTE)),
    ),
    "seance-null": (_seance_3_de_la_ligne_2_en(None), lambda m: _verifier_rang(m, ligne=2, seance=3)),
    "seance-nombre": (
        _seance_3_de_la_ligne_2_en(18),
        lambda m: (_verifier_rang(m, ligne=2, seance=3), _verifier_valeur(m, 18)),
    ),
    "heures-geantes": (
        _modifier_seance_3_ligne_2("heures", 10**400),  # JSON valide ; `math.isfinite` lève OverflowError
        lambda m: (_exiger(m, "heures"), _verifier_rang(m, ligne=2, seance=3)),
    ),
}


@pytest.mark.parametrize(("modifier", "verifier"), list(CAS_TYPES_INATTENDUS.values()), ids=list(CAS_TYPES_INATTENDUS))
def test_t_fac_27_types_inattendus_refuses_sans_trace(racine_fictive, tmp_path, modifier, verifier):
    """T-FAC-27 — une spec dont `lignes` n'est pas une liste de lignes, une ligne qui n'est pas un objet, `seances`
    qui n'est pas une liste, une séance qui n'est pas un objet, ou des heures entières trop grandes pour un float sont
    refusées avant toute écriture, par le refus de structure (`ValueError` ou `EcritureRefusee`) qui nomme le rang.
    Avant : AttributeError (« 'str' object has no attribute 'get' »), TypeError (« 'heures' in None »), OverflowError
    de `math.isfinite`, ou, pour `seances` en objet, le message trompeur « séance 1 : clé heures absente » alors que la
    clé est présente."""
    facture = _module()
    spec = _spec_seances()
    modifier(spec)

    message = _refus_avant_ecriture(facture, spec, racine_fictive, tmp_path)

    verifier(message)


def _supprimer_de_la_seance_3_ligne_2(cle: str):
    def modifier(spec: dict):
        del spec["lignes"][1]["seances"][2][cle]
    return modifier


def _supprimer_de_la_ligne_2(cle: str):
    def modifier(spec: dict):
        del spec["lignes"][1][cle]
    return modifier


# (modification de `_spec_seances()`, vérification du message) : dates et clés de ligne que la garde ne regardait
# pas ; elles levaient KeyError ou la ValueError de `strptime`, sans rang, plus loin dans `generer`.
CAS_DATES_ET_CLES_DE_LIGNE = {
    "seance-sans-date": (
        _supprimer_de_la_seance_3_ligne_2("date"),
        lambda m: (_exiger(m, r"\bdate\b"), _verifier_rang(m, ligne=2, seance=3)),
    ),
    "date-mal-formee": (
        _modifier_seance_3_ligne_2("date", "15/09/2026"),
        lambda m: (_verifier_rang(m, ligne=2, seance=3), _verifier_valeur(m, "15/09/2026")),
    ),
    "ligne-sans-promotion": (
        _supprimer_de_la_ligne_2("promotion"),
        lambda m: (_exiger(m, "promotion"), _verifier_rang(m, ligne=2)),
    ),
    "ligne-sans-libelle": (
        _supprimer_de_la_ligne_2("libelle"),
        lambda m: (_exiger(m, r"libell[eé]"), _verifier_rang(m, ligne=2)),
    ),
    "date-emission-mal-formee": (
        _poser("date_emission", "30/09/2026"),
        lambda m: (_exiger(m, "date_emission"), _verifier_valeur(m, "30/09/2026")),
    ),
}


@pytest.mark.parametrize(
    ("modifier", "verifier"), list(CAS_DATES_ET_CLES_DE_LIGNE.values()), ids=list(CAS_DATES_ET_CLES_DE_LIGNE),
)
def test_t_fac_28_dates_et_cles_de_ligne_refusees_avec_leur_rang(racine_fictive, tmp_path, modifier, verifier):
    """T-FAC-28 — une séance sans `date` ou à la date hors format AAAA-MM-JJ (« 15/09/2026 »), une ligne sans
    `promotion` ou sans `libelle`, une `date_emission` hors format : refus avant toute écriture, dans le même style
    que les autres refus de structure, par un message qui nomme la clé, le rang de la ligne et de la séance, et la
    valeur. Avant : KeyError brute ou ValueError de `strptime` (« time data '15/09/2026' does not match format »),
    sans rang, que le script affichait en trace Python."""
    facture = _module()
    spec = _spec_seances()
    modifier(spec)

    message = _refus_avant_ecriture(facture, spec, racine_fictive, tmp_path)

    verifier(message)
