"""Facture dans le paquet (P-0a ; banque de tests : T-FAC-01, T-FAC-03 à T-FAC-13).

T-FAC-02 (aide du script, en sous-processus) est à la fin de `tests/integration/test_scripts.py`.

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
