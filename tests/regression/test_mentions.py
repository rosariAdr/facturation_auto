"""Mentions légales et échéance de la facture (F-24, ADR030 ; banque de tests : T-MEN-01 à T-MEN-20).

Tests d'acceptation écrits avant l'implémentation. Ils fixent le contrat de
`calculer_echeance`, `emetteur_porte_ei`, `apply_metadata(ref_contrat=…)`,
`apply_mentions`, de la zone d'impression, de `post_mortem` et de `verifier_facture`.

Plan des cellules visé (gabarit standard, references/FORMAT.md « Mentions et totaux ») :
O80 Total HT ; C83 Net à payer (E83 = Q80) ; C84 vidée ; C85 mention de TVA ;
C86 n° de déclaration d'activité ; C91 échéance ; C92 escompte ; C93 pénalités ;
C94 indemnité forfaitaire ; ligne 15 n° de contrat ; impression jusqu'à la ligne 94.
Au-delà de 60 séances, ces lignes descendent d'autant que les lignes insérées.

Données fictives uniquement.
"""

from __future__ import annotations

from datetime import date, timedelta
from pathlib import Path

import pytest
from openpyxl import Workbook

from facturation import pcf_format as fmt
from tests.conftest import NDA_FICTIF, creer_master

pytestmark = pytest.mark.regression

ENTITE = "ROSARI"
EMISSION = date(2026, 12, 1)
CONTRAT = "CT-EXEMPLE-0001"
ECOLE = "ÉCOLE FICTIVE B"
GROUPES = [("M1 S1", "Module A", [(date(2026, 11, 3), 3.5, 60), (date(2026, 11, 10), 3.5, 60)])]

TOTAL_HT = "Total HT"
NET_A_PAYER = "Net à payer :"
TVA = "TVA non applicable, article 293 B du Code général des impôts."
ESCOMPTE = "Pas d'escompte en cas de paiement anticipé."
PENALITES = (
    "Pénalités de retard : trois fois le taux d'intérêt légal, exigibles dès le "
    "lendemain de l'échéance, sans rappel préalable."
)
INDEMNITE = "Indemnité forfaitaire pour frais de recouvrement en cas de retard de paiement : 40 €."
ANCIENS_LIBELLES = ("Montant non assujetti", "À payer TTC", "Payable Fin de Mois")


def _classeur() -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    return wb


def _facture_minimale(groupes: list | None = None):
    """Gabarit cloné et séances placées : l'ordre d'appel du script, avant les mentions."""
    ws = fmt.clone_template(_classeur(), "R_ECO001 Nov 2026", entite=ENTITE)
    fmt.place_sessions(ws, ecole_name=ECOLE, groups=groupes or GROUPES, entite=ENTITE)
    return ws


def _facture_complete(*, ref_contrat: str | None, echeance: date | None):
    """Facture fictive entière, construite par l'API dans l'ordre du script."""
    ws = fmt.clone_template(_classeur(), "R_ECO001 Nov 2026", entite=ENTITE)
    dest = fmt.get_destinataire("ECO", entite=ENTITE)
    fmt.apply_emetteur(ws, entite=ENTITE)
    fmt.apply_destinataire(ws, dest=dest, entite=ENTITE)
    fmt.apply_metadata(ws, n_facture="R_ECO001", date_emission=EMISSION, entite=ENTITE, ref_contrat=ref_contrat)
    fmt.place_sessions(ws, ecole_name=ECOLE, groups=GROUPES, entite=ENTITE)
    fmt.apply_rib(ws, entite=ENTITE)
    fmt.apply_mentions(ws, entite=ENTITE, date_emission=EMISSION, echeance=echeance)
    fmt.apply_print_setup(ws, entite=ENTITE)
    return ws, dest


def _anciens_libelles(ws) -> list[str]:
    """Textes de l'onglet qui contiennent encore un libellé retiré par ADR030."""
    return [
        cellule.value
        for ligne in ws.iter_rows()
        for cellule in ligne
        if isinstance(cellule.value, str) and any(ancien in cellule.value for ancien in ANCIENS_LIBELLES)
    ]


# ──────────────────────────────────────────────────────────────────────
# Fonctions pures
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.unit
@pytest.mark.parametrize(
    ("emission", "attendue"),
    [
        (date(2026, 12, 1), date(2026, 12, 31)),
        (date(2026, 12, 15), date(2027, 1, 14)),  # changement d'année
        (date(2027, 1, 31), date(2027, 3, 2)),  # février de 28 jours
        (date(2028, 1, 31), date(2028, 3, 1)),  # février bissextile
    ],
)
def test_t_men_01_echeance_par_defaut_emission_plus_30_jours(emission, attendue):
    """T-MEN-01 — sans délai donné, l'échéance est la date d'émission plus 30 jours (ADR030 §4)."""
    assert fmt.calculer_echeance(emission) == attendue
    assert fmt.calculer_echeance(emission, echeance=None) == attendue


@pytest.mark.unit
def test_t_men_02_echeance_donnee_prioritaire():
    """T-MEN-02 — le délai du contrat ou de l'école, quand il est donné, prime sur le défaut."""
    assert fmt.calculer_echeance(EMISSION, echeance=date(2027, 1, 15)) == date(2027, 1, 15)
    assert fmt.calculer_echeance(EMISSION, date(2026, 12, 10)) == date(2026, 12, 10)


@pytest.mark.unit
def test_t_men_18_echeance_anterieure_a_l_emission_refusee():
    """T-MEN-18 — une échéance antérieure à la date d'émission est refusée ; le jour même est accepté."""
    with pytest.raises(ValueError):
        fmt.calculer_echeance(EMISSION, echeance=date(2026, 11, 30))
    assert fmt.calculer_echeance(EMISSION, echeance=EMISSION) == EMISSION


@pytest.mark.unit
@pytest.mark.parametrize(
    "nom",
    [
        "PRÉNOM NOM EI",
        "Prénom Nom ei",
        "Prénom Nom - EI",
        "Prénom Nom EI ",
        "  Prénom Nom EI",
        "Prénom Nom Entrepreneur individuel",
        "Prénom Nom, entrepreneur individuel",
        "Prénom Nom (ENTREPRENEUR INDIVIDUEL)",
    ],
    ids=["majuscules", "minuscules", "tiret", "espace-final", "espaces-en-tete", "toutes-lettres",
         "toutes-lettres-virgule", "toutes-lettres-majuscules"],
)
def test_t_men_03_emetteur_porte_ei(nom):
    """T-MEN-03 — « EI » en fin de nom, ou « Entrepreneur individuel » n'importe où, sans tenir compte
    de la casse ni des espaces en début et en fin de nom."""
    assert fmt.emetteur_porte_ei(nom) is True


@pytest.mark.unit
@pytest.mark.parametrize(
    "nom",
    ["Prénom Nom", "", "Prénom NOMEI", "Prénom Nom EIRL", "EI Prénom Nom", "Prénom Entrepreneur", "Prénom Nom E.I."],
    ids=["sans-mention", "vide", "ei-colle", "eirl", "ei-en-tete", "entrepreneur-seul", "e-point-i-point"],
)
def test_t_men_04_emetteur_sans_ei(nom):
    """T-MEN-04 — « EI » doit être un mot entier en fin de nom ; « E.I. » ou un nom sans la mention est refusé."""
    assert fmt.emetteur_porte_ei(nom) is False


# ──────────────────────────────────────────────────────────────────────
# apply_mentions — textes, échéance, anciens libellés
# ──────────────────────────────────────────────────────────────────────


def test_t_men_05_textes_exacts_des_mentions(racine_fictive):
    """T-MEN-05 — totaux, TVA, déclaration d'activité et conditions de paiement au mot près."""
    ws = _facture_minimale()
    fmt.apply_mentions(ws, entite=ENTITE, date_emission=EMISSION)

    assert ws["O80"].value == TOTAL_HT
    assert ws["C83"].value == NET_A_PAYER
    assert ws["E83"].value == "=Q80"
    assert ws["C84"].value is None
    assert ws["C85"].value == TVA
    assert ws["C86"].value == NDA_FICTIF
    assert ws["C92"].value == ESCOMPTE
    assert ws["C93"].value == PENALITES
    assert ws["C94"].value == INDEMNITE


def test_t_men_06_anciens_libelles_absents(racine_fictive):
    """T-MEN-06 — « Montant non assujetti », « À payer TTC » et « Payable Fin de Mois » ne restent nulle part."""
    # La recherche voit bien un ancien libellé quand il est là : la garde ne passe pas à vide.
    temoin = _classeur().create_sheet("TEMOIN")
    temoin["C83"] = "À payer TTC*:"
    assert _anciens_libelles(temoin) == ["À payer TTC*:"]

    ws = _facture_minimale()
    fmt.apply_mentions(ws, entite=ENTITE, date_emission=EMISSION)
    assert _anciens_libelles(ws) == []


def test_t_men_07_echeance_par_defaut_imprimee(racine_fictive):
    """T-MEN-07 — sans échéance donnée : « Échéance : » suivi de la date d'émission plus 30 jours."""
    ws = _facture_minimale()
    fmt.apply_mentions(ws, entite=ENTITE, date_emission=EMISSION)
    assert ws["C91"].value == "Échéance : 31/12/2026"


def test_t_men_08_echeance_donnee_imprimee(racine_fictive):
    """T-MEN-08 — échéance donnée : c'est elle qui est imprimée, au format JJ/MM/AAAA."""
    ws = _facture_minimale()
    fmt.apply_mentions(ws, entite=ENTITE, date_emission=EMISSION, echeance=date(2027, 1, 15))
    assert ws["C91"].value == "Échéance : 15/01/2027"


def test_t_men_19_indemnite_fusionnee_c94_k94(racine_fictive):
    """T-MEN-19 — la ligne de l'indemnité est fusionnée sur C94:K94, comme les pénalités sur C93:K93."""
    ws = _facture_minimale()
    fmt.apply_mentions(ws, entite=ENTITE, date_emission=EMISSION)
    fusions = [str(plage) for plage in ws.merged_cells.ranges]
    assert "C94:K94" in fusions, fusions
    assert "C93:K93" in fusions, fusions


# ──────────────────────────────────────────────────────────────────────
# apply_metadata — ligne du n° de contrat
# ──────────────────────────────────────────────────────────────────────


def _fusions_ligne_15(ws) -> list[str]:
    return [str(plage) for plage in ws.merged_cells.ranges if plage.min_row <= 15 <= plage.max_row]


def test_t_men_09_ligne_de_contrat_presente(racine_fictive):
    """T-MEN-09 — n° de contrat donné : libellé en C15, valeur en D15, fusion D15:H15 (comme la ligne 14)."""
    ws = fmt.clone_template(_classeur(), "R_ECO001 Nov 2026", entite=ENTITE)
    fmt.apply_metadata(ws, n_facture="R_ECO001", date_emission=EMISSION, entite=ENTITE, ref_contrat=CONTRAT)
    assert ws["C15"].value == "N° de contrat :"
    assert ws["D15"].value == CONTRAT
    assert _fusions_ligne_15(ws) == ["D15:H15"]


def test_t_men_10_ligne_de_contrat_absente(racine_fictive):
    """T-MEN-10 — sans n° de contrat, la ligne 15 reste vide : ni libellé, ni valeur, ni fusion."""
    ws = fmt.clone_template(_classeur(), "R_ECO001 Nov 2026", entite=ENTITE)
    fmt.apply_metadata(ws, n_facture="R_ECO001", date_emission=EMISSION, entite=ENTITE, ref_contrat=None)
    assert [ws.cell(row=15, column=c).value for c in range(1, 19)] == [None] * 18
    assert _fusions_ligne_15(ws) == []


# ──────────────────────────────────────────────────────────────────────
# Mise en page
# ──────────────────────────────────────────────────────────────────────


def test_t_men_11_zone_d_impression_jusqu_a_la_ligne_94(racine_fictive):
    """T-MEN-11 — la nouvelle ligne de l'indemnité (94) est imprimée."""
    assert fmt.LAYOUT_STANDARD["derniere_ligne_utile"] == 94
    ws = fmt.clone_template(_classeur(), "R_ECO001 Nov 2026", entite=ENTITE)
    fmt.apply_print_setup(ws, entite=ENTITE)
    assert ws.print_area.endswith("!$A$1:$R$94"), ws.print_area


SEANCES_62 = [(date(2026, 11, 2) + timedelta(days=i % 28), 1.0, 50.0) for i in range(62)]


def test_t_men_12_decalage_au_dela_de_60_seances(racine_fictive):
    """T-MEN-12 — 62 séances : deux lignes insérées, les mentions descendent de deux lignes."""
    ws = _facture_minimale([("M1 S1", "Module A", SEANCES_62)])
    fmt.apply_mentions(ws, entite=ENTITE, date_emission=EMISSION, echeance=date(2027, 1, 15))

    assert ws["O82"].value == TOTAL_HT
    assert ws["C85"].value == NET_A_PAYER
    assert ws["C86"].value is None
    assert ws["C87"].value == TVA
    assert ws["C88"].value == NDA_FICTIF
    assert ws["C93"].value == "Échéance : 15/01/2027"
    assert ws["C94"].value == ESCOMPTE
    assert ws["C95"].value == PENALITES
    assert ws["C96"].value == INDEMNITE
    assert _anciens_libelles(ws) == []


def test_t_men_16_net_a_payer_suit_le_total_decale(racine_fictive):
    """T-MEN-16 — 62 séances : le net à payer (E85) reprend le total décalé (Q82), pas l'ancienne cellule Q80."""
    ws = _facture_minimale([("M1 S1", "Module A", SEANCES_62)])
    assert ws["Q82"].value == "=SUM(Q20:R81)"  # total recalculé par place_sessions
    fmt.apply_mentions(ws, entite=ENTITE, date_emission=EMISSION)
    assert ws["E85"].value == "=Q82"


def test_t_men_17_zone_d_impression_suit_le_decalage(racine_fictive):
    """T-MEN-17 — 62 séances : l'impression va jusqu'à la dernière mention décalée (ligne 96).

    Ordre d'appel du script : place_sessions, apply_mentions, puis apply_print_setup.
    """
    ws = _facture_minimale([("M1 S1", "Module A", SEANCES_62)])
    fmt.apply_mentions(ws, entite=ENTITE, date_emission=EMISSION)
    fmt.apply_print_setup(ws, entite=ENTITE)
    assert ws.print_area.endswith("!$A$1:$R$96"), ws.print_area


# ──────────────────────────────────────────────────────────────────────
# Donnée obligatoire manquante
# ──────────────────────────────────────────────────────────────────────


@pytest.fixture
def racine_sans_nda(tmp_path: Path, monkeypatch) -> Path:
    """MASTER fictif complet, sauf la ligne `Info` / `NDA` ; état du module rendu à la fin du test."""
    monkeypatch.setattr(fmt, "_DATA_ROOT", fmt._DATA_ROOT)
    monkeypatch.setattr(fmt, "_MASTER_DATA_CACHE", {})
    racine = tmp_path / "donnees"
    creer_master(racine, nda=None)
    fmt.set_data_root(racine)
    return racine


def test_t_men_13_nda_manquant_bloque(racine_sans_nda):
    """T-MEN-13 — sans `Info` / `NDA` dans le MASTER, les mentions ne s'écrivent pas : ValueError qui le nomme."""
    ws = _facture_minimale()
    with pytest.raises(ValueError, match=r"Info / NDA"):
        fmt.apply_mentions(ws, entite=ENTITE, date_emission=EMISSION)


# ──────────────────────────────────────────────────────────────────────
# Contrôle après écriture (PCF étape 7)
# ──────────────────────────────────────────────────────────────────────


def test_t_men_14_post_mortem_echeance_nda_contrat(racine_fictive):
    """T-MEN-14 — le post-mortem relit l'échéance, la déclaration d'activité et le n° de contrat."""
    ws, _ = _facture_complete(ref_contrat=CONTRAT, echeance=date(2027, 1, 15))
    pm = fmt.post_mortem(ws, entite=ENTITE)
    assert pm["echeance"] == "Échéance : 15/01/2027"
    assert pm["nda"] == NDA_FICTIF
    assert pm["contrat"] == CONTRAT

    sans_contrat, _ = _facture_complete(ref_contrat=None, echeance=None)
    assert fmt.post_mortem(sans_contrat, entite=ENTITE)["contrat"] is None


def _verifier(ws, dest):
    return fmt.verifier_facture(
        ws, entite=ENTITE, n_facture="R_ECO001", date_emission=EMISSION, dest=dest,
        ecole_name=ECOLE, groups=GROUPES, echeance=date(2027, 1, 15), ref_contrat=CONTRAT,
    )


@pytest.mark.parametrize(
    ("cellule", "valeur", "cle"),
    [
        (None, None, None),
        ("C91", "Échéance : 16/01/2027", "echeance"),
        ("D15", "CT-EXEMPLE-9999", "contrat"),
        ("C86", "Déclaration d'activité modifiée", "nda"),
    ],
    ids=["conforme", "echeance", "contrat", "nda"],
)
def test_t_men_15_verifier_facture_echeance_nda_contrat(racine_fictive, cellule, valeur, cle):
    """T-MEN-15 — facture conforme : aucun écart ; une seule cellule altérée : un seul écart, qui la nomme."""
    ws, dest = _facture_complete(ref_contrat=CONTRAT, echeance=date(2027, 1, 15))
    if cellule is None:
        assert _verifier(ws, dest) == []
        return
    ws[cellule] = valeur
    ecarts = _verifier(ws, dest)
    assert len(ecarts) == 1 and ecarts[0].startswith(f"ERREUR sur cellule {cle}"), ecarts


@pytest.mark.parametrize(
    ("cellule", "valeur", "cle"),
    [
        (None, None, None),
        ("C91", "Échéance : 01/01/2027", "echeance"),
        ("D15", "CT-INATTENDU", "contrat"),
    ],
    ids=["conforme", "echeance-par-defaut", "contrat-inattendu"],
)
def test_t_men_20_verifier_facture_controle_toujours(racine_fictive, cellule, valeur, cle):
    """T-MEN-20 — sans `echeance` ni `ref_contrat`, le contrôle a lieu quand même : échéance attendue
    = émission + 30 jours, ligne 15 attendue vide."""
    ws, dest = _facture_complete(ref_contrat=None, echeance=None)
    if cellule is not None:
        ws[cellule] = valeur
    ecarts = fmt.verifier_facture(
        ws, entite=ENTITE, n_facture="R_ECO001", date_emission=EMISSION, dest=dest,
        ecole_name=ECOLE, groups=GROUPES,
    )
    if cle is None:
        assert ecarts == []
    else:
        assert len(ecarts) == 1 and ecarts[0].startswith(f"ERREUR sur cellule {cle}"), ecarts
