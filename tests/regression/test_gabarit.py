"""Fiabilisation du gabarit (F-18 ; banque de tests : T-FMT-09 à T-FMT-18).

Tests d'acceptation écrits avant l'implémentation, avec T-FMT-07a et 07b (`test_format.py`),
dont les marqueurs xfail sont retirés. Données fictives uniquement.

Mesures retenues, déterministes :

- **Pagination** (T-FMT-10, T-FMT-11) : la facture est rendue en PDF par LibreOffice et ses
  pages sont comptées dans le PDF. Une facture de plus d'une page ne laisse aucun saut de
  page au hasard : chaque saut est un saut manuel déclaré dans `ws.row_breaks`, ce qui se
  vérifie par « pages du PDF = sauts manuels + 1 ». On sait alors où tombe chaque saut, et
  aucun ne doit tomber dans le bloc du bas, de « Total HT » à l'indemnité forfaitaire.
  `Break(id=k)` place le saut entre les lignes k et k + 1.

- **Texte visible** (T-FMT-15, T-FMT-16) : la largeur du texte est estimée selon la police
  (Calibri : 0,62 em par majuscule, 0,52 em par autre caractère, 0,23 em par espace, +3 %
  en gras ; 1 pt = 96/72 px) et comparée à la place disponible : largeur des colonnes de la
  fusion qui contient la cellule (7 px par unité de largeur Excel), ou de sa colonne et des
  cellules vides qui la suivent si le texte, aligné à gauche, peut déborder. Une cellule en
  « réduire pour ajuster » est visible ; une cellule « renvoyer à la ligne » l'est si la
  hauteur de ses lignes couvre le nombre de lignes de texte (1,25 × corps par ligne).
"""

from __future__ import annotations

import math
import re
import subprocess
from datetime import date, timedelta
from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

from facturation import pcf_format as fmt
from facturation.verification import trouver_libreoffice
from tests.conftest import NDA_FICTIF

pytestmark = pytest.mark.regression

ENTITE = "ROSARI"
EMISSION = date(2026, 12, 1)
ECOLE = "ÉCOLE FICTIVE B"
BON = "BC-EXEMPLE-0001"
CONTRAT = "CT-EXEMPLE-0001"
IBAN = "FR76 0000 0000 0000 0000 0000 000"
BIC = "TESTFRPP"
LIGNE_TOTAL, LIGNE_INDEMNITE = 80, 94  # bloc du bas, sans décalage
COLS = fmt.LAYOUT_STANDARD["cols_table"]
SANS_LIBREOFFICE = pytest.mark.skipif(trouver_libreoffice() is None, reason="LibreOffice absent")


def _groupes(n: int) -> list:
    """n séances fictives, en modules de 10 séances au plus, sur deux promotions en alternance."""
    seances = [(date(2026, 11, 2) + timedelta(days=i % 28), 3.5, 60) for i in range(n)]
    return [
        (f"M{1 + k % 2} S1", f"Module {chr(65 + k)}", seances[debut:debut + 10])
        for k, debut in enumerate(range(0, n, 10))
    ]


def _facture(n: int, *, ref_contrat: str | None = CONTRAT, dest: dict | None = None):
    """Facture fictive complète de n séances, construite dans l'ordre du script."""
    wb = Workbook()
    wb.remove(wb.active)
    ws = fmt.clone_template(wb, f"R_ECO{n:03d} Nov 2026", entite=ENTITE)
    dest = dest or fmt.get_destinataire("ECO", entite=ENTITE)
    groupes = _groupes(n)
    fmt.apply_emetteur(ws, entite=ENTITE)
    fmt.apply_destinataire(ws, dest=dest, entite=ENTITE)
    fmt.apply_metadata(
        ws, n_facture=f"R_ECO{n:03d}", date_emission=EMISSION, entite=ENTITE,
        ref_bon_commande=BON, ref_contrat=ref_contrat,
    )
    fmt.place_sessions(ws, ecole_name=ECOLE, groups=groupes, entite=ENTITE)
    fmt.apply_mentions(ws, entite=ENTITE, date_emission=EMISSION)
    fmt.apply_rib(ws, entite=ENTITE)
    fmt.apply_print_setup(ws, entite=ENTITE)
    return wb, ws, dest, groupes


def _verifier(ws, dest, groupes, n: int, *, ref_contrat: str | None = CONTRAT) -> list:
    return fmt.verifier_facture(
        ws, entite=ENTITE, n_facture=f"R_ECO{n:03d}", date_emission=EMISSION, dest=dest,
        ecole_name=ECOLE, groups=groupes, ref_contrat=ref_contrat,
    )


def _pages_pdf(wb: Workbook, dossier: Path) -> int:
    """Nombre de pages du PDF rendu par LibreOffice."""
    xlsx = dossier / "facture.xlsx"
    wb.save(xlsx)
    subprocess.run(
        [str(trouver_libreoffice()), "--headless", "--convert-to", "pdf", "--outdir", str(dossier), str(xlsx)],
        check=True, capture_output=True, timeout=180,
    )
    return len(re.findall(rb"/Type\s*/Page(?!s)", (dossier / "facture.pdf").read_bytes()))


def _sauts_manuels(ws) -> list[int]:
    """Lignes après lesquelles un saut de page manuel est déclaré."""
    return sorted(saut.id for saut in ws.row_breaks.brk)


# ──────────────────────────────────────────────────────────────────────
# Mesure de la place d'un texte (T-FMT-15, T-FMT-16)
# ──────────────────────────────────────────────────────────────────────


def _largeur_texte_px(texte, police) -> float:
    em = sum(0.23 if car == " " else 0.62 if car.isupper() else 0.52 for car in str(texte))
    if police.b:
        em *= 1.03
    return em * (police.sz or 11) * 96 / 72


def _largeur_colonne_px(ws, col: int) -> float:
    dim = ws.column_dimensions[get_column_letter(col)]
    return 0.0 if dim.hidden else 7 * (dim.width or 8.43)


def _fusion(ws, ligne: int, col: int):
    return next((p for p in ws.merged_cells.ranges if p.min_row <= ligne <= p.max_row and p.min_col <= col <= p.max_col), None)


def _texte_visible(ws, coord: str) -> tuple[bool, str]:
    """La cellule montre-t-elle tout son texte ? Renvoie le verdict et le détail de la mesure."""
    cellule = ws[coord]
    if cellule.alignment.shrink_to_fit:
        return True, "réduire pour ajuster"
    fusion = _fusion(ws, cellule.row, cellule.column)
    cmin, cmax = (fusion.min_col, fusion.max_col) if fusion else (cellule.column, cellule.column)
    rmin, rmax = (fusion.min_row, fusion.max_row) if fusion else (cellule.row, cellule.row)
    place = sum(_largeur_colonne_px(ws, c) for c in range(cmin, cmax + 1))
    deborde = fusion is None and not cellule.alignment.wrap_text and cellule.alignment.horizontal in (None, "general", "left")
    c = cmax + 1
    while deborde and c <= COLS["total"] + 1 and ws.cell(row=cellule.row, column=c).value is None and not _fusion(ws, cellule.row, c):
        place += _largeur_colonne_px(ws, c)
        c += 1
    texte = _largeur_texte_px(cellule.value, cellule.font)
    if cellule.alignment.wrap_text:
        lignes = math.ceil(texte / place)
        besoin = lignes * (cellule.font.sz or 11) * 1.25
        hauteur = sum(ws.row_dimensions[r].height or 15 for r in range(rmin, rmax + 1))
        return hauteur >= besoin, f"{lignes} ligne(s) de texte : {besoin:.0f} pt nécessaires, {hauteur:.0f} pt disponibles"
    return texte <= place, f"texte ≈ {texte:.0f} px pour {place:.0f} px de place"


def test_mesure_du_texte_visible_temoins():
    """Témoins de la mesure : elle voit un texte coupé, un texte qui tient, un débordement et une réduction."""
    ws = Workbook().active
    ws.column_dimensions["C"].width, ws.column_dimensions["D"].width = 19.57, 12.14
    for coord, texte in (("C1", "Bon de commande :"), ("C2", "Bon de commande :"), ("C3", "BC :"), ("C4", "Bon de commande :")):
        ws[coord] = texte
        ws[coord].font = Font(name="Calibri", sz=18, b=True)
    ws["D1"] = ws["D3"] = ws["D4"] = "BC-EXEMPLE-0001"  # C2 seule peut déborder vers D2, vide
    ws["C4"].alignment = Alignment(shrink_to_fit=True)
    assert [_texte_visible(ws, c)[0] for c in ("C1", "C2", "C3", "C4")] == [False, True, True, True]


# ──────────────────────────────────────────────────────────────────────
# 1. Formats de nombre
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize("n", [60, 62])
def test_t_fmt_09_formats_homogenes_sur_toutes_les_lignes_de_cours(racine_fictive, n):
    """T-FMT-09 — heures, taux et totaux ont le format de la ligne 20 sur toutes les lignes de
    cours : 20 à 79 pour 60 séances, 20 à 81 avec les deux lignes insérées pour 62."""
    _, ws, _, _ = _facture(n)
    colonnes = {nom: COLS[nom] for nom in ("heures", "taux", "total")}
    reference = {nom: ws.cell(row=20, column=c).number_format for nom, c in colonnes.items()}
    assert "General" not in reference.values(), reference
    ecarts = [
        (r, nom, ws.cell(row=r, column=c).number_format)
        for r in range(20, 20 + n)
        for nom, c in colonnes.items()
        if ws.cell(row=r, column=c).number_format != reference[nom]
    ]
    assert not ecarts, f"{len(ecarts)} cellules hors format, dont {ecarts[:4]}"


# ──────────────────────────────────────────────────────────────────────
# 2. Pagination
# ──────────────────────────────────────────────────────────────────────


@SANS_LIBREOFFICE
@pytest.mark.parametrize("n", [1, 7])
def test_t_fmt_10_rendu_court_sur_une_page(racine_fictive, tmp_path, n):
    """T-FMT-10 — 1 et 7 séances : une page (17 séances : T-FMT-07b)."""
    wb, _, _, _ = _facture(n)
    assert _pages_pdf(wb, tmp_path) == 1


@SANS_LIBREOFFICE
def test_t_fmt_11_rendu_de_40_seances_bloc_du_bas_entier(racine_fictive, tmp_path):
    """T-FMT-11 — 40 séances : au plus 2 pages ; chaque saut de page est un saut manuel connu,
    et aucun ne coupe le bloc du bas (lignes 80 à 94, « Total HT » à l'indemnité)."""
    wb, ws, _, _ = _facture(40)
    pages, sauts = _pages_pdf(wb, tmp_path), _sauts_manuels(ws)
    assert pages <= 2, f"{pages} pages"
    assert pages == len(sauts) + 1, f"{pages} pages pour les sauts manuels {sauts} : un saut automatique est tombé au hasard"
    assert not [s for s in sauts if LIGNE_TOTAL <= s < LIGNE_INDEMNITE], f"saut dans le bloc du bas : {sauts}"


# ──────────────────────────────────────────────────────────────────────
# 3. Au-delà de 60 séances, tout le bas de facture suit le décalage
# ──────────────────────────────────────────────────────────────────────


def test_t_fmt_12_rib_decale(racine_fictive):
    """T-FMT-12 — 62 séances : IBAN et BIC écrits en D90 et D91, en face de leurs libellés décalés."""
    _, ws, _, _ = _facture(62)
    assert (ws["C90"].value, ws["D90"].value) == ("IBAN", IBAN)
    assert (ws["C91"].value, ws["D91"].value) == ("BIC", BIC)
    assert (ws["D88"].value, ws["D89"].value) == (None, None)


def test_t_fmt_13_post_mortem_decale(racine_fictive):
    """T-FMT-13 — 62 séances : le post-mortem relit total, RIB, échéance, NDA et contrat aux bonnes cellules."""
    _, ws, _, _ = _facture(62)
    pm = fmt.post_mortem(ws, entite=ENTITE)
    assert pm["total_formule"] == "=SUM(Q20:R81)"
    assert (pm["iban"], pm["bic"]) == (IBAN, BIC)
    assert pm["echeance"] == "Échéance : 31/12/2026"
    assert pm["nda"] == NDA_FICTIF
    assert pm["contrat"] == CONTRAT


@pytest.mark.parametrize(
    ("cellule", "valeur", "cle"),
    [
        (None, None, None),
        ("Q82", "=SUM(Q20:R79)", "total_formule"),
        ("D90", "FR76 9999", "iban"),
        ("C93", "Échéance : 01/01/2027", "echeance"),
    ],
    ids=["conforme", "total", "iban", "echeance"],
)
def test_t_fmt_14_verifier_facture_decalee(racine_fictive, cellule, valeur, cle):
    """T-FMT-14 — 62 séances : facture conforme sans écart ; une cellule décalée altérée, un seul écart."""
    _, ws, dest, groupes = _facture(62)
    if cellule is not None:
        ws[cellule] = valeur
    ecarts = _verifier(ws, dest, groupes, 62)
    if cle is None:
        assert ecarts == []
    else:
        assert len(ecarts) == 1 and ecarts[0].startswith(f"ERREUR sur cellule {cle}"), ecarts


# ──────────────────────────────────────────────────────────────────────
# 4. Rien n'est coupé à l'affichage
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize(("coord", "libelle"), [("C14", "Bon de commande :"), ("C15", "N° de contrat :")])
def test_t_fmt_15_libelles_de_reference_entiers(racine_fictive, coord, libelle):
    """T-FMT-15 — les libellés du bon de commande et du contrat sont entièrement visibles."""
    _, ws, _, _ = _facture(1)
    assert ws[coord].value == libelle
    visible, mesure = _texte_visible(ws, coord)
    assert visible, f"{coord} « {libelle} » coupé : {mesure}"


NOM_LONG = "SOCIÉTÉ FICTIVE D'ENSEIGNEMENT SUPÉRIEUR PRIVÉ EXEMPLE SAS"


def test_t_fmt_16_denomination_longue_du_client_entiere(racine_fictive):
    """T-FMT-16 — une dénomination juridique longue reste entière en M7 : réduite ou renvoyée à la ligne."""
    dest = {**fmt.get_destinataire("ECO", entite=ENTITE), "nom": NOM_LONG}
    _, ws, _, _ = _facture(1, dest=dest)
    assert ws["M7"].value == NOM_LONG
    visible, mesure = _texte_visible(ws, "M7")
    assert visible, f"M7 coupé : {mesure}"


# ──────────────────────────────────────────────────────────────────────
# 5. Le contrôle après écriture relit les textes fixes du bas de facture
# ──────────────────────────────────────────────────────────────────────

TEXTES_ALTERES = [
    ("O", LIGNE_TOTAL, "TOTAL"),
    ("C", 83, "À payer TTC*:"),
    ("C", 85, "TVA non applicable-art. 293 B du CGI"),
    ("C", 93, "Pénalités de retard modifiées"),
    ("C", 94, "Indemnité forfaitaire modifiée"),
]


@pytest.mark.parametrize("n", [2, 62], ids=["2-seances", "62-seances"])
@pytest.mark.parametrize(
    ("col", "ligne", "valeur", "ref_contrat"),
    [*((c, r, v, CONTRAT) for c, r, v in TEXTES_ALTERES), ("C", 15, "Contrat :", CONTRAT), ("C", 15, "N° de contrat :", None)],
    ids=["total-ht", "net-a-payer", "tva", "penalites", "indemnite", "libelle-contrat", "libelle-sans-contrat"],
)
def test_t_fmt_17_verifier_facture_controle_les_textes_fixes(racine_fictive, n, col, ligne, valeur, ref_contrat):
    """T-FMT-17 — un texte fixe du bas de facture, ou le libellé de C15, modifié : un seul écart, qui le montre.

    Les lignes du bas de facture descendent de deux lignes avec 62 séances ; la ligne 15 ne bouge pas.
    """
    _, ws, dest, groupes = _facture(n, ref_contrat=ref_contrat)
    decalage = max(0, n - 60) if ligne >= LIGNE_TOTAL else 0
    ws[f"{col}{ligne + decalage}"] = valeur
    ecarts = _verifier(ws, dest, groupes, n, ref_contrat=ref_contrat)
    assert len(ecarts) == 1 and ecarts[0].startswith("ERREUR sur cellule") and repr(valeur) in ecarts[0], ecarts


@pytest.mark.parametrize(("n", "cellule", "fausse"), [(2, "E83", "=Q79"), (62, "E85", "=Q80")], ids=["2-seances", "62-seances"])
def test_t_fmt_18_verifier_facture_controle_le_net_a_payer(racine_fictive, n, cellule, fausse):
    """T-FMT-18 (proposé) — la formule du net à payer qui ne vise plus le total est signalée."""
    _, ws, dest, groupes = _facture(n)
    ws[cellule] = fausse
    ecarts = _verifier(ws, dest, groupes, n)
    assert len(ecarts) == 1 and ecarts[0].startswith("ERREUR sur cellule") and repr(fausse) in ecarts[0], ecarts
