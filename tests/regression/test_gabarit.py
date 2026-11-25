"""Fiabilisation du gabarit (F-18 ; banque de tests : T-FMT-09 à T-FMT-23).

Tests d'acceptation écrits avant l'implémentation, avec T-FMT-07a et 07b (`test_format.py`),
dont les marqueurs xfail sont retirés. T-FMT-19 à 23 viennent de la revue croisée de F-18
(25/11/2026). Données fictives uniquement.

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
  cellules vides qui la suivent si le texte, aligné à gauche, peut déborder. Une colonne sans
  largeur déclarée compte pour la largeur par défaut de la feuille, sinon 8,43, comme dans
  Excel et LibreOffice (et non les 13 qu'openpyxl prête à une dimension créée à la lecture).
  Une cellule en « réduire pour ajuster » est visible ; une cellule « renvoyer à la ligne »
  l'est si la hauteur de ses lignes couvre le nombre de lignes de texte (1,25 × corps par ligne).

- **Taille imprimée** (T-FMT-19) : corps de la police × échelle d'impression. L'estimation
  (T-FMT-19a, sans LibreOffice) prend l'échelle d'une page ajustée en largeur ; elle majore la
  taille réelle. La mesure qui fait foi (T-FMT-19b) lit le corps de chaque texte dans le PDF
  rendu par LibreOffice, qui y écrit les corps déjà réduits.
"""

from __future__ import annotations

import math
import re
import subprocess
import zlib
from datetime import date, timedelta
from pathlib import Path

import pytest
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import column_index_from_string, get_column_letter

from facturation import pcf_format as fmt
from facturation.verification import trouver_libreoffice
from tests.conftest import NDA_FICTIF, creer_master

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
    """Largeur affichée d'une colonne, en px ; 0 si elle est masquée.

    Lit les dimensions déclarées sans en créer : `ws.column_dimensions[lettre]` créerait la
    colonne avec la largeur par défaut d'openpyxl (13), alors qu'Excel et LibreOffice affichent
    une colonne non déclarée à la largeur par défaut de la feuille, sinon 8,43.
    """
    for dim in ws.column_dimensions.values():
        debut = dim.min or column_index_from_string(dim.index)
        if debut <= col <= (dim.max or debut):
            if dim.hidden:
                return 0.0
            if dim.width:
                return 7 * dim.width
            break
    return 7 * (ws.sheet_format.defaultColWidth or 8.43)


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


def test_mesure_largeur_par_defaut_temoin():
    """Témoin de la largeur par défaut : une colonne non déclarée mesure 8,43 (59 px), pas les 13
    d'openpyxl (91 px) ; la largeur par défaut de la feuille prime ; la mesure ne crée aucune colonne."""
    ws = Workbook().active
    ws["B1"], ws["C1"] = "Exemple ABC", "valeur"  # ≈ 86 px en Calibri 11 ; C1 empêche le débordement
    ws["B1"].font = Font(name="Calibri", sz=11)
    assert not _texte_visible(ws, "B1")[0]
    assert "B" not in ws.column_dimensions
    ws.sheet_format.defaultColWidth = 20
    assert _texte_visible(ws, "B1")[0]


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


@pytest.mark.parametrize(
    ("coord", "libelle"),
    [("C14", "Bon de commande :"), ("C15", "N° de contrat :"), ("C16", "N° SIRET :"), ("C17", "N° de Facture :")],
)
def test_t_fmt_15_libelles_de_reference_entiers(racine_fictive, coord, libelle):
    """T-FMT-15 — les libellés du bon de commande, du contrat, du SIRET et du n° de facture
    (C14 à C17) sont entièrement visibles."""
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


# ──────────────────────────────────────────────────────────────────────
# Lecture du PDF rendu par LibreOffice (T-FMT-19b, T-FMT-20b)
# ──────────────────────────────────────────────────────────────────────

_OBJET_PDF = re.compile(rb"(\d+)\s+0\s+obj(.*?)endobj", re.S)
_FLUX_PDF = re.compile(rb"stream\r?\n(.*?)\r?\nendstream", re.S)
_JETON_PDF = re.compile(
    rb"<[0-9A-Fa-f\s]*>|\[|\]|/[^\s/\[\]<>()]+|\((?:\\.|[^\\)])*\)|[-+]?(?:\d+\.?\d*|\.\d+)|[A-Za-z'\"*]+"
)


def _flux_pdf(corps: bytes) -> bytes | None:
    m = _FLUX_PDF.search(corps)
    if m is None:
        return None
    return zlib.decompress(m.group(1)) if b"FlateDecode" in corps[: m.start()] else m.group(1)


def _cmap_pdf(flux: bytes) -> tuple[dict, int]:
    """Table code de glyphe → texte d'une CMap ToUnicode, et longueur des codes en octets."""
    table, octets = {}, 1
    for bloc in re.findall(rb"beginbfchar(.*?)endbfchar", flux, re.S):
        for src, dst in re.findall(rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", bloc):
            octets = len(src) // 2
            table[int(src, 16)] = bytes.fromhex(dst.decode()).decode("utf-16-be")
    for bloc in re.findall(rb"beginbfrange(.*?)endbfrange", flux, re.S):
        for src, fin, dst in re.findall(rb"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>", bloc):
            octets = len(src) // 2
            for k, code in enumerate(range(int(src, 16), int(fin, 16) + 1)):
                table[code] = chr(int(dst, 16) + k)
    return table, octets


def _fragments_pdf(pdf: bytes) -> list[tuple[str, float, str]]:
    """(police, corps imprimé en points, texte) de chaque texte affiché par le PDF.

    Lecture minimale, suffisante pour les PDF de LibreOffice : polices et CMap ToUnicode,
    opérateurs Tf, Tj et TJ, échelles des matrices cm et Tm (le corps imprimé est le corps
    déclaré multiplié par ces échelles).
    """
    objets = {int(m.group(1)): m.group(2) for m in _OBJET_PDF.finditer(pdf)}
    polices = {}
    for corps in objets.values():
        for nom, num in re.findall(rb"/(F\d+)\s+(\d+)\s+0\s+R", corps):
            police = objets.get(int(num), b"")
            base = re.search(rb"/BaseFont\s*/([^\s/<>\[\]]+)", police)
            uni = re.search(rb"/ToUnicode\s+(\d+)\s+0\s+R", police)
            if base:
                cmap = _cmap_pdf(_flux_pdf(objets[int(uni.group(1))]) or b"") if uni else ({}, 1)
                polices[nom.decode()] = (base.group(1).decode().split("+")[-1], *cmap)
    fragments = []
    for corps in objets.values():
        flux = _flux_pdf(corps)
        if flux is None or b"Tf" not in flux:
            continue
        pile, ctm, tm, police, taille, operandes = [], 1.0, 1.0, None, 0.0, []
        for jeton in _JETON_PDF.findall(flux):
            if not (jeton[:1].isalpha() or jeton in (b"'", b'"')):
                if jeton not in (b"[", b"]"):
                    operandes.append(jeton)
                continue
            if jeton == b"q":
                pile.append(ctm)
            elif jeton == b"Q" and pile:
                ctm = pile.pop()
            elif jeton in (b"cm", b"Tm") and len(operandes) >= 6:
                a, b, c, d = (float(x) for x in operandes[-6:-2])
                if jeton == b"cm":
                    ctm *= math.sqrt(abs(a * d - b * c))
                else:
                    tm = math.sqrt(abs(a * d - b * c))
            elif jeton == b"BT":
                tm = 1.0
            elif jeton == b"Tf" and len(operandes) >= 2:
                police, taille = operandes[-2].decode().lstrip("/"), float(operandes[-1])
            elif jeton in (b"Tj", b"TJ", b"'", b'"') and police in polices:
                nom, table, octets = polices[police]
                texte = ""
                for chaine in operandes:
                    if chaine[:1] == b"<":
                        brut = bytes.fromhex(re.sub(rb"\s", b"", chaine[1:-1]).decode())
                        texte += "".join(
                            table.get(int.from_bytes(brut[i:i + octets], "big"), "?")
                            for i in range(0, len(brut), octets)
                        )
                fragments.append((nom, taille * tm * ctm, texte))
            operandes = []
    return fragments


def test_lecture_pdf_temoin():
    """Témoin de la lecture du PDF : police, corps réduit par la matrice cm, texte décodé par la CMap."""
    cmap = b"beginbfchar\n<01> <0054>\n<02> <00C9>\nendbfchar\nbeginbfrange\n<03> <04> <0061>\nendbfrange"
    contenu = b"q 0.5 0 0 0.5 0 0 cm BT /F1 12 Tf <01020304>Tj ET Q BT /F1 10 Tf [<03> -20 <04>]TJ ET"
    pdf = (
        b"1 0 obj<</Font<</F1 2 0 R>>>>endobj\n"
        b"2 0 obj<</Type/Font/BaseFont/ABCDEF+Calibri/ToUnicode 3 0 R>>endobj\n"
        b"3 0 obj<</Length " + str(len(cmap)).encode() + b">>stream\n" + cmap + b"\nendstream endobj\n"
        b"4 0 obj<</Length " + str(len(contenu)).encode() + b">>stream\n" + contenu + b"\nendstream endobj\n"
    )
    assert _fragments_pdf(pdf) == [("Calibri", 6.0, "TÉab"), ("Calibri", 10.0, "ab")]


@pytest.fixture(scope="module")
def rendu_17(tmp_path_factory):
    """Facture fictive de 17 séances et son rendu PDF, produits une fois pour le module.

    Renvoie (onglet, pages du PDF, fragments de texte du PDF).
    """
    dossier = tmp_path_factory.mktemp("rendu_17")
    creer_master(dossier / "donnees")
    ancienne = fmt._DATA_ROOT
    fmt.set_data_root(dossier / "donnees")
    try:
        wb, ws, _, _ = _facture(17)
        pages = _pages_pdf(wb, dossier)
    finally:
        fmt._DATA_ROOT = ancienne
        fmt._MASTER_DATA_CACHE.clear()
    return ws, pages, _fragments_pdf((dossier / "facture.pdf").read_bytes())


# ──────────────────────────────────────────────────────────────────────
# 6. Mentions lisibles à l'impression
# ──────────────────────────────────────────────────────────────────────

MENTIONS_IMPRIMEES = {
    "C85": "TVA", "C86": "déclaration d'activité", "C91": "échéance",
    "C92": "escompte", "C93": "pénalités", "C94": "indemnité",
}
CORPS_IMPRIME_MIN = 6.0  # points
A4_LARGEUR_PT = 595.28


def _zone(ws) -> tuple[int, int, int, int]:
    """Colonne et ligne de début, colonne et ligne de fin de la zone d'impression."""
    col1, lig1, col2, lig2 = re.search(r"\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$", ws.print_area).groups()
    return column_index_from_string(col1), int(lig1), column_index_from_string(col2), int(lig2)


def _echelle_ajustee_en_largeur(ws) -> float:
    """Échelle d'impression d'une page A4 portrait ajustée en largeur : un majorant de l'échelle réelle.

    Largeur utile du papier divisée par la largeur des colonnes imprimées, arrondie au
    pourcentage entier inférieur comme le font Excel et LibreOffice. Majorant : une facture
    ajustée aussi en hauteur peut être réduite davantage, et LibreOffice compte des colonnes
    un peu plus larges que 7 px par unité (échelle réelle de 0,42 pour 0,44 estimée, 25/11/2026).
    """
    assert (ws.page_setup.paperSize, ws.page_setup.orientation, ws.page_setup.fitToWidth) == (9, "portrait", 1)
    cmin, _, cmax, _ = _zone(ws)
    largeur_pt = 0.75 * sum(_largeur_colonne_px(ws, c) for c in range(cmin, cmax + 1))
    utile_pt = A4_LARGEUR_PT - 72 * (ws.page_margins.left + ws.page_margins.right)
    return math.floor(100 * min(1.0, utile_pt / largeur_pt)) / 100


def test_t_fmt_19a_mentions_lisibles_estimation(racine_fictive):
    """T-FMT-19a — 17 séances, sur une page : les mentions de bas de facture (TVA, déclaration
    d'activité, échéance, escompte, pénalités, indemnité) s'impriment en 6 points au moins,
    d'après le majorant de l'échelle (page ajustée en largeur) ; joué même sans LibreOffice."""
    _, ws, _, _ = _facture(17)
    assert not _sauts_manuels(ws) and ws.page_setup.fitToHeight == 1, "la facture de 17 séances doit tenir sur une page"
    echelle = _echelle_ajustee_en_largeur(ws)
    trop_petites = [
        f"{coord} ({nom}) : {ws[coord].font.sz} pt × {echelle:.2f} = {ws[coord].font.sz * echelle:.1f} pt au plus"
        for coord, nom in MENTIONS_IMPRIMEES.items()
        if ws[coord].font.sz * echelle < CORPS_IMPRIME_MIN
    ]
    assert not trop_petites, f"mentions sous {CORPS_IMPRIME_MIN} pt imprimés : {trop_petites}"


@SANS_LIBREOFFICE
def test_t_fmt_19b_mentions_lisibles_rendu_pdf(rendu_17):
    """T-FMT-19b — même exigence, mesurée dans le PDF : la facture de 17 séances tient sur une page
    et chaque mention y est imprimée en 6 points au moins."""
    ws, pages, fragments = rendu_17
    assert pages == 1, f"{pages} pages"
    absentes, trop_petites = [], []
    for coord, nom in MENTIONS_IMPRIMEES.items():
        debut = " ".join(str(ws[coord].value).split())[:15]
        corps = [taille for _, taille, texte in fragments if debut in " ".join(texte.split())]
        if not corps:
            absentes.append(f"{coord} ({nom}) « {debut}… »")
        elif min(corps) < CORPS_IMPRIME_MIN:
            trop_petites.append(f"{coord} ({nom}) : {min(corps):.1f} pt")
    assert not absentes, f"mentions introuvables dans le PDF : {absentes}"
    assert not trop_petites, f"mentions sous {CORPS_IMPRIME_MIN} pt imprimés : {trop_petites}"


# ──────────────────────────────────────────────────────────────────────
# 7. Polices nommées
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize("n", [17, 40, 62])
def test_t_fmt_20a_polices_nommees_calibri(racine_fictive, n):
    """T-FMT-20a — toute cellule stylée de la zone d'impression porte une police nommée, Calibri
    comme les en-têtes du gabarit, jamais None : 17 séances, 40 (fusions scindées au saut de
    page) et 62 (lignes insérées)."""
    _, ws, _, _ = _facture(n)
    cmin, rmin, cmax, rmax = _zone(ws)
    autres = [
        (cellule.coordinate, cellule.font.name)
        for ligne in ws.iter_rows(min_row=rmin, max_row=rmax, min_col=cmin, max_col=cmax)
        for cellule in ligne
        if cellule.has_style and cellule.font.name != "Calibri"
    ]
    assert not autres, f"{len(autres)} cellules sans police Calibri, dont {autres[:6]}"


@SANS_LIBREOFFICE
def test_t_fmt_20b_rendu_pdf_dans_la_police_des_en_tetes(rendu_17):
    """T-FMT-20b — dans le PDF, tout texte de la facture est dans la famille de police de l'en-tête
    « ECOLE » du gabarit : aucune police de substitution (Cambria sous Windows au 25/11/2026)."""
    _, _, fragments = rendu_17

    def famille(police: str) -> str:
        return police.split("-")[0]

    reference = {famille(police) for police, _, texte in fragments if texte.strip() == "ECOLE"}
    assert len(reference) == 1, reference
    autres = sorted({(famille(police), texte.strip()) for police, _, texte in fragments if famille(police) not in reference})
    assert not autres, f"{len(autres)} textes hors de la police {reference}, dont {autres[:6]}"


# ──────────────────────────────────────────────────────────────────────
# 8. Au-delà de 60 séances, le bas de facture garde ses hauteurs
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize("n", [61, 75])
def test_t_fmt_21_hauteurs_du_bas_de_facture_decalees(racine_fictive, n):
    """T-FMT-21 — au-delà de 60 séances, chaque ligne du bas de facture (« Total HT » à l'indemnité)
    a la hauteur et la visibilité de la même ligne d'une facture de 60 séances, décalée."""
    _, ref, _, _ = _facture(60)
    _, ws, _, _ = _facture(n)
    decalage = n - 60

    def hauteur(feuille, r):
        return feuille.row_dimensions[r].height, bool(feuille.row_dimensions[r].hidden)

    ecarts = [
        (r, hauteur(ref, r), r + decalage, hauteur(ws, r + decalage))
        for r in range(LIGNE_TOTAL, LIGNE_INDEMNITE + 1)
        if hauteur(ref, r) != hauteur(ws, r + decalage)
    ]
    assert not ecarts, f"(ligne à 60 séances, hauteur, ligne à {n}, hauteur) : {ecarts}"


# ──────────────────────────────────────────────────────────────────────
# 9. Pagination au-delà d'une page (revue de F-18)
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize("n", [22, 24, 25, 40, 55])
def test_t_fmt_22_derniere_page_garde_des_lignes_de_cours(racine_fictive, n):
    """T-FMT-22 (proposé) — au-delà d'une page, la dernière page porte au moins 3 lignes de cours
    au-dessus du total, comme l'annonce `_plan_de_pages` (MIN_LIGNES_AVEC_TOTAL).

    22 séances : le bloc du bas ne tient pas sous la page, 3 lignes partent avec lui. 24, 25 et
    55 : les lignes débordent seules, et la dernière page n'en garde qu'une ou deux.
    """
    _, ws, _, _ = _facture(n)
    sauts = _sauts_manuels(ws)
    assert sauts, f"{n} séances sur une page"
    total = LIGNE_TOTAL + max(0, n - 60)
    dernieres = [r for r in range(sauts[-1] + 1, total) if not ws.row_dimensions[r].hidden]
    assert len(dernieres) >= 3, f"{n} séances, sauts après {sauts} : {len(dernieres)} ligne(s) de cours sur la dernière page"


FRONTIERES_DU_TABLEAU = [
    (COLS["ecole"] - 1, COLS["ecole"]),  # bord gauche
    *((c - 1, c) for c in (COLS["promo"], COLS["libelle"], COLS["date"], COLS["heures"], COLS["taux"], COLS["total"])),
    (COLS["total"] + 1, COLS["total"] + 2),  # bord droit
]


@pytest.mark.parametrize("n", [17, 40])
def test_t_fmt_23_traits_verticaux_du_tableau_sur_chaque_page(racine_fictive, n):
    """T-FMT-23 (proposé) — chaque ligne de cours garde ses traits verticaux (bord gauche, séparateurs
    de colonnes, bord droit), y compris sous un saut de page, où les fusions école, promotion et
    libellé sont scindées : 17 séances (une page, témoin) et 40 (saut après la ligne 42)."""
    _, ws, _, _ = _facture(n)

    def trait(r: int, gauche: int, droite: int) -> bool:
        cotes = (ws.cell(row=r, column=gauche).border.right, ws.cell(row=r, column=droite).border.left)
        return any(cote is not None and cote.style for cote in cotes)

    manquants = [
        (r, f"{get_column_letter(g)}|{get_column_letter(d)}")
        for r in range(20, 20 + n)
        for g, d in FRONTIERES_DU_TABLEAU
        if not trait(r, g, d)
    ]
    assert not manquants, f"sauts après {_sauts_manuels(ws)} ; {len(manquants)} traits manquants, dont {manquants[:8]}"
