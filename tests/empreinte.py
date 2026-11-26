"""Empreintes comparables d'une feuille openpyxl, pour vérifier qu'une copie est fidèle.

Les listes d'attributs comparés sont prises dans les classes d'openpyxl, et non dans
`pcf_format` : un attribut que le clonage oublierait fait donc échouer la comparaison.
Chaque style est réduit à sa sérialisation XML, calculée une fois par style : comparer
les ~27 000 cellules du gabarit reste rapide, et deux classeurs se comparent champ par champ.
"""

from __future__ import annotations

from openpyxl.cell.cell import MergedCell
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.page import PrintPageSetup
from openpyxl.worksheet.views import SheetView
from openpyxl.xml.functions import tostring

# Volontairement hors comparaison : `id` (lien vers les réglages d'imprimante du fichier
# source) et la sélection d'onglet, propre au classeur qui contient la feuille.
_MISE_EN_PAGE = tuple(a for a in PrintPageSetup.__attrs__ if a != "id")
_VUE = tuple(a for a in SheetView.__attrs__ if a not in ("tabSelected", "workbookViewId"))
_OBJETS = ("sheet_format", "sheet_properties", "page_margins", "print_options", "HeaderFooter", "protection")


def _xml(objet) -> bytes:
    arbre = objet.to_tree()
    return b"" if arbre is None else tostring(arbre)


def empreinte_cellules(ws) -> dict:
    """{(ligne, colonne): (continuation de fusion ?, valeur, style)} pour chaque cellule qui
    porte une valeur ou un style. Une cellule en trop, en moins ou différente change le dict."""
    styles: dict[int, tuple] = {}
    resultat = {}
    for ligne in ws.iter_rows():
        for c in ligne:
            if c.value is None and not c.has_style:
                continue
            style = ()
            if c.has_style:
                style = styles.get(c.style_id)
                if style is None:
                    style = styles[c.style_id] = (
                        _xml(c.font), _xml(c.fill), _xml(c.border), _xml(c.alignment),
                        _xml(c.protection), c.number_format, c.quotePrefix, c.pivotButton, c.style,
                    )
            resultat[(c.row, c.column)] = (isinstance(c, MergedCell), c.value, style)
    return resultat


def largeurs_effectives(ws) -> dict:
    """{lettre: (largeur, masquée, niveau de plan)}, une dimension sur plusieurs colonnes
    (<col min="5" max="7">) étant dépliée colonne par colonne."""
    colonnes = {}
    for lettre, dim in ws.column_dimensions.items():
        debut = dim.min or column_index_from_string(lettre)
        etat = (dim.width or None, bool(dim.hidden), dim.outlineLevel or 0)
        if etat == (None, False, 0):
            continue
        for c in range(debut, (dim.max or debut) + 1):
            colonnes[get_column_letter(c)] = etat
    return colonnes


def empreinte_feuille(ws) -> dict:
    """Tout ce qui, hors cellules, change le rendu ou l'impression de la feuille."""
    zone = ws.print_area
    return {
        "fusions": sorted(str(p) for p in ws.merged_cells.ranges),
        "colonnes": largeurs_effectives(ws),
        "lignes": {
            r: (d.height or None, bool(d.hidden), d.outlineLevel or 0)
            for r, d in ws.row_dimensions.items()
            if d.height or d.hidden or d.outlineLevel
        },
        "mise_en_page": {a: getattr(ws.page_setup, a) for a in _MISE_EN_PAGE},
        "vue": {a: getattr(ws.sheet_view, a) for a in _VUE},
        "volets": ws.freeze_panes,
        "zone_impression": [p.rsplit("!", 1)[-1] for p in zone.split(",")] if zone else [],
        "titres_impression": (ws.print_title_rows, ws.print_title_cols),
        **{nom: _xml(getattr(ws, nom)) for nom in _OBJETS},
    }


def ecarts(attendu: dict, obtenu: dict, limite: int = 5) -> list:
    """Clés dont la valeur diffère (ou qui manquent d'un côté), pour des messages lisibles."""
    cles = sorted(set(attendu) | set(obtenu), key=str)
    return [(k, attendu.get(k), obtenu.get(k)) for k in cles if attendu.get(k) != obtenu.get(k)][:limite]
