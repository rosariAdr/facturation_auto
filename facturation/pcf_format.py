"""pcf_format.py — Module canonique de mise en forme des factures.

Source unique du format : ce module + le gabarit `facturation/assets/` > MODELE.
Toute spec visuelle ailleurs est obsolète (cf. references/FORMAT.md).

Usage : voir references/PROC00_pcf.md étape 5.2.

Gabarit (ADR015) : une mise en page unique, `PCF_facture_standard.xlsx`, vide de
toute coordonnée ; seules les informations de l'émetteur changent d'une entité à
l'autre. Une entité peut avoir son propre gabarit `PCF_facture_<ENTITE>.xlsx`,
prioritaire s'il existe.

Source des données :
- Coordonnées émetteur, destinataires : `MASTER.Coordonnées` (Type=Émetteur/Destinataire)
- SIRET, IBAN, BIC, TVA, NDA : `MASTER.Coordonnées` (Type=Info, Code=<ENTITE>)
- Aucune duplication de données dans ce module — tout est lu depuis MASTER au runtime.
  Le gabarit versionné est volontairement vide de toute coordonnée : elles sont
  injectées par apply_emetteur / apply_metadata / apply_mentions / apply_rib.

Mentions légales (ADR030) : les textes fixes du bas de facture (Total HT, Net à payer,
TVA, échéance, escompte, pénalités, indemnité de recouvrement) sont écrits par
apply_mentions, pas par le gabarit.

Où le module cherche les données (racine de données) :
1. `set_data_root(chemin)` si appelé explicitement ;
2. sinon la variable d'environnement `PCF_DATA_ROOT` ;
3. sinon le répertoire de travail courant.
La racine de données doit contenir `<ENTITE>/<ENTITE>_MASTER.xlsx`.

Entité explicite (ADR001) : toute l'API publique exige `entite=` en argument
nommé, sans valeur par défaut. Un script qui l'oublie lève une erreur au lieu
de lire en silence le MASTER d'une autre entité.
"""

import os
import re
from copy import copy as _copy
from datetime import date, timedelta
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

# ──────────────────────────────────────────────────────────────────────
# Chemins
# ──────────────────────────────────────────────────────────────────────

ASSETS = Path(__file__).resolve().parent / "assets"
GABARIT_STANDARD = ASSETS / "PCF_facture_standard.xlsx"

# Code d'entité : majuscules et chiffres, 2 à 16 caractères (ex. ROSARI, KAUR).
_ENTITE_RE = re.compile(r"^[A-Z][A-Z0-9]{1,15}$")

_DATA_ROOT: Optional[Path] = None


def valider_entite(entite: str) -> str:
    """Vérifie le format d'un code d'entité et le renvoie.

    Raises:
        ValueError: si le code n'est pas en majuscules (2 à 16 caractères) ou
            s'il vaut STANDARD, réservé au gabarit commun.
    """
    if not isinstance(entite, str) or not _ENTITE_RE.match(entite):
        raise ValueError(
            f"Code d'entité invalide : {entite!r}. Attendu : majuscules et "
            "chiffres, 2 à 16 caractères (ex. ROSARI, KAUR)."
        )
    if entite == "STANDARD":
        raise ValueError("Le code d'entité STANDARD est réservé au gabarit commun.")
    return entite


def gabarit_pour(entite: str) -> Path:
    """Gabarit à utiliser pour une entité : le sien s'il existe, sinon le standard."""
    propre = ASSETS / f"PCF_facture_{valider_entite(entite)}.xlsx"
    return propre if propre.exists() else GABARIT_STANDARD


def set_data_root(chemin) -> Path:
    """Fixe la racine des données métier pour la session Python en cours.

    Args:
        chemin: dossier contenant `<ENTITE>/<ENTITE>_MASTER.xlsx`.

    Returns:
        Le chemin résolu.
    """
    global _DATA_ROOT
    _DATA_ROOT = Path(chemin).expanduser().resolve()
    _MASTER_DATA_CACHE.clear()
    return _DATA_ROOT


def get_data_root() -> Path:
    """Racine des données : set_data_root() > $PCF_DATA_ROOT > dossier courant."""
    if _DATA_ROOT is not None:
        return _DATA_ROOT
    env = os.environ.get("PCF_DATA_ROOT")
    if env:
        return Path(env).expanduser().resolve()
    return Path.cwd()

# ──────────────────────────────────────────────────────────────────────
# Cache lecture MASTER.Coordonnées (lazy loading)
# ──────────────────────────────────────────────────────────────────────

_MASTER_DATA_CACHE: dict = {}


def _master_path_for(entite: str) -> Path:
    """Retourne le chemin du MASTER de l'entité, sous la racine de données."""
    valider_entite(entite)
    return get_data_root() / entite / f"{entite}_MASTER.xlsx"


def _load_master_data(entite: str, refresh: bool = False) -> dict:
    """Charge MASTER.Coordonnées et retourne un dict structuré.

    Format :
    {
        "emetteur": {"nom": ..., "adresse_l1": ..., "adresse_l2": ..., "email": ..., "tel": ...},
        "destinataires": {
            "ABC": {...},
            "ECO": {...},
            ...
        },
        "info": {"SIRET": ..., "IBAN": ..., "BIC": ..., "TVA": ..., "NDA": ...},
    }

    Args:
        entite: "ROSARI" ou "KAUR".
        refresh: si True, force la relecture du MASTER (ignore le cache).
    """
    if not refresh and entite in _MASTER_DATA_CACHE:
        return _MASTER_DATA_CACHE[entite]

    master_path = _master_path_for(entite)
    if not master_path.exists():
        raise FileNotFoundError(
            f"MASTER introuvable : {master_path}\n"
            f"Racine de données utilisée : {get_data_root()}\n"
            "Corrige-la via set_data_root(...) ou la variable d'env PCF_DATA_ROOT."
        )

    wb = load_workbook(master_path, data_only=True)
    if "Coordonnées" not in wb.sheetnames:
        raise RuntimeError(
            f"Onglet 'Coordonnées' manquant dans {master_path.name}"
        )
    ws = wb["Coordonnées"]

    # Header attendu : Type | Code | Nom | Adresse L1 | Adresse L2 | Adresse L3 | Email | Téléphone
    data = {"emetteur": None, "destinataires": {}, "info": {}}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        type_ = row[0]
        code = row[1]
        nom = row[2]
        adr_l1 = row[3]
        adr_l2 = row[4]
        adr_l3 = row[5]
        email = row[6]
        tel = row[7]

        if type_ == "Émetteur" and code == entite:
            data["emetteur"] = {
                "nom": nom or "",
                "adresse_l1": adr_l1 or "",
                "adresse_l2": adr_l2 or "",
                "email": email or "",
                "tel": tel or "",
            }
        elif type_ == "Destinataire" and code:
            # Stocker None si toutes les valeurs sont vides
            if not nom and not adr_l1 and not adr_l2 and not email:
                data["destinataires"][code] = None
            else:
                data["destinataires"][code] = {
                    "nom": nom or "",
                    "adresse_l1": adr_l1 or "",
                    "adresse_l2": adr_l2 or "",
                    "adresse_l3": adr_l3 or "",
                    "email": email or "",
                }
        elif type_ == "Info" and code == entite:
            # Info row : champ Nom contient le label (SIRET/IBAN/BIC/TVA)
            # et Adresse L1 contient la valeur
            if nom:
                data["info"][nom] = adr_l1 or ""

    if data["emetteur"] is None:
        raise RuntimeError(
            f"Pas d'émetteur '{entite}' trouvé dans {master_path.name}.Coordonnées"
        )
    required_info = {"SIRET", "IBAN", "BIC"}
    missing = required_info - set(data["info"].keys())
    if missing:
        raise RuntimeError(
            f"Info manquantes pour {entite} dans MASTER.Coordonnées : {missing}. "
            f"Attendu lignes : Type=Info, Code={entite}, Nom=SIRET/IBAN/BIC/TVA."
        )

    _MASTER_DATA_CACHE[entite] = data
    return data


# Champs sans lesquels un destinataire est incomplet : PROC00 étape 4 (raison sociale,
# adresse complète) et MASTER_SCHEMA (Adresse L1 = rue, Adresse L2 = code postal et ville).
# L'email et le pays (Adresse L3) restent facultatifs, comme dans apply_destinataire.
CHAMPS_DESTINATAIRE_OBLIGATOIRES = (
    ("nom", "Nom / Raison sociale"),
    ("adresse_l1", "Adresse L1"),
    ("adresse_l2", "Adresse L2"),
)


def champs_destinataire_manquants(dest: dict) -> list:
    """Colonnes de MASTER.Coordonnées obligatoires mais vides pour ce destinataire (FA-09).

    Returns:
        Les libellés des colonnes manquantes ; liste vide si le destinataire est complet.
    """
    return [
        libelle
        for cle, libelle in CHAMPS_DESTINATAIRE_OBLIGATOIRES
        if not str(dest.get(cle) or "").strip()
    ]


def get_destinataire(code_ecole: str, *, entite: str) -> Optional[dict]:
    """Helper public : retourne le destinataire pour une école donnée, ou None.

    Args:
        code_ecole: code école (ex. "ABC", "ECO").
        entite: "ROSARI" ou "KAUR".
    """
    data = _load_master_data(entite)
    return data["destinataires"].get(code_ecole)


def get_emetteur(*, entite: str) -> dict:
    """Helper public : retourne l'émetteur de l'entité, lu dans MASTER.Coordonnées.

    Args:
        entite: code de l'entité, obligatoire et nommé (ADR001).

    Returns:
        Une copie du dict nom, adresse_l1, adresse_l2, email, tel.
    """
    return dict(_emetteur_for(entite))


def get_nda(*, entite: str) -> str:
    """N° de déclaration d'activité de l'entité, imprimé tel quel sur la facture (ADR030).

    Lu dans `MASTER.Coordonnées`, ligne Type=Info, Code=<ENTITE>, Nom=NDA : la colonne
    Adresse L1 porte le texte complet de la mention.

    Args:
        entite: code de l'entité, obligatoire et nommé (ADR001).

    Raises:
        ValueError: si la ligne `Info` / `NDA` manque ou si son texte est vide.
    """
    nda = _load_master_data(entite)["info"].get("NDA")
    if not str(nda or "").strip():
        raise ValueError(
            f"N° de déclaration d'activité absent de MASTER.Coordonnées pour {entite} : "
            f"ajouter la ligne Info / NDA (Type=Info, Code={entite}, Nom=NDA, "
            "Adresse L1 = texte imprimé sur la facture)."
        )
    return nda


# « EI » en dernier mot du nom (ADR030) : « Prénom Nom EI », « Prénom Nom - EI ».
_EI_FINAL_RE = re.compile(r"\bEI\Z", re.IGNORECASE)


def emetteur_porte_ei(nom: str) -> bool:
    """Vrai si le nom de l'émetteur porte la mention d'entrepreneur individuel (ADR030).

    Accepté, sans tenir compte de la casse ni des espaces en début et en fin : « EI »
    comme dernier mot du nom, ou « Entrepreneur individuel » n'importe où. Refusé :
    « E.I. », « EI » collé au nom (« NOMEI ») ou placé en tête, « EIRL », « Entrepreneur » seul.
    """
    nom = str(nom or "").strip()
    return bool(_EI_FINAL_RE.search(nom)) or "entrepreneur individuel" in nom.casefold()


DELAI_ECHEANCE_PAR_DEFAUT = timedelta(days=30)


def calculer_echeance(date_emission: date, echeance: Optional[date] = None) -> date:
    """Date d'échéance imprimée sur la facture (ADR030 §4).

    Args:
        date_emission: date d'émission de la facture.
        echeance: délai du contrat ou de l'école, déjà converti en date ; None pour
            le défaut, date d'émission plus 30 jours.

    Raises:
        ValueError: si l'échéance donnée précède la date d'émission (le jour même est accepté).
    """
    if echeance is None:
        return date_emission + DELAI_ECHEANCE_PAR_DEFAUT
    if echeance < date_emission:
        raise ValueError(
            f"échéance {echeance.strftime('%d/%m/%Y')} antérieure à la date d'émission "
            f"{date_emission.strftime('%d/%m/%Y')}"
        )
    return echeance


# Plan des cellules du gabarit standard, commun à toutes les entités (ADR015).
LAYOUT_STANDARD = {
    "bandeau_row": 2,
    "emetteur_rows": (7, 11),
    "emetteur_col": 3,  # C
    "destinataire_rows": (7, 10),
    "destinataire_col": 13,  # M
    "bon_commande_row": 14,
    "bon_commande_label_col": 3,  # C
    "bon_commande_value_cols": (4, 8),  # D:H (merge)
    "contrat_row": 15,  # mêmes colonnes que le bon de commande
    "siret_label_cell": "C16",
    "siret_cell": "D16",
    "n_facture_label_cell": "C17",
    "n_facture_cell": "D17",
    "table_header_row": 19,
    "cours_first_row": 20,
    "cours_last_row": 79,  # 60 rows
    "cols_table": {
        "ecole": 3,    # C (merge C:D)
        "promo": 5,    # E (merge E:F)
        "libelle": 7,  # G (merge G:J)
        "date": 11,    # K (merge K:L)
        "heures": 13,  # M (merge M:N)
        "taux": 15,    # O (merge O:P)
        "total": 17,   # Q (merge Q:R)
    },
    "total_row": 80,
    # Mentions légales (ADR030), écrites par apply_mentions en colonne C. Au-delà de
    # 60 séances, elles descendent d'autant que les lignes insérées par place_sessions.
    "mentions_col": 3,  # C
    "mentions_fin_courte": 7,  # G : échéance, escompte (fusion C:G du gabarit)
    "mentions_fin_longue": 11,  # K : TVA, pénalités, indemnité (fusion C:K)
    "net_a_payer_row": 83,  # libellé C83:D83, montant E83:F83
    "ancienne_asterisque_row": 84,  # « *Montant non assujetti… », vidée
    "tva_row": 85,
    "nda_row": 86,
    "echeance_row": 91,
    "escompte_row": 92,
    "penalites_row": 93,
    "indemnite_row": 94,
    "rib_iban_cell": "D88",
    "rib_bic_cell": "D89",
    "derniere_ligne_utile": 94,  # dernière mention légale → borne du print_area
}

LAYOUT_ROSARI = LAYOUT_STANDARD  # nom historique, conservé pour les scripts existants

# Plans propres à une entité dont le gabarit diffère du standard (aucun pour l'instant).
LAYOUTS: dict = {}

THIN = Side(style="thin", color="000000")

# Textes des mentions légales (ADR030), au mot près.
MENTION_TOTAL_HT = "Total HT"
MENTION_NET_A_PAYER = "Net à payer :"
MENTION_TVA = "TVA non applicable, article 293 B du Code général des impôts."
MENTION_ESCOMPTE = "Pas d'escompte en cas de paiement anticipé."
MENTION_PENALITES = (
    "Pénalités de retard : trois fois le taux d'intérêt légal, exigibles dès le "
    "lendemain de l'échéance, sans rappel préalable."
)
MENTION_INDEMNITE = (
    "Indemnité forfaitaire pour frais de recouvrement en cas de retard de paiement : 40 €."
)
LIBELLE_CONTRAT = "N° de contrat :"

# ──────────────────────────────────────────────────────────────────────
# Helpers internes
# ──────────────────────────────────────────────────────────────────────


def _set_if_not_merged(ws: Worksheet, row: int, col: int, value):
    """Écrit value si la cell n'est pas une MergedCell continuation."""
    cell = ws.cell(row=row, column=col)
    if not isinstance(cell, MergedCell):
        cell.value = value


def _unmerge_in_range(
    ws: Worksheet, rmin: int, rmax: int, cmin: int = 1, cmax: int = 30
):
    """Démerge toutes les plages incluses dans la zone."""
    to_remove = []
    for mr in list(ws.merged_cells.ranges):
        if (
            mr.min_row >= rmin
            and mr.max_row <= rmax
            and mr.min_col >= cmin
            and mr.max_col <= cmax
        ):
            to_remove.append(str(mr))
    for mr_str in to_remove:
        try:
            ws.unmerge_cells(mr_str)
        except Exception:
            pass


def _frame_box(ws: Worksheet, rmin: int, rmax: int, cmin: int, cmax: int):
    """Trace un encadré propre : cadre extérieur fin, aucun trait horizontal interne.

    Le template DOR001 d'origine porte des bordures résiduelles au milieu des
    encadrés émetteur/destinataire (la zone faisait 3 lignes avant d'être
    étendue à 5). Sans ce nettoyage, un trait traverse l'encadré.
    """
    # Note : les bordures sont assignables sur les MergedCell (contrairement
    # aux valeurs), il faut donc parcourir toute la zone et pas seulement les
    # cellules d'ancrage, sinon le cadre droit manque sur les zones fusionnées.
    for r in range(rmin, rmax + 1):
        for c in range(cmin, cmax + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = Border(
                top=THIN if r == rmin else None,
                bottom=THIN if r == rmax else None,
                left=THIN if c == cmin else None,
                right=THIN if c == cmax else None,
            )


def _xworkbook_clone_modele(
    template_path: Path, dst_wb: Workbook, new_name: str
) -> Worksheet:
    """Charge MODELE depuis template_path et le recopie cellule par cellule dans dst_wb.

    Le gabarit compte environ 26 800 cellules stylées mais une soixantaine de styles
    distincts. Chaque style source n'est donc recopié qu'une fois (police, remplissage,
    bordure, alignement, format de nombre) ; les cellules suivantes qui portent le même
    style reçoivent directement les index du style déjà créé dans dst_wb. Le résultat est
    identique à une copie attribut par attribut, en 20 à 50 fois moins de temps.
    """
    src_wb = load_workbook(template_path)
    src_ws = src_wb["MODELE"]

    new_ws = dst_wb.create_sheet(new_name)

    # Copie des cellules : un style source -> un style destination, recopié une seule fois.
    styles_recopies: dict[tuple, object] = {}
    for row in src_ws.iter_rows():
        for cell in row:
            if cell.value is None and not cell.has_style:
                continue
            new_cell = new_ws.cell(row=cell.row, column=cell.column)
            new_cell.value = cell.value
            if not cell.has_style:
                continue
            cle = tuple(cell._style)
            deja = styles_recopies.get(cle)
            if deja is not None:
                new_cell._style = _copy(deja)
                continue
            new_cell.font = _copy(cell.font)
            new_cell.fill = _copy(cell.fill)
            new_cell.border = _copy(cell.border)
            new_cell.alignment = _copy(cell.alignment)
            new_cell.number_format = cell.number_format
            styles_recopies[cle] = _copy(new_cell._style)

    # Column widths
    for col_letter, dim in src_ws.column_dimensions.items():
        if dim.width:
            new_ws.column_dimensions[col_letter].width = dim.width

    # Row heights
    for r, dim in src_ws.row_dimensions.items():
        if dim.height:
            new_ws.row_dimensions[r].height = dim.height

    # Merges
    for mr in src_ws.merged_cells.ranges:
        try:
            new_ws.merge_cells(str(mr))
        except Exception:
            pass

    return new_ws


def _titre_facture(date_emission: date) -> str:
    """Texte du bandeau (B2), écrit par apply_metadata et contrôlé par verifier_facture."""
    return f"FACTURE\nDate d'émission: {date_emission.strftime('%d/%m/%Y')}"


def _texte_echeance(echeance: date) -> str:
    """Ligne d'échéance, écrite par apply_mentions et contrôlée par verifier_facture."""
    return f"Échéance : {echeance.strftime('%d/%m/%Y')}"


def _decalage(ws: Worksheet, layout: dict) -> int:
    """Lignes insérées par place_sessions au-delà des 60 lignes de cours (0 sinon).

    place_sessions insère les lignes juste avant le total et y réécrit la formule
    `=SUM(Q<première ligne>:R<ligne au-dessus>)` : la ligne qui porte cette formule
    donne le décalage de tout ce qui suit (total, mentions, zone d'impression).

    Raises:
        RuntimeError: si la formule du total est introuvable sous le tableau.
    """
    col = layout["cols_table"]["total"]
    debut = (
        f"=SUM({get_column_letter(col)}{layout['cours_first_row']}:"
        f"{get_column_letter(col + 1)}"
    )
    for r in range(layout["total_row"], ws.max_row + 1):
        if ws.cell(row=r, column=col).value == f"{debut}{r - 1})":
            return r - layout["total_row"]
    raise RuntimeError(
        f"Formule du total introuvable en colonne {get_column_letter(col)} à partir de la "
        f"ligne {layout['total_row']} : la facture ne vient pas du gabarit."
    )


def _fusionner_ligne(ws: Worksheet, row: int, cmin: int, cmax: int):
    """Fusionne row sur les colonnes cmin à cmax, après avoir retiré les fusions qu'elles contiennent."""
    _unmerge_in_range(ws, row, row, cmin, cmax)
    ws.merge_cells(start_row=row, start_column=cmin, end_row=row, end_column=cmax)


def _ecrire_reference(ws: Worksheet, layout: dict, row: int, libelle: str, valeur: str):
    """Ligne de référence (bon de commande, contrat) : libellé en C, valeur fusionnée sur D:H.

    Libellé et valeur prennent la police du libellé « N° de Facture ».
    """
    label_col = layout["bon_commande_label_col"]
    vmin, vmax = layout["bon_commande_value_cols"]

    modele = ws[layout["n_facture_label_cell"]]
    for col, texte in ((label_col, libelle), (vmin, valeur)):
        cell = ws.cell(row=row, column=col)
        cell.value = texte
        cell.font = _copy(modele.font)
        cell.alignment = _copy(modele.alignment)

    _fusionner_ligne(ws, row, vmin, vmax)
    # Hauteur auto : la ligne fait 14.25pt dans le template, illisible en 18pt
    ws.row_dimensions[row].height = None


def _zone_impression(ws: Worksheet, layout: dict):
    """Zone d'impression : de A1 à la dernière mention légale, décalage compris."""
    derniere_col = get_column_letter(layout["cols_table"]["total"] + 1)
    derniere_ligne = layout["derniere_ligne_utile"] + _decalage(ws, layout)
    ws.print_area = f"A1:{derniere_col}{derniere_ligne}"


def _layout_for(entite: str) -> dict:
    """Plan des cellules de l'entité : le sien s'il est déclaré, sinon le standard."""
    return LAYOUTS.get(valider_entite(entite), LAYOUT_STANDARD)


def _emetteur_for(entite: str) -> dict:
    """Retourne l'émetteur lu depuis MASTER.Coordonnées."""
    return _load_master_data(entite)["emetteur"]


def _rib_for(entite: str) -> dict:
    """Retourne {IBAN, BIC} lu depuis MASTER.Coordonnées (Type=Info)."""
    info = _load_master_data(entite)["info"]
    return {"IBAN": info["IBAN"], "BIC": info["BIC"]}


def _siret_for(entite: str) -> str:
    """Retourne le SIRET lu depuis MASTER.Coordonnées (Type=Info)."""
    return _load_master_data(entite)["info"]["SIRET"]


# ──────────────────────────────────────────────────────────────────────
# API publique
# ──────────────────────────────────────────────────────────────────────


def clone_template(
    dst_wb: Workbook, sheet_name: str, *, entite: str
) -> Worksheet:
    """Charge MODELE depuis le template de l'entité et le copie dans dst_wb.

    Args:
        dst_wb: workbook de destination (doit déjà exister, peut être vide ou avoir d'autres sheets).
        sheet_name: nom de la nouvelle feuille (ex. "R_ABC001 Mars 2026").
        entite: code de l'entité (ex. "ROSARI").

    Returns:
        Le nouveau worksheet, prêt à recevoir les données via les autres fonctions.
    """
    gabarit = gabarit_pour(entite)
    if not gabarit.exists():
        raise FileNotFoundError(
            f"Gabarit introuvable : {gabarit}. "
            "Le régénérer avec scripts/build_template.py."
        )
    return _xworkbook_clone_modele(gabarit, dst_wb, sheet_name)


def apply_emetteur(ws: Worksheet, *, entite: str):
    """Re-écrit l'émetteur dans le bloc col C (ROSARI) ou équivalent.

    Idempotent : peut être appelé sur un template déjà rempli.

    Args:
        ws: worksheet cible (créé via clone_template).
        entite: "ROSARI" ou "KAUR".
    """
    layout = _layout_for(entite)
    emetteur = _emetteur_for(entite)
    rmin, rmax = layout["emetteur_rows"]
    col = layout["emetteur_col"]

    lignes = [
        emetteur["nom"],
        emetteur["adresse_l1"],
        emetteur["adresse_l2"],
        emetteur["email"],
        emetteur["tel"],
    ]
    for i, valeur in enumerate(lignes):
        r = rmin + i
        _set_if_not_merged(ws, r, col, valeur)
        cell = ws.cell(row=r, column=col)
        if not isinstance(cell, MergedCell):
            if i == 0:
                cell.font = Font(bold=True, size=12)
            cell.alignment = Alignment(horizontal="center", vertical="center")

    _frame_box(ws, rmin, rmax, col, col + 2)


def apply_destinataire(
    ws: Worksheet, dest: Optional[dict], *, entite: str
):
    """Écrit le destinataire ou vide la zone si dest=None.

    Si dest=None : vide les valeurs (l'encadré reste, c'est un placeholder).
    Si dest=dict : écrit nom, adresse_l1, adresse_l2 (+ l3), email.

    Le template fusionne M7:P8 (le nom sur 2 lignes), ce qui faisait tomber la
    1re ligne d'adresse dans une MergedCell et la supprimait silencieusement.
    On démerge donc la zone et on refusionne une plage par ligne, pour disposer
    d'autant de slots que de lignes d'adresse.

    Args:
        ws: worksheet cible.
        dest: dict avec keys nom/adresse_l1/adresse_l2/adresse_l3/email, ou None.
        entite: "ROSARI" ou "KAUR".

    dest dict structure :
        {"nom": str, "adresse_l1": str, "adresse_l2": str,
         "adresse_l3": str (optional), "email": str (optional)}
    """
    layout = _layout_for(entite)
    rmin, rmax = layout["destinataire_rows"]
    col = layout["destinataire_col"]
    col_end = col + 3  # M:P

    # Normaliser la zone : une plage fusionnée par ligne
    _unmerge_in_range(ws, rmin, rmax, col, col_end)
    for r in range(rmin, rmax + 1):
        ws.merge_cells(
            start_row=r, start_column=col, end_row=r, end_column=col_end
        )
        ws.cell(row=r, column=col).value = None

    _frame_box(ws, rmin, rmax, col, col_end)

    if dest is None:
        # Encadré vide conservé comme placeholder
        return

    # Ville et pays sur la même ligne : la zone n'offre que rmax-rmin+1 slots
    ville = dest.get("adresse_l2", "")
    pays = dest.get("adresse_l3", "")
    if pays:
        ville = f"{ville}, {pays}" if ville else pays

    lignes = [
        v
        for v in (
            dest.get("nom", ""),
            dest.get("adresse_l1", ""),
            ville,
            dest.get("email", ""),
        )
        if v
    ]
    if len(lignes) > rmax - rmin + 1:
        raise RuntimeError(
            f"Destinataire à {len(lignes)} lignes pour une zone de "
            f"{rmax - rmin + 1} lignes (rows {rmin}-{rmax})."
        )

    for i, valeur in enumerate(lignes):
        r = rmin + i
        cell = ws.cell(row=r, column=col)
        cell.value = valeur
        cell.font = Font(bold=True, size=12) if i == 0 else Font(size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center")


def apply_metadata(
    ws: Worksheet,
    n_facture: str,
    date_emission: date,
    *,
    entite: str,
    ref_bon_commande: Optional[str] = None,
    ref_contrat: Optional[str] = None,
):
    """Écrit titre row 2 + Bon de commande row 14 + Contrat row 15 + SIRET D16 + N° Facture D17.

    Args:
        ws: worksheet cible.
        n_facture: ex. "R_ABC001".
        date_emission: date d'émission de la facture.
        entite: "ROSARI" ou "KAUR".
        ref_bon_commande: réf. du bon de commande client (ex. "BC-EXEMPLE-0001").
            Si None, la ligne n'est pas écrite. Obligatoire pour une école dont
            la comptabilité rapproche facture et bon de commande.
        ref_contrat: n° du contrat (ex. "CT-EXEMPLE-0001"), sur sa propre ligne (ADR030 §3).
            Si None, la ligne 15 reste entièrement vide.
    """
    layout = _layout_for(entite)
    siret = _siret_for(entite)

    _set_if_not_merged(ws, layout["bandeau_row"], 2, _titre_facture(date_emission))

    ws[layout["siret_label_cell"]] = "N° SIRET :"
    ws[layout["siret_cell"]] = siret

    ws[layout["n_facture_label_cell"]] = "N° de Facture :"
    ws[layout["n_facture_cell"]] = n_facture

    if ref_bon_commande:
        _ecrire_reference(ws, layout, layout["bon_commande_row"], "Bon de commande :", ref_bon_commande)
    if ref_contrat:
        _ecrire_reference(ws, layout, layout["contrat_row"], LIBELLE_CONTRAT, ref_contrat)


def place_sessions(
    ws: Worksheet,
    ecole_name: str,
    groups: list,
    *,
    entite: str,
    masquer_jour: bool = False,
):
    """Place sessions groupées avec merges Promotion/Libellé/Jour/Heures/Taux/Total.

    Args:
        ws: worksheet cible (template cloné).
        ecole_name: nom de l'école pour la colonne ECOLE (ex. "ÉCOLE FICTIVE B").
        groups: list of (promotion: str, libelle: str, sessions: list).
            Chaque session = (date: date | datetime, heures: float, taux: float).
        entite: "ROSARI" ou "KAUR".
        masquer_jour: masque les colonnes Jour (K:L). À activer quand la facture
            agrège un volume horaire sans séance datée (ex. une école facturée à
            1 ligne par UE de bon de commande). La date reste saisie, simplement
            non imprimée.

    Comportement :
        - Démerge la zone cours du template
        - Vide les valeurs cours
        - Place chaque session : date, heures, taux, formule total
        - Applique merges : Promotion E:F par groupe, Libellé G:J par sous-groupe,
          ECOLE C:D global sur les rows utilisées
        - Si nb_sessions < 60 : masque les rows excédentaires (jamais delete)
        - Si nb_sessions > 60 : insert_rows pour étendre, avec démerge préalable
    """
    layout = _layout_for(entite)
    cols = layout["cols_table"]
    first_row = layout["cours_first_row"]
    last_row = layout["cours_last_row"]
    capacity = last_row - first_row + 1  # 60 rows

    # Aplatir les groupes
    flat = []
    libelle_subranges = []
    promo_ranges = []
    cur_row = first_row
    last_promo = None
    promo_start = None
    for promo, libelle, sessions in groups:
        libelle_start = cur_row
        for sess in sessions:
            d, heures, taux = sess
            flat.append((promo, libelle, d, heures, taux))
            cur_row += 1
        libelle_end = cur_row - 1
        libelle_subranges.append((libelle, libelle_start, libelle_end))
        if promo != last_promo:
            if last_promo is not None:
                promo_ranges.append(
                    (last_promo, promo_start, libelle_start - 1)
                )
            last_promo = promo
            promo_start = libelle_start
    if last_promo is not None:
        promo_ranges.append((last_promo, promo_start, cur_row - 1))

    n_sessions = len(flat)
    last_session_row = first_row + n_sessions - 1

    # Extension dynamique si > capacity
    if n_sessions > capacity:
        n_extra = n_sessions - capacity
        # Démerger les plages sous la zone cours (TOTAL etc.)
        _unmerge_in_range(ws, last_row + 1, ws.max_row, 1, 30)
        # Insérer extra rows juste avant TOTAL
        ws.insert_rows(last_row + 1, n_extra)
        # Mettre à jour last_row
        new_last_row = last_row + n_extra
        # Appliquer pattern cours sur les nouvelles rows
        for r in range(last_row + 1, new_last_row + 1):
            ws.merge_cells(start_row=r, start_column=cols["promo"], end_row=r, end_column=cols["promo"] + 1)
            ws.merge_cells(start_row=r, start_column=cols["libelle"], end_row=r, end_column=cols["libelle"] + 3)
            ws.merge_cells(start_row=r, start_column=cols["date"], end_row=r, end_column=cols["date"] + 1)
            ws.merge_cells(start_row=r, start_column=cols["heures"], end_row=r, end_column=cols["heures"] + 1)
            ws.merge_cells(start_row=r, start_column=cols["taux"], end_row=r, end_column=cols["taux"] + 1)
            ws.merge_cells(start_row=r, start_column=cols["total"], end_row=r, end_column=cols["total"] + 1)
            ws.row_dimensions[r].height = 46.5
        last_row = new_last_row
        # Mettre à jour les formules TOTAL et À payer
        total_row = layout["total_row"] + n_extra
        ws.cell(row=total_row, column=cols["total"]).value = (
            f"=SUM({chr(64 + cols['total'])}{first_row}:{chr(64 + cols['total'] + 1)}{last_row})"
        )

    # Démerger zone cours actuelle (rows first_row à last_row, cols 3-18)
    # Important : démerger AUSSI la merge ECOLE globale C20:D79 si présente
    _unmerge_in_range(ws, first_row, last_row, 3, 18)

    # Vider toutes les valeurs cours
    for r in range(first_row, last_row + 1):
        for c in range(3, 19):
            cell = ws.cell(row=r, column=c)
            if not isinstance(cell, MergedCell):
                cell.value = None

    # Re-merger pattern uniforme
    for r in range(first_row, last_row + 1):
        ws.merge_cells(start_row=r, start_column=cols["promo"], end_row=r, end_column=cols["promo"] + 1)
        ws.merge_cells(start_row=r, start_column=cols["libelle"], end_row=r, end_column=cols["libelle"] + 3)
        ws.merge_cells(start_row=r, start_column=cols["date"], end_row=r, end_column=cols["date"] + 1)
        ws.merge_cells(start_row=r, start_column=cols["heures"], end_row=r, end_column=cols["heures"] + 1)
        ws.merge_cells(start_row=r, start_column=cols["taux"], end_row=r, end_column=cols["taux"] + 1)
        ws.merge_cells(start_row=r, start_column=cols["total"], end_row=r, end_column=cols["total"] + 1)

    # ECOLE merge global C:D sur toutes les rows utilisées (et pas plus)
    if n_sessions > 0:
        ws.merge_cells(
            start_row=first_row,
            start_column=cols["ecole"],
            end_row=last_session_row,
            end_column=cols["ecole"] + 1,
        )

    # Placer chaque session
    for i, (promo, libelle, d, heures, taux) in enumerate(flat):
        r = first_row + i

        # Date col K (mergée K:L)
        cd = ws.cell(row=r, column=cols["date"])
        cd.value = d
        cd.number_format = "DD/MM/YYYY"
        cd.font = Font(size=14)
        cd.alignment = Alignment(horizontal="center", vertical="center")

        # Heures col M (mergée M:N)
        ch = ws.cell(row=r, column=cols["heures"])
        ch.value = heures
        ch.font = Font(size=14)
        ch.alignment = Alignment(horizontal="center", vertical="center")

        # Taux col O (mergée O:P)
        ct = ws.cell(row=r, column=cols["taux"])
        ct.value = taux
        ct.font = Font(size=14)
        ct.alignment = Alignment(horizontal="center", vertical="center")

        # Total col Q (mergée Q:R) — formule
        col_taux_letter = chr(64 + cols["taux"])
        col_heures_letter = chr(64 + cols["heures"])
        cq = ws.cell(row=r, column=cols["total"])
        cq.value = f"={col_taux_letter}{r}*{col_heures_letter}{r}"
        cq.font = Font(size=14)
        cq.alignment = Alignment(horizontal="center", vertical="center")

    # ECOLE valeur en C20
    if n_sessions > 0:
        ce = ws.cell(row=first_row, column=cols["ecole"])
        ce.value = ecole_name
        ce.font = Font(bold=True, size=18)
        ce.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    # Promotion + Libellé valeurs (avec merges multi-rows si group >= 2)
    # Note : on a fait les merges 1-row par défaut. Pour les groupes multi-rows,
    # on démerge les rows correspondantes et on re-merge multi-row.
    for promo, ps, pe in promo_ranges:
        if pe > ps:
            # Multi-row : démerger les rows individuelles, re-merger multi-row
            for r in range(ps, pe + 1):
                for mr in list(ws.merged_cells.ranges):
                    if (
                        mr.min_row == r
                        and mr.max_row == r
                        and mr.min_col == cols["promo"]
                        and mr.max_col == cols["promo"] + 1
                    ):
                        ws.unmerge_cells(str(mr))
                        break
            ws.merge_cells(
                start_row=ps,
                start_column=cols["promo"],
                end_row=pe,
                end_column=cols["promo"] + 1,
            )
        cell = ws.cell(row=ps, column=cols["promo"])
        cell.value = promo
        cell.font = Font(size=16)
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    for libelle, ls, le in libelle_subranges:
        if le > ls:
            for r in range(ls, le + 1):
                for mr in list(ws.merged_cells.ranges):
                    if (
                        mr.min_row == r
                        and mr.max_row == r
                        and mr.min_col == cols["libelle"]
                        and mr.max_col == cols["libelle"] + 3
                    ):
                        ws.unmerge_cells(str(mr))
                        break
            ws.merge_cells(
                start_row=ls,
                start_column=cols["libelle"],
                end_row=le,
                end_column=cols["libelle"] + 3,
            )
        cell = ws.cell(row=ls, column=cols["libelle"])
        cell.value = libelle
        cell.font = Font(size=14)
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    # Masquer rows excédentaires (jamais delete)
    for r in range(last_session_row + 1, last_row + 1):
        ws.row_dimensions[r].hidden = True

    if masquer_jour:
        for c in (cols["date"], cols["date"] + 1):
            ws.column_dimensions[get_column_letter(c)].hidden = True


def apply_mentions(
    ws: Worksheet,
    *,
    entite: str,
    date_emission: date,
    echeance: Optional[date] = None,
):
    """Écrit les totaux et les mentions légales du bas de facture (ADR030, F-24).

    À appeler après place_sessions : au-delà de 60 séances, les lignes insérées
    décalent d'autant le total, les mentions et la zone d'impression.

    Écrit (plan du gabarit standard, sans décalage) : « Total HT » en O80 ; « Net à
    payer : » en C83, avec en E83 la formule qui reprend le total ; C84 vidée ; la
    mention de TVA en C85 ; le n° de déclaration d'activité en C86 ; l'échéance en
    C91 ; escompte, pénalités et indemnité forfaitaire en C92, C93 et C94.

    Args:
        ws: worksheet cible.
        entite: code de l'entité, qui fixe le MASTER où lire `Info` / `NDA`.
        date_emission: date d'émission de la facture.
        echeance: délai du contrat ou de l'école ; None pour la date d'émission plus 30 jours.

    Raises:
        ValueError: `Info` / `NDA` absent du MASTER, ou échéance antérieure à la date
            d'émission. Rien n'est écrit dans ce cas.
    """
    layout = _layout_for(entite)
    # Les données obligatoires sont lues avant toute écriture : en cas d'erreur, l'onglet reste intact.
    nda = get_nda(entite=entite)
    texte_echeance = _texte_echeance(calculer_echeance(date_emission, echeance))
    decalage = _decalage(ws, layout)
    cols = layout["cols_table"]
    col = layout["mentions_col"]
    courte, longue = layout["mentions_fin_courte"], layout["mentions_fin_longue"]
    total_row = layout["total_row"] + decalage

    def ligne(cle: str) -> int:
        return layout[cle] + decalage

    ws.cell(row=total_row, column=cols["taux"]).value = MENTION_TOTAL_HT

    net = ligne("net_a_payer_row")
    _fusionner_ligne(ws, net, col, col + 1)
    _fusionner_ligne(ws, net, col + 2, col + 3)
    ws.cell(row=net, column=col).value = MENTION_NET_A_PAYER
    ws.cell(row=net, column=col + 2).value = f"={get_column_letter(cols['total'])}{total_row}"

    ws.cell(row=ligne("ancienne_asterisque_row"), column=col).value = None

    # (ligne, texte, dernière colonne fusionnée ou None, ligne dont la cellule C donne la police).
    # Le NDA n'est pas fusionné : un texte officiel long déborde jusqu'au bord de la zone d'impression.
    for cle, texte, col_fin, modele in (
        ("tva_row", MENTION_TVA, longue, "tva_row"),
        ("nda_row", nda, None, "tva_row"),
        ("echeance_row", texte_echeance, courte, "echeance_row"),
        ("escompte_row", MENTION_ESCOMPTE, courte, "escompte_row"),
        ("penalites_row", MENTION_PENALITES, longue, "penalites_row"),
        ("indemnite_row", MENTION_INDEMNITE, longue, "penalites_row"),
    ):
        r = ligne(cle)
        if col_fin is not None:
            _fusionner_ligne(ws, r, col, col_fin)
        cell = ws.cell(row=r, column=col)
        cell.value = texte
        if modele != cle:
            source = ws.cell(row=ligne(modele), column=col)
            cell.font = _copy(source.font)
            cell.alignment = _copy(source.alignment)

    _zone_impression(ws, layout)


def apply_print_setup(ws: Worksheet, *, entite: str):
    """Fige la mise en page : portrait, ajusté à 1 page de large.

    Le template fait 18 colonnes ; sans ajustement, l'export PDF coupe les
    colonnes Taux/Total sur une 2e page. La zone d'impression va jusqu'à la
    dernière mention légale, décalée comme elle au-delà de 60 séances.

    Args:
        ws: worksheet cible.
        entite: "ROSARI" ou "KAUR".
    """
    layout = _layout_for(entite)
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    _zone_impression(ws, layout)


def apply_rib(ws: Worksheet, *, entite: str):
    """Re-écrit IBAN/BIC dans les cells du template (idempotent).

    Args:
        ws: worksheet cible.
        entite: "ROSARI" ou "KAUR".
    """
    layout = _layout_for(entite)
    rib = _rib_for(entite)
    ws[layout["rib_iban_cell"]] = rib["IBAN"]
    ws[layout["rib_bic_cell"]] = rib["BIC"]


def post_mortem(ws: Worksheet, *, entite: str) -> dict:
    """Retourne les cellules clés pour vérification après écriture.

    Args:
        ws: worksheet à inspecter.
        entite: "ROSARI" ou "KAUR".

    Returns:
        dict avec : titre, emetteur_nom, destinataire_nom, n_facture, siret,
                    ecole, premier_libelle, total_formule, iban, bic,
                    echeance (texte de la ligne), nda, contrat.
    """
    layout = _layout_for(entite)
    cols = layout["cols_table"]
    rmin_em, _ = layout["emetteur_rows"]
    rmin_dest, _ = layout["destinataire_rows"]
    col_mentions = layout["mentions_col"]

    return {
        "titre": ws.cell(row=layout["bandeau_row"], column=2).value,
        "emetteur_nom": ws.cell(row=rmin_em, column=layout["emetteur_col"]).value,
        "destinataire_nom": ws.cell(
            row=rmin_dest, column=layout["destinataire_col"]
        ).value,
        "n_facture": ws[layout["n_facture_cell"]].value,
        "siret": ws[layout["siret_cell"]].value,
        "ecole": ws.cell(row=layout["cours_first_row"], column=cols["ecole"]).value,
        "premier_libelle": ws.cell(
            row=layout["cours_first_row"], column=cols["libelle"]
        ).value,
        "total_formule": ws.cell(
            row=layout["total_row"], column=cols["total"]
        ).value,
        "iban": ws[layout["rib_iban_cell"]].value,
        "bic": ws[layout["rib_bic_cell"]].value,
        "echeance": ws.cell(row=layout["echeance_row"], column=col_mentions).value,
        "nda": ws.cell(row=layout["nda_row"], column=col_mentions).value,
        "contrat": ws.cell(
            row=layout["contrat_row"], column=layout["bon_commande_value_cols"][0]
        ).value,
    }


def verifier_facture(
    ws: Worksheet,
    *,
    entite: str,
    n_facture: str,
    date_emission: date,
    dest: Optional[dict],
    ecole_name: str,
    groups: list,
    echeance: Optional[date] = None,
    ref_contrat: Optional[str] = None,
) -> list:
    """Compare les cellules relues d'une facture aux valeurs attendues (PCF étape 7.3, FA-12).

    Contrôle les cellules du post-mortem et, ligne par ligne, la date, les heures et
    le taux de chaque séance. Les valeurs attendues viennent des mêmes sources que
    l'écriture : la spec (numéro, date, échéance, contrat, séances) et
    MASTER.Coordonnées (émetteur, SIRET, IBAN, BIC, NDA). L'échéance, le NDA et la
    ligne de contrat sont toujours contrôlés : sans `echeance`, l'échéance attendue est
    la date d'émission plus 30 jours ; sans `ref_contrat`, la ligne de contrat doit être vide.

    Args:
        ws: onglet de la facture, de préférence relu depuis le fichier écrit.
        entite: code de l'entité.
        n_facture, date_emission, dest, ecole_name, groups: ceux passés à
            apply_metadata, apply_destinataire et place_sessions.
        echeance, ref_contrat: ceux passés à apply_mentions et apply_metadata.

    Returns:
        Une ligne « ERREUR sur cellule … : attendu …, trouvé … » par écart ; liste vide
        si la facture est conforme.

    Raises:
        ValueError: `Info` / `NDA` absent du MASTER, ou échéance antérieure à l'émission.
    """
    layout = _layout_for(entite)
    cols = layout["cols_table"]
    donnees = _load_master_data(entite)
    pm = post_mortem(ws, entite=entite)
    lettre_total = get_column_letter(cols["total"])
    lettre_total_fin = get_column_letter(cols["total"] + 1)

    attendu = {
        "titre": _titre_facture(date_emission),
        "emetteur_nom": donnees["emetteur"]["nom"],
        "destinataire_nom": dest["nom"] if dest else None,
        "n_facture": n_facture,
        "siret": donnees["info"]["SIRET"],
        "ecole": ecole_name,
        "premier_libelle": groups[0][1] if groups else None,
        "total_formule": (
            f"=SUM({lettre_total}{layout['cours_first_row']}:"
            f"{lettre_total_fin}{layout['cours_last_row']})"
        ),
        "iban": donnees["info"]["IBAN"],
        "bic": donnees["info"]["BIC"],
        "echeance": _texte_echeance(calculer_echeance(date_emission, echeance)),
        "nda": get_nda(entite=entite),
        "contrat": ref_contrat or None,  # apply_metadata n'écrit pas de contrat vide
    }
    ecarts = [
        f"ERREUR sur cellule {cle} : attendu {valeur!r}, trouvé {pm[cle]!r}"
        for cle, valeur in attendu.items()
        if pm[cle] != valeur
    ]

    seances = [seance for _, _, sessions in groups for seance in sessions]
    for i, (jour, heures, taux) in enumerate(seances):
        ligne = layout["cours_first_row"] + i
        for nom, col, valeur in (
            ("date", cols["date"], jour), ("heures", cols["heures"], heures), ("taux", cols["taux"], taux),
        ):
            lu = ws.cell(row=ligne, column=col).value
            if nom == "date" and hasattr(lu, "date") and not hasattr(valeur, "hour"):
                lu = lu.date()  # une date relue revient en datetime
            if lu != valeur:
                coord = f"{get_column_letter(col)}{ligne}"
                ecarts.append(f"ERREUR sur cellule {nom} ({coord}) : attendu {valeur!r}, trouvé {lu!r}")
    return ecarts
