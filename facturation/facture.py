"""facture.py — Une facture décrite par une spec JSON : aperçu, refus, écriture et contrôle (P-0a).

Le code vient de `scripts/nouvelle_facture.py`, qui n'en est plus que l'enveloppe en ligne de
commande : F-16 et F-17 importent d'ici l'aperçu, les refus et l'écriture d'un onglet, sans toucher
à `sys.path`. Une facture s'écrit par un seul chemin de code et se contrôle par un seul :

- `ecrire_onglet` écrit l'onglet de la facture dans un classeur ouvert (PCF étape 5), sans rien
  enregistrer ; elle refuse une autre entité que celle de la spec, un onglet ou un numéro déjà
  présents, et retire l'onglet commencé si une erreur survient en cours d'écriture ;
- `verifier_onglet` compare un onglet, titre compris, aux valeurs déduites de la spec (PCF
  étape 7.3) ;
- `generer` enchaîne les contrôles, l'aperçu (PCF étape 6), `ecrire_onglet`, l'enregistrement, le
  post-mortem (PCF étape 7.2) et `verifier_onglet` sur l'onglet relu du fichier. L'émission (F-17)
  ne passe pas par `generer` : elle appelle `ecrire_onglet` et `verifier_onglet` sous verrou.

Les étapes humaines du protocole — périmètre, coordonnées, aperçu — restent à la charge de
l'appelant (cf. references/PROC00_pcf.md).

Format de la spec (voir examples/spec_exemple.json) :
{
  "entite": "ROSARI",
  "ecole": "ABC",
  "ecole_label": "ÉCOLE FICTIVE A",
  "fichier_sortie": "ROSARI/ABC/ROSARI_ABC.xlsx",
  "onglet": "R_ABC001 Sept 2026",
  "n_facture": "R_ABC001",
  "date_emission": "2026-09-30",
  "ref_bon_commande": "BC-EXEMPLE-0001",
  "ref_contrat": "CT-EXEMPLE-0001",        (facultatif)
  "echeance": "2026-10-30",                (facultatif : défaut émission + 30 jours)
  "masquer_jour": true,                    (facultatif : true ou false, défaut false)
  "lignes": [
    {"promotion": "M1 S1", "libelle": "[2026][M1 S1][Module A][...]",
     "seances": [{"date": "2026-09-15", "heures": 18, "taux": 90}]}
  ]
}
La spec s'enregistre en UTF-8 (avec ou sans BOM). Le titre d'onglet est un texte non vide d'au plus
31 caractères, sans \\ / ? * [ ] : (règles d'Excel). Les dates s'écrivent AAAA-MM-JJ. `lignes` et
`seances` sont des listes, même d'un seul élément ; chaque ligne porte `promotion`, `libelle` (texte)
et au moins une séance ; heures et taux sont des nombres écrits comme tels (jamais une chaîne),
finis, les heures strictement positives et le taux positif ou nul.
"""

from __future__ import annotations

import json
import math
import re
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path, PurePosixPath, PureWindowsPath

from openpyxl import Workbook, load_workbook
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.worksheet.worksheet import Worksheet

from facturation import config
from facturation import pcf_format as fmt

__all__ = [
    "EcritureRefusee",
    "PostMortemEnEcart",
    "echeance_de_la_spec",
    "charger_spec",
    "controler_spec",
    "construire_groupes",
    "conflits_ecriture",
    "apercu",
    "generer",
    "ecrire_onglet",
    "verifier_onglet",
    "fmt",
    "config",
]


class EcritureRefusee(RuntimeError):
    """La facture ne peut pas être écrite : rien n'a été modifié sur le disque."""


class PostMortemEnEcart(RuntimeError):
    """La facture est écrite, mais des cellules relues diffèrent des valeurs attendues."""

    def __init__(self, ecarts: list[str]):
        self.ecarts = ecarts
        super().__init__("\n".join(ecarts))


def _parse_date(valeur) -> date:
    """Accepte 'YYYY-MM-DD' ou un objet date déjà construit."""
    if isinstance(valeur, date):
        return valeur
    return datetime.strptime(str(valeur), "%Y-%m-%d").date()


def echeance_de_la_spec(spec: dict) -> date:
    """Échéance de la facture (ADR030) : `echeance` de la spec, sinon date d'émission + 30 jours.

    Raises:
        ValueError: échéance mal formée (attendu AAAA-MM-JJ) ou antérieure à la date d'émission.
    """
    donnee = spec.get("echeance")
    return fmt.calculer_echeance(
        _parse_date(spec["date_emission"]), _parse_date(donnee) if donnee else None
    )


_LONGUEUR_MAX_ONGLET = 31  # limite d'Excel ; openpyxl n'émet qu'un avertissement au-delà
# Caractères qu'Excel refuse dans un titre d'onglet ; openpyxl lève une ValueError qui ne cite pas le titre.
_CARACTERES_INTERDITS_ONGLET = re.compile(r"[\\/?*\[\]:]")
_LONGUEUR_MAX_CITATION = 60


def _citer(valeur) -> str:
    """La valeur fautive telle que Python l'écrit (`repr`), abrégée si elle est longue (ligne entière, nombre géant)."""
    try:
        texte = repr(valeur)
    except ValueError:  # entier trop long pour être écrit en décimal (limite de Python)
        return "un nombre entier de plus de 4 300 chiffres"
    if len(texte) <= _LONGUEUR_MAX_CITATION:
        return texte
    return f"{texte[:_LONGUEUR_MAX_CITATION // 2]}… ({len(texte)} caractères)"


def _nombre_fini(valeur) -> bool:
    """Nombre écrit comme tel (ni chaîne, ni booléen, ni None), fini : ni NaN ni l'infini, ni un entier trop
    grand pour un nombre décimal (10**400 est du JSON valide, et `math.isfinite` lève alors OverflowError)."""
    if isinstance(valeur, bool) or not isinstance(valeur, (int, float)):
        return False
    try:
        return math.isfinite(valeur)
    except OverflowError:
        return False


def _date_lisible(valeur) -> bool:
    """Date AAAA-MM-JJ, ou objet date déjà construit : ce que `_parse_date` accepte."""
    try:
        _parse_date(valeur)
    except (TypeError, ValueError):
        return False
    return True


def _refuser_titre_d_onglet_invalide(onglet) -> None:
    """Refuse un titre d'onglet qu'Excel n'accepte pas : autre chose qu'un texte (null, nombre : jamais converti
    en silence, et openpyxl nomme « Sheet » un onglet sans titre), vide, plus de 31 caractères, ou l'un des
    caractères \\ / ? * [ ] :.

    Raises:
        ValueError: le message nomme l'onglet et la valeur fautive.
    """
    if not isinstance(onglet, str):
        raise ValueError(
            f"onglet = {_citer(onglet)} : attendu le titre de l'onglet en texte, entre guillemets "
            "(ex. \"R_ABC001 Sept 2026\")."
        )
    if not onglet:
        raise ValueError("titre d'onglet vide : donner un titre à l'onglet dans la spec (ex. \"R_ABC001 Sept 2026\").")
    interdit = _CARACTERES_INTERDITS_ONGLET.search(onglet)
    if interdit:
        raise ValueError(
            f"titre d'onglet « {onglet} » : caractère « {interdit.group()} » interdit par Excel "
            "(\\ / ? * [ ] :) — le remplacer dans la spec."
        )
    if len(onglet) > _LONGUEUR_MAX_ONGLET:
        raise ValueError(
            f"titre d'onglet « {onglet} » trop long : {len(onglet)} caractères, pour "
            f"{_LONGUEUR_MAX_ONGLET} au plus (limite d'Excel) — le raccourcir dans la spec."
        )


def _refuser_seance_mal_formee(seance, ou: str) -> None:
    """Refuse une séance qui n'est pas un objet, sans date AAAA-MM-JJ, ou aux heures ou au taux invalides.

    Raises:
        ValueError: le message commence par `ou` (rangs de la ligne et de la séance) et cite la valeur fautive.
    """
    if not isinstance(seance, dict):
        raise ValueError(
            f"{ou} : attendu un objet {{\"date\": …, \"heures\": …, \"taux\": …}}, trouvé {_citer(seance)}."
        )
    if "date" not in seance:
        raise ValueError(f"{ou} : clé date absente — écrire la date de la séance (AAAA-MM-JJ).")
    if not _date_lisible(seance["date"]):
        raise ValueError(
            f"{ou} : date = {_citer(seance['date'])} — attendu une date AAAA-MM-JJ (ex. 2026-09-15)."
        )
    for champ in ("heures", "taux"):
        if champ not in seance:
            raise ValueError(f"{ou} : clé {champ} absente.")
        valeur = seance[champ]
        if isinstance(valeur, int) and not isinstance(valeur, bool) and not _nombre_fini(valeur):
            raise ValueError(
                f"{ou} : {champ} = {_citer(valeur)} — nombre trop grand, hors de portée d'un nombre décimal."
            )
        if not _nombre_fini(valeur):
            raise ValueError(
                f"{ou} : {champ} = {_citer(valeur)} n'est pas un nombre fini — écrire un nombre, "
                "sans guillemets, avec un point décimal (ex. 1.5)."
            )
        if champ == "heures" and valeur <= 0:
            raise ValueError(f"{ou} : heures = {valeur!r} — attendu un nombre d'heures "
                             "strictement positif.")
        if champ == "taux" and valeur < 0:
            raise ValueError(f"{ou} : taux = {valeur!r} — un taux ne peut pas être négatif.")


def _refuser_spec_mal_formee(spec: dict) -> None:
    """Refuse, avant toute écriture, une spec que l'écriture convertirait en silence, fausserait ou ferait échouer
    par une erreur Python sans rang.

    - titre d'onglet qu'Excel refuse : autre chose qu'un texte, vide, plus de 31 caractères, caractère interdit ;
    - `masquer_jour` présent qui n'est pas un booléen (`bool("false")` vaut vrai) ;
    - `date_emission` qui n'est pas une date AAAA-MM-JJ ;
    - `lignes` qui n'est pas une liste, ou liste vide ;
    - ligne qui n'est pas un objet, sans `promotion`, sans `libelle` ou au libellé qui n'est pas un texte ;
    - ligne sans séance (`seances` vide ou absente) : elle disparaîtrait de la facture ; `seances` qui n'est pas
      une liste ;
    - séance qui n'est pas un objet, sans date AAAA-MM-JJ, ou dont les heures ou le taux manquent ou ne sont pas
      des nombres finis écrits comme tels (jamais de conversion d'une chaîne, même « 18 »), heures nulles ou
      négatives, taux négatif (un taux à 0 reste admis).

    Les rangs des lignes, et des séances dans leur ligne, se comptent à partir de 1. L'échéance reste contrôlée
    par `generer` et `echeance_de_la_spec`.

    Raises:
        ValueError: le message nomme la cause, la ligne, la séance et la valeur fautive.
        KeyError: `onglet`, `date_emission` ou `lignes` absent de la spec.
    """
    _refuser_titre_d_onglet_invalide(spec["onglet"])
    if "masquer_jour" in spec and not isinstance(spec["masquer_jour"], bool):
        raise ValueError(
            f"masquer_jour = {spec['masquer_jour']!r} : attendu true ou false (booléen, sans "
            "guillemets), ou la clé absente."
        )
    if not _date_lisible(spec["date_emission"]):
        raise ValueError(
            f"date_emission = {_citer(spec['date_emission'])} — attendu une date AAAA-MM-JJ (ex. 2026-09-30)."
        )
    lignes = spec["lignes"]
    if not lignes:
        raise ValueError("La spec ne contient aucune ligne de cours.")
    if not isinstance(lignes, list):
        raise ValueError(
            f"lignes = {_citer(lignes)} : attendu une liste de lignes, entre crochets "
            "([{\"promotion\": …, \"libelle\": …, \"seances\": […]}, …])."
        )
    for rang, ligne in enumerate(lignes, start=1):
        if not isinstance(ligne, dict):
            raise ValueError(
                f"ligne {rang} : attendu un objet {{\"promotion\": …, \"libelle\": …, \"seances\": […]}}, "
                f"trouvé {_citer(ligne)}."
            )
        if "libelle" not in ligne:
            raise ValueError(f"ligne {rang} : clé libelle absente — écrire le libellé de la ligne.")
        libelle = ligne["libelle"]
        if not isinstance(libelle, str):
            raise ValueError(f"ligne {rang} : libelle = {_citer(libelle)} — attendu un texte, entre guillemets.")
        if "promotion" not in ligne:
            raise ValueError(f"ligne {rang} « {libelle} » : clé promotion absente — écrire la promotion de la ligne.")
        seances = ligne.get("seances")
        if not seances:
            raise ValueError(
                f"ligne {rang} « {libelle} » sans séance (clé seances vide ou absente) : lui "
                "ajouter ses séances, ou la retirer de la spec."
            )
        if not isinstance(seances, list):
            raise ValueError(
                f"ligne {rang} « {libelle} » : seances = {_citer(seances)} — attendu une liste de séances, "
                "entre crochets, même pour une seule séance ([{\"date\": …, \"heures\": …, \"taux\": …}])."
            )
        for n, seance in enumerate(seances, start=1):
            _refuser_seance_mal_formee(seance, f"ligne {rang} « {libelle} », séance {n}")


def charger_spec(chemin: Path) -> dict:
    """Lit la spec JSON (UTF-8, avec ou sans BOM) et vérifie sa structure avant tout usage.

    Raises:
        ValueError: fichier qui n'est pas en UTF-8 (UTF-16 ou ANSI compris), JSON illisible, spec qui
            n'est pas un objet, clés obligatoires manquantes, aucune ligne de cours, ou spec mal formée
            (`_refuser_spec_mal_formee` : titre d'onglet, `masquer_jour`, date d'émission, lignes et
            séances) ; le message, en français, nomme la cause et ce qu'il faut corriger.
        OSError: fichier introuvable ou illisible (`FileNotFoundError` s'il n'existe pas).
    """
    brut = chemin.read_bytes()
    try:
        # « utf-8-sig » : le Bloc-notes et PowerShell 5 écrivent l'UTF-8 avec un BOM.
        texte = brut.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        # PowerShell 5.1 écrit en UTF-16 avec `>` ou `Out-File` ; un éditeur ancien, en ANSI (cp1252).
        nature = (
            "elle est en UTF-16"
            if brut.startswith((b"\xff\xfe", b"\xfe\xff"))
            else f"octet 0x{exc.object[exc.start]:02X} hors UTF-8, sans doute de l'ANSI (Windows-1252)"
        )
        raise ValueError(
            f"{chemin.name} n'est pas encodée en UTF-8 ({nature}) : l'enregistrer en UTF-8 (Bloc-notes : "
            "Enregistrer sous, Encodage UTF-8 ; PowerShell : Out-File -Encoding utf8), puis relancer."
        ) from None
    try:
        spec = json.loads(texte)
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"{chemin.name} n'est pas un JSON lisible, ligne {exc.lineno}, colonne {exc.colno} ({exc.msg}) : "
            "corriger la syntaxe (virgules, guillemets, accolades)."
        ) from None
    if not isinstance(spec, dict):
        raise ValueError(
            f"{chemin.name} : attendu un objet JSON entre accolades ({{\"entite\": …, \"lignes\": […]}}), "
            f"trouvé {_citer(spec)}."
        )
    obligatoires = {
        "entite", "ecole", "fichier_sortie", "onglet",
        "n_facture", "date_emission", "lignes",
    }
    manquantes = obligatoires - set(spec)
    if manquantes:
        raise ValueError(f"Clés manquantes dans {chemin.name} : {sorted(manquantes)}")
    _refuser_spec_mal_formee(spec)
    return spec


def controler_spec(spec: dict, *, prefixe: str | None, racine: Path) -> Path:
    """Vérifie que le numéro et le fichier de sortie concordent avec l'entité (FA-03).

    Les entités ne se mélangent jamais : le numéro commence par le préfixe de
    l'entité, et le classeur se range sous `<racine>/<ENTITE>/`, en chemin relatif,
    sans « .. ». Les séparateurs « / » et « \\ » sont acceptés l'un comme l'autre.

    Args:
        spec: spec de la facture (clés `entite`, `n_facture`, `fichier_sortie`).
        prefixe: préfixe des numéros de l'entité (`prefixe_facture` du profil) ; None
            pour la règle par défaut (initiale de l'entité + « _ »).
        racine: racine de données.

    Returns:
        Le chemin du classeur de sortie, sous `<racine>/<ENTITE>/`.

    Raises:
        EcritureRefusee: entité invalide, ou numéro ou fichier de sortie incohérent avec elle.
    """
    try:
        entite = fmt.valider_entite(spec["entite"])
    except ValueError as exc:
        raise EcritureRefusee(f"entite {spec['entite']!r} : {exc}") from exc
    prefixe = prefixe or config.prefixe_par_defaut(entite)
    n_facture = str(spec["n_facture"])
    if not n_facture.startswith(prefixe):
        raise EcritureRefusee(
            f"le numéro {n_facture} ne commence pas par le préfixe {prefixe} de l'entité {entite}"
        )

    brut = str(spec["fichier_sortie"])
    chemin = PureWindowsPath(brut)  # reconnaît aussi les lecteurs (C:) et les chemins réseau
    parties = chemin.parts
    relatif_sous_entite = (
        not chemin.anchor
        and not PurePosixPath(brut).is_absolute()
        and ".." not in parties
        and len(parties) >= 2
        and parties[0] == entite
    )
    sortie = racine.joinpath(*parties) if relatif_sous_entite else None
    # Dernier filet : le chemin résolu (liens symboliques compris) reste sous <racine>/<ENTITE>/.
    if sortie is None or not sortie.resolve().is_relative_to((racine / entite).resolve()):
        raise EcritureRefusee(
            f"fichier_sortie {brut!r} invalide : attendu un chemin relatif sous {entite}/ "
            f"dans la racine de données, sans « .. » (ex. {entite}/<ECOLE>/{entite}_<ECOLE>.xlsx)"
        )
    return sortie


def construire_groupes(lignes: list) -> list:
    """Transforme les lignes de la spec en `groups` pour place_sessions()."""
    groupes = []
    for ligne in lignes:
        seances = [
            (_parse_date(s["date"]), float(s["heures"]), float(s["taux"]))
            for s in ligne["seances"]
        ]
        groupes.append((ligne["promotion"], ligne["libelle"], seances))
    return groupes


def conflits_ecriture(
    wb: Workbook | None, onglet: str, n_facture: str, *, entite: str
) -> list[str]:
    """Raisons qui interdisent d'écrire cette facture dans le classeur de l'école (FA-02).

    Une facture émise ne se modifie pas : elle s'annule par un avoir. L'écriture est
    donc refusée si l'onglet existe déjà, y compris à la casse près (Excel ne
    distingue pas « R_X001 » de « r_x001 », et openpyxl créerait en silence un
    onglet suffixé « 1 »), ou si le numéro figure déjà dans un onglet du classeur :
    dans son titre (factures renommées « PAID … ») ou dans sa cellule N° de facture.

    Args:
        wb: classeur cible, ou None s'il n'existe pas encore.
        onglet: titre de l'onglet à créer.
        n_facture: numéro de la facture à créer.
        entite: code de l'entité, qui fixe la cellule du N° de facture.

    Returns:
        Une raison par onglet en conflit ; liste vide si l'écriture est possible.
    """
    if wb is None:
        return []
    numero = re.compile(rf"(?<![A-Za-z0-9]){re.escape(n_facture)}(?![A-Za-z0-9])", re.IGNORECASE)
    plan = fmt.LAYOUTS.get(fmt.valider_entite(entite), fmt.LAYOUT_STANDARD)
    ligne_n, col_n = coordinate_to_tuple(plan["n_facture_cell"])
    raisons = []
    for nom in wb.sheetnames:
        if nom.casefold() == onglet.casefold():
            casse = "" if nom == onglet else " (à la casse près)"
            raisons.append(f"l'onglet '{nom}' existe déjà{casse}")
            continue
        feuille = wb[nom]
        # Lire la seule cellule N°, et seulement si elle est dans l'étendue de l'onglet :
        # openpyxl crée toute cellule lue, ce qui agrandirait l'onglet au prochain save.
        cellule = (
            feuille.cell(row=ligne_n, column=col_n).value
            if isinstance(feuille, Worksheet)  # une feuille graphique n'a pas de cellule
            and ligne_n <= feuille.max_row and col_n <= feuille.max_column
            else None
        )
        dans_cellule = isinstance(cellule, str) and cellule.strip().casefold() == n_facture.casefold()
        if numero.search(nom) or dans_cellule:
            raisons.append(f"le numéro {n_facture} figure déjà dans l'onglet '{nom}'")
    return raisons


def _au_centime(montant: Decimal) -> Decimal:
    """Montant arrondi au centime, la moitié vers le haut (49,995 → 50.00)."""
    return montant.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def apercu(
    spec: dict,
    groupes: list,
    dest: dict,
    conflits: list | None = None,
    *,
    emetteur: dict,
    echeance: date | None = None,
    nda: str | None = None,
) -> str:
    """Construit l'aperçu du PCF étape 6 (gate humain avant écriture).

    Le libellé s'affiche en entier (il recopie à l'identique le code produit du bon
    de commande) ; les colonnes s'élargissent au plus long libellé et à la plus
    longue promotion (FA-05). Les raisons de `conflits` (FA-02) s'affichent dans
    l'état de l'onglet : nouvel onglet, ou écriture refusée.

    Les montants se calculent en décimal à partir de la valeur écrite (`str`, jamais le float
    binaire : 1,5 h × 33,33 € font 49,995 €, pas 49,99499…) et s'arrondissent au centime, la
    moitié vers le haut. Le TOTAL HT arrondit une seule fois la somme exacte des produits, comme le
    `=SUM` du gabarit qui additionne des produits non arrondis. La colonne taux s'arrondit de même
    (45,125 € s'affiche 45.13 €, comme le montant d'une heure), et le TOTAL HEURES se somme en
    décimal (3 × 0,7 h font 2.1 h, pas 2.0999999999999996 h), au format inchangé.

    Args:
        echeance: échéance à annoncer ; par défaut, celle de la spec (ADR030).
        nda: n° de déclaration d'activité lu dans le MASTER ; « — » s'il n'est pas fourni.
    """
    echeance = echeance or echeance_de_la_spec(spec)
    seances_txt = [
        (d, str(promo), libelle.replace("\n", " "), heures, taux)
        for promo, libelle, seances in groupes
        for d, heures, taux in seances
    ]
    largeur_promo = max([12, *(len(s[1]) for s in seances_txt)])
    largeur_libelle = max([0, *(len(s[2]) for s in seances_txt)])
    lignes_txt = []
    total_h = Decimal(0)
    total_eur = Decimal(0)
    for d, promo, libelle, heures, taux in seances_txt:
        montant = Decimal(str(heures)) * Decimal(str(taux))
        total_h += Decimal(str(heures))
        total_eur += montant
        lignes_txt.append(
            f"  | {d.strftime('%d/%m/%Y')} | {promo:<{largeur_promo}} | {libelle:<{largeur_libelle}} "
            f"| {heures:>5}h | {_au_centime(Decimal(str(taux))):>6} € | {_au_centime(montant):>9} €"
        )
    bc = spec.get("ref_bon_commande") or "—"
    contrat = spec.get("ref_contrat") or "—"
    if conflits:
        nature = "REMPLACEMENT" if any(r.startswith("l'onglet") for r in conflits) else "DOUBLON DE NUMÉRO"
        etat_onglet = [f"Onglet        : {nature} — écriture refusée",
                       *(f"                - {raison}" for raison in conflits)]
    else:
        etat_onglet = ["Onglet        : nouvel onglet"]
    return "\n".join([
        "====== APERÇU FACTURE ======",
        f"Entité        : {spec['entite']}",
        f"N°            : {spec['n_facture']}",
        f"Date émission : {_parse_date(spec['date_emission']).strftime('%d/%m/%Y')}",
        f"Échéance      : {echeance.strftime('%d/%m/%Y')}",
        f"Émetteur      : {emetteur['nom']} — {emetteur.get('adresse_l1', '')}, "
        f"{emetteur.get('adresse_l2', '')}",
        f"NDA           : {nda or '—'}",
        f"Destinataire  : {dest['nom']} — {dest.get('adresse_l1', '')}, "
        f"{dest.get('adresse_l2', '')}",
        f"Bon de commande : {bc}",
        f"Contrat       : {contrat}",
        f"Fichier       : {spec['fichier_sortie']} > onglet '{spec['onglet']}'",
        *etat_onglet,
        "",
        "DÉTAIL DES SÉANCES :",
        *lignes_txt,
        "",
        # Somme exacte, affichée comme avant (écriture d'un float : « 7.5h », « 18.0h »).
        f"TOTAL HEURES : {float(total_h)}h",
        f"TOTAL HT     : {_au_centime(total_eur)} €",
        "============================",
    ])


def _refuser_autre_entite(spec: dict, entite: str) -> None:
    """Les entités ne se mélangent jamais (ADR001) : `entite` doit être celle de la spec.

    Raises:
        ValueError: `entite` diffère de `spec["entite"]` ; le message nomme les deux.
    """
    if spec["entite"] != entite:
        raise ValueError(
            f"entité {entite!r} demandée pour la facture {spec.get('n_facture')!r}, qui est celle "
            f"de l'entité {spec['entite']!r} : les entités ne se mélangent jamais (ADR001)."
        )


def ecrire_onglet(wb: Workbook, spec: dict, dest: dict, *, entite: str) -> Worksheet:
    """Écrit l'onglet de la facture à la fin de `wb`, sans rien enregistrer (PCF étape 5).

    Séquence d'écriture de `generer` : `clone_template` (onglet `spec["onglet"]`), émetteur,
    destinataire, métadonnées (numéro, date d'émission, bon de commande, contrat), séances
    (libellé `ecole_label`, sinon le code de l'école ; colonne Jour masquée ou non), mentions
    avec l'échéance de la spec, RIB, mise en page.

    Avant la première écriture, elle se garde elle-même : `entite` doit être celle de la spec
    (ADR001) ; la spec doit être bien formée, comme à sa lecture par `charger_spec` (titre
    d'onglet, `masquer_jour`, date d'émission, lignes et séances) ; la spec est lue en entier
    (dates, échéance, séances, numéro, école) ; l'onglet ne doit pas exister, même à la casse
    près, ni le numéro figurer dans un onglet du classeur (`conflits_ecriture`, FA-02). Si une
    erreur survient ensuite, en cours d'écriture (MASTER incomplet, par exemple), l'onglet
    commencé est retiré avant que l'erreur ne remonte : le classeur garde ses onglets tels qu'ils
    étaient, et un appelant qui attrape l'erreur ne peut pas enregistrer une facture à moitié
    écrite. Le destinataire complet et les mentions de l'ADR030 restent contrôlés par l'appelant,
    comme dans `generer`. Ni la spec ni le destinataire ne sont modifiés.

    Args:
        wb: classeur ouvert qui reçoit l'onglet ; les autres onglets ne sont pas touchés.
        spec: spec de la facture.
        dest: destinataire lu dans MASTER.Coordonnées (`fmt.get_destinataire`).
        entite: code de l'entité, qui fixe le gabarit et l'émetteur ; celui de la spec.

    Returns:
        L'onglet écrit.

    Raises:
        ValueError: `entite` différente de `spec["entite"]`, spec mal formée (titre d'onglet
            qu'Excel refuse, `masquer_jour` non booléen, date d'émission ou de séance mal formée,
            `lignes` ou `seances` qui ne sont pas des listes, ligne ou séance qui n'est pas un
            objet, clé de ligne absente, ligne sans séance, heures ou taux invalides), échéance
            mal formée ou antérieure à la date d'émission (`echeance_de_la_spec`), ou `Info` /
            `NDA` absent du MASTER.
        KeyError: clé obligatoire absente de la spec.
        EcritureRefusee: onglet déjà présent, ou numéro déjà présent dans un onglet ; le message
            nomme l'onglet en conflit.
    """
    _refuser_autre_entite(spec, entite)
    _refuser_spec_mal_formee(spec)
    date_emission = _parse_date(spec["date_emission"])
    echeance = echeance_de_la_spec(spec)
    groupes = construire_groupes(spec["lignes"])
    onglet = spec["onglet"]
    n_facture = spec["n_facture"]
    ecole_name = spec.get("ecole_label", spec["ecole"])
    masquer_jour = spec.get("masquer_jour", False)  # booléen : contrôlé ci-dessus
    ref_bon_commande = spec.get("ref_bon_commande")
    ref_contrat = spec.get("ref_contrat") or None

    conflits = conflits_ecriture(wb, onglet, n_facture, entite=entite)
    if conflits:
        raise EcritureRefusee(
            f"l'onglet '{onglet}' n'a pas été écrit : " + " ; ".join(conflits) + ". "
            "Une facture émise ne se modifie pas : elle s'annule par un avoir."
        )

    onglets_avant = {id(feuille) for feuille in [*wb.worksheets, *wb.chartsheets]}
    try:
        ws = fmt.clone_template(wb, onglet, entite=entite)
        fmt.apply_emetteur(ws, entite=entite)
        fmt.apply_destinataire(ws, dest=dest, entite=entite)
        fmt.apply_metadata(
            ws,
            n_facture=n_facture,
            date_emission=date_emission,
            entite=entite,
            ref_bon_commande=ref_bon_commande,
            ref_contrat=ref_contrat,
        )
        fmt.place_sessions(
            ws,
            ecole_name=ecole_name,
            groups=groupes,
            entite=entite,
            masquer_jour=masquer_jour,
        )
        fmt.apply_mentions(ws, entite=entite, date_emission=date_emission, echeance=echeance)
        fmt.apply_rib(ws, entite=entite)
        fmt.apply_print_setup(ws, entite=entite)
    except BaseException:
        # Retirer l'onglet commencé, y compris s'il a été créé avant l'erreur de clone_template.
        for feuille in [*wb.worksheets, *wb.chartsheets]:
            if id(feuille) not in onglets_avant:
                wb.remove(feuille)
        raise
    return ws


def verifier_onglet(ws: Worksheet, spec: dict, dest: dict, *, entite: str) -> list[str]:
    """Contrôle après écriture (PCF étape 7.3, FA-12) : écarts entre l'onglet et la spec.

    Les valeurs attendues se déduisent de la spec comme à l'écriture : titre de l'onglet
    (`spec["onglet"]`, à l'identique : openpyxl suffixe en silence un titre déjà pris), numéro,
    date d'émission, libellé de l'école (`ecole_label`, sinon le code), séances
    (`construire_groupes`), échéance (`echeance_de_la_spec` : sans échéance, émission + 30 jours)
    et contrat (sans `ref_contrat`, la ligne de contrat doit être vide). Le destinataire attendu
    est celui reçu ; l'émetteur, le SIRET, le RIB et le NDA sont lus dans le MASTER de `entite`
    par `fmt.verifier_facture`, d'où le refus d'une autre entité que celle de la spec. Ni la spec
    ni le destinataire ne sont modifiés.

    Args:
        ws: onglet à contrôler, de préférence relu depuis le fichier enregistré.
        spec: spec de la facture.
        dest: destinataire attendu, celui passé à `ecrire_onglet`.
        entite: code de l'entité ; celui de la spec.

    Returns:
        Une ligne « ERREUR sur cellule … : attendu …, trouvé … » par écart (pour le titre, clé
        `onglet`) ; liste vide si l'onglet est conforme.

    Raises:
        ValueError: `entite` différente de `spec["entite"]`, date mal formée, échéance antérieure
            à la date d'émission, ou `Info` / `NDA` absent du MASTER.
    """
    _refuser_autre_entite(spec, entite)
    ecarts = []
    if ws.title != spec["onglet"]:
        ecarts.append(f"ERREUR sur cellule onglet : attendu {spec['onglet']!r}, trouvé {ws.title!r}")
    return ecarts + fmt.verifier_facture(
        ws,
        entite=entite,
        n_facture=spec["n_facture"],
        date_emission=_parse_date(spec["date_emission"]),
        dest=dest,
        ecole_name=spec.get("ecole_label", spec["ecole"]),
        groups=construire_groupes(spec["lignes"]),
        echeance=echeance_de_la_spec(spec),
        ref_contrat=spec.get("ref_contrat") or None,
    )


def generer(spec: dict, dry_run: bool = False, *, prefixe: str | None = None) -> dict:
    """Applique le PCF étapes 5 et 7. Retourne le post-mortem (dict vide si dry-run).

    Args:
        spec: spec de la facture.
        dry_run: aperçu seul, rien n'est écrit.
        prefixe: préfixe des numéros de l'entité ; par défaut, la règle du profil
            (initiale de l'entité + « _ »).

    Raises:
        ValueError: spec mal formée (comme à sa lecture par `charger_spec`), refusée avant
            l'aperçu.
        EcritureRefusee: refus avant toute écriture, dont le classeur qui ne peut pas être
            enregistré (`PermissionError` : ouvert dans Excel ou LibreOffice) ; il reste intact.
        PostMortemEnEcart: facture écrite, mais non conforme à l'aperçu.
    """
    entite = spec["entite"]
    _refuser_spec_mal_formee(spec)
    sortie = controler_spec(spec, prefixe=prefixe, racine=fmt.get_data_root())
    dest = fmt.get_destinataire(spec["ecole"], entite=entite)
    if dest is None:
        raise RuntimeError(
            f"Destinataire '{spec['ecole']}' absent de MASTER.Coordonnées — "
            "gate étape 4 du PCF : compléter avant de générer "
            "(cf. references/ONBOARDING.md)."
        )
    manquants = fmt.champs_destinataire_manquants(dest)
    if manquants:
        raise EcritureRefusee(
            f"destinataire '{spec['ecole']}' incomplet dans MASTER.Coordonnées "
            f"(manque : {', '.join(manquants)}) — gate étape 4 du PCF : compléter avant "
            "de générer (cf. references/ONBOARDING.md)."
        )

    # Mentions obligatoires (ADR030) : contrôlées avant l'aperçu, donc aussi en --dry-run.
    emetteur = fmt.get_emetteur(entite=entite)
    if not fmt.emetteur_porte_ei(emetteur["nom"]):
        raise EcritureRefusee(
            f"le nom de l'émetteur « {emetteur['nom']} » ne porte pas la mention « EI » ni "
            "« Entrepreneur individuel » (ADR030) — compléter la ligne Émetteur de "
            "MASTER.Coordonnées avant de générer."
        )
    try:
        nda = fmt.get_nda(entite=entite)
    except ValueError as exc:
        raise EcritureRefusee(f"{exc} Mention obligatoire (ADR030).") from exc
    # La date d'émission est déjà contrôlée par `_refuser_spec_mal_formee` : une erreur ici ne
    # peut venir que de l'échéance.
    try:
        echeance = echeance_de_la_spec(spec)
    except ValueError as exc:
        raise EcritureRefusee(
            f"echeance {spec.get('echeance')!r} invalide : {exc} (attendu AAAA-MM-JJ, "
            "au plus tôt la date d'émission)"
        ) from exc

    groupes = construire_groupes(spec["lignes"])
    onglet = spec["onglet"]
    # Le classeur cible est ouvert dès l'aperçu, pour annoncer l'état de l'onglet.
    wb = load_workbook(sortie) if sortie.exists() else None
    conflits = conflits_ecriture(wb, onglet, spec["n_facture"], entite=entite)
    print(apercu(spec, groupes, dest, conflits, emetteur=emetteur, echeance=echeance, nda=nda))
    if conflits:
        raise EcritureRefusee(
            f"{sortie.name} n'a pas été modifié : " + " ; ".join(conflits) + ". "
            "Une facture émise ne se modifie pas : elle s'annule par un avoir."
        )
    if dry_run:
        print("\n[dry-run] Aucun fichier écrit.")
        return {}

    sortie.parent.mkdir(parents=True, exist_ok=True)
    if wb is None:
        wb = Workbook()
        wb.remove(wb.active)

    ws = ecrire_onglet(wb, spec, dest, entite=entite)
    try:
        wb.save(sortie)
    except PermissionError as exc:
        # Excel ou LibreOffice verrouille le classeur ouvert : l'ouverture en écriture échoue avant
        # de rien y écrire, le fichier garde son contenu.
        raise EcritureRefusee(
            f"{sortie.name} n'a pas pu être enregistré (accès refusé) : le classeur est sans doute "
            f"ouvert dans Excel ou LibreOffice. Le fermer, puis relancer la commande ({sortie})."
        ) from exc
    print(f"\nÉcrit : {sortie}")

    # Post-mortem sur les cellules relues dans le fichier écrit, pas sur l'objet en mémoire.
    ws_relu = load_workbook(sortie)[ws.title]
    pm = fmt.post_mortem(ws_relu, entite=entite)
    print("\n── Post-mortem (PCF étape 7.2) ──")
    for cle, valeur in pm.items():
        print(f"  {cle:<18} : {valeur!r}")

    ecarts = verifier_onglet(ws_relu, spec, dest, entite=entite)
    if ecarts:
        raise PostMortemEnEcart(ecarts)
    print(f"\nOK — facture {spec['n_facture']} conforme à l'aperçu (post-mortem vérifié, PCF étape 7.3).")
    return pm
