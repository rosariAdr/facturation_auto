"""nouvelle_facture.py — Génère une facture depuis une spec JSON.

Remplace les scripts one-shot : la spec décrit le quoi, ce script applique le PCF
étape 5 (format) et étape 7 (écriture + post-mortem). Les étapes humaines du
protocole — périmètre, coordonnées, aperçu — restent à la charge de l'appelant
(cf. references/PROC00_pcf.md).

Usage :
    python scripts/nouvelle_facture.py spec.json
    python scripts/nouvelle_facture.py spec.json --data-root "/chemin/Factures"
    python scripts/nouvelle_facture.py spec.json --dry-run    # aperçu seul

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
  "masquer_jour": true,
  "lignes": [
    {"promotion": "M1 S1", "libelle": "[2026][M1 S1][Module A][...]",
     "seances": [{"date": "2026-09-15", "heures": 18, "taux": 90}]}
  ]
}
"""
import argparse
import json
import re
import sys
from datetime import date, datetime
from pathlib import Path, PurePosixPath, PureWindowsPath
from typing import Optional

from openpyxl import Workbook, load_workbook
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.worksheet.worksheet import Worksheet

# Le paquet de ce dépôt passe toujours en premier, même si l'interpréteur porte
# l'installation éditable d'une autre copie (FA-13).
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from facturation import config, pcf_format as fmt  # noqa: E402


class EcritureRefusee(RuntimeError):
    """La facture ne peut pas être écrite : rien n'a été modifié sur le disque."""


def _parse_date(valeur) -> date:
    """Accepte 'YYYY-MM-DD' ou un objet date déjà construit."""
    if isinstance(valeur, date):
        return valeur
    return datetime.strptime(str(valeur), "%Y-%m-%d").date()


def charger_spec(chemin: Path) -> dict:
    """Lit la spec JSON et vérifie que les clés obligatoires sont présentes."""
    spec = json.loads(chemin.read_text(encoding="utf-8"))
    obligatoires = {
        "entite", "ecole", "fichier_sortie", "onglet",
        "n_facture", "date_emission", "lignes",
    }
    manquantes = obligatoires - set(spec)
    if manquantes:
        raise ValueError(f"Clés manquantes dans {chemin.name} : {sorted(manquantes)}")
    if not spec["lignes"]:
        raise ValueError("La spec ne contient aucune ligne de cours.")
    return spec


def controler_spec(spec: dict, *, prefixe: Optional[str], racine: Path) -> Path:
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
    wb: Optional[Workbook], onglet: str, n_facture: str, *, entite: str
) -> list:
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


def apercu(
    spec: dict, groupes: list, dest: dict, conflits: Optional[list] = None, *, emetteur: dict
) -> str:
    """Construit l'aperçu du PCF étape 6 (gate humain avant écriture).

    Le libellé s'affiche en entier (il recopie à l'identique le code produit du bon
    de commande) ; les colonnes s'élargissent au plus long libellé et à la plus
    longue promotion (FA-05). Les raisons de `conflits` (FA-02) s'affichent dans
    l'état de l'onglet : nouvel onglet, ou écriture refusée.
    """
    seances_txt = [
        (d, str(promo), libelle.replace("\n", " "), heures, taux)
        for promo, libelle, seances in groupes
        for d, heures, taux in seances
    ]
    largeur_promo = max([12, *(len(s[1]) for s in seances_txt)])
    largeur_libelle = max([0, *(len(s[2]) for s in seances_txt)])
    lignes_txt = []
    total_h = total_eur = 0.0
    for d, promo, libelle, heures, taux in seances_txt:
        total_h += heures
        total_eur += heures * taux
        lignes_txt.append(
            f"  | {d.strftime('%d/%m/%Y')} | {promo:<{largeur_promo}} | {libelle:<{largeur_libelle}} "
            f"| {heures:>5}h | {taux:>6.2f} € | {heures * taux:>9.2f} €"
        )
    bc = spec.get("ref_bon_commande") or "—"
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
        f"Émetteur      : {emetteur['nom']} — {emetteur.get('adresse_l1', '')}, "
        f"{emetteur.get('adresse_l2', '')}",
        f"Destinataire  : {dest['nom']} — {dest.get('adresse_l1', '')}, "
        f"{dest.get('adresse_l2', '')}",
        f"Bon de commande : {bc}",
        f"Fichier       : {spec['fichier_sortie']} > onglet '{spec['onglet']}'",
        *etat_onglet,
        "",
        "DÉTAIL DES SÉANCES :",
        *lignes_txt,
        "",
        f"TOTAL HEURES : {total_h}h",
        f"TOTAL HT     : {total_eur:.2f} €",
        "============================",
    ])


class PostMortemEnEcart(RuntimeError):
    """La facture est écrite, mais des cellules relues diffèrent des valeurs attendues."""

    def __init__(self, ecarts: list):
        self.ecarts = ecarts
        super().__init__("\n".join(ecarts))


def generer(spec: dict, dry_run: bool = False, *, prefixe: Optional[str] = None) -> dict:
    """Applique le PCF étapes 5 et 7. Retourne le post-mortem (dict vide si dry-run).

    Args:
        spec: spec de la facture.
        dry_run: aperçu seul, rien n'est écrit.
        prefixe: préfixe des numéros de l'entité ; par défaut, la règle du profil
            (initiale de l'entité + « _ »).
    """
    entite = spec["entite"]
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

    groupes = construire_groupes(spec["lignes"])
    onglet = spec["onglet"]
    # Le classeur cible est ouvert dès l'aperçu, pour annoncer l'état de l'onglet.
    wb = load_workbook(sortie) if sortie.exists() else None
    conflits = conflits_ecriture(wb, onglet, spec["n_facture"], entite=entite)
    print(apercu(spec, groupes, dest, conflits, emetteur=fmt.get_emetteur(entite=entite)))
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

    ws = fmt.clone_template(wb, onglet, entite=entite)
    fmt.apply_emetteur(ws, entite=entite)
    fmt.apply_destinataire(ws, dest=dest, entite=entite)
    fmt.apply_metadata(
        ws,
        n_facture=spec["n_facture"],
        date_emission=_parse_date(spec["date_emission"]),
        entite=entite,
        ref_bon_commande=spec.get("ref_bon_commande"),
    )
    fmt.place_sessions(
        ws,
        ecole_name=spec.get("ecole_label", spec["ecole"]),
        groups=groupes,
        entite=entite,
        masquer_jour=bool(spec.get("masquer_jour", False)),
    )
    fmt.apply_rib(ws, entite=entite)
    fmt.apply_print_setup(ws, entite=entite)
    wb.save(sortie)
    print(f"\nÉcrit : {sortie}")

    # Post-mortem sur les cellules relues dans le fichier écrit, pas sur l'objet en mémoire.
    ws_relu = load_workbook(sortie)[ws.title]
    pm = fmt.post_mortem(ws_relu, entite=entite)
    print("\n── Post-mortem (PCF étape 7.2) ──")
    for cle, valeur in pm.items():
        print(f"  {cle:<18} : {valeur!r}")

    ecarts = fmt.verifier_facture(
        ws_relu,
        entite=entite,
        n_facture=spec["n_facture"],
        date_emission=_parse_date(spec["date_emission"]),
        dest=dest,
        ecole_name=spec.get("ecole_label", spec["ecole"]),
        groups=groupes,
    )
    if ecarts:
        raise PostMortemEnEcart(ecarts)
    print(f"\nOK — facture {spec['n_facture']} conforme à l'aperçu (post-mortem vérifié, PCF étape 7.3).")
    return pm


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("spec", type=Path, help="Fichier JSON de spec")
    parser.add_argument(
        "--data-root",
        type=Path,
        default=None,
        help="Racine des données métier (défaut : profil, $PCF_DATA_ROOT ou dossier courant)",
    )
    parser.add_argument(
        "--profil",
        type=Path,
        default=None,
        help="Profil utilisateur (défaut : profil standard s'il existe)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Affiche l'aperçu (PCF étape 6) sans rien écrire",
    )
    args = parser.parse_args()

    profil = config.resoudre_racine(args.data_root, args.profil)
    print(f"Racine de données : {fmt.get_data_root()}")

    spec = charger_spec(args.spec)
    # Le préfixe du profil ne vaut que pour son entité ; sinon, règle par défaut.
    prefixe = profil.prefixe_facture if profil and profil.entite == spec["entite"] else None
    try:
        generer(spec, dry_run=args.dry_run, prefixe=prefixe)
    except EcritureRefusee as exc:
        print(f"\nÉcriture refusée — {exc}", file=sys.stderr)
        sys.exit(1)
    except PostMortemEnEcart as exc:
        print("\n" + "\n".join(exc.ecarts), file=sys.stderr)
        print(
            "\nFacture écrite mais NON CONFORME (PCF étape 7.3) : ne pas l'envoyer ni la "
            "marquer finalisée ; corriger la cause, puis la régénérer.",
            file=sys.stderr,
        )
        sys.exit(2)


if __name__ == "__main__":
    main()
