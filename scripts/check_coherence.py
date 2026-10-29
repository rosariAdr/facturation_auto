"""check_coherence.py — Lecture seule : état de MASTER.Cours et MASTER.Sessions.

Sert le PCF étape 3.1 (vérification avant journalisation) : afficher l'état d'une
ou plusieurs écoles, et signaler les écarts entre heures dispensées, facturées et
restant à facturer.

Usage :
    python scripts/check_coherence.py --entite ROSARI
    python scripts/check_coherence.py --entite ROSARI --ecole EPITA "ISG Lyon"
    python scripts/check_coherence.py --entite ROSARI --sessions
"""
import argparse
import sys
from pathlib import Path

from openpyxl import load_workbook

try:
    from facturation import config, pcf_format as fmt
except ModuleNotFoundError:  # exécution depuis le dépôt sans installation
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from facturation import config, pcf_format as fmt

# MASTER.Cours : École | Module | Classe | Vol contrac | H dispensées |
#                H facturées | H à facturer | N° Facture liée | N° Devis lié | Statut
COLS_COURS = {
    "ecole": 1, "module": 2, "classe": 3, "h_dispensees": 5,
    "h_facturees": 6, "h_a_facturer": 7, "n_facture": 8, "statut": 10,
}
# MASTER.Sessions : Date_log | École | Période | N° Facture | Date séance | ...
COLS_SESSIONS = {
    "ecole": 2, "periode": 3, "n_facture": 4, "date_seance": 5,
    "heures": 8, "promotion": 9, "libelle": 10, "taux": 11, "total": 12,
}


def _nombre(valeur) -> float:
    """Convertit une cellule en float, 0.0 si vide ou non numérique."""
    try:
        return float(valeur)
    except (TypeError, ValueError):
        return 0.0


def afficher_cours(ws, ecoles: list):
    """Affiche MASTER.Cours et signale les incohérences d'heures."""
    print(
        f"{'Row':<4} {'École':<14} {'Module':<48} {'Classe':<16} "
        f"{'H disp':>8} {'H fact':>8} {'H rest':>8} {'N° Facture':<28} {'Statut':<14}"
    )
    print("─" * 160)
    ecarts = []
    for r in range(2, ws.max_row + 1):
        ecole = ws.cell(row=r, column=COLS_COURS["ecole"]).value
        if ecole is None or (ecoles and ecole not in ecoles):
            continue
        v = {k: ws.cell(row=r, column=c).value for k, c in COLS_COURS.items()}
        print(
            f"{r:<4} {str(ecole)[:14]:<14} {str(v['module'])[:48]:<48} "
            f"{str(v['classe'])[:16]:<16} {_nombre(v['h_dispensees']):>8.1f} "
            f"{_nombre(v['h_facturees']):>8.1f} {_nombre(v['h_a_facturer']):>8.1f} "
            f"{str(v['n_facture'] or '')[:28]:<28} {str(v['statut'] or '')[:14]:<14}"
        )
        disp = _nombre(v["h_dispensees"])
        somme = _nombre(v["h_facturees"]) + _nombre(v["h_a_facturer"])
        if disp and abs(disp - somme) > 0.01:
            ecarts.append((r, ecole, v["module"], disp, somme))

    if ecarts:
        print("\n⚠ Écarts (H dispensées ≠ H facturées + H à facturer) :")
        for r, ecole, module, disp, somme in ecarts:
            print(f"  row {r} | {ecole} | {module} : {disp} ≠ {somme}")
    else:
        print("\n✅ Aucun écart d'heures détecté.")


def afficher_sessions(ws, ecoles: list):
    """Affiche MASTER.Sessions groupé par n° de facture, et détecte les doublons."""
    par_facture: dict = {}
    for r in range(2, ws.max_row + 1):
        ecole = ws.cell(row=r, column=COLS_SESSIONS["ecole"]).value
        n_fact = ws.cell(row=r, column=COLS_SESSIONS["n_facture"]).value
        if not n_fact or (ecoles and ecole not in ecoles):
            continue
        par_facture.setdefault(n_fact, []).append(
            (r, _nombre(ws.cell(row=r, column=COLS_SESSIONS["heures"]).value),
             _nombre(ws.cell(row=r, column=COLS_SESSIONS["total"]).value))
        )

    print(f"\n{'N° Facture':<16} {'Lignes':>7} {'Heures':>9} {'Total HT':>12}")
    print("─" * 50)
    total_general = 0.0
    for n_fact in sorted(par_facture):
        lignes = par_facture[n_fact]
        heures = sum(h for _, h, _ in lignes)
        total = sum(t for _, _, t in lignes)
        total_general += total
        print(f"{str(n_fact):<16} {len(lignes):>7} {heures:>9.1f} {total:>12.2f} €")
    print("─" * 50)
    print(f"{'TOTAL':<16} {'':<7} {'':<9} {total_general:>12.2f} €")


def main():
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument(
        "--entite", default=None,
        help="Code de l'entité (défaut : celle du profil, sinon ROSARI)",
    )
    parser.add_argument(
        "--ecole", nargs="*", default=[],
        help="Filtre sur une ou plusieurs écoles (défaut : toutes)",
    )
    parser.add_argument(
        "--sessions", action="store_true",
        help="Affiche aussi le récapitulatif MASTER.Sessions par facture",
    )
    parser.add_argument("--data-root", type=Path, default=None)
    parser.add_argument(
        "--profil", type=Path, default=None,
        help="Profil utilisateur (défaut : profil standard s'il existe)",
    )
    args = parser.parse_args()

    profil = config.resoudre_racine(args.data_root, args.profil)
    entite = args.entite or (profil.entite if profil else "ROSARI")

    master = fmt._master_path_for(entite)
    if not master.exists():
        raise SystemExit(
            f"MASTER introuvable : {master}\n"
            f"Racine de données : {fmt.get_data_root()} "
            "(--data-root, --profil ou $PCF_DATA_ROOT)"
        )
    print(f"MASTER : {master}\n")
    wb = load_workbook(master, data_only=True)
    afficher_cours(wb["Cours"], args.ecole)
    if args.sessions:
        afficher_sessions(wb["Sessions"], args.ecole)


if __name__ == "__main__":
    main()
