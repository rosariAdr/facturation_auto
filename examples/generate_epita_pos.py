"""generate_epita_pos.py — Génère 7 factures EPITA (1 par UE) selon POs.

Sources :
- PO-26-ASSOTA-01791 (10/07/2026) : 1 UE — 1 620 €
- PO-26-ASSOTA-01377 (11/05/2026) : 6 UEs — 8 820 €
Total : 10 440 € HT

Convention :
- 1 facture par UE = 1 ligne cours = qté H × 90 €
- Date d'émission : 31/07/2026 (fin de mois englobant les deux POs)
- Date "séance" = date de livraison souhaitée du PO. Colonne Jour masquée :
  la facture agrège un volume horaire, pas des séances datées.
- Libellé = code produit du PO à l'identique (rapprochement Basware)
- Réf PO en ligne "Bon de commande", comme sur les factures EPITA001-004 payées
"""
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook

try:
    from facturation import pcf_format as fmt
except ModuleNotFoundError:  # exécution depuis le dépôt sans installation
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from facturation import pcf_format as fmt

# Racine des données : $PCF_DATA_ROOT, ou dossier courant. Ne jamais coder en dur
# un chemin de données dans un script versionné.
ROOT = fmt.get_data_root()
MASTER = ROOT / "ROSARI" / "ROSARI_MASTER.xlsx"
EPITA_FILE = ROOT / "ROSARI" / "EPITA" / "ROSARI_EPITA.xlsx"

DATE_EMISSION = date(2026, 7, 31)
PERIODE_LABEL = "Juillet 2026"
ECOLE_LABEL = "EPITA"
TAUX = 90.0

# (n_facture, sheet_name, promotion, libelle_ue, ref_po, heures, date_liv)
# libelle_ue = code produit recopié à l'identique du PO, sans reformatage.
INVOICES = [
    ("R_EPITA003", "R_EPITA003 Juil 2026",
     "DSA S2", "[2026][PI][DSA S2][MathsforData][18H][S8_25-26][26.07]",
     "PO-26-ASSOTA-01791", 18.0, date(2026, 7, 8)),
    ("R_EPITA004", "R_EPITA004 Juil 2026",
     "DSA S2", "[2026][PI][DSA S2][Maths Data Science][98H][S8_25-26][26.02-03]",
     "PO-26-ASSOTA-01377 (ligne 1)", 18.0, date(2026, 7, 31)),
    ("R_EPITA005", "R_EPITA005 Juil 2026",
     "AIS S1", "[2026][PI][AIS S1][Numerical Algorithms][98H][S8_25-26][26.03]",
     "PO-26-ASSOTA-01377 (ligne 2)", 18.0, date(2026, 7, 31)),
    ("R_EPITA006", "R_EPITA006 Juil 2026",
     "AIS S1+S2",
     "[2026][PI][AIS S1+S2][OperationsResearchIOptimization][98H][S8_25-26][26.04]",
     "PO-26-ASSOTA-01377 (ligne 3)", 14.0, date(2026, 7, 31)),
    ("R_EPITA007", "R_EPITA007 Juil 2026",
     "AIS S1",
     "[2026][PI][AIS S1][Operations Research I: Linear Programming][98H][S8_25-26][26.05]",
     "PO-26-ASSOTA-01377 (ligne 4)", 12.0, date(2026, 7, 31)),
    ("R_EPITA008", "R_EPITA008 Juil 2026",
     "AIS S2", "[2026][PI][AIS S2][Spark & Python for Big Data][98H][S8_25-26][26.05]",
     "PO-26-ASSOTA-01377 (ligne 5)", 15.0, date(2026, 7, 31)),
    ("R_EPITA009", "R_EPITA009 Juil 2026",
     "AIS S1", "[2026][PI][AIS S1][Algorithmic Workshop][98H][S8_25-26][26.06]",
     "PO-26-ASSOTA-01377 (ligne 6)", 21.0, date(2026, 7, 31)),
]


def log_sessions_to_master():
    """Écrit 1 ligne par facture dans MASTER.Sessions.

    Idempotent : si le n° de facture est déjà journalisé, la ligne est mise à
    jour sur place au lieu d'en créer une seconde.
    """
    print("\n══ Étape 3 — Journalisation MASTER.Sessions ══")
    wb = load_workbook(MASTER)
    ws_s = wb["Sessions"]
    date_log = datetime.now().strftime("%Y-%m-%d")

    # n_facture (col 4) → row, pour retrouver les lignes déjà journalisées
    existantes = {
        ws_s.cell(row=r, column=4).value: r
        for r in range(2, ws_s.max_row + 1)
        if ws_s.cell(row=r, column=4).value
    }

    n_new = n_upd = 0
    for n_fact, _, promo, libelle, ref_po, heures, d in INVOICES:
        total = heures * TAUX
        libelle_full = f"{libelle} — {ref_po}"
        row = existantes.get(n_fact)
        if row is None:
            row = ws_s.max_row + 1
            n_new += 1
        else:
            n_upd += 1
        ws_s.cell(row=row, column=1, value=date_log)
        ws_s.cell(row=row, column=2, value="EPITA")
        ws_s.cell(row=row, column=3, value=PERIODE_LABEL)
        ws_s.cell(row=row, column=4, value=n_fact)
        ws_s.cell(row=row, column=5, value=d).number_format = "DD/MM/YYYY"
        ws_s.cell(row=row, column=8, value=heures)
        ws_s.cell(row=row, column=9, value=promo)
        ws_s.cell(row=row, column=10, value=libelle_full)
        ws_s.cell(row=row, column=11, value=TAUX)
        ws_s.cell(row=row, column=12, value=total)
        ws_s.cell(row=row, column=13, value=f"{promo} – {libelle} – {ref_po}")
    wb.save(MASTER)
    print(f"  {n_new} sessions ajoutées, {n_upd} mises à jour")


def generate_all():
    print(f"\n══ Étape 7 — Génération {len(INVOICES)} factures EPITA ══")
    dest = fmt.get_destinataire("EPITA", entite="ROSARI")
    if dest is None:
        raise RuntimeError("Destinataire EPITA absent de MASTER.Coordonnées")
    print(f"  Destinataire : {dest['nom']} / {dest['adresse_l1']} / {dest['adresse_l2']}")

    wb = load_workbook(EPITA_FILE)
    total_all = 0.0

    for n_fact, sheet_name, promo, libelle, ref_po, heures, d in INVOICES:
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
            print(f"  Ancien onglet '{sheet_name}' supprimé")

        ws = fmt.clone_template(wb, sheet_name, entite="ROSARI")
        fmt.apply_emetteur(ws, entite="ROSARI")
        fmt.apply_destinataire(ws, dest=dest, entite="ROSARI")
        fmt.apply_metadata(
            ws,
            n_facture=n_fact,
            date_emission=DATE_EMISSION,
            entite="ROSARI",
            ref_bon_commande=ref_po,
        )

        groups = [(promo, libelle, [(d, heures, TAUX)])]
        fmt.place_sessions(
            ws,
            ecole_name=ECOLE_LABEL,
            groups=groups,
            entite="ROSARI",
            masquer_jour=True,
        )
        fmt.apply_rib(ws, entite="ROSARI")
        fmt.apply_print_setup(ws, entite="ROSARI")

        total = heures * TAUX
        total_all += total
        print(f"  ✅ {n_fact} — {promo} — {libelle} — {heures}h × {TAUX}€ = {total:.0f}€")

    wb.save(EPITA_FILE)
    print(f"\n  Sauvegardé : {EPITA_FILE.name}")
    print(f"  TOTAL : {total_all:.2f} € HT")
    return total_all


def main():
    print(f"PCF — Génération factures EPITA depuis POs")
    print(f"Date émission : {DATE_EMISSION.strftime('%d/%m/%Y')}")
    log_sessions_to_master()
    generate_all()


if __name__ == "__main__":
    main()
