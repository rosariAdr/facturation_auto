"""factures_par_bon_de_commande.py — Plusieurs factures depuis des bons de commande.

Exemple de script dédié pour une école qui émet des bons de commande (BC) :
une facture par ligne de BC. Toutes les données ci-dessous sont fictives
(école « ABC », références de BC, libellés, heures, taux) : adapter la liste
INVOICES avant usage, avec les vraies valeurs lues dans le MASTER et les BC.

Convention :
- 1 facture par ligne de BC = 1 ligne de cours = heures × taux
- Date d'émission : le 1er du mois suivant (ADR027)
- Date « séance » = date de livraison souhaitée du BC. Colonne Jour masquée :
  la facture porte un volume horaire, pas des séances datées.
- Libellé = code produit du BC recopié à l'identique (rapprochement comptable)
- Référence du BC sur la ligne « Bon de commande », jamais dans le libellé

Attention : generate_all() supprime puis recrée un onglet qui porte déjà le même
nom. Ne l'utiliser que sur des brouillons : une facture émise ne se modifie pas,
elle se corrige par un avoir (ADR024, ADR031).
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
ENTITE = "ROSARI"
CODE_ECOLE = "ABC"
MASTER = ROOT / ENTITE / f"{ENTITE}_MASTER.xlsx"
ECOLE_FILE = ROOT / ENTITE / CODE_ECOLE / f"{ENTITE}_{CODE_ECOLE}.xlsx"

DATE_EMISSION = date(2026, 8, 1)
PERIODE_LABEL = "Juillet 2026"
ECOLE_LABEL = CODE_ECOLE
TAUX = 50.0

# (n_facture, sheet_name, promotion, libelle_ue, ref_bc, heures, date_liv)
# libelle_ue = code produit recopié à l'identique du BC, sans reformatage.
INVOICES = [
    ("R_ABC001", "R_ABC001 Juil 2026",
     "M1 S1", "[2026][M1 S1][Module A][20H]",
     "BC-EXEMPLE-0001", 20.0, date(2026, 7, 10)),
    ("R_ABC002", "R_ABC002 Juil 2026",
     "M1 S2", "[2026][M1 S2][Module B][30H]",
     "BC-EXEMPLE-0002 (ligne 1)", 15.0, date(2026, 7, 31)),
    ("R_ABC003", "R_ABC003 Juil 2026",
     "M2 S1", "[2026][M2 S1][Module C][30H]",
     "BC-EXEMPLE-0002 (ligne 2)", 12.0, date(2026, 7, 31)),
]


def log_sessions_to_master():
    """Écrit 1 ligne par facture dans MASTER.Sessions.

    Idempotent : si le n° de facture est déjà journalisé, la ligne est mise à
    jour sur place au lieu d'en créer une seconde (ADR003 amendé).
    """
    print("\n══ Journal MASTER.Sessions ══")
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
    for n_fact, _, promo, libelle, ref_bc, heures, d in INVOICES:
        total = heures * TAUX
        libelle_full = f"{libelle} — {ref_bc}"
        row = existantes.get(n_fact)
        if row is None:
            row = ws_s.max_row + 1
            n_new += 1
        else:
            n_upd += 1
        ws_s.cell(row=row, column=1, value=date_log)
        ws_s.cell(row=row, column=2, value=CODE_ECOLE)
        ws_s.cell(row=row, column=3, value=PERIODE_LABEL)
        ws_s.cell(row=row, column=4, value=n_fact)
        ws_s.cell(row=row, column=5, value=d).number_format = "DD/MM/YYYY"
        ws_s.cell(row=row, column=8, value=heures)
        ws_s.cell(row=row, column=9, value=promo)
        ws_s.cell(row=row, column=10, value=libelle_full)
        ws_s.cell(row=row, column=11, value=TAUX)
        ws_s.cell(row=row, column=12, value=total)
        ws_s.cell(row=row, column=13, value=f"{promo} – {libelle} – {ref_bc}")
    wb.save(MASTER)
    print(f"  {n_new} lignes ajoutées, {n_upd} mises à jour")


def generate_all():
    print(f"\n══ Génération de {len(INVOICES)} factures ══")
    dest = fmt.get_destinataire(CODE_ECOLE, entite=ENTITE)
    if dest is None:
        raise RuntimeError(f"Destinataire {CODE_ECOLE} absent de MASTER.Coordonnées")
    print(f"  Destinataire : {dest['nom']} / {dest['adresse_l1']} / {dest['adresse_l2']}")

    wb = load_workbook(ECOLE_FILE)
    total_all = 0.0

    for n_fact, sheet_name, promo, libelle, ref_bc, heures, d in INVOICES:
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
            print(f"  Ancien onglet '{sheet_name}' supprimé")

        ws = fmt.clone_template(wb, sheet_name, entite=ENTITE)
        fmt.apply_emetteur(ws, entite=ENTITE)
        fmt.apply_destinataire(ws, dest=dest, entite=ENTITE)
        fmt.apply_metadata(
            ws,
            n_facture=n_fact,
            date_emission=DATE_EMISSION,
            entite=ENTITE,
            ref_bon_commande=ref_bc,
        )

        groups = [(promo, libelle, [(d, heures, TAUX)])]
        fmt.place_sessions(
            ws,
            ecole_name=ECOLE_LABEL,
            groups=groups,
            entite=ENTITE,
            masquer_jour=True,
        )
        fmt.apply_rib(ws, entite=ENTITE)
        fmt.apply_print_setup(ws, entite=ENTITE)

        total = heures * TAUX
        total_all += total
        print(f"  ✅ {n_fact} — {promo} — {libelle} — {heures}h × {TAUX}€ = {total:.0f}€")

    wb.save(ECOLE_FILE)
    print(f"\n  Sauvegardé : {ECOLE_FILE.name}")
    print(f"  TOTAL : {total_all:.2f} € HT")
    return total_all


def main():
    print("PCF — Factures depuis des bons de commande (exemple fictif)")
    print(f"Date d'émission : {DATE_EMISSION.strftime('%d/%m/%Y')}")
    log_sessions_to_master()
    generate_all()


if __name__ == "__main__":
    main()
