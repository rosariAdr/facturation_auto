"""nouvelle_facture.py — Génère une facture depuis une spec JSON.

Remplace les scripts one-shot : la spec décrit le quoi, ce script applique le PCF
étape 5 (format) et étape 7 (écriture + post-mortem). Les étapes humaines du
protocole — périmètre, coordonnées, aperçu — restent à la charge de l'appelant
(cf. references/PROC00_pcf.md).

Ce script n'est que l'enveloppe en ligne de commande de `facturation/facture.py`,
qui porte l'aperçu, les refus, l'écriture et le contrôle, et décrit le format de
la spec (exemple : examples/spec_exemple.json).

Usage :
    python scripts/nouvelle_facture.py spec.json
    python scripts/nouvelle_facture.py spec.json --data-root "/chemin/Factures"
    python scripts/nouvelle_facture.py spec.json --dry-run    # aperçu seul
"""
import argparse
import sys
from pathlib import Path

# Le paquet de ce dépôt passe toujours en premier, même si l'interpréteur porte
# l'installation éditable d'une autre copie (FA-13).
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
# Le script réexporte les noms publics du paquet, `fmt` et `config` compris : un
# patch de `fmt.apply_rib` vise le module `pcf_format` lui-même.
from facturation.facture import *  # noqa: E402, F403


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
