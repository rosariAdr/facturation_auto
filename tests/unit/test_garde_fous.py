"""Contrôles purs de `scripts/nouvelle_facture.py`, sur des classeurs fictifs en mémoire.

Revue system design du 18/11/2026 : FA-02, FA-03, FA-05, FA-09, FA-12.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest
from openpyxl import Workbook

pytestmark = pytest.mark.unit

DEPOT = Path(__file__).resolve().parents[2]


def _charger_script():
    """Importe `scripts/nouvelle_facture.py`, qui n'est pas un paquet, par son chemin."""
    spec = importlib.util.spec_from_file_location("nouvelle_facture", DEPOT / "scripts" / "nouvelle_facture.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


nf = _charger_script()


def _classeur(*onglets: tuple[str, str | None]) -> Workbook:
    """Classeur en mémoire : un onglet par (titre, numéro écrit en D17)."""
    wb = Workbook()
    wb.remove(wb.active)
    for titre, numero in onglets:
        wb.create_sheet(titre)["D17"] = numero
    return wb


# ──────────────────────────────────────────────────────────────────────
# FA-02 — conflits avec le classeur cible
# ──────────────────────────────────────────────────────────────────────


def test_classeur_absent_aucun_conflit():
    assert nf.conflits_ecriture(None, "A_ECOLEA001 Sept 2026", "A_ECOLEA001", entite="ALPHA") == []


def test_onglet_identique_ou_a_la_casse_pres():
    wb = _classeur(("A_ECOLEA001 Sept 2026", "A_ECOLEA001"))
    assert nf.conflits_ecriture(wb, "A_ECOLEA001 Sept 2026", "A_ECOLEA009", entite="ALPHA")
    assert nf.conflits_ecriture(wb, "a_ecolea001 SEPT 2026", "A_ECOLEA009", entite="ALPHA")


def test_numero_dans_le_titre_ou_dans_la_cellule():
    par_titre = _classeur(("PAID A_ECOLEA001 Sept 2026", None))
    par_cellule = _classeur(("Facture réglée", "A_ECOLEA001"))
    for wb in (par_titre, par_cellule):
        conflits = nf.conflits_ecriture(wb, "A_ECOLEA001 Oct 2026", "A_ECOLEA001", entite="ALPHA")
        assert len(conflits) == 1 and "A_ECOLEA001" in conflits[0]


def test_controle_ne_cree_aucune_cellule_dans_les_autres_onglets():
    """Le contrôle lit les onglets sans les agrandir : le save ne doit pas les réécrire autrement."""
    wb = _classeur()
    notes = wb.create_sheet("Notes")
    notes["A1"] = "note fictive"
    assert nf.conflits_ecriture(wb, "A_ECOLEA001 Sept 2026", "A_ECOLEA001", entite="ALPHA") == []
    assert (notes.max_row, notes.max_column) == (1, 1)


def test_numero_voisin_n_est_pas_un_conflit():
    """A_ECOLEA001 ne doit pas être confondu avec A_ECOLEA0010 ni avec A_ECOLEA00."""
    wb = _classeur(("A_ECOLEA0010 Sept 2026", "A_ECOLEA0010"), ("A_ECOLEA00 Août 2026", "A_ECOLEA00"))
    assert nf.conflits_ecriture(wb, "A_ECOLEA001 Oct 2026", "A_ECOLEA001", entite="ALPHA") == []


# ──────────────────────────────────────────────────────────────────────
# FA-03 — entité explicite (ADR001), préfixe et dossier de sortie
# ──────────────────────────────────────────────────────────────────────

import inspect  # noqa: E402

from facturation import pcf_format as fmt  # noqa: E402

API_PUBLIQUE = [
    fmt.get_destinataire, fmt.clone_template, fmt.apply_emetteur, fmt.apply_destinataire,
    fmt.apply_metadata, fmt.place_sessions, fmt.apply_print_setup, fmt.apply_rib, fmt.post_mortem,
]


@pytest.mark.parametrize("fonction", API_PUBLIQUE, ids=lambda f: f.__name__)
def test_entite_obligatoire_et_nommee(fonction):
    """ADR001 : aucune entité par défaut ; `entite=` se passe toujours par son nom."""
    parametre = inspect.signature(fonction).parameters["entite"]
    assert parametre.kind is inspect.Parameter.KEYWORD_ONLY
    assert parametre.default is inspect.Parameter.empty


def test_oublier_l_entite_leve_une_erreur():
    with pytest.raises(TypeError, match="entite"):
        fmt.clone_template(Workbook(), "SANS_ENTITE")


RACINE = Path("donnees")


def test_spec_coherente_acceptee():
    spec = {"entite": "ALPHA", "n_facture": "A_ECOLEA001", "fichier_sortie": "ALPHA/ECOLEA/ALPHA_ECOLEA.xlsx"}
    assert nf.controler_spec(spec, prefixe="A_", racine=RACINE) == RACINE / "ALPHA" / "ECOLEA" / "ALPHA_ECOLEA.xlsx"


@pytest.mark.parametrize("entite", ["alpha", "", "STANDARD"])
def test_entite_invalide_refusee_proprement(entite):
    """Un code d'entité invalide donne un refus nommé, pas une trace d'exception."""
    spec = {"entite": entite, "n_facture": "A_ECOLEA001", "fichier_sortie": "ALPHA/ECOLEA/ALPHA_ECOLEA.xlsx"}
    with pytest.raises(nf.EcritureRefusee, match="entite"):
        nf.controler_spec(spec, prefixe=None, racine=RACINE)


def test_prefixe_par_defaut_quand_aucun_n_est_donne():
    spec = {"entite": "ALPHA", "n_facture": "B_ECOLEA001", "fichier_sortie": "ALPHA/ECOLEA/ALPHA_ECOLEA.xlsx"}
    with pytest.raises(nf.EcritureRefusee, match="préfixe A_"):
        nf.controler_spec(spec, prefixe=None, racine=RACINE)


@pytest.mark.parametrize("n_facture", ["B_ECOLEA001", "a_ECOLEA001", "ECOLEA001", "XA_ECOLEA001"])
def test_numero_hors_prefixe_refuse(n_facture):
    spec = {"entite": "ALPHA", "n_facture": n_facture, "fichier_sortie": "ALPHA/ECOLEA/ALPHA_ECOLEA.xlsx"}
    with pytest.raises(nf.EcritureRefusee, match="préfixe A_"):
        nf.controler_spec(spec, prefixe="A_", racine=RACINE)


@pytest.mark.parametrize("fichier_sortie", [
    "/tmp/ALPHA/x.xlsx", "C:/ALPHA/x.xlsx", "C:ALPHA/x.xlsx", r"\\serveur\partage\ALPHA\x.xlsx",
    "ALPHA/../BETA/x.xlsx", r"ALPHA\..\BETA\x.xlsx", "ALPHA/ECOLEA/../../x.xlsx",
    "BETA/ECOLEA/x.xlsx", "x.xlsx", "ALPHA", "alpha/ECOLEA/x.xlsx", "",
])
def test_fichier_de_sortie_hors_du_dossier_de_l_entite_refuse(fichier_sortie):
    spec = {"entite": "ALPHA", "n_facture": "A_ECOLEA001", "fichier_sortie": fichier_sortie}
    with pytest.raises(nf.EcritureRefusee, match="fichier_sortie"):
        nf.controler_spec(spec, prefixe="A_", racine=RACINE)


# ──────────────────────────────────────────────────────────────────────
# FA-05 — aperçu complet
# ──────────────────────────────────────────────────────────────────────

from datetime import date  # noqa: E402

EMETTEUR = {"nom": "EMETTEUR FICTIF ALPHA", "adresse_l1": "1 rue Fictive", "adresse_l2": "00000 VILLE-FICTIVE"}
DEST = {"nom": "SERVICE COMPTABLE ECOLE A", "adresse_l1": "1 avenue Fictive", "adresse_l2": "00000 VILLE-A"}
SPEC = {"entite": "ALPHA", "n_facture": "A_ECOLEA001", "date_emission": "2026-10-01",
        "fichier_sortie": "ALPHA/ECOLEA/ALPHA_ECOLEA.xlsx", "onglet": "A_ECOLEA001 Sept 2026"}


def test_apercu_libelles_complets_et_colonnes_alignees():
    court = "Module court"
    long = "[2026][CODE-PRODUIT-FICTIF-0002][PROMO 2][Libellé recopié à l'identique du bon de commande]"
    groupes = [("PROMO 1", court, [(date(2026, 9, 15), 3.0, 50.0)]),
               ("PROMO 2", long, [(date(2026, 9, 16), 2.0, 50.0), (date(2026, 9, 17), 2.0, 50.0)])]
    texte = nf.apercu(SPEC, groupes, DEST, emetteur=EMETTEUR)

    lignes = [ligne for ligne in texte.splitlines() if ligne.startswith("  | ")]
    assert len(lignes) == 3
    assert court in lignes[0] and all(long in ligne for ligne in lignes[1:])
    assert len({ligne.index("h |") for ligne in lignes}) == 1, "colonnes décalées"
    assert "Émetteur      : EMETTEUR FICTIF ALPHA — 1 rue Fictive, 00000 VILLE-FICTIVE" in texte


def test_apercu_promotion_numerique_acceptee():
    """Une promotion saisie comme nombre dans la spec JSON s'affiche, comme avant le correctif."""
    groupes = [(2026, "Module fictif", [(date(2026, 9, 15), 3.0, 50.0)])]
    assert "| 2026 " in nf.apercu(SPEC, groupes, DEST, emetteur=EMETTEUR)


def test_apercu_libelle_sur_plusieurs_lignes_reste_sur_une_ligne():
    groupes = [("PROMO 1", "Partie 1\nPartie 2", [(date(2026, 9, 15), 3.0, 50.0)])]
    texte = nf.apercu(SPEC, groupes, DEST, emetteur=EMETTEUR)
    assert "Partie 1 Partie 2" in texte


# ──────────────────────────────────────────────────────────────────────
# FA-10 — l'adresse iCal secrète n'est jamais recopiée
# ──────────────────────────────────────────────────────────────────────

from facturation import config, verification  # noqa: E402

JETON = "private-0123456789abcdef"
URL_FICTIVE = f"https://calendar.example.org/calendar/ical/personne%40example.org/{JETON}/basic.ics"


def _brut(fichier_url_secrete) -> dict:
    return {
        "utilisateur": {"nom": "Personne A", "entite": "ALPHA"},
        "agenda": {"fichier_url_secrete": fichier_url_secrete},
        "donnees": {"racine": "/donnees"},
    }


@pytest.mark.parametrize("valeur", [URL_FICTIVE, f"webcal://calendar.example.org/{JETON}.ics"])
def test_adresse_collee_a_la_place_du_fichier_refusee_sans_recopie(tmp_path, valeur):
    with pytest.raises(config.ProfilInvalide) as exc:
        config.valider(_brut(valeur), tmp_path / "profil.toml")
    assert any(e.startswith("[agenda] fichier_url_secrete") for e in exc.value.erreurs)
    assert JETON not in str(exc.value)
    assert "calendar.example.org" not in str(exc.value)


def test_nom_de_fichier_accepte(tmp_path):
    profil = config.valider(_brut("agenda.secret"), tmp_path / "profil.toml")
    assert profil.agenda_fichier_url_secrete == tmp_path / "agenda.secret"


def _profil(fichier: Path) -> config.Profil:
    return config.Profil(
        chemin=Path("profil.toml"), nom="Personne A", entite="ALPHA", prefixe_facture="A_",
        agenda_source="ical", agenda_fichier_url_secrete=fichier, fuseau="Europe/Paris",
        racine=Path("donnees"), lancement_jour=25, lancement_heure="16:00",
        date_emission="premier_du_mois_suivant", pdf_moteur="libreoffice",
    )


def test_verification_n_affiche_que_le_nom_du_fichier(tmp_path):
    """Fichier introuvable : seul son nom s'affiche, pas le dossier qui le contient."""
    dossier = tmp_path / "dossier-personnel-0042"
    (constat,) = verification._verifier_agenda(_profil(dossier / "agenda.secret"))
    assert "agenda.secret" in constat.message
    assert "dossier-personnel-0042" not in constat.message


def test_verification_ne_recopie_pas_une_adresse_passee_en_chemin():
    """Défense en profondeur : un Profil construit sans `valider` ne fait pas fuir le jeton."""
    (constat,) = verification._verifier_agenda(_profil(Path(URL_FICTIVE)))
    assert JETON not in constat.message


# ──────────────────────────────────────────────────────────────────────
# FA-09 — destinataire incomplet
# ──────────────────────────────────────────────────────────────────────

from facturation import pcf_format as fmt  # noqa: E402

COMPLET = {"nom": "SERVICE COMPTABLE ECOLE A", "adresse_l1": "1 avenue Fictive", "adresse_l2": "00000 VILLE-A",
           "adresse_l3": "", "email": ""}


def test_destinataire_complet_sans_email_accepte():
    """L'email et le pays (L3) restent facultatifs, comme dans apply_destinataire."""
    assert fmt.champs_destinataire_manquants(COMPLET) == []


@pytest.mark.parametrize(
    ("champ", "libelle"),
    [("nom", "Nom / Raison sociale"), ("adresse_l1", "Adresse L1"), ("adresse_l2", "Adresse L2")],
)
@pytest.mark.parametrize("vide", ["", "   ", None])
def test_champ_obligatoire_manquant(champ, libelle, vide):
    assert fmt.champs_destinataire_manquants({**COMPLET, champ: vide}) == [libelle]


def test_destinataire_sans_adresse():
    sans_adresse = {**COMPLET, "adresse_l1": "", "adresse_l2": "", "email": "compta@ecole-b.example"}
    assert fmt.champs_destinataire_manquants(sans_adresse) == ["Adresse L1", "Adresse L2"]
