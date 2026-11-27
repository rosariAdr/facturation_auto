"""verification.py — Contrôle d'une installation à partir d'un profil (F-02).

`facturation profil verifier` passe ces contrôles avant tout lancement : un
problème bloquant (ERREUR) empêcherait de préparer les factures, un point
d'ATTENTION ne bloque pas encore mais le deviendra à un lot suivant.

Depuis P-0b, la vérification signale aussi en ATTENTION chaque onglet de suivi que l'outil
n'a pas pu lire (`Master.suivi_invalide`) et chaque perte qu'un enregistrement du MASTER par
l'outil entraînerait (`master.fidelite_aller_retour`, qui travaille sur une copie temporaire).
Un MASTER que `lire_master` refuse (colonne obligatoire manquante, fichier illisible…) est une
ERREUR : toute préparation échouerait (revue de P-0b). Un MASTER qui ne s'ouvre pas parce qu'il
est ouvert dans Excel ou verrouillé reste une ATTENTION, comme le classeur ouvert.

Rien n'est écrit : la vérification ne fait que lire.
"""

from __future__ import annotations

import shutil
import sys
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook

from facturation import config, master
from facturation import pcf_format as fmt
from facturation.config import Profil

OK, ATTENTION, ERREUR = "OK", "ATTENTION", "ERREUR"

ONGLETS_OBLIGATOIRES = ("Coordonnées", "Sessions")
ONGLETS_SUIVI = ("Cours", "Tarifs", "Écoles", "Devis")

_EMPLACEMENTS_LIBREOFFICE_WINDOWS = (
    Path(r"C:\Program Files\LibreOffice\program\soffice.exe"),
    Path(r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"),
)


@dataclass(frozen=True)
class Constat:
    niveau: str
    sujet: str
    message: str


def trouver_libreoffice() -> Path | None:
    """Chemin de l'exécutable LibreOffice, ou None s'il est introuvable."""
    for nom in ("soffice", "libreoffice"):
        trouve = shutil.which(nom)
        if trouve:
            return Path(trouve)
    if sys.platform == "win32":
        for chemin in _EMPLACEMENTS_LIBREOFFICE_WINDOWS:
            if chemin.exists():
                return chemin
    return None


def verifier(profil: Profil) -> list[Constat]:
    """Contrôle le profil, les données et les outils. Renvoie la liste des constats."""
    constats: list[Constat] = [
        Constat(OK, "Profil", f"{profil.chemin} — {profil.nom}, entité {profil.entite}"),
    ]
    constats += _verifier_donnees(profil)
    constats += _verifier_gabarit(profil)
    constats += _verifier_agenda(profil)
    constats += _verifier_pdf(profil)
    return constats


def _verifier_donnees(profil: Profil) -> list[Constat]:
    if not profil.racine.is_dir():
        return [Constat(
            ERREUR, "Données",
            f"dossier introuvable : {profil.racine}. Si c'est un lecteur Google Drive "
            "pour ordinateur, vérifiez que l'application est lancée et la session ouverte.",
        )]
    constats = [Constat(OK, "Données", f"dossier trouvé : {profil.racine}")]

    if not profil.master.is_file():
        constats.append(Constat(ERREUR, "MASTER", f"fichier introuvable : {profil.master}"))
        return constats

    verrou = profil.master.with_name(f"~${profil.master.name}")
    if verrou.exists():
        constats.append(Constat(
            ATTENTION, "MASTER",
            "le fichier semble ouvert dans Excel : fermez-le avant toute écriture",
        ))

    try:
        wb = load_workbook(profil.master, read_only=True)
        onglets = set(wb.sheetnames)
        wb.close()
    except Exception as exc:  # fichier corrompu, verrouillé, non synchronisé…
        constats.append(Constat(ERREUR, "MASTER", f"illisible : {exc}"))
        return constats

    manquants = [o for o in ONGLETS_OBLIGATOIRES if o not in onglets]
    if manquants:
        constats.append(Constat(ERREUR, "MASTER", f"onglets obligatoires absents : {', '.join(manquants)}"))
        return constats
    absents_suivi = [o for o in ONGLETS_SUIVI if o not in onglets]
    if absents_suivi:
        constats.append(Constat(ATTENTION, "MASTER", f"onglets de suivi absents : {', '.join(absents_suivi)}"))
    constats.append(Constat(OK, "MASTER", f"{profil.master.name} lisible, {len(onglets)} onglets"))
    constats += _verifier_lecture_et_fidelite(profil)

    fmt.set_data_root(profil.racine)
    try:
        donnees = fmt._load_master_data(profil.entite, refresh=True)
    except RuntimeError as exc:
        constats.append(Constat(ERREUR, "Coordonnées", str(exc)))
        return constats
    constats.append(Constat(
        OK, "Coordonnées",
        f"émetteur « {donnees['emetteur']['nom']} », SIRET, IBAN et BIC présents",
    ))

    # Mentions obligatoires de la facture (ADR030) : sans elles, toute facture est bloquée.
    if not fmt.emetteur_porte_ei(donnees["emetteur"]["nom"]):
        constats.append(Constat(
            ERREUR, "Coordonnées",
            f"le nom de l'émetteur « {donnees['emetteur']['nom']} » ne porte pas la mention "
            "« EI » ni « Entrepreneur individuel » (ADR030) : factures bloquées",
        ))
    try:
        fmt.get_nda(entite=profil.entite)
    except ValueError as exc:
        constats.append(Constat(ERREUR, "Coordonnées", f"{exc} Factures bloquées (ADR030)."))

    incomplets = sorted(code for code, dest in donnees["destinataires"].items() if dest is None)
    if incomplets:
        constats.append(Constat(
            ATTENTION, "Coordonnées",
            f"destinataires sans coordonnées (factures bloquées) : {', '.join(incomplets)}",
        ))
    partiels = sorted(
        f"{code} ({', '.join(manquants)})"
        for code, dest in donnees["destinataires"].items()
        if dest is not None and (manquants := fmt.champs_destinataire_manquants(dest))
    )
    if partiels:
        constats.append(Constat(
            ATTENTION, "Coordonnées",
            f"destinataires incomplets (factures bloquées) : {', '.join(partiels)}",
        ))
    return constats


def _verifier_lecture_et_fidelite(profil: Profil) -> list[Constat]:
    """MASTER refusé par l'outil : une ERREUR. Onglets de suivi illisibles et pertes d'un
    enregistrement par l'outil : une ATTENTION chacun."""
    constats: list[Constat] = []
    try:
        lu = master.lire_master(profil.master)
    except master.MasterVerrouille as exc:  # classeur ouvert ou verrouillé : passager
        constats.append(Constat(ATTENTION, "MASTER", f"lecture par l'outil impossible : {exc}"))
    except (master.MasterInvalide, OSError) as exc:
        constats.append(Constat(
            ERREUR, "MASTER", f"refusé par l'outil, aucune facture ne peut être préparée : {exc}",
        ))
    else:
        for onglet, message in lu.suivi_invalide.items():
            constats.append(Constat(ATTENTION, "MASTER", f"onglet de suivi « {onglet} » ignoré : {message}"))
    try:
        pertes = master.fidelite_aller_retour(profil.master)
    except (master.MasterInvalide, OSError) as exc:
        constats.append(Constat(ATTENTION, "MASTER", f"contrôle de fidélité impossible : {exc}"))
    else:
        for perte in pertes:
            constats.append(Constat(ATTENTION, "MASTER", f"un enregistrement par l'outil changerait le classeur : {perte}"))
    return constats


def _verifier_gabarit(profil: Profil) -> list[Constat]:
    gabarit = fmt.gabarit_pour(profil.entite)
    if not gabarit.is_file():
        return [Constat(ERREUR, "Gabarit", f"introuvable : {gabarit}")]
    sorte = "standard" if gabarit == fmt.GABARIT_STANDARD else "propre à l'entité"
    return [Constat(OK, "Gabarit", f"{gabarit.name} ({sorte})")]


def _verifier_agenda(profil: Profil) -> list[Constat]:
    if profil.agenda_source == "api":
        return [Constat(ATTENTION, "Agenda", "source « api » : prévue en option, pas encore disponible")]
    fichier = profil.agenda_fichier_url_secrete
    if fichier is None:
        return [Constat(
            ATTENTION, "Agenda",
            "aucune adresse iCal secrète déclarée : lancez `facturation profil agenda` "
            "(nécessaire à partir du lot 1)",
        )]
    if not fichier.is_file():
        # Seul le nom s'affiche : le chemin complet pourrait contenir l'adresse elle-même (FA-10).
        return [Constat(
            ATTENTION, "Agenda",
            f"fichier d'adresse iCal introuvable : {fichier.name} "
            "(déclaré dans le profil, [agenda] fichier_url_secrete)",
        )]
    try:
        # utf-8-sig : un fichier retouché au Bloc-notes ou sous PowerShell 5 peut commencer par un BOM.
        adresse = fichier.read_text(encoding="utf-8-sig").strip()
    except (OSError, UnicodeDecodeError):
        # Ni l'erreur d'origine ni le contenu : ils pourraient citer l'adresse.
        return [Constat(
            ERREUR, "Agenda",
            f"{fichier.name} illisible (attendu : du texte UTF-8) : relancez `facturation profil agenda`",
        )]
    # L'adresse est un secret : on ne l'affiche jamais, on contrôle seulement sa forme.
    if not config.adresse_ical_valide(adresse):
        return [Constat(
            ERREUR, "Agenda",
            f"le contenu de {fichier.name} n'est pas une adresse iCal (https://….ics, sans espace) : "
            "relancez `facturation profil agenda`",
        )]
    return [Constat(OK, "Agenda", f"adresse iCal secrète présente dans {fichier.name} (non affichée)")]


def _verifier_pdf(profil: Profil) -> list[Constat]:
    if profil.pdf_moteur == "excel":
        return [Constat(ATTENTION, "PDF", "moteur « excel » : prévu en option (ADR021), pas encore disponible")]
    soffice = trouver_libreoffice()
    if soffice is None:
        return [Constat(ATTENTION, "PDF", "LibreOffice introuvable : nécessaire pour exporter les PDF")]
    return [Constat(OK, "PDF", f"LibreOffice trouvé : {soffice}")]
