"""config.py — Profil de l'utilisateur (F-02, ADR014).

Chaque personne qui facture a un profil TOML, rangé hors du dépôt, qui dit :
qui facture (entité), où lire l'agenda, où sont les données, quand lancer la
préparation et avec quel moteur PDF. Aucune de ces valeurs n'est écrite dans le
code : deux utilisateurs = deux profils, et leurs données ne se croisent jamais.

Emplacement du profil, dans cet ordre :
1. chemin donné explicitement (option `--profil`) ;
2. variable d'environnement `FACTURATION_PROFIL` ;
3. emplacement par défaut : `%APPDATA%\\facturation\\profil.toml` sous Windows,
   `~/.config/facturation/profil.toml` ailleurs.

Exemple commenté : `examples/profil_exemple.toml`.
"""

from __future__ import annotations

import os
import re
import sys
import tomllib
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from facturation.pcf_format import valider_entite

NOM_FICHIER = "profil.toml"

SOURCES_AGENDA = ("ical", "api")
MOTEURS_PDF = ("libreoffice", "excel")
DATES_EMISSION = ("premier_du_mois_suivant",)

# Valeurs par défaut validées au cadrage (ADR024 : lancement le 25, émission au 1er).
DEFAUTS = {
    "agenda": {"source": "ical", "fuseau": "Europe/Paris"},
    "lancement": {"jour": 25, "heure": "16:00", "date_emission": "premier_du_mois_suivant"},
    "pdf": {"moteur": "libreoffice"},
}

# Clés acceptées par section ; toute autre clé est refusée (faute de frappe probable).
SCHEMA = {
    "utilisateur": {"nom", "entite", "prefixe_facture"},
    "agenda": {"source", "fichier_url_secrete", "fuseau"},
    "donnees": {"racine"},
    "lancement": {"jour", "heure", "date_emission"},
    "pdf": {"moteur"},
}

_HEURE_RE = re.compile(r"^([01]\d|2[0-3]):[0-5]\d$")
_PREFIXE_RE = re.compile(r"^[A-Z]{1,3}_$")


class ProfilInvalide(ValueError):
    """Profil absent ou mal rempli. `erreurs` liste chaque problème, champ par champ."""

    def __init__(self, chemin: Path, erreurs: list[str]):
        self.chemin = chemin
        self.erreurs = erreurs
        detail = "\n".join(f"  - {e}" for e in erreurs)
        super().__init__(f"Profil invalide ({chemin}) :\n{detail}")


@dataclass(frozen=True)
class Profil:
    """Profil validé d'un utilisateur."""

    chemin: Path
    nom: str
    entite: str
    prefixe_facture: str
    agenda_source: str
    agenda_fichier_url_secrete: Optional[Path]
    fuseau: str
    racine: Path
    lancement_jour: int
    lancement_heure: str
    date_emission: str
    pdf_moteur: str

    @property
    def dossier_entite(self) -> Path:
        """Dossier des données de l'entité : `<racine>/<ENTITE>/`."""
        return self.racine / self.entite

    @property
    def master(self) -> Path:
        """Chemin du MASTER de l'entité."""
        return self.dossier_entite / f"{self.entite}_MASTER.xlsx"

    @property
    def dossier_sorties(self) -> Path:
        """Dossier des sorties de lancement (brouillons, aperçus, audits)."""
        return self.dossier_entite / "_sorties"


def emplacement_par_defaut() -> Path:
    """Emplacement du profil quand aucun chemin n'est donné."""
    env = os.environ.get("FACTURATION_PROFIL")
    if env:
        return Path(env).expanduser()
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA") or Path.home() / "AppData" / "Roaming")
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME") or Path.home() / ".config")
    return base / "facturation" / NOM_FICHIER


def charger(chemin: Optional[Path] = None) -> Profil:
    """Lit et valide un profil.

    Raises:
        ProfilInvalide: fichier absent, TOML illisible, champ manquant ou invalide.
    """
    chemin = Path(chemin).expanduser() if chemin else emplacement_par_defaut()
    if not chemin.exists():
        raise ProfilInvalide(
            chemin,
            ["fichier introuvable : créez-le avec `facturation profil init`"],
        )
    try:
        brut = tomllib.loads(chemin.read_text(encoding="utf-8"))
    except tomllib.TOMLDecodeError as exc:
        raise ProfilInvalide(chemin, [f"TOML illisible : {exc}"]) from exc
    return valider(brut, chemin)


def valider(brut: dict[str, Any], chemin: Path) -> Profil:
    """Valide le contenu d'un profil déjà lu. Toutes les erreurs sont remontées d'un coup."""
    erreurs: list[str] = []

    for section, valeur in brut.items():
        if section not in SCHEMA:
            erreurs.append(f"[{section}] : section inconnue")
        elif not isinstance(valeur, dict):
            erreurs.append(f"[{section}] : doit être une section, pas une valeur")
        else:
            for cle in valeur:
                if cle not in SCHEMA[section]:
                    erreurs.append(f"[{section}] {cle} : clé inconnue")

    def lire(section: str, cle: str, obligatoire: bool = False) -> Any:
        bloc = brut.get(section) if isinstance(brut.get(section), dict) else {}
        if cle in bloc:
            return bloc[cle]
        if cle in DEFAUTS.get(section, {}):
            return DEFAUTS[section][cle]
        if obligatoire:
            erreurs.append(f"[{section}] {cle} : obligatoire")
        return None

    nom = lire("utilisateur", "nom", obligatoire=True)
    if nom is not None and (not isinstance(nom, str) or not nom.strip()):
        erreurs.append("[utilisateur] nom : doit être un texte non vide")

    entite = lire("utilisateur", "entite", obligatoire=True)
    if entite is not None:
        try:
            valider_entite(entite)
        except ValueError as exc:
            erreurs.append(f"[utilisateur] entite : {exc}")
            entite = None

    prefixe = lire("utilisateur", "prefixe_facture")
    if prefixe is None and isinstance(entite, str):
        prefixe = f"{entite[0]}_"
    if prefixe is not None and (not isinstance(prefixe, str) or not _PREFIXE_RE.match(prefixe)):
        erreurs.append(
            f"[utilisateur] prefixe_facture : {prefixe!r} invalide, attendu 1 à 3 "
            "majuscules suivies de « _ » (ex. R_)"
        )

    source = lire("agenda", "source")
    if source not in SOURCES_AGENDA:
        erreurs.append(f"[agenda] source : {source!r} invalide, attendu {' ou '.join(SOURCES_AGENDA)}")

    url_secrete = lire("agenda", "fichier_url_secrete")
    fichier_url: Optional[Path] = None
    if url_secrete is not None:
        if not isinstance(url_secrete, str) or not url_secrete.strip():
            erreurs.append("[agenda] fichier_url_secrete : doit être un chemin de fichier")
        else:
            fichier_url = Path(url_secrete).expanduser()
            if not fichier_url.is_absolute():
                fichier_url = chemin.parent / fichier_url

    fuseau = lire("agenda", "fuseau")
    try:
        ZoneInfo(str(fuseau))
    except (ZoneInfoNotFoundError, ValueError):
        erreurs.append(f"[agenda] fuseau : {fuseau!r} inconnu (ex. Europe/Paris)")

    racine_brute = lire("donnees", "racine", obligatoire=True)
    racine: Optional[Path] = None
    if racine_brute is not None:
        if not isinstance(racine_brute, str) or not racine_brute.strip():
            erreurs.append("[donnees] racine : doit être un chemin de dossier")
        else:
            racine = Path(racine_brute).expanduser()

    jour = lire("lancement", "jour")
    if not isinstance(jour, int) or isinstance(jour, bool) or not 1 <= jour <= 28:
        erreurs.append(f"[lancement] jour : {jour!r} invalide, attendu un entier de 1 à 28")

    heure = lire("lancement", "heure")
    if not isinstance(heure, str) or not _HEURE_RE.match(heure):
        erreurs.append(f"[lancement] heure : {heure!r} invalide, attendu HH:MM (ex. 16:00)")

    date_emission = lire("lancement", "date_emission")
    if date_emission not in DATES_EMISSION:
        erreurs.append(
            f"[lancement] date_emission : {date_emission!r} invalide, attendu "
            f"{' ou '.join(DATES_EMISSION)}"
        )

    moteur = lire("pdf", "moteur")
    if moteur not in MOTEURS_PDF:
        erreurs.append(f"[pdf] moteur : {moteur!r} invalide, attendu {' ou '.join(MOTEURS_PDF)}")

    if erreurs:
        raise ProfilInvalide(chemin, erreurs)

    return Profil(
        chemin=chemin,
        nom=nom.strip(),
        entite=entite,
        prefixe_facture=prefixe,
        agenda_source=source,
        agenda_fichier_url_secrete=fichier_url,
        fuseau=fuseau,
        racine=racine,
        lancement_jour=jour,
        lancement_heure=heure,
        date_emission=date_emission,
        pdf_moteur=moteur,
    )


def resoudre_racine(
    data_root: Optional[Path] = None, profil: Optional[Path] = None
) -> Optional[Profil]:
    """Fixe la racine de données des scripts et renvoie le profil utilisé, s'il y en a un.

    Ordre : `--data-root` > `--profil` > `$PCF_DATA_ROOT` > profil par défaut s'il
    existe > dossier courant (comportement historique de `pcf_format`).
    """
    from facturation import pcf_format as fmt

    if data_root:
        fmt.set_data_root(data_root)
        return None
    if profil:
        p = charger(profil)
        fmt.set_data_root(p.racine)
        return p
    if os.environ.get("PCF_DATA_ROOT"):
        return None
    defaut = emplacement_par_defaut()
    if defaut.exists():
        p = charger(defaut)
        fmt.set_data_root(p.racine)
        return p
    return None


NOM_FICHIER_ICAL = "agenda.secret"
_ICAL_RE = re.compile(r"^https://\S+\.ics$")
_LIGNE_ICAL_RE = re.compile(r"^#?\s*fichier_url_secrete\s*=.*$", re.MULTILINE)


def adresse_ical_valide(adresse: str) -> bool:
    """Forme attendue d'une adresse iCal : https://….ics, sans espace."""
    return bool(_ICAL_RE.match(adresse.strip()))


def enregistrer_adresse_ical(profil: Profil, adresse: str) -> Path:
    """Range l'adresse iCal secrète à côté du profil et la déclare dans le profil.

    L'adresse est écrite dans `agenda.secret` (jamais dans le profil lui-même ni
    dans le dépôt), puis la ligne `fichier_url_secrete` du profil est activée.

    Returns:
        Le chemin du fichier secret.

    Raises:
        ValueError: adresse mal formée.
        ProfilInvalide: le profil modifié ne se relit plus (cas anormal).
    """
    adresse = adresse.strip()
    if not adresse_ical_valide(adresse):
        raise ValueError(
            "adresse iCal invalide : elle doit commencer par https:// et finir par .ics "
            "(Google Agenda → Paramètres et partage → Adresse secrète au format iCal)"
        )
    secret = profil.chemin.parent / NOM_FICHIER_ICAL
    secret.write_text(adresse + "\n", encoding="utf-8")

    texte = profil.chemin.read_text(encoding="utf-8")
    ligne = f'fichier_url_secrete = "{NOM_FICHIER_ICAL}"'
    if _LIGNE_ICAL_RE.search(texte):
        texte = _LIGNE_ICAL_RE.sub(ligne, texte, count=1)
    elif re.search(r"^\[agenda\]\s*$", texte, re.MULTILINE):
        texte = re.sub(r"^\[agenda\]\s*$", f"[agenda]\n{ligne}", texte, count=1, flags=re.MULTILINE)
    else:
        texte = texte.rstrip() + f"\n\n[agenda]\n{ligne}\n"
    profil.chemin.write_text(texte, encoding="utf-8")
    charger(profil.chemin)  # le profil doit rester valide
    return secret


def _texte_toml(valeur: str) -> str:
    """Chaîne TOML entre guillemets, avec échappement des caractères spéciaux."""
    return '"' + valeur.replace("\\", "\\\\").replace('"', '\\"') + '"'


def rediger(nom: str, entite: str, racine: str) -> str:
    """Contenu d'un profil neuf, commenté, prêt à être écrit sur disque."""
    valider_entite(entite)
    racine_posix = racine.replace("\\", "/")
    return f"""# Profil de facturation — un fichier par personne qui facture.
# Rangé hors du dépôt git : il ne doit jamais être versionné.

[utilisateur]
nom = {_texte_toml(nom)}
entite = {_texte_toml(entite)}
# prefixe_facture = "{entite[0]}_"     # par défaut : initiale de l'entité + « _ »

[agenda]
source = "ical"                  # "ical" (adresse secrète) ou "api" (Google Agenda)
# fichier_url_secrete = "agenda.secret"   # fichier séparé, contenant l'adresse iCal secrète
fuseau = "Europe/Paris"

[donnees]
# Dossier qui contient {entite}/{entite}_MASTER.xlsx (barres obliques acceptées sous Windows)
racine = {_texte_toml(racine_posix)}

[lancement]
jour = 25                        # prépare tout le mois en cours (ADR024)
heure = "16:00"                  # heure du lancement le 25 (Q6)
date_emission = "premier_du_mois_suivant"

[pdf]
moteur = "libreoffice"           # ou "excel"
"""
