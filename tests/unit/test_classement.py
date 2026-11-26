"""Classement des séances de l'agenda (F-12, ADR016, ADR028 ; banque de tests : T-CLA-01 à T-CLA-52).

Tests d'acceptation écrits avant l'implémentation : chaque test importe `facturation.classement`
lui-même, pour que son absence fasse échouer les tests sans bloquer la collecte des autres.
Séances (`facturation.agenda.Seance`) et lignes de `Correspondances` (dictionnaires
{en-tête : valeur}, forme rendue par `facturation.master`) construites à la main, avec des
données fictives : écoles ABC, DEF et XYZ, intitulés et codes inventés.

Ordre de recherche (contrat de la vague 2, ADR016) : code d'événement `#ÉCOLE-PROMO-MODULE` dans
la description, cherché parmi les lignes de type `code` (le premier code trouvé compte) ; sinon
intitulé exact, sans tenir compte de la casse ni des espaces en début et en fin ; sinon le plus
long préfixe ; sinon file à classer (« non reconnue »). Code absent de Correspondances : « code
inconnu ». Deux lignes qui conviennent au même niveau : « ambiguë ». Un créneau à classer n'est
jamais facturé par défaut (ADR028).

Lectures retenues là où le contrat se tait (`💡 [proposé — à valider]`) :
- un niveau qui a des candidats tranche : une ambiguïté (code, exact) ne retombe pas sur le
  niveau suivant, pas plus qu'un code inconnu ne retombe sur l'intitulé ;
- clés et types de clé comparés sans tenir compte de la casse ni des espaces en début et en
  fin, pour les trois types (le contrat ne le dit que pour l'intitulé exact) ;
- un dièse suivi de minuscules (`#visio`) n'est pas un code ; un code écrit dans l'intitulé
  n'est pas lu comme un code ;
- une ligne de Correspondances sans type de clé, d'un type inconnu ou sans clé lève
  `ValueError` (une clé vide de type `préfixe` reconnaîtrait toutes les séances) ;
- valeurs de la ligne retenue rendues sans espaces autour, cellule vide → chaîne vide ;
- `ligne_correspondance` : clé = intitulé (types `exact` et `préfixe`) ou premier code de la
  description (type `code`) ; type de clé, mode d'heures, école, clé tarif, clé introuvable et
  valeur absente en mode `coefficient` ou `forfait` contrôlés par `ValueError`.

T-CLA-41 à T-CLA-52 viennent de la revue croisée de F-12. Leurs lectures `💡 [proposé — à
valider]` sont dites dans chaque test.
"""

from __future__ import annotations

import copy
import dataclasses
import importlib
import itertools
import unicodedata
from datetime import date, datetime
from zoneinfo import ZoneInfo

import pytest

from facturation.agenda import Seance

pytestmark = pytest.mark.unit

PARIS = ZoneInfo("Europe/Paris")
COLONNES_CORRESPONDANCES = (  # MASTER_SCHEMA, onglet Correspondances
    "Clé", "Type de clé", "École", "Module", "Promotion", "Clé tarif", "Mode heures", "Valeur",
    "Gardée le", "Notes",
)
GARDEE_LE = date(2026, 11, 26)
_NUMEROS = itertools.count(1)


def _module():
    """Module testé, importé dans le test : son absence fait échouer le test, pas la collecte."""
    return importlib.import_module("facturation.classement")


def _seance(
    intitule: str = "Cours fictif",
    description: str = "",
    jour: date = date(2026, 11, 3),
    debut: tuple[int, int] = (9, 0),
    fin: tuple[int, int] = (12, 30),
) -> Seance:
    """Séance fictive de l'agenda, à l'heure de Paris, d'identifiant unique."""
    moment_debut = datetime(jour.year, jour.month, jour.day, *debut, tzinfo=PARIS)
    moment_fin = datetime(jour.year, jour.month, jour.day, *fin, tzinfo=PARIS)
    return Seance(
        uid=f"seance-fictive-{next(_NUMEROS)}@example.org",
        date=jour,
        debut=moment_debut,
        fin=moment_fin,
        duree_h=(moment_fin - moment_debut).total_seconds() / 3600,
        intitule=intitule,
        description=description,
    )


def _ligne(
    cle,
    type_cle,
    ecole="ABC",
    *,
    module="Module fictif",
    promotion="M1",
    cle_tarif=None,
    mode="agenda",
    valeur=None,
    notes=None,
) -> dict:
    """Ligne fictive de l'onglet Correspondances, avec toutes les colonnes de MASTER_SCHEMA."""
    return {
        "Clé": cle,
        "Type de clé": type_cle,
        "École": ecole,
        "Module": module,
        "Promotion": promotion,
        "Clé tarif": cle_tarif if cle_tarif is not None else f"{ecole}|{module}|{promotion}",
        "Mode heures": mode,
        "Valeur": valeur,
        "Gardée le": None,
        "Notes": notes,
    }


def _classee(seance, lignes):
    """La séance classée (et elle seule) ; échoue si elle part dans la file à classer."""
    classees, a_classer = _module().classer([seance], lignes)
    assert a_classer == [], f"séance mise à classer : {[a.raison for a in a_classer]}"
    assert len(classees) == 1
    assert classees[0].seance == seance
    return classees[0]


def _raison(seance, lignes) -> str:
    """La raison de la mise en file à classer ; échoue si la séance est classée."""
    classees, a_classer = _module().classer([seance], lignes)
    assert classees == [], f"séance classée à tort : {[(c.ecole, c.par) for c in classees]}"
    assert len(a_classer) == 1
    assert a_classer[0].seance == seance
    return a_classer[0].raison


def _message(exc_info) -> str:
    return str(exc_info.value).casefold()


def _ligne_proposee(seance, **options) -> dict:
    """`ligne_correspondance` avec des valeurs fictives par défaut, surchargées par `options`."""
    arguments = {
        "ecole": "XYZ",
        "module": "Atelier fictif",
        "promotion": "L3",
        "cle_tarif": "XYZ|Atelier fictif|L3",
        "garde_le": GARDEE_LE,
        **options,
    }
    return _module().ligne_correspondance(seance, **arguments)


# ──────────────────────────────────────────────────────────────────────
# Forme du résultat
# ──────────────────────────────────────────────────────────────────────


def test_T_CLA_01_classes_de_resultat_figees():
    """T-CLA-01 — SeanceClassee et SeanceAClasser : dataclasses figées aux champs du contrat ;
    une séance à classer ne porte ni école ni tarif (jamais facturée par défaut)."""
    m = _module()
    assert dataclasses.is_dataclass(m.SeanceClassee)
    assert dataclasses.is_dataclass(m.SeanceAClasser)
    assert {f.name for f in dataclasses.fields(m.SeanceClassee)} == {
        "seance", "ecole", "module", "promotion", "cle_tarif", "correspondance", "par",
    }
    assert {f.name for f in dataclasses.fields(m.SeanceAClasser)} == {"seance", "raison"}

    seance = _seance()
    classee = m.SeanceClassee(
        seance=seance, ecole="ABC", module="Module fictif", promotion="M1",
        cle_tarif="ABC|Module fictif|M1", correspondance=_ligne("Cours fictif", "exact"), par="exact",
    )
    a_classer = m.SeanceAClasser(seance=seance, raison="non reconnue")
    with pytest.raises(dataclasses.FrozenInstanceError):
        classee.ecole = "DEF"
    with pytest.raises(dataclasses.FrozenInstanceError):
        a_classer.raison = "ambiguë"


def test_T_CLA_02_entrees_vides_et_sans_correspondances():
    """T-CLA-02 — `classer` rend un couple de listes ; sans séance, deux listes vides ; sans
    aucune correspondance, chaque séance va dans la file à classer, raison « non reconnue »."""
    m = _module()
    resultat = m.classer([], [])
    assert isinstance(resultat, tuple) and len(resultat) == 2
    assert resultat == ([], [])
    assert m.classer([], [_ligne("Cours fictif", "exact")]) == ([], [])

    seances = [_seance("Cours fictif"), _seance("Autre cours fictif", jour=date(2026, 11, 4))]
    classees, a_classer = m.classer(seances, [])
    assert classees == []
    assert [(a.seance, a.raison) for a in a_classer] == [(s, "non reconnue") for s in seances]


# ──────────────────────────────────────────────────────────────────────
# Niveau 1 — code d'événement dans la description
# ──────────────────────────────────────────────────────────────────────


def test_T_CLA_03_code_reconnu_parmi_les_lignes_de_type_code():
    """T-CLA-03 — Un code de la description, présent parmi les lignes de type `code`, classe la
    séance (`par` = « code ») avec l'école, le module, la promotion et la clé tarif de la ligne,
    qui est rendue entière (mode d'heures et valeur compris, pour F-13)."""
    ligne = _ligne(
        "#ABC-M1-STATS", "code", "ABC", module="Statistiques fictives", promotion="M1",
        cle_tarif="ABC|Statistiques fictives|M1", mode="forfait", valeur=7.5, notes="ligne du code",
    )
    autre = _ligne("#DEF-M2-ECO", "code", "DEF", module="Économie fictive", promotion="M2")
    seance = _seance("Intitulé quelconque", description="#ABC-M1-STATS")

    classee = _classee(seance, [autre, ligne])
    assert classee.par == "code"
    assert (classee.ecole, classee.module, classee.promotion, classee.cle_tarif) == (
        "ABC", "Statistiques fictives", "M1", "ABC|Statistiques fictives|M1",
    )
    assert classee.correspondance == ligne
    assert classee.correspondance["Mode heures"] == "forfait"
    assert classee.correspondance["Valeur"] == 7.5


@pytest.mark.parametrize(
    "description",
    [
        "Salle 12\nCode : #ABC-M1-STATS\nApporter le vidéoprojecteur",
        "Séance de rattrapage (#ABC-M1-STATS).",
        "#ABC-M1-STATS, groupe complet",
        "   #ABC-M1-STATS   ",
    ],
    ids=["plusieurs-lignes", "parentheses-point", "virgule", "espaces"],
)
def test_T_CLA_04_code_au_milieu_du_texte(description):
    """T-CLA-04 — Le code est trouvé où qu'il soit dans la description, quelle que soit la
    ponctuation qui l'entoure (le code s'arrête au premier caractère hors majuscules, chiffres
    et tirets)."""
    classee = _classee(_seance(description=description), [_ligne("#ABC-M1-STATS", "code", "ABC")])
    assert classee.par == "code"
    assert classee.ecole == "ABC"


def test_T_CLA_05_le_code_prime_sur_l_intitule():
    """T-CLA-05 — Un code connu l'emporte sur un intitulé exact et sur un préfixe qui
    conviennent aussi (ADR016 : le code fait foi)."""
    lignes = [
        _ligne("Cours fictif", "préfixe", "XYZ"),
        _ligne("Cours fictif avancé", "exact", "DEF"),
        _ligne("#ABC-M1-STATS", "code", "ABC"),
    ]
    classee = _classee(_seance("Cours fictif avancé", description="#ABC-M1-STATS"), lignes)
    assert (classee.par, classee.ecole) == ("code", "ABC")


def test_T_CLA_06_code_inconnu_sans_repli_sur_l_intitule():
    """T-CLA-06 — Un code présent mais absent de Correspondances met la séance à classer,
    raison « code inconnu », même si son intitulé convient à une ligne exacte ou à un préfixe."""
    lignes = [
        _ligne("#ABC-M1-STATS", "code", "ABC"),
        _ligne("Cours fictif", "exact", "DEF"),
        _ligne("Cours", "préfixe", "XYZ"),
    ]
    seance = _seance("Cours fictif", description="Nouveau groupe : #ZZZ-X9-INCONNU")
    assert _raison(seance, lignes) == "code inconnu"
    assert _raison(_seance(description="#ZZZ-X9-INCONNU"), []) == "code inconnu"


@pytest.mark.parametrize(
    ("code_decrit", "cle"),
    [("#ABC-M1-STAT", "#ABC-M1-STATS"), ("#ABC-M1-STATS", "#ABC-M1-STAT")],
    ids=["code-plus-court", "code-plus-long"],
)
def test_T_CLA_07_code_compare_en_entier(code_decrit, cle):
    """T-CLA-07 — Un code se compare en entier, jamais par préfixe : un code plus court ou plus
    long que la clé reste « code inconnu »."""
    assert _raison(_seance(description=code_decrit), [_ligne(cle, "code", "ABC")]) == "code inconnu"


def test_T_CLA_08_le_premier_code_trouve_compte():
    """T-CLA-08 — Avec plusieurs codes dans la description, seul le premier compte : inconnu, il
    met la séance à classer même si le second est connu ; connu, il classe la séance sans
    ambiguïté avec le second."""
    lignes = [_ligne("#ABC-M1-STATS", "code", "ABC"), _ligne("#DEF-M2-ECO", "code", "DEF")]
    inconnu_puis_connu = _seance(description="#ZZZ-X9-INCONNU, sinon #ABC-M1-STATS")
    assert _raison(inconnu_puis_connu, lignes) == "code inconnu"

    connu_puis_connu = _seance(description="#DEF-M2-ECO (ancien code : #ABC-M1-STATS)")
    classee = _classee(connu_puis_connu, lignes)
    assert (classee.par, classee.ecole) == ("code", "DEF")


def test_T_CLA_09_seules_les_lignes_de_type_code_portent_les_codes():
    """T-CLA-09 — Le code est cherché parmi les lignes de type `code` seulement : une ligne
    `exact` (ou `préfixe`) dont la clé est ce code ne le reconnaît pas."""
    lignes = [_ligne("#ABC-M1-STATS", "exact", "ABC"), _ligne("#ABC", "préfixe", "DEF")]
    assert _raison(_seance("Cours fictif", description="#ABC-M1-STATS"), lignes) == "code inconnu"


def test_T_CLA_10_cle_de_code_sans_casse_ni_espaces():
    """T-CLA-10 — La clé d'une ligne `code` se compare sans tenir compte de la casse ni des
    espaces en début et en fin (saisie à la main dans le MASTER).
    💡 [proposé — à valider] : le contrat ne le précise que pour l'intitulé exact."""
    classee = _classee(_seance(description="#ABC-M1-STATS"), [_ligne("  #abc-m1-stats ", "code", "ABC")])
    assert (classee.par, classee.ecole) == ("code", "ABC")


@pytest.mark.parametrize(
    "description",
    ["#visio", "Lien #visio envoyé la veille", "#abc-m1-stats"],
    ids=["mot-diese", "mot-diese-dans-le-texte", "code-en-minuscules"],
)
def test_T_CLA_11_diese_en_minuscules_n_est_pas_un_code(description):
    """T-CLA-11 — Un code s'écrit en lettres majuscules : un dièse suivi de minuscules n'est pas
    un code (ni « code inconnu ») et la séance est reconnue par son intitulé."""
    lignes = [_ligne("#ABC-M1-STATS", "code", "ABC"), _ligne("Cours fictif", "exact", "DEF")]
    classee = _classee(_seance("Cours fictif", description=description), lignes)
    assert (classee.par, classee.ecole) == ("exact", "DEF")


def test_T_CLA_12_code_dans_l_intitule_non_lu_comme_code():
    """T-CLA-12 — Le code se lit dans la description : un code écrit dans l'intitulé ne rend pas
    la séance « code inconnu », et l'intitulé est rapproché normalement (ici par préfixe)."""
    classee = _classee(_seance("Cours fictif #ZZZ-X9-INCONNU"), [_ligne("Cours fictif", "préfixe", "DEF")])
    assert (classee.par, classee.ecole) == ("préfixe", "DEF")


def test_T_CLA_13_code_ambigu_sans_repli():
    """T-CLA-13 — Deux lignes `code` de même clé : « ambiguë », sans retomber sur l'intitulé
    exact qui conviendrait."""
    lignes = [
        _ligne("#ABC-M1-STATS", "code", "ABC"),
        _ligne("#ABC-M1-STATS", "code", "DEF"),
        _ligne("Cours fictif", "exact", "XYZ"),
    ]
    assert _raison(_seance("Cours fictif", description="#ABC-M1-STATS"), lignes) == "ambiguë"


# ──────────────────────────────────────────────────────────────────────
# Niveau 2 — intitulé exact
# ──────────────────────────────────────────────────────────────────────


def test_T_CLA_14_intitule_exact():
    """T-CLA-14 — Sans code, l'intitulé égal à la clé d'une ligne `exact` classe la séance
    (`par` = « exact ») avec les valeurs de cette ligne."""
    economie = _ligne("Cours fictif d'économie", "exact", "DEF", module="Économie fictive", promotion="M2")
    lignes = [_ligne("Cours fictif de statistiques", "exact", "ABC"), economie]
    classee = _classee(_seance("Cours fictif d'économie"), lignes)
    assert classee.par == "exact"
    assert (classee.ecole, classee.module, classee.promotion, classee.cle_tarif) == (
        "DEF", "Économie fictive", "M2", "DEF|Économie fictive|M2",
    )
    assert classee.correspondance == economie


@pytest.mark.parametrize(
    ("intitule", "cle"),
    [
        ("ÉCONOMÉTRIE FICTIVE", "Économétrie fictive"),
        ("  économétrie fictive  ", "Économétrie fictive"),
        ("Économétrie fictive", " \tÉCONOMÉTRIE fictive  "),
    ],
    ids=["majuscules-accentuees", "espaces-intitule", "espaces-et-casse-cle"],
)
def test_T_CLA_15_exact_sans_casse_ni_espaces_autour(intitule, cle):
    """T-CLA-15 — L'intitulé exact se compare sans tenir compte de la casse (accents compris) ni
    des espaces en début et en fin, d'un côté comme de l'autre."""
    classee = _classee(_seance(intitule), [_ligne(cle, "exact", "ABC")])
    assert (classee.par, classee.ecole) == ("exact", "ABC")


@pytest.mark.parametrize(
    "intitule",
    ["Cours fictif — groupe 2", "Rattrapage cours fictif", "Cours"],
    ids=["texte-en-plus", "texte-avant", "intitule-partiel"],
)
def test_T_CLA_16_exact_ne_vaut_pas_contient(intitule):
    """T-CLA-16 — Une ligne `exact` ne reconnaît ni un intitulé plus long, ni un intitulé qui la
    contient, ni un intitulé partiel : « non reconnue »."""
    assert _raison(_seance(intitule), [_ligne("Cours fictif", "exact", "ABC")]) == "non reconnue"


def test_T_CLA_17_exact_prime_sur_prefixe():
    """T-CLA-17 — Un intitulé exact l'emporte sur les préfixes, même sur un préfixe égal à
    l'intitulé entier."""
    lignes = [
        _ligne("Cours fictif", "préfixe", "ABC"),
        _ligne("Cours fictif avancé", "préfixe", "DEF"),
        _ligne("Cours fictif avancé", "exact", "XYZ"),
    ]
    classee = _classee(_seance("Cours fictif avancé"), lignes)
    assert (classee.par, classee.ecole) == ("exact", "XYZ")


def test_T_CLA_18_exact_ambigu_sans_repli():
    """T-CLA-18 — Deux lignes `exact` qui conviennent (même clé à la casse et aux espaces près) :
    « ambiguë », sans retomber sur un préfixe qui conviendrait."""
    lignes = [
        _ligne("Cours fictif", "exact", "ABC"),
        _ligne("COURS FICTIF ", "exact", "DEF"),
        _ligne("Cours", "préfixe", "XYZ"),
    ]
    assert _raison(_seance("Cours fictif"), lignes) == "ambiguë"


# ──────────────────────────────────────────────────────────────────────
# Niveau 3 — préfixe
# ──────────────────────────────────────────────────────────────────────


def test_T_CLA_19_prefixe():
    """T-CLA-19 — Sans code ni intitulé exact, une ligne `préfixe` par laquelle commence
    l'intitulé classe la séance (`par` = « préfixe »)."""
    ligne = _ligne("Cours fictif", "préfixe", "ABC", module="Module fictif", promotion="M1")
    classee = _classee(_seance("Cours fictif — séance 4 sur 10"), [ligne])
    assert classee.par == "préfixe"
    assert (classee.ecole, classee.cle_tarif) == ("ABC", "ABC|Module fictif|M1")
    assert classee.correspondance == ligne


def test_T_CLA_20_le_plus_long_prefixe_gagne_quel_que_soit_l_ordre():
    """T-CLA-20 — Parmi plusieurs préfixes qui conviennent, le plus long gagne, quel que soit
    l'ordre des lignes dans l'onglet."""
    lignes = [
        _ligne("Cours", "préfixe", "ABC"),
        _ligne("Cours fictif", "préfixe", "DEF"),
        _ligne("Cours fictif avancé", "préfixe", "XYZ"),
    ]
    attendus = {
        "Cours fictif avancé — TD 3": "XYZ",
        "Cours fictif — TD 1": "DEF",
        "Cours magistral": "ABC",
    }
    for ordre in itertools.permutations(lignes):
        for intitule, ecole in attendus.items():
            classee = _classee(_seance(intitule), list(ordre))
            assert (classee.par, classee.ecole) == ("préfixe", ecole), (intitule, [ligne["Clé"] for ligne in ordre])


@pytest.mark.parametrize(
    "intitule",
    ["Rattrapage du cours fictif", "Cours fict"],
    ids=["prefixe-au-milieu", "intitule-plus-court-que-la-cle"],
)
def test_T_CLA_21_prefixe_est_un_debut(intitule):
    """T-CLA-21 — Un préfixe doit ouvrir l'intitulé : une clé trouvée au milieu, ou plus longue
    que l'intitulé, ne reconnaît pas la séance."""
    assert _raison(_seance(intitule), [_ligne("Cours fictif", "préfixe", "ABC")]) == "non reconnue"


def test_T_CLA_22_prefixe_sans_casse_ni_espaces_autour():
    """T-CLA-22 — Le préfixe se compare sans tenir compte de la casse ni des espaces en début et
    en fin (de la clé comme de l'intitulé).
    💡 [proposé — à valider] : le contrat ne le précise que pour l'intitulé exact."""
    classee = _classee(_seance("  Cours fictif — TD 2"), [_ligne(" cours FICTIF ", "préfixe", "ABC")])
    assert (classee.par, classee.ecole) == ("préfixe", "ABC")


def test_T_CLA_23_egalite_au_plus_long_prefixe():
    """T-CLA-23 — Deux préfixes de même longueur au plus long niveau : « ambiguë » ; une égalité
    entre préfixes plus courts ne compte pas, le plus long gagne."""
    egalite_au_plus_long = [
        _ligne("Cours", "préfixe", "ABC"),
        _ligne("Cours fictif", "préfixe", "DEF"),
        _ligne("cours fictif ", "préfixe", "XYZ"),
    ]
    assert _raison(_seance("Cours fictif — TD 1"), egalite_au_plus_long) == "ambiguë"

    egalite_plus_courte = [
        _ligne("Cours", "préfixe", "ABC"),
        _ligne("COURS", "préfixe", "DEF"),
        _ligne("Cours fictif", "préfixe", "XYZ"),
    ]
    classee = _classee(_seance("Cours fictif — TD 1"), egalite_plus_courte)
    assert (classee.par, classee.ecole) == ("préfixe", "XYZ")


# ──────────────────────────────────────────────────────────────────────
# Valeurs vides, types de clé, lignes invalides
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    "intitule",
    ["", "   ", "Réunion pédagogique fictive"],
    ids=["intitule-vide", "intitule-espaces", "intitule-sans-regle"],
)
def test_T_CLA_24_non_reconnue(intitule):
    """T-CLA-24 — Sans code, sans intitulé exact ni préfixe qui convienne (intitulé vide compris) :
    file à classer, raison « non reconnue » ; la séance n'est jamais classée."""
    lignes = [
        _ligne("#ABC-M1-STATS", "code", "ABC"),
        _ligne("Cours fictif", "exact", "ABC"),
        _ligne("Cours", "préfixe", "DEF"),
    ]
    assert _raison(_seance(intitule, description=""), lignes) == "non reconnue"


def test_T_CLA_25_type_de_cle_sans_casse_ni_espaces():
    """T-CLA-25 — Le type de clé se lit sans tenir compte de la casse ni des espaces autour ; `par`
    rend toujours la forme du contrat (« code », « exact », « préfixe »).
    💡 [proposé — à valider]."""
    lignes = [
        _ligne("#ABC-M1-STATS", "CODE", "ABC"),
        _ligne("Cours fictif d'économie", " Exact ", "DEF"),
        _ligne("Atelier fictif", "PRÉFIXE ", "XYZ"),
    ]
    assert _classee(_seance(description="#ABC-M1-STATS"), lignes).par == "code"
    assert _classee(_seance("Cours fictif d'économie"), lignes).par == "exact"
    assert _classee(_seance("Atelier fictif — groupe B"), lignes).par == "préfixe"


def test_T_CLA_26_type_de_cle_inconnu():
    """T-CLA-26 — Une ligne d'un type de clé inconnu lève `ValueError`, dont le message cite le
    type fautif et la clé de la ligne. 💡 [proposé — à valider]."""
    lignes = [_ligne("Cours fictif", "exact", "ABC"), _ligne("Atelier fictif", "regex", "DEF")]
    with pytest.raises(ValueError) as exc:
        _module().classer([_seance("Séance quelconque")], lignes)
    assert "regex" in _message(exc)
    assert "atelier fictif" in _message(exc)


@pytest.mark.parametrize("type_cle", [None, "", "   "], ids=["none", "vide", "espaces"])
def test_T_CLA_27_type_de_cle_vide(type_cle):
    """T-CLA-27 — Une ligne sans type de clé lève `ValueError`, dont le message nomme le type
    manquant et la clé de la ligne. 💡 [proposé — à valider]."""
    lignes = [_ligne("Cours fictif", "exact", "ABC"), _ligne("Atelier fictif", type_cle, "DEF")]
    with pytest.raises(ValueError) as exc:
        _module().classer([_seance("Séance quelconque")], lignes)
    assert "type" in _message(exc)
    assert "atelier fictif" in _message(exc)


@pytest.mark.parametrize(
    ("type_cle", "cle"),
    [("préfixe", None), ("préfixe", ""), ("préfixe", "   "), ("exact", ""), ("code", None)],
    ids=["prefixe-none", "prefixe-vide", "prefixe-espaces", "exact-vide", "code-none"],
)
def test_T_CLA_28_cle_vide(type_cle, cle):
    """T-CLA-28 — Une ligne sans clé lève `ValueError`, dont le message nomme la clé manquante :
    une clé vide de type `préfixe` reconnaîtrait toutes les séances. 💡 [proposé — à valider]."""
    lignes = [_ligne(cle, type_cle, "DEF")]
    with pytest.raises(ValueError) as exc:
        _module().classer([_seance("Cours fictif", description="#ABC-M1-STATS")], lignes)
    # « sans clé », et non « clé » seul, que tout message citant le « type de clé » contiendrait.
    assert "sans clé" in _message(exc)


def test_T_CLA_29_valeurs_de_la_ligne_retenue():
    """T-CLA-29 — École, module, promotion et clé tarif sont rendus sans espaces autour ; une
    cellule vide donne une chaîne vide ; `correspondance` reste la ligne telle que lue.
    💡 [proposé — à valider]."""
    ligne = _ligne(
        "Cours fictif", "exact", " ABC ", module="Module fictif ", promotion=None,
        cle_tarif=" ABC|Module fictif|M1",
    )
    lue = copy.deepcopy(ligne)
    classee = _classee(_seance("Cours fictif"), [ligne])
    assert (classee.ecole, classee.module, classee.promotion, classee.cle_tarif) == (
        "ABC", "Module fictif", "", "ABC|Module fictif|M1",
    )
    assert classee.correspondance == lue


# ──────────────────────────────────────────────────────────────────────
# Critère d'acceptation F-12 (ROADMAP) — toutes les séances classées ou à classer
# ──────────────────────────────────────────────────────────────────────


def test_T_CLA_30_toutes_les_seances_classees_ou_a_classer():
    """T-CLA-30 — Critère F-12 : 100 % des séances de la période sont classées ou mises dans la
    file à classer, chacune une seule fois, dans l'ordre reçu ; aucune séance non reconnue,
    ambiguë ou de code inconnu n'est classée (donc facturée)."""
    lignes = [
        _ligne("#ABC-M1-STATS", "code", "ABC"),
        _ligne("Cours fictif d'économie", "exact", "DEF"),
        _ligne("Atelier fictif", "préfixe", "XYZ"),
        _ligne("Cours en double", "exact", "ABC"),
        _ligne("Cours en double", "exact", "DEF"),
    ]
    s1 = _seance("Statistiques", description="#ABC-M1-STATS", jour=date(2026, 11, 2))
    s2 = _seance("Cours fictif d'économie", jour=date(2026, 11, 3))
    s3 = _seance("Réunion de rentrée fictive", jour=date(2026, 11, 4))
    s4 = _seance("Atelier fictif — groupe A", jour=date(2026, 11, 5))
    s5 = _seance("Cours fictif d'économie", description="#ZZZ-X9-INCONNU", jour=date(2026, 11, 6))
    s6 = _seance("Cours en double", jour=date(2026, 11, 9))
    s7 = _seance("Atelier fictif — groupe B", jour=date(2026, 11, 10))
    seances = [s1, s2, s3, s4, s5, s6, s7]

    classees, a_classer = _module().classer(seances, lignes)
    assert [(c.seance, c.par, c.ecole) for c in classees] == [
        (s1, "code", "ABC"), (s2, "exact", "DEF"), (s4, "préfixe", "XYZ"), (s7, "préfixe", "XYZ"),
    ]
    assert [(a.seance, a.raison) for a in a_classer] == [
        (s3, "non reconnue"), (s5, "code inconnu"), (s6, "ambiguë"),
    ]
    uids = [c.seance.uid for c in classees] + [a.seance.uid for a in a_classer]
    assert sorted(uids) == sorted(s.uid for s in seances)


def test_T_CLA_31_classer_ne_modifie_pas_ses_entrees():
    """T-CLA-31 — `classer` calcule sans rien écrire : la liste des séances et les lignes de
    Correspondances sont rendues intactes."""
    lignes = [
        _ligne("#ABC-M1-STATS", "code", "ABC"),
        _ligne(" Cours fictif ", " Exact ", " DEF ", promotion=None, cle_tarif="DEF|Module fictif|"),
        _ligne("Atelier", "préfixe", "XYZ"),
    ]
    seances = [
        _seance(description="#ABC-M1-STATS"),
        _seance("Cours fictif"),
        _seance("Atelier fictif"),
        _seance("Réunion fictive"),
    ]
    lignes_avant, seances_avant = copy.deepcopy(lignes), list(seances)
    _module().classer(seances, lignes)
    assert lignes == lignes_avant
    assert seances == seances_avant


# ──────────────────────────────────────────────────────────────────────
# ligne_correspondance — la règle proposée quand Adrian classe une séance (ADR028)
# ──────────────────────────────────────────────────────────────────────


def test_T_CLA_32_ligne_exacte_par_defaut():
    """T-CLA-32 — Par défaut, la ligne proposée est de type `exact`, de clé l'intitulé sans
    espaces autour, mode d'heures `agenda`, sans valeur, gardée à la date donnée ; elle porte
    toutes les colonnes de Correspondances (MASTER_SCHEMA)."""
    seance = _seance("  Cours fictif d'économie ", description="Salle 12")
    ligne = _module().ligne_correspondance(
        seance, ecole="DEF", module="Économie fictive", promotion="M2",
        cle_tarif="DEF|Économie fictive|M2", garde_le=GARDEE_LE,
    )
    assert isinstance(ligne, dict)
    assert set(ligne) == set(COLONNES_CORRESPONDANCES)
    assert ligne["Clé"] == "Cours fictif d'économie"
    assert ligne["Type de clé"] == "exact"
    assert (ligne["École"], ligne["Module"], ligne["Promotion"], ligne["Clé tarif"]) == (
        "DEF", "Économie fictive", "M2", "DEF|Économie fictive|M2",
    )
    assert ligne["Mode heures"] == "agenda"
    assert ligne["Valeur"] is None
    assert ligne["Gardée le"] == GARDEE_LE


def test_T_CLA_33_arguments_nommes_et_date_obligatoire():
    """T-CLA-33 — Les arguments après la séance sont nommés, et la date `garde_le` est
    obligatoire (la ligne dit quand Adrian a gardé la règle)."""
    m = _module()
    seance = _seance("Atelier fictif")
    with pytest.raises(TypeError):
        m.ligne_correspondance(
            seance, ecole="XYZ", module="Atelier fictif", promotion="L3", cle_tarif="XYZ|Atelier fictif|L3",
        )
    with pytest.raises(TypeError):
        m.ligne_correspondance(seance, "XYZ", "Atelier fictif", "L3", "XYZ|Atelier fictif|L3", garde_le=GARDEE_LE)


def test_T_CLA_34_ligne_de_type_code():
    """T-CLA-34 — De type `code`, la clé est le premier code de la description ; mode d'heures
    et valeur sont repris tels quels."""
    seance = _seance("Atelier fictif", description="Salle 12 — #XYZ-L3-ATELIER (puis #DEF-M2-ECO)")
    ligne = _ligne_proposee(seance, type_cle="code", mode_heures="forfait", valeur=7.5)
    assert ligne["Clé"] == "#XYZ-L3-ATELIER"
    assert ligne["Type de clé"] == "code"
    assert (ligne["Mode heures"], ligne["Valeur"]) == ("forfait", 7.5)


def test_T_CLA_35_ligne_de_type_prefixe():
    """T-CLA-35 — De type `préfixe`, la clé est l'intitulé sans espaces autour.
    💡 [proposé — à valider] : la signature ne prend pas de clé, l'intitulé est la seule source."""
    ligne = _ligne_proposee(_seance(" Atelier fictif "), type_cle="préfixe", mode_heures="coefficient", valeur=0.875)
    assert (ligne["Clé"], ligne["Type de clé"]) == ("Atelier fictif", "préfixe")
    assert (ligne["Mode heures"], ligne["Valeur"]) == ("coefficient", 0.875)


@pytest.mark.parametrize(
    ("type_cle", "octobre", "novembre", "raison_avant"),
    [
        ("exact", ("Atelier fictif de rentrée", ""), ("Atelier fictif de rentrée", "Salle 4"), "non reconnue"),
        ("code", ("Atelier", "#XYZ-L3-ATELIER"), ("Atelier — suite", "Rappel : #XYZ-L3-ATELIER"), "code inconnu"),
        ("préfixe", ("Atelier fictif", ""), ("Atelier fictif — suite", ""), "non reconnue"),
    ],
    ids=["exact", "code", "prefixe"],
)
def test_T_CLA_36_regle_gardee_reconnue_au_lancement_suivant(type_cle, octobre, novembre, raison_avant):
    """T-CLA-36 — Critère F-12 (ADR028) : une séance mise à classer, classée par Adrian et gardée
    comme règle, est reconnue au lancement suivant ; sans règle gardée (correction ponctuelle),
    elle reste à classer."""
    existantes = [_ligne("Cours fictif", "exact", "ABC")]
    seance_octobre = _seance(*octobre, jour=date(2026, 10, 6))
    seance_novembre = _seance(*novembre, jour=date(2026, 11, 3))
    assert _raison(seance_octobre, existantes) == raison_avant

    regle = _ligne_proposee(seance_octobre, type_cle=type_cle)

    assert _raison(seance_novembre, existantes) == raison_avant  # correction ponctuelle : non gardée
    classee = _classee(seance_novembre, existantes + [regle])
    assert classee.par == type_cle
    assert (classee.ecole, classee.module, classee.promotion, classee.cle_tarif) == (
        "XYZ", "Atelier fictif", "L3", "XYZ|Atelier fictif|L3",
    )
    assert classee.correspondance == regle


@pytest.mark.parametrize(
    ("options", "cite"),
    [({"type_cle": "regex"}, "regex"), ({"mode_heures": "horaire"}, "horaire")],
    ids=["type-de-cle-inconnu", "mode-heures-inconnu"],
)
def test_T_CLA_37_type_ou_mode_inconnu(options, cite):
    """T-CLA-37 — Un type de clé ou un mode d'heures hors du contrat lève `ValueError`, dont le
    message cite la valeur fautive. 💡 [proposé — à valider]."""
    with pytest.raises(ValueError) as exc:
        _ligne_proposee(_seance("Atelier fictif", description="#XYZ-L3-ATELIER"), **options)
    assert cite in _message(exc)


@pytest.mark.parametrize(
    ("type_cle", "intitule", "description", "cause"),
    [
        # « aucun code », et non « code » seul, que le rappel du type « code » contiendrait.
        ("code", "Atelier fictif", "Salle 12", "aucun code"),
        ("code", "Atelier fictif", "#visio", "aucun code"),
        ("exact", "   ", "", "intitulé"),
        ("préfixe", "", "#XYZ-L3-ATELIER", "intitulé"),
    ],
    ids=["code-absent", "code-en-minuscules", "exact-intitule-vide", "prefixe-intitule-vide"],
)
def test_T_CLA_38_cle_introuvable_dans_la_seance(type_cle, intitule, description, cause):
    """T-CLA-38 — Pas de ligne sans clé : un type `code` sans code dans la description, ou un
    type `exact` ou `préfixe` sur un intitulé vide, lève `ValueError`, dont le message nomme ce
    qui manque. 💡 [proposé — à valider] : `classer` refuserait cette ligne (T-CLA-28)."""
    with pytest.raises(ValueError) as exc:
        _ligne_proposee(_seance(intitule, description=description), type_cle=type_cle)
    assert cause in _message(exc)


@pytest.mark.parametrize("mode", ["coefficient", "forfait"])
def test_T_CLA_39_valeur_obligatoire_en_coefficient_ou_forfait(mode):
    """T-CLA-39 — En mode `coefficient` ou `forfait`, une valeur absente lève `ValueError`, dont le
    message cite le mode (ne jamais inventer un volume horaire ; F-13 la refuserait).
    💡 [proposé — à valider]."""
    with pytest.raises(ValueError) as exc:
        _ligne_proposee(_seance("Atelier fictif"), mode_heures=mode)
    assert mode in _message(exc)
    assert "sans valeur" in _message(exc)  # le mode seul passerait aussi sur « mode inconnu »


@pytest.mark.parametrize(
    ("options", "cause"),
    [({"ecole": ""}, "école"), ({"ecole": "   "}, "école"), ({"cle_tarif": ""}, "tarif")],
    ids=["ecole-vide", "ecole-espaces", "cle-tarif-vide"],
)
def test_T_CLA_40_ecole_et_cle_tarif_obligatoires(options, cause):
    """T-CLA-40 — Une règle sans école ou sans clé tarif lève `ValueError`, dont le message nomme
    le champ manquant (ne jamais inventer un taux). 💡 [proposé — à valider]."""
    with pytest.raises(ValueError) as exc:
        _ligne_proposee(_seance("Atelier fictif"), **options)
    assert cause in _message(exc)


# ──────────────────────────────────────────────────────────────────────
# Constats de la revue croisée de F-12 (T-CLA-41 à T-CLA-52)
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    ("intitule", "description", "existantes", "type_cle", "code"),
    [
        (
            "Statistiques fictives", "#ABC-M1-STATS",
            [_ligne("#ABC-M1-STATS", "code", "ABC", module="Statistiques fictives")],
            "exact", "#ABC-M1-STATS",
        ),
        ("Atelier fictif", "#XYZ-L3-ATELIER", [_ligne("Cours fictif", "exact", "ABC")], "exact", "#XYZ-L3-ATELIER"),
        ("Atelier fictif", "#XYZ-L3-ATELIER", [_ligne("Cours fictif", "exact", "ABC")], "préfixe", "#XYZ-L3-ATELIER"),
    ],
    ids=["code-connu-corrige-en-exact", "code-inconnu-garde-en-exact", "code-inconnu-garde-en-prefixe"],
)
def test_T_CLA_41_regle_sans_code_sur_une_seance_qui_porte_un_code(intitule, description, existantes, type_cle, code):
    """T-CLA-41 — Constat de revue (critère F-12, ADR028 : une règle gardée s'applique au
    lancement suivant). Une séance dont la description porte un code est classée au niveau du
    code, sans repli sur l'intitulé (T-CLA-05, T-CLA-06) : une règle `exact` ou `préfixe`
    gardée sur elle ne s'appliquerait jamais. Aujourd'hui, la règle est acceptée sans un mot et
    la séance reste classée par le code (mauvaise école facturée) ou « code inconnu ».
    Deux issues admises, 💡 [proposé — à valider] : `ligne_correspondance` refuse la règle par
    une `ValueError` qui cite le code (choisir le type `code`, ou corriger la ligne du code), ou
    la règle classe bien la séance au lancement suivant."""
    seance_octobre = _seance(intitule, description=description, jour=date(2026, 10, 6))
    seance_novembre = _seance(intitule, description=description, jour=date(2026, 11, 3))
    try:
        regle = _ligne_proposee(seance_octobre, type_cle=type_cle)
    except ValueError as exc:
        assert code.casefold() in str(exc).casefold()
        return
    classee = _classee(seance_novembre, existantes + [regle])
    assert (classee.ecole, classee.cle_tarif) == ("XYZ", "XYZ|Atelier fictif|L3")


@pytest.mark.parametrize(
    ("description", "ecole_du_code_tronque"),
    [
        ("#ABC-M1-STATS_2 (ancien : #DEF-M2-ECO)", "ABC"),
        ("#ABC-M1-STATSbis puis #DEF-M2-ECO", "ABC"),
        ("#DEF-M2-Eco", None),
        ("#DEF-M2-ECObis", "DEF"),
    ],
    ids=["souligne-puis-ancien-code", "minuscules-puis-ancien-code", "dernier-segment-en-minuscules",
         "suffixe-en-minuscules"],
)
def test_T_CLA_42_code_prolonge_jamais_ignore(description, ecole_du_code_tronque):
    """T-CLA-42 — Constat de revue (ADR016 : un code écrit dans l'événement fait foi ; contrat :
    code fait de lettres majuscules, chiffres et tirets, le premier code trouvé compte). Un code
    prolongé par d'autres caractères n'est jamais ignoré : la séance n'est classée ni par un code
    écrit plus loin (l'ancien code), ni par son intitulé. Aujourd'hui, le jeton entier est écarté
    sans bruit et la séance part vers DEF par l'ancien code, ou vers XYZ par l'intitulé.
    Deux lectures admises, à trancher 💡 [proposé — à valider] : le code se lit jusqu'au premier
    caractère hors majuscules, chiffres et tirets (docstring de T-CLA-04 : la séance est classée
    par ce code tronqué, s'il est connu) ; ou le jeton n'est pas un code valide et la séance est
    mise à classer (docstring du module : un code « se lit en entier, jamais tronqué »)."""
    lignes = [
        _ligne("#ABC-M1-STATS", "code", "ABC"),
        _ligne("#DEF-M2-ECO", "code", "DEF"),
        _ligne("Cours fictif", "exact", "XYZ"),
    ]
    seance = _seance("Cours fictif", description=description)
    classees, a_classer = _module().classer([seance], lignes)
    if a_classer:
        assert [a.seance for a in a_classer] == [seance]
        return
    assert len(classees) == 1
    assert (classees[0].par, classees[0].ecole) == ("code", ecole_du_code_tronque), (
        "ni l'ancien code, ni l'intitulé ne classent une séance dont le premier code est prolongé"
    )


@pytest.mark.parametrize(
    ("ecole", "cle_tarif", "cause"),
    [
        (None, "ABC|Module fictif|M1", "école"),
        ("   ", "ABC|Module fictif|M1", "école"),
        ("ABC", "", "tarif"),
        ("ABC", "   ", "tarif"),
    ],
    ids=["ecole-none", "ecole-espaces", "cle-tarif-vide", "cle-tarif-espaces"],
)
def test_T_CLA_43_ligne_sans_ecole_ou_sans_cle_tarif(ecole, cle_tarif, cause):
    """T-CLA-43 — Constat de revue. Une ligne de Correspondances sans école ou sans clé tarif lève
    `ValueError`, dont le message nomme le champ manquant et la clé de la ligne, comme
    `ligne_correspondance` le fait pour une règle gardée (T-CLA-40) : ne jamais inventer un
    taux. Aujourd'hui, la séance est classée avec une école ou une clé tarif vide, et l'erreur
    n'apparaît qu'en aval (F-13, F-16), loin de la ligne en cause.
    💡 [proposé — à valider] : une ligne invalide lève `ValueError`, comme en T-CLA-26 à T-CLA-28."""
    lignes = [_ligne("Cours fictif", "exact", ecole, cle_tarif=cle_tarif)]
    with pytest.raises(ValueError) as exc:
        _module().classer([_seance("Cours fictif")], lignes)
    assert cause in _message(exc)
    assert "cours fictif" in _message(exc)


@pytest.mark.parametrize(
    ("type_cle", "intitule", "description", "existantes", "raison_avant"),
    [
        (
            "exact", "Atelier fictif de rentrée", "",
            [_ligne("Atelier fictif de rentrée", "exact", "ABC"), _ligne("Atelier fictif de rentrée", "exact", "DEF")],
            "ambiguë",
        ),
        ("exact", "Atelier fictif de rentrée", "", [_ligne("Atelier fictif de rentrée", "exact", "ABC")], None),
        (
            "code", "Atelier fictif", "#XYZ-L3-ATELIER",
            [_ligne("#XYZ-L3-ATELIER", "code", "ABC"), _ligne("#XYZ-L3-ATELIER", "code", "DEF")],
            "ambiguë",
        ),
        ("code", "Atelier fictif", "#XYZ-L3-ATELIER", [_ligne("#XYZ-L3-ATELIER", "code", "ABC")], None),
    ],
    ids=["exact-ambigue", "exact-corrigee", "code-ambigue", "code-corrigee"],
)
def test_T_CLA_44_regle_gardee_au_meme_niveau_qu_une_ligne_existante(
    type_cle, intitule, description, existantes, raison_avant
):
    """T-CLA-44 — Constat de revue (critère F-12, ADR028 : quand Adrian classe *ou corrige* une
    séance, la règle gardée s'applique au lancement suivant). Une séance ambiguë, ou classée par
    une ligne qu'Adrian corrige vers XYZ, reçoit une règle gardée de même type et de même clé :
    au lancement suivant, la règle gardée la classe. Aujourd'hui, la ligne ajoutée rend la
    séance « ambiguë » à chaque lancement (rien n'est facturé à tort, mais la règle ne sert pas).
    💡 [proposé — à valider] : trou du contrat, à trancher par Adrian. `ligne_correspondance`
    rend « la ligne à ajouter » sans voir les lignes existantes, et deux lignes au même niveau
    font une ambiguïté. Le test ajoute la ligne, seule opération du contrat actuel ; à adapter si
    la décision est de remplacer la ligne existante ou de demander de la modifier."""
    seance_octobre = _seance(intitule, description=description, jour=date(2026, 10, 6))
    seance_novembre = _seance(intitule, description=description, jour=date(2026, 11, 3))
    if raison_avant:
        assert _raison(seance_octobre, existantes) == raison_avant
    else:
        avant = _classee(seance_octobre, existantes)
        assert (avant.par, avant.ecole) == (type_cle, "ABC")

    regle = _ligne_proposee(seance_octobre, type_cle=type_cle)

    classee = _classee(seance_novembre, existantes + [regle])
    assert (classee.par, classee.ecole, classee.cle_tarif) == (type_cle, "XYZ", "XYZ|Atelier fictif|L3")


@pytest.mark.parametrize(
    "cle",
    ["ABC-M1-STATS", "#visio", "Statistiques fictives"],
    ids=["diese-oublie", "mot-diese", "intitule"],
)
def test_T_CLA_45_cle_de_type_code_sans_forme_de_code(cle):
    """T-CLA-45 — Constat de revue. Une ligne de type `code` dont la clé n'a pas la forme
    `#ÉCOLE-PROMO-MODULE` (dièse oublié, mot-dièse, intitulé) ne pourrait jamais servir : elle
    lève `ValueError`, dont le message cite la clé et la forme attendue (ou le dièse).
    Aujourd'hui, elle est acceptée sans un mot et la séance qui porte le code part en « code
    inconnu », alors qu'Adrian voit la ligne dans le MASTER.
    💡 [proposé — à valider] : une clé de code en minuscules reste admise (T-CLA-10), la forme
    se contrôle donc sans tenir compte de la casse."""
    with pytest.raises(ValueError) as exc:
        _module().classer([_seance(description="#ABC-M1-STATS")], [_ligne(cle, "code", "ABC")])
    assert cle.casefold() in _message(exc)
    assert "#école-promo-module" in _message(exc) or "dièse" in _message(exc)


@pytest.mark.parametrize(
    ("cle", "intitule"),
    [
        ("Statistiques M1", "Statistiques M12 — TD 1"),
        ("Stat", "Statut de la réunion"),
        ("Cours", "Courses d'orientation fictives"),
    ],
    ids=["autre-promotion", "reunion", "autre-mot"],
)
def test_T_CLA_46_prefixe_suivi_d_une_frontiere_de_mot(cle, intitule):
    """T-CLA-46 — Constat de revue (ADR028 : ne pas facturer une réunion ; critère F-12 : aucune
    séance facturée par erreur). Un préfixe ne reconnaît un intitulé que s'il y est suivi de la
    fin du texte ou d'un caractère qui n'est ni une lettre ni un chiffre : « Statistiques M1 » ne
    classe pas « Statistiques M12 » au tarif M1, « Stat » ne classe pas une réunion. Aujourd'hui,
    le préfixe se compare caractère par caractère et ces séances sont classées.
    💡 [proposé — à valider] : le contrat dit « préfixe » sans préciser la frontière de mot."""
    lignes = [_ligne(cle, "préfixe", "ABC")]
    assert _raison(_seance(intitule), lignes) == "non reconnue"
    assert _classee(_seance(f"{cle} — TD 1"), lignes).par == "préfixe"
    assert _classee(_seance(f"{cle}, groupe B"), lignes).par == "préfixe"


@pytest.mark.parametrize(
    ("ligne", "repere"),
    [
        ({"Clé": None, "Type de clé": None, "École": None, "Notes": "à revoir"}, "à revoir"),
        ({"Clé": "  ", "Module": "Module fictif", "Clé tarif": "ABC|Module fictif|M1"}, "abc|module fictif|m1"),
    ],
    ids=["notes-seules", "cle-tarif"],
)
def test_T_CLA_47_ligne_sans_cle_reperable(ligne, repere):
    """T-CLA-47 — Constat de revue. Une ligne sans clé qui ne porte ni type de clé ni école (une
    ligne de commentaire, par exemple) lève `ValueError`, dont le message donne de quoi la
    retrouver : ses notes ou sa clé tarif. Aujourd'hui, le message ne cite rien et la ligne
    fautive bloque tout le classement sans qu'on puisse la retrouver.
    💡 [proposé — à valider] : `master.py` ne transmet pas de numéro de ligne."""
    lignes = [_ligne("Cours fictif", "exact", "ABC"), ligne]
    with pytest.raises(ValueError) as exc:
        _module().classer([_seance("Cours fictif")], lignes)
    assert "sans clé" in _message(exc)
    assert repere in _message(exc)


@pytest.mark.parametrize("garde_le", [None, "hier", "26/11/2026"], ids=["none", "texte", "date-en-texte"])
def test_T_CLA_48_date_gardee_le_controlee(garde_le):
    """T-CLA-48 — Constat de revue. `garde_le` doit être une date : la colonne « Gardée le » du
    MASTER ne reçoit ni texte ni cellule vide. Aujourd'hui, n'importe quelle valeur est
    recopiée telle quelle. 💡 [proposé — à valider] : `TypeError` ou `ValueError`, dont le
    message nomme le champ."""
    with pytest.raises((TypeError, ValueError)) as exc:
        _ligne_proposee(_seance("Atelier fictif"), garde_le=garde_le)
    assert "gardée le" in _message(exc) or "garde_le" in _message(exc)


@pytest.mark.parametrize(
    ("mode", "valeur"),
    [
        ("forfait", -1),
        ("coefficient", -0.5),
        ("forfait", float("nan")),
        ("coefficient", float("inf")),
        ("forfait", True),
        ("coefficient", "0.875"),
        ("agenda", 1.0),
    ],
    ids=["forfait-negatif", "coefficient-negatif", "nan", "infini", "booleen", "texte", "valeur-en-mode-agenda"],
)
def test_T_CLA_49_valeur_du_mode_d_heures_controlee(mode, valeur):
    """T-CLA-49 — Constat de revue (couverture : contrôles annoncés par la docstring mais non
    testés). Une valeur négative, non finie, booléenne ou en texte en mode `coefficient` ou
    `forfait`, ou une valeur donnée en mode `agenda` qui l'ignorerait, lève `ValueError`, dont le
    message cite le mode (contrat F-13 : valeur négative refusée)."""
    with pytest.raises(ValueError) as exc:
        _ligne_proposee(_seance("Atelier fictif"), mode_heures=mode, valeur=valeur)
    assert mode in _message(exc)


@pytest.mark.parametrize("mode_heures", ["", "   ", None, " Agenda "], ids=["vide", "espaces", "none", "casse"])
def test_T_CLA_50_mode_d_heures_vide_vaut_agenda(mode_heures):
    """T-CLA-50 — Constat de revue (couverture). Un mode d'heures vide vaut `agenda` (contrat
    F-13 : « cellule vide comprise ») et se lit sans tenir compte de la casse ni des espaces."""
    ligne = _ligne_proposee(_seance("Atelier fictif"), mode_heures=mode_heures)
    assert (ligne["Mode heures"], ligne["Valeur"]) == ("agenda", None)


@pytest.mark.parametrize(
    ("description", "par", "ecole"),
    [
        ("#VISIO", "exact", "DEF"),
        ("Rendu le #2026-11-03", "exact", "DEF"),
        ("page#ABC-M1-STATS", "exact", "DEF"),
        ("Code : #ABC-M1-STATS-", "code", "ABC"),
    ],
    ids=["un-seul-segment", "premier-segment-numerique", "diese-colle-a-un-mot", "tiret-final"],
)
def test_T_CLA_51_formes_de_code_ecartees_ou_admises(description, par, ecole):
    """T-CLA-51 — Constat de revue (couverture des formes écrites dans la docstring du module).
    Un seul segment, un premier segment numérique ou un dièse collé à un mot ne font pas un code :
    la séance est rapprochée par son intitulé. Un tiret final est de la ponctuation : le code
    qui le précède est lu."""
    lignes = [_ligne("#ABC-M1-STATS", "code", "ABC"), _ligne("Cours fictif", "exact", "DEF")]
    classee = _classee(_seance("Cours fictif", description=description), lignes)
    assert (classee.par, classee.ecole) == (par, ecole)


def test_T_CLA_52_normalisation_unicode_et_prefixe_sans_accent():
    """T-CLA-52 — Constat de revue (couverture). Un intitulé ou une clé en Unicode décomposé (NFD,
    fréquent dans un export iCal) se compare comme sa forme composée (NFC), dans les deux sens ;
    le type de clé « prefixe » sans accent vaut « préfixe »."""
    nfc = unicodedata.normalize("NFC", "Économétrie fictive")
    nfd = unicodedata.normalize("NFD", nfc)
    assert nfc != nfd
    assert _classee(_seance(nfd), [_ligne(nfc, "exact", "ABC")]).par == "exact"
    assert _classee(_seance(nfc), [_ligne(nfd, "exact", "ABC")]).par == "exact"
    assert _classee(_seance("Atelier fictif — groupe A"), [_ligne("Atelier", "prefixe", "ABC")]).par == "préfixe"
