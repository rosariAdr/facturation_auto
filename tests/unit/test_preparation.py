"""Préparation multi-écoles des brouillons : logique (F-16a, `facturation/preparation.py` et
`facturation/decisions.py` ; banque de tests : T-PRP-01 à T-PRP-37, T-PRP-45 à T-PRP-47 et, pour les
constats de revue de F-16a, T-PRP-48 à T-PRP-60 ; les autres dans `tests/integration/test_preparation.py`).

Tests d'acceptation écrits avant l'implémentation : chaque test importe `facturation.preparation`
(ou `facturation.decisions`) lui-même, pour que l'absence du module fasse échouer les tests sans
bloquer la collecte des autres. Données fictives créées à la volée (`tests/aide_preparation.py`) ;
l'agenda est un contenu .ics fictif injecté, jamais téléchargé : tout accès réseau fait échouer le test.

Contrat (contrats des vagues 3 à 5, § 1 et 4, avec les § 15 et 16, qui priment) :
- `charger_entrees` lit le MASTER, les coordonnées (une coordonnée manquante devient
  `erreur_coordonnees`, jamais une erreur globale), l'agenda de la fenêtre de lecture, les décisions
  actives du mois, les pièces d'émission et `Sessions` ; il n'écrit rien ;
- `preparer` est pure et déterministe : un brouillon par école du lot (ou par contrat), des blocages
  cumulés sans rien lever, les créneaux non reconnus jamais facturés, l'anti-doublon contre
  `Sessions`, les brouillons engagés jamais régénérés ; une erreur propre à une école ne retient
  jamais les autres ;
- montants exacts (`Decimal`), total arrondi une seule fois au centime ; empreinte sans `tenue`,
  `numero_probable`, avertissements, blocages ni date de préparation ;
- § 16 (Q-B1) : une école peut avoir une ligne de `Règles_facturation` par contrat, chacune avec son
  `N° de contrat` ; le n° imprimé vient de cette ligne, `Tarifs.Réf. contrat` y relie la ligne de tarif.

Valeurs décidées par l'orchestrateur (formulaire d'Adrian sans réponse, valeurs « Recommandé ») :
aperçu en texte seulement (Q-D7) ; « 45 jours fin de mois » = date de facture + 45 jours, puis fin de
ce mois-là (Q-A4) ; l'école à cycle préparée 5 jours avant la date limite relève de F-20 (Q-A1).

Lectures retenues (`💡 [proposé — à valider]`) :
- `EntreesPreparation.regles` et `fenetre_de_lecture` reçoivent le dictionnaire que rend
  `regles.regles_par_ecole` (clé (école, n° de contrat) depuis le § 16) ;
- une décision `classer` porte ses valeurs en texte (`valeur` vide en mode « agenda ») ; une décision
  `bon_commande` porte sa référence dans `valeurs["bon_commande"]` ; `exclure` porte `valeurs["raison"]` ;
- un brouillon engagé (pièce « en émission » ou « émise ») n'est plus dans `Preparation.brouillons` :
  il est listé dans `deja_emises` ;
- une décision `inclure` ne réintègre pas une séance retrouvée exactement dans `Sessions` ;
- les numéros courts `S<n>` couvrent aussi les séances déjà facturées (cible de `corriger --inclure S12`) ;
- suffixe d'un identifiant par contrat : référence en majuscules réduite à `[A-Z0-9]`, 20 caractères ;
  en cas de collision, chaque identifiant concerné reçoit `-` et les 4 premiers caractères
  hexadécimaux du sha256 (UTF-8) de sa référence ;
- « période déjà facturée hors outil » : `Factures.Période` reconnue sous le libellé de l'onglet
  (« Déc 2026 », T-PRP-14) et sous la forme longue de l'historique (« Décembre 2026 », § 15, T-PRP-36) ;
- ajouts de l'orchestrateur : `ecole_label` vaut le code quand l'école n'a pas de libellé (T-PRP-21,
  § 18 hors périmètre) ; une séance et une ligne `Sessions` de la même école le même jour à d'autres
  heures donnent l'avertissement « possible doublon avec Sessions, ligne N », dans « À FAIRE », sans
  exclusion ni blocage (T-PRP-22, constat écarté de la revue de P-0b).
"""

from __future__ import annotations

import importlib
import json
import re
import sys
import types
import urllib.request
from datetime import date, datetime
from decimal import Decimal

import pytest

from facturation import pcf_format as fmt
from tests.aide_preparation import (
    BIC_FICTIF,
    IBAN_FICTIF,
    MAINTENANT,
    MOIS,
    PARIS,
    POSTE_TEST,
    SEANCES_STANDARD,
    SIRET_FICTIF,
    SeanceFictive,
    arborescence,
    bloc,
    brouillon,
    changer_coordonnee,
    changer_info,
    cle,
    cles_facturees,
    correspondance,
    correspondances_standard,
    ecoles_standard,
    ecrire_decision,
    ecrire_piece,
    entrees,
    ics,
    identifiants,
    installer,
    ligne_sessions,
    master_standard,
    preparer,
    prp,
    regle,
    regles_standard,
    retirer_info,
    seances_sauf,
    standard,
    tarif,
    tarifs_standard,
    textes,
    uids_factures,
)
from tests.conftest import NDA_FICTIF

pytestmark = pytest.mark.unit


@pytest.fixture(autouse=True)
def _isolement(monkeypatch):
    """Aucun accès réseau (l'agenda est injecté) ; racine de `pcf_format` restaurée après le test."""

    def refuser(*args, **kwargs):
        raise AssertionError("accès réseau interdit dans les tests : l'agenda est injecté (contenu_ics)")

    monkeypatch.setattr(urllib.request, "urlopen", refuser)
    ancienne = fmt._DATA_ROOT
    yield
    fmt._DATA_ROOT = ancienne
    fmt._MASTER_DATA_CACHE.clear()


def _par_uid(b) -> dict:
    return {ligne.uid: ligne for ligne in b.lignes}


def _groupes(b) -> list:
    """`Brouillon.groupes()` sous une forme comparable (listes et tuples normalisés)."""
    return [(promo, libelle, [tuple(s) for s in seances]) for promo, libelle, seances in b.groupes()]


# ──────────────────────────────────────────────────────────────────────
# T-PRP-01 à T-PRP-03 : brouillons, contrats, montants
# ──────────────────────────────────────────────────────────────────────


def test_t_prp_01_un_brouillon_par_ecole_du_lot(tmp_path):
    """T-PRP-01 — Préparation le 20/12 à 16 h, sur un MASTER fictif (2 écoles `Lot 1`, 1 `Plus tard`,
    1 école sans règle) : un brouillon par école du lot, du 01/12 au 31/12 ; séances tenues et à venir
    distinguées ; pas de brouillon pour l'école `Plus tard` ; « écoles sans règle » listée."""
    P = prp()
    prep = preparer(installer(tmp_path))

    assert prep.format == P.FORMAT_PREPARATION == 1
    assert (prep.entite, prep.mois, prep.prepare_le) == ("ROSARI", MOIS, MAINTENANT)
    assert identifiants(prep) == {"ABC-2026-12", "ECO-2026-12"}
    for b in prep.brouillons:
        assert b.statut == "prêt", b.blocages
        assert not b.blocages
        assert (b.entite, b.mois) == ("ROSARI", MOIS)
        assert b.periode == (date(2026, 12, 1), date(2026, 12, 31))
        assert b.periode_label == "Déc 2026"
        assert b.date_emission == date(2027, 1, 1)
        assert b.masquer_jour is False and b.regroupement == "séance"
        assert b.contrat is None and b.bon_commande is None
        assert b.info["SIRET"] == SIRET_FICTIF and b.info["IBAN"] == IBAN_FICTIF and b.info["BIC"] == BIC_FICTIF
        assert b.nda == NDA_FICTIF
        assert b.emetteur["nom"] == "FORMATEUR FICTIF ROSARI EI"

    abc, eco = brouillon(prep, "ABC-2026-12"), brouillon(prep, "ECO-2026-12")
    assert {u: ligne.tenue for u, ligne in _par_uid(abc).items()} == {"abc-1": True, "abc-2": True, "abc-3": False}
    assert {u: ligne.tenue for u, ligne in _par_uid(eco).items()} == {"eco-1": True, "eco-2": False}
    assert (abc.ecole, abc.ecole_label, abc.destinataire["nom"]) == ("ABC", "ÉCOLE FICTIVE A", "ÉCOLE FICTIVE A")
    assert (eco.ecole, eco.ecole_label) == ("ECO", "ÉCOLE FICTIVE B")
    assert (abc.total_heures, abc.total_ht) == (9.0, Decimal("540.00"))
    assert (eco.total_heures, eco.total_ht) == (6.0, Decimal("270.00"))
    assert (abc.numero_probable, eco.numero_probable) == ("R_ABC001", "R_ECO001")

    ligne = _par_uid(abc)["abc-1"]
    assert ligne.cle_seance == cle(standard("abc-1")) == "abc-1@2026-12-03T09:00+0100"
    assert ligne.date == date(2026, 12, 3)
    assert ligne.debut == datetime(2026, 12, 3, 9, 0, tzinfo=PARIS)
    assert ligne.fin == datetime(2026, 12, 3, 12, 0, tzinfo=PARIS)
    assert ligne.intitule == "Statistiques M1"
    assert (ligne.ecole, ligne.promotion, ligne.module, ligne.libelle, ligne.cle_tarif) == (
        "ABC", "M1", "Statistiques", "Statistiques", "ABC|Statistiques|M1")
    assert ligne.classee_par == "code"
    assert (ligne.duree_agenda_h, ligne.heures, ligne.taux, ligne.montant) == (3.0, 3.0, 60.0, Decimal("180"))
    assert ligne.ref_contrat is None
    # Taux saisi en texte « 45 €/h » (§ 15), séance reconnue par son intitulé exact.
    assert {(l.classee_par, l.taux) for l in eco.lignes} == {("exact", 45.0)}

    # École « Plus tard » : aucun brouillon, ses séances en hors lot ; école sans règle listée.
    assert not any(b.ecole == "FUT" for b in prep.brouillons)
    assert [s.ecole for s in prep.hors_lot] == ["FUT"]
    assert tuple(prep.ecoles_sans_regle) == ("SRG",)
    assert not {"fut-1", "srg-1"} & uids_factures(prep)
    assert {a.seance.uid for a in prep.a_classer} == {"inconnu-1"}


def _contrats_abc() -> dict:
    """Deux lignes de règles pour ABC, une par contrat (§ 16), et une école ECO sans contrat."""
    return dict(
        regles=[
            regle("ABC", **{"N° de contrat": "CT-0001", "Une facture par contrat": "oui"}),
            regle("ABC", **{"N° de contrat": "CT-0002", "Une facture par contrat": "oui"}),
            regle("ECO"),
        ],
        tarifs=[
            tarif("ABC", "Statistiques", "M1", 60, "CT-0001"),
            tarif("ABC", "Économétrie", "M2", 70, "CT-0002"),
            tarif("ECO", "Économie", "L3", 45),
        ],
        correspondances=[
            correspondance("#ABC-M1-STATS", "code", "ABC", "Statistiques", "M1"),
            correspondance("#ABC-M2-ECONO", "code", "ABC", "Économétrie", "M2"),
            correspondance("Économie L3", "exact", "ECO", "Économie", "L3"),
        ],
    )


SEANCE_M2 = SeanceFictive("abc-m2", "2026-12-14", "09:00", "11:00", "Économétrie M2", "#ABC-M2-ECONO")


def test_t_prp_02_une_facture_par_contrat(tmp_path):
    """T-PRP-02 — « Une facture par contrat » avec deux `Réf. contrat` : deux brouillons, chacun avec
    son contrat (§ 16 : une ligne de règles par contrat ; identifiant `ABC-2026-12-CT0001`)."""
    p = installer(tmp_path, **_contrats_abc())
    prep = preparer(p, ics([standard("abc-1"), standard("abc-2"), SEANCE_M2, standard("eco-1")]))

    assert identifiants(prep) == {"ABC-2026-12-CT0001", "ABC-2026-12-CT0002", "ECO-2026-12"}
    c1, c2 = brouillon(prep, "ABC-2026-12-CT0001"), brouillon(prep, "ABC-2026-12-CT0002")
    assert (c1.contrat, set(_par_uid(c1))) == ("CT-0001", {"abc-1", "abc-2"})
    assert (c2.contrat, set(_par_uid(c2))) == ("CT-0002", {"abc-m2"})
    assert {l.ref_contrat for l in c1.lignes} == {"CT-0001"}
    assert c1.statut == c2.statut == "prêt"
    assert (c1.total_ht, c2.total_ht) == (Decimal("360.00"), Decimal("140.00"))
    assert c1.spec("R_ABC001")["ref_contrat"] == "CT-0001"
    assert c2.spec("R_ABC002")["ref_contrat"] == "CT-0002"
    assert brouillon(prep, "ECO-2026-12").contrat is None


def test_t_prp_02_plusieurs_contrats_sans_regle_ni_numero_bloque(tmp_path):
    """T-PRP-02 — Deux références sans « Une facture par contrat » ni `N° de contrat` : un seul
    brouillon, bloqué (« plusieurs contrats (…) : renseignez … (Règles_facturation) ») ; la mention du
    contrat est obligatoire (ADR030 §3). L'autre école reste prête."""
    tarifs = [
        tarif("ABC", "Statistiques", "M1", 60),
        tarif("ECO", "Économie", "L3", 45, "CT-0003"),
        tarif("ECO", "Gestion", "L3", 45, "CT-0004"),
    ]
    corr = correspondances_standard() + [correspondance("Gestion L3", "exact", "ECO", "Gestion", "L3")]
    gestion = SeanceFictive("eco-g", "2026-12-09", "14:00", "16:00", "Gestion L3")
    prep = preparer(installer(tmp_path, tarifs=tarifs, correspondances=corr),
                    ics([standard("abc-1"), standard("eco-1"), gestion]))

    assert identifiants(prep) == {"ABC-2026-12", "ECO-2026-12"}
    eco = brouillon(prep, "ECO-2026-12")
    assert eco.statut == "bloqué"
    bl = textes(eco.blocages)
    assert "plusieurs contrats" in bl
    assert "CT-0003" in bl and "CT-0004" in bl and "Règles_facturation" in bl
    assert set(_par_uid(eco)) == {"eco-1", "eco-g"}
    assert brouillon(prep, "ABC-2026-12").statut == "prêt"


def test_t_prp_02_seances_sans_reference_de_contrat(tmp_path):
    """T-PRP-02 — Une facture par contrat : les séances dont la ligne de tarif n'a pas de `Réf. contrat`
    forment le brouillon `…-SANS-CONTRAT`, `contrat=None`, bloqué (« référence de contrat absente
    (Tarifs, colonne « Réf. contrat », clé … ») ; le brouillon du contrat reste prêt."""
    donnees = _contrats_abc()
    donnees["regles"] = [donnees["regles"][0], regle("ECO")]
    donnees["tarifs"][1] = tarif("ABC", "Économétrie", "M2", 70, None)
    prep = preparer(installer(tmp_path, **donnees), ics([standard("abc-1"), SEANCE_M2, standard("eco-1")]))

    assert identifiants(prep) == {"ABC-2026-12-CT0001", "ABC-2026-12-SANS-CONTRAT", "ECO-2026-12"}
    sans = brouillon(prep, "ABC-2026-12-SANS-CONTRAT")
    assert sans.contrat is None and sans.statut == "bloqué"
    assert set(_par_uid(sans)) == {"abc-m2"}
    bl = textes(sans.blocages)
    assert "Réf. contrat" in bl and "ABC|Économétrie|M2" in bl
    assert brouillon(prep, "ABC-2026-12-CT0001").statut == "prêt"


def test_t_prp_02_suffixe_des_identifiants_par_contrat(tmp_path):
    """T-PRP-02 (cas limites) — Suffixe = référence réduite à `[A-Z0-9]` et à 20 caractères ; deux
    références réduites au même suffixe reçoivent chacune `-` et les 4 premiers caractères
    hexadécimaux du sha256 de leur référence (💡 lecture de « leur sha256 »)."""
    import hashlib

    longue = "CONTRAT-CADRE-2026-ECOLE-A-0001"
    references = ("CT-0001", "CT 0001", longue)
    regles = [regle("ABC", **{"N° de contrat": r, "Une facture par contrat": "oui"}) for r in references]
    modules = ("Statistiques", "Économétrie", "Algèbre")
    tarifs = [tarif("ABC", m, "M1", 60, r) for m, r in zip(modules, references)]
    codes = ("#ABC-M1-STATS", "#ABC-M1-ECONO", "#ABC-M1-ALGEBRE")
    corr = [correspondance(c, "code", "ABC", m, "M1") for c, m in zip(codes, modules)]
    seances = [SeanceFictive(f"abc-{i}", f"2026-12-0{i + 1}", "09:00", "10:00", m, c)
               for i, (m, c) in enumerate(zip(modules, codes))]
    prep = preparer(installer(tmp_path, regles=regles + [regle("ECO")], tarifs=tarifs, correspondances=corr),
                    ics(seances))

    def h(ref: str) -> str:
        return hashlib.sha256(ref.encode("utf-8")).hexdigest()[:4]

    assert identifiants(prep) == {
        f"ABC-2026-12-CT0001-{h('CT-0001')}",
        f"ABC-2026-12-CT0001-{h('CT 0001')}",
        "ABC-2026-12-CONTRATCADRE2026ECOL",
    }
    assert {b.contrat for b in prep.brouillons} == set(references)


def test_t_prp_03_total_exact_arrondi_a_la_fin(tmp_path):
    """T-PRP-03 — Total = Σ heures × taux exact, arrondi une seule fois à la fin ; il comprend un
    coefficient 7/8 (9 h–17 h → 7 h) et un forfait (7 h → 7,5 h), et égale le total de `facture.apercu`.

    À 33,33 €/h : 1,5 h → 49,995 ; 7 h → 233,31 ; 7,5 h → 249,975. Somme exacte 533,28 ; la somme des
    lignes arrondies donnerait 533,29."""
    from facturation import facture
    from facturation.heures import arrondi_centime

    tarifs = [tarif("ABC", m, "M1", 33.33) for m in ("Statistiques", "Journée intensive", "Jury")]
    corr = [
        correspondance("#ABC-M1-STATS", "code", "ABC", "Statistiques", "M1"),
        correspondance("#ABC-M1-JOUR", "code", "ABC", "Journée intensive", "M1", mode="coefficient", valeur=0.875),
        correspondance("#ABC-M1-JURY", "code", "ABC", "Jury", "M1", mode="forfait", valeur=7.5),
    ]
    seances = [
        SeanceFictive("abc-s", "2026-12-03", "09:00", "10:30", "Statistiques M1", "#ABC-M1-STATS"),
        SeanceFictive("abc-j", "2026-12-07", "09:00", "17:00", "Journée intensive M1", "#ABC-M1-JOUR"),
        SeanceFictive("abc-y", "2026-12-09", "09:00", "16:00", "Jury M1", "#ABC-M1-JURY"),
    ]
    abc = brouillon(preparer(installer(tmp_path, tarifs=tarifs, correspondances=corr), ics(seances)), "ABC-2026-12")

    lignes = _par_uid(abc)
    assert {u: (l.duree_agenda_h, l.heures) for u, l in lignes.items()} == {
        "abc-s": (1.5, 1.5), "abc-j": (8.0, 7.0), "abc-y": (7.0, 7.5)}
    assert {u: l.montant for u, l in lignes.items()} == {
        "abc-s": Decimal("49.995"), "abc-j": Decimal("233.31"), "abc-y": Decimal("249.975")}
    assert abc.total_heures == 16.0
    assert abc.total_ht == Decimal("533.28")
    # Garde : la somme des lignes arrondies diffère du total exact arrondi une fois.
    assert sum(arrondi_centime(l.montant) for l in abc.lignes) == Decimal("533.29") != abc.total_ht

    spec = abc.spec("R_ABC001")
    texte = facture.apercu(spec, facture.construire_groupes(spec["lignes"]), abc.destinataire,
                           emetteur=abc.emetteur, echeance=abc.echeance, nda=abc.nda)
    total = re.search(r"TOTAL HT\s*:\s*([0-9.]+) €", texte)
    assert total and Decimal(total.group(1)) == abc.total_ht


# ──────────────────────────────────────────────────────────────────────
# T-PRP-05 à T-PRP-08 : file à classer, taux, Sessions, mentions
# ──────────────────────────────────────────────────────────────────────

AMBIGUE = SeanceFictive("ambigu-1", "2026-12-14", "10:00", "12:00", "Séminaire")
CODE_INCONNU = SeanceFictive("code-1", "2026-12-15", "14:00", "15:00", "Atelier", "#ZZZ-M9-ATELIER")


def test_t_prp_05_creneaux_a_classer_jamais_factures(tmp_path):
    """T-PRP-05 — Créneau non reconnu, ambigu ou de code inconnu : dans `a_classer`, absent de toute
    ligne. Une décision `classer` injectée le fait facturer (`classee_par="décision"`) ; c'est aussi
    la garde de l'assertion négative : le même contrôle voit la séance dès qu'elle est facturée."""
    corr = correspondances_standard() + [
        correspondance("Séminaire", "exact", "ABC", "Statistiques", "M1"),
        correspondance("Séminaire", "exact", "ECO", "Économie", "L3"),
    ]
    seances = [*SEANCES_STANDARD, AMBIGUE, CODE_INCONNU]
    p = installer(tmp_path, correspondances=corr)
    prep = preparer(p, ics(seances))

    raisons = {a.seance.uid: a.raison for a in prep.a_classer}
    assert raisons == {"inconnu-1": "non reconnue", "ambigu-1": "ambiguë", "code-1": "code inconnu"}
    a_classer = {cle(standard("inconnu-1")), cle(AMBIGUE), cle(CODE_INCONNU)}
    assert not a_classer & cles_facturees(prep)

    inconnu = standard("inconnu-1")
    ecrire_decision(p, "classer", cle(inconnu), {
        "ecole": "ABC", "module": "Statistiques", "promotion": "M1", "cle_tarif": "ABC|Statistiques|M1",
        "mode_heures": "agenda", "valeur": "",
    })
    apres = preparer(p, ics(seances))
    assert cle(inconnu) in cles_facturees(apres)
    ligne = _par_uid(brouillon(apres, "ABC-2026-12"))["inconnu-1"]
    assert (ligne.classee_par, ligne.ecole, ligne.heures, ligne.taux) == ("décision", "ABC", 1.0, 60.0)
    assert "inconnu-1" not in {a.seance.uid for a in apres.a_classer}
    assert not {cle(AMBIGUE), cle(CODE_INCONNU)} & cles_facturees(apres)


@pytest.mark.parametrize(
    "tarifs_eco", [[], [tarif("ECO", "Économie", "L3", "à négocier")]], ids=["ligne absente", "taux illisible"],
)
def test_t_prp_06_taux_absent_bloque_sans_rien_inventer(tmp_path, tarifs_eco):
    """T-PRP-06 — Taux absent : brouillon bloqué avec la clé citée, aucun taux inventé (taux et montant
    à None, lignes visibles) ; les autres brouillons sont prêts."""
    tarifs = [t for t in tarifs_standard() if t["École"] != "ECO"] + tarifs_eco
    prep = preparer(installer(tmp_path, tarifs=tarifs))

    eco = brouillon(prep, "ECO-2026-12")
    assert eco.statut == "bloqué"
    bl = textes(eco.blocages)
    assert "ECO|Économie|L3" in bl and "Tarifs" in bl and re.search(r"taux", bl, re.IGNORECASE)
    assert set(_par_uid(eco)) == {"eco-1", "eco-2"}
    assert all(l.taux is None and l.montant is None for l in eco.lignes)
    assert all(l.heures == 3.0 for l in eco.lignes)
    assert eco.total_ht is None
    assert brouillon(prep, "ABC-2026-12").statut == "prêt"


def test_t_prp_07_seance_deja_dans_sessions_exclue(tmp_path):
    """T-PRP-07 — Séance présente dans `Sessions` (exacte, ou même jour sans heure) : exclue et listée.
    Une décision `inclure` réintègre la séance « même jour sans heure » ; la correspondance exacte,
    elle, reste exclue (💡 : `inclure` ne vaut que pour une séance « même jour sans heure »)."""
    a1, a2 = standard("abc-1"), standard("abc-2")
    sessions = [
        ligne_sessions(a1, ecole="ABC", n_facture="R_ABC001", debut_="09:00", fin_="12:00"),
        ligne_sessions(a2, ecole="ABC", n_facture="R_ABC001"),
    ]
    p = installer(tmp_path, sessions=sessions)
    prep = preparer(p)

    abc = brouillon(prep, "ABC-2026-12")
    assert set(_par_uid(abc)) == {"abc-3"}
    deja = {d.cle_seance: d for d in prep.deja_facturees}
    assert set(deja) == {cle(a1), cle(a2)}
    assert (deja[cle(a1)].certitude, deja[cle(a1)].n_facture, deja[cle(a1)].ecole, deja[cle(a1)].date) == (
        "exacte", "R_ABC001", "ABC", date(2026, 12, 3))
    assert deja[cle(a2)].certitude == "même jour sans heure"
    assert abc.numero_probable == "R_ABC002"

    ecrire_decision(p, "inclure", cle(a2))
    ecrire_decision(p, "inclure", cle(a1))
    apres = preparer(p)
    assert set(_par_uid(brouillon(apres, "ABC-2026-12"))) == {"abc-2", "abc-3"}
    assert cle(a1) in {d.cle_seance for d in apres.deja_facturees}


@pytest.mark.parametrize("cas", ["émetteur sans EI", "NDA absent", "SIRET absent"])
def test_t_prp_08_mention_obligatoire_absente_bloque_tous_les_brouillons(tmp_path, cas):
    """T-PRP-08 — Émetteur sans « EI », NDA absent ou SIRET absent : tous les brouillons bloqués, aucune
    exception, aperçus produits."""
    if cas == "émetteur sans EI":
        p, motif = installer(tmp_path, nom_emetteur="FORMATEUR FICTIF ROSARI"), r"\bEI\b"
    elif cas == "NDA absent":
        p, motif = installer(tmp_path, nda=None), r"NDA|déclaration d'activité"
    else:
        p, motif = installer(tmp_path), r"SIRET"
        retirer_info(p.master, "ROSARI", "SIRET")
    P = prp()
    prep = preparer(p)

    assert identifiants(prep) == {"ABC-2026-12", "ECO-2026-12"}
    for b in prep.brouillons:
        assert b.statut == "bloqué"
        assert re.search(motif, textes(b.blocages), re.IGNORECASE), b.blocages
        texte = P.apercu_facture(b)
        assert b.blocages[0] in texte
    assert "À FAIRE" in P.apercu_global(prep)


# ──────────────────────────────────────────────────────────────────────
# T-PRP-10, T-PRP-12, T-PRP-13 : empreintes, cycle, brouillons engagés
# ──────────────────────────────────────────────────────────────────────


def test_t_prp_10_empreintes_stables_et_sensibles(tmp_path):
    """T-PRP-10 — Même agenda à deux heures différentes : mêmes empreintes (la marque `tenue` change,
    l'empreinte non). Un créneau à classer de plus (avertissement) ne la change pas. Un taux modifié
    change l'empreinte de son brouillon, et seulement de lui."""
    P = prp()
    p = installer(tmp_path)
    avant = preparer(p)
    plus_tard = preparer(p, maintenant=datetime(2026, 12, 29, 9, 0, tzinfo=PARIS))

    def empreintes(prep):
        return {b.identifiant: b.empreinte for b in prep.brouillons}

    assert empreintes(plus_tard) == empreintes(avant)
    assert _par_uid(brouillon(plus_tard, "ABC-2026-12"))["abc-3"].tenue is True
    assert _par_uid(brouillon(avant, "ABC-2026-12"))["abc-3"].tenue is False
    for b in avant.brouillons:
        assert re.fullmatch(r"[0-9a-f]{64}", b.empreinte)
        assert P.empreinte_brouillon(b) == b.empreinte

    reunion = SeanceFictive("inconnu-2", "2026-12-18", "13:00", "14:00", "Réunion de rentrée")
    assert empreintes(preparer(p, ics([*SEANCES_STANDARD, reunion]))) == empreintes(avant)

    tarifs = [tarif("ABC", "Statistiques", "M1", 62) if t["École"] == "ABC" else t for t in tarifs_standard()]
    autre = preparer(installer(tmp_path / "taux_modifie", tarifs=tarifs))
    assert brouillon(autre, "ABC-2026-12").empreinte != brouillon(avant, "ABC-2026-12").empreinte
    assert brouillon(autre, "ECO-2026-12").empreinte == brouillon(avant, "ECO-2026-12").empreinte


def test_t_prp_12_ecole_a_cycle(tmp_path):
    """T-PRP-12 — École à cycle (début 16) : période du 16/11 au 15/12, date d'émission le 01/01,
    avertissement, aucun blocage. La fenêtre de lecture couvre le cycle et le mois civil ; la séance du
    18/12 appartient au cycle suivant."""
    regles = [regle("ABC"), regle("ECO", Rythme="cycle", **{"Début de cycle": 16})]
    seances = [
        standard("abc-1"),
        SeanceFictive("eco-a", "2026-11-20", "14:00", "17:00", "Économie L3"),
        standard("eco-1"),
        SeanceFictive("eco-c", "2026-12-18", "14:00", "17:00", "Économie L3"),
    ]
    p = installer(tmp_path, regles=regles)
    assert entrees(p, ics(seances)).fenetre == (date(2026, 11, 16), date(2026, 12, 31))
    prep = preparer(p, ics(seances))

    eco = brouillon(prep, "ECO-2026-12")
    assert eco.periode == (date(2026, 11, 16), date(2026, 12, 15))
    assert eco.periode_label == "Déc 2026"
    assert set(_par_uid(eco)) == {"eco-a", "eco-1"}
    assert eco.date_emission == date(2027, 1, 1)
    assert re.search(r"cycle", textes(eco.avertissements), re.IGNORECASE)
    assert not eco.blocages and eco.statut == "prêt"
    assert prep.fenetre == (date(2026, 11, 16), date(2026, 12, 31))
    assert brouillon(prep, "ABC-2026-12").periode == (date(2026, 12, 1), date(2026, 12, 31))


@pytest.mark.parametrize("finale", [False, True], ids=["en émission", "émise"])
def test_t_prp_13_brouillon_engage_jamais_regenere(tmp_path, finale):
    """T-PRP-13 — Brouillon engagé (pièce présente) : non régénéré, listé « déjà émise » ; un contenu
    changé est marqué, et la séance ajoutée dans la période émise est signalée sans être facturée
    (renvoi à `facturation recontroler`, F-23)."""
    P = prp()
    p = installer(tmp_path)
    abc = brouillon(preparer(p), "ABC-2026-12")
    ecrire_piece(p, "R_ABC001", identifiant="ABC-2026-12", empreinte=abc.empreinte, brouillon_=abc, finale=finale)

    prep = preparer(p)
    assert identifiants(prep) == {"ECO-2026-12"}
    assert [tuple(d) for d in prep.deja_emises] == [("ABC-2026-12", "R_ABC001", False)]
    assert not prep.non_facturees_periode_emise

    ajout = SeanceFictive("abc-4", "2026-12-17", "09:00", "12:00", "Statistiques M1", "#ABC-M1-STATS")
    change = preparer(p, ics([*SEANCES_STANDARD, ajout]))
    assert identifiants(change) == {"ECO-2026-12"}
    assert [tuple(d) for d in change.deja_emises] == [("ABC-2026-12", "R_ABC001", True)]
    assert {s.seance.uid for s in change.non_facturees_periode_emise} == {"abc-4"}
    assert "abc-4" not in uids_factures(change)
    assert "facturation recontroler" in P.apercu_global(change)


# ──────────────────────────────────────────────────────────────────────
# T-PRP-14 à T-PRP-17 : blocages, bon de commande, fin de module, regroupement
# ──────────────────────────────────────────────────────────────────────


def test_t_prp_14_doublon_d_agenda_bloque(tmp_path):
    """T-PRP-14 — Doublon d'agenda (deux séances de la même école, même début et même fin) : blocage ;
    les autres écoles restent prêtes."""
    copie = SeanceFictive("abc-1-copie", "2026-12-03", "09:00", "12:00", "Statistiques M1", "#ABC-M1-STATS")
    prep = preparer(installer(tmp_path), ics([*SEANCES_STANDARD, copie]))

    abc = brouillon(prep, "ABC-2026-12")
    assert abc.statut == "bloqué"
    assert re.search(r"doublon", textes(abc.blocages), re.IGNORECASE)
    assert brouillon(prep, "ECO-2026-12").statut == "prêt"


def _facture_registre(n: str, ecole: str, periode: str, contrat: str | None = None) -> dict:
    return {"N° facture": n, "École": ecole, "Contrat": contrat, "Période": periode,
            "Date d'émission": datetime(2026, 12, 1), "Total HT": 100, "Statut": "émise"}


def test_t_prp_14_periode_deja_facturee_hors_outil(tmp_path):
    """T-PRP-14 — Ligne `Factures` de la même école, de la même période (« Déc 2026 ») et du même
    contrat (vide égal vide), sans pièce d'émission : blocage, levé par une décision `inclure` sur
    l'identifiant. Garde : la ligne d'une autre période (ECO, « Nov 2026 ») ne bloque pas."""
    factures = [_facture_registre("R_ABC004", "ABC", "Déc 2026"), _facture_registre("R_ECO004", "ECO", "Nov 2026")]
    p = installer(tmp_path, factures=factures)
    prep = preparer(p)

    abc = brouillon(prep, "ABC-2026-12")
    assert abc.statut == "bloqué"
    assert re.search(r"déjà facturée", textes(abc.blocages), re.IGNORECASE)
    assert not re.search(r"déjà facturée", textes(brouillon(prep, "ECO-2026-12").blocages), re.IGNORECASE)

    ecrire_decision(p, "inclure", "ABC-2026-12")
    abc = brouillon(preparer(p), "ABC-2026-12")
    assert not re.search(r"déjà facturée", textes(abc.blocages), re.IGNORECASE)
    assert abc.statut == "prêt"


def test_t_prp_15_ecole_a_bon_de_commande(tmp_path):
    """T-PRP-15 — École à bon de commande : préparée puis bloquée, avec le renvoi à
    `examples/factures_par_bon_de_commande.py` (Q-C1). Avec une décision `bon_commande`, le brouillon
    est prêt et `spec()` porte la référence ; avant, `spec()` lève `ValueError`."""
    regles = [regle("ABC", **{"Bon de commande requis": "oui"}), regle("ECO")]
    p = installer(tmp_path, regles=regles)
    abc = brouillon(preparer(p), "ABC-2026-12")

    assert abc.statut == "bloqué" and abc.bon_commande is None
    bl = textes(abc.blocages)
    assert re.search(r"bon de commande", bl, re.IGNORECASE)
    assert "factures_par_bon_de_commande" in bl
    with pytest.raises(ValueError):
        abc.spec("R_ABC001")

    ecrire_decision(p, "bon_commande", "ABC-2026-12", {"bon_commande": "BC-FICTIF-0001"})
    abc = brouillon(preparer(p), "ABC-2026-12")
    assert abc.statut == "prêt" and abc.bon_commande == "BC-FICTIF-0001"
    assert abc.spec("R_ABC001")["ref_bon_commande"] == "BC-FICTIF-0001"


def test_t_prp_16_fin_de_module_non_preparee(tmp_path):
    """T-PRP-16 — École en fin de module : dans `ecoles_non_preparees` (« fin de module : à préparer à
    la main »), sans exception ; aucune de ses séances n'est facturée, les autres écoles sont préparées."""
    regles = [regle("ABC"), regle("ECO", Rythme="fin de module")]
    prep = preparer(installer(tmp_path, regles=regles))

    assert identifiants(prep) == {"ABC-2026-12"}
    non_preparees = {e.ecole: e.raison for e in prep.ecoles_non_preparees}
    assert set(non_preparees) == {"ECO"}
    assert "fin de module" in non_preparees["ECO"]
    assert not {"eco-1", "eco-2"} & uids_factures(prep)


def test_t_prp_17_regroupement_par_jour(tmp_path):
    """T-PRP-17 — Regroupement « jour » : deux séances du même jour fusionnent en une ligne dans
    `groupes()`, mais restent deux `LigneBrouillon`. En regroupement « séance », une ligne par séance,
    triées par début."""
    regles = [regle("ABC", Regroupement="jour"), regle("ECO")]
    seances = [
        standard("abc-1"),
        SeanceFictive("abc-1b", "2026-12-03", "13:00", "16:00", "Statistiques M1", "#ABC-M1-STATS"),
        standard("abc-2"),
        standard("eco-1"),
        SeanceFictive("eco-1b", "2026-12-08", "09:00", "11:00", "Économie L3"),
    ]
    prep = preparer(installer(tmp_path, regles=regles), ics(seances))

    abc = brouillon(prep, "ABC-2026-12")
    assert abc.regroupement == "jour"
    assert len(abc.lignes) == 3
    assert _groupes(abc) == [("M1", "Statistiques", [(date(2026, 12, 3), 6.0, 60.0), (date(2026, 12, 10), 3.0, 60.0)])]
    seances_spec = abc.spec("R_ABC001")["lignes"][0]["seances"]
    assert [(s["date"], s["heures"]) for s in seances_spec] == [("2026-12-03", 6.0), ("2026-12-10", 3.0)]

    eco = brouillon(prep, "ECO-2026-12")
    assert _groupes(eco) == [("L3", "Économie", [(date(2026, 12, 8), 2.0, 45.0), (date(2026, 12, 8), 3.0, 45.0)])]


# ──────────────────────────────────────────────────────────────────────
# T-PRP-20 : sections annexes
# ──────────────────────────────────────────────────────────────────────


def _module_factice(monkeypatch, nom: str, section_apercu) -> None:
    module = types.ModuleType(nom)
    module.section_apercu = section_apercu
    monkeypatch.setitem(sys.modules, nom, module)


def test_t_prp_20_sections_annexes(tmp_path, monkeypatch):
    """T-PRP-20 — `sections_annexes` avec un module factice (`facturation.registre`) : sa section est
    rendue, appelée avec le profil et la date ; un module absent (`facturation.apres_emission`) est
    ignoré."""
    p = installer(tmp_path)
    appels = []

    def section_apercu(profil, *, aujourd_hui):
        appels.append((profil.entite, aujourd_hui))
        return "FACTURES ÉCHUES NON PAYÉES : 1", ["R_ABC001 — 12 jours de retard"]

    _module_factice(monkeypatch, "facturation.registre", section_apercu)
    monkeypatch.setitem(sys.modules, "facturation.apres_emission", None)  # import → ModuleNotFoundError

    sections = prp().sections_annexes(p, aujourd_hui=date(2026, 12, 20))
    assert [(titre, list(lignes)) for titre, lignes in sections] == [
        ("FACTURES ÉCHUES NON PAYÉES : 1", ["R_ABC001 — 12 jours de retard"])]
    assert appels == [("ROSARI", date(2026, 12, 20))]


def test_t_prp_20_section_en_erreur_ou_vide(tmp_path, monkeypatch):
    """T-PRP-20 — Toute autre exception donne la section « section indisponible : <type> », y compris
    un `ModuleNotFoundError` sur un autre nom que celui du module (dépendance absente) ; une section
    `None` n'est pas rendue."""

    def en_panne(profil, *, aujourd_hui):
        raise RuntimeError("panne fictive")

    def dependance_absente(profil, *, aujourd_hui):
        raise ModuleNotFoundError("No module named 'dependance_fictive'", name="dependance_fictive")

    p = installer(tmp_path)
    _module_factice(monkeypatch, "facturation.registre", en_panne)
    _module_factice(monkeypatch, "facturation.apres_emission", dependance_absente)
    sections = prp().sections_annexes(p, aujourd_hui=date(2026, 12, 20))
    texte = "\n".join(f"{titre}\n" + "\n".join(lignes) for titre, lignes in sections)
    assert "section indisponible : RuntimeError" in texte
    assert "section indisponible : ModuleNotFoundError" in texte

    _module_factice(monkeypatch, "facturation.registre", lambda profil, *, aujourd_hui: None)
    monkeypatch.setitem(sys.modules, "facturation.apres_emission", None)
    assert list(prp().sections_annexes(p, aujourd_hui=date(2026, 12, 20))) == []


# ──────────────────────────────────────────────────────────────────────
# T-PRP-21, T-PRP-22 : ajouts de l'orchestrateur
# ──────────────────────────────────────────────────────────────────────


@pytest.mark.parametrize(
    "ecoles",
    [
        [e for e in ecoles_standard() if e["Code facturation"] != "ECO"],
        [*(e for e in ecoles_standard() if e["Code facturation"] != "ECO"), {"École": None, "Code facturation": "ECO"}],
    ],
    ids=["absente de Écoles", "libellé vide"],
)
def test_t_prp_21_ecole_sans_libelle_le_code_sert_de_libelle(tmp_path, ecoles):
    """T-PRP-21 — Un brouillon dont l'école n'a pas de libellé produit une spec dont `ecole_label`
    égale le code école, jamais None ni vide (§ 18, hors périmètre de P-0a) ; avertissement « libellé
    de l'école absent de Écoles : le code est utilisé », sans blocage."""
    eco = brouillon(preparer(installer(tmp_path, ecoles=ecoles)), "ECO-2026-12")

    assert eco.ecole_label == "ECO"
    assert eco.spec("R_ECO001")["ecole_label"] == "ECO"
    assert re.search(r"libellé", textes(eco.avertissements), re.IGNORECASE)
    assert eco.statut == "prêt"


def test_t_prp_22_possible_doublon_avec_sessions_avertit_sans_exclure(tmp_path):
    """T-PRP-22 (💡 [proposé — à valider], constat écarté de la revue de P-0b) — Une séance de l'agenda
    et une ligne `Sessions` de la même école, le même jour, à d'autres heures (09:00–12:00 contre
    09:00–12:30) : avertissement « possible doublon avec Sessions, ligne N » dans l'aperçu (section
    « À FAIRE »), sans exclure la séance ni bloquer le brouillon. Garde : la ligne `Sessions` d'un
    autre jour (ligne 3, 04/12) ne donne aucun avertissement, et les séances d'autres jours (abc-2,
    abc-3) non plus.

    Constat de revue de F-16a : la ligne 3 avait les mêmes heures que les séances (09:00–12:00) ; elle
    était écartée comme « identique » même sans le filtre de date, et l'assertion négative passait à
    vide. Elle chevauche désormais (09:00–12:30) : sans le filtre de date, elle et la ligne 2
    donneraient un avertissement pour abc-1, abc-2 et abc-3 (garde cassée une fois)."""
    a1 = standard("abc-1")
    sessions = [
        ligne_sessions(a1, ecole="ABC", n_facture="R_ABC001", debut_="09:00", fin_="12:30", heures=3.5),
        ligne_sessions(a1._replace(jour="2026-12-04"), ecole="ABC", n_facture="R_ABC001",
                       debut_="09:00", fin_="12:30", heures=3.5),
    ]
    prep = preparer(installer(tmp_path, sessions=sessions))

    abc = brouillon(prep, "ABC-2026-12")
    assert cle(a1) in {l.cle_seance for l in abc.lignes}
    assert cle(a1) not in {d.cle_seance for d in prep.deja_facturees}
    assert abc.statut == "prêt" and "Sessions" not in textes(abc.blocages)
    avertissements = textes(abc.avertissements)
    assert "possible doublon avec Sessions, ligne 2" in avertissements
    assert "ligne 3" not in avertissements
    assert sum(a.startswith("possible doublon") for a in abc.avertissements) == 1, abc.avertissements
    a_faire = bloc(prp().apercu_global(prep), "À FAIRE", "FACTURES")
    assert "possible doublon avec Sessions, ligne 2" in a_faire


# ──────────────────────────────────────────────────────────────────────
# T-PRP-23 à T-PRP-30 : sélection, dates, entrées, décisions, erreurs, spec
# ──────────────────────────────────────────────────────────────────────


def test_t_prp_23_selection_d_ecoles(tmp_path):
    """T-PRP-23 — `ecoles` restreint la préparation ; un code inconnu ou hors du lot lève `ValueError`
    (« école « X » inconnue ou hors du Lot 1 (onglet Règles_facturation) »)."""
    P = prp()
    p = installer(tmp_path)
    assert identifiants(preparer(p, ecoles=["ECO"])) == {"ECO-2026-12"}
    e = entrees(p)
    for code in ("ZZZ", "FUT"):
        with pytest.raises(ValueError, match=code) as exc:
            P.preparer(entite="ROSARI", mois=MOIS, entrees=e, prefixe="R_", maintenant=MAINTENANT, ecoles=[code])
        assert "Règles_facturation" in str(exc.value)


def test_t_prp_24_dates_d_emission_et_echeances(tmp_path):
    """T-PRP-24 — Date d'émission par défaut : le 1er du mois suivant (ADR027) ; « 45 jours fin de
    mois » = date de facture + 45 jours, puis fin de ce mois-là (01/01/2027 → 28/02/2027, Q-A4) ; délai
    vide = 30 jours ; dates en ISO dans la spec."""
    P = prp()
    assert P.date_emission_par_defaut("2026-12") == date(2027, 1, 1)
    assert P.date_emission_par_defaut("2026-11") == date(2026, 12, 1)
    with pytest.raises(ValueError, match="2026-13"):
        P.date_emission_par_defaut("2026-13")

    prep = preparer(installer(tmp_path))
    abc = brouillon(prep, "ABC-2026-12")
    assert abc.echeance == date(2027, 2, 28)
    assert brouillon(prep, "ECO-2026-12").echeance == date(2027, 1, 31)
    spec = abc.spec("R_ABC001")
    assert (spec["date_emission"], spec["echeance"]) == ("2027-01-01", "2027-02-28")


def test_t_prp_25_fenetre_de_lecture(tmp_path):
    """T-PRP-25 — `fenetre_de_lecture` : union des périodes des écoles (fin de module ignorée) et du
    mois civil ; `charger_entrees` la restreint aux écoles demandées."""
    from facturation.master import lire_master
    from facturation.regles import regles_par_ecole

    P = prp()
    p = installer(tmp_path, regles=[regle("ABC"), regle("ECO", Rythme="cycle", **{"Début de cycle": 16}),
                                    regle("FIN", Rythme="fin de module")])
    regles = regles_par_ecole(lire_master(p.master).regles_facturation)
    assert P.fenetre_de_lecture(regles, "2026-12", ["ABC", "ECO", "FIN"]) == (date(2026, 11, 16), date(2026, 12, 31))
    assert P.fenetre_de_lecture(regles, "2026-12", ["ABC", "FIN"]) == (date(2026, 12, 1), date(2026, 12, 31))
    assert P.fenetre_de_lecture(regles, "2027-01", ["ECO"]) == (date(2026, 12, 16), date(2027, 1, 31))
    assert entrees(p, ecoles=["ABC"]).fenetre == (date(2026, 12, 1), date(2026, 12, 31))
    assert entrees(p).fenetre == (date(2026, 11, 16), date(2026, 12, 31))


def test_t_prp_26_charger_entrees(tmp_path):
    """T-PRP-26 — `charger_entrees` : séances de la fenêtre, coordonnées, N° du MASTER et des pièces
    (abandonnées comprises), pièces engagées du mois par identifiant, empreinte d'agenda stable d'un
    chargement à l'autre et sensible à un changement de séance."""
    p = installer(tmp_path, factures=[_facture_registre("R_ECO002", "ECO", "Oct 2026")])
    ecrire_piece(p, "R_ABC007", identifiant="ABC-2026-11", mois="2026-11", abandonnee=True)
    ecrire_piece(p, "R_ABC006", identifiant="ABC-2026-12")
    e = entrees(p)

    assert e.fenetre == (date(2026, 12, 1), date(2026, 12, 31))
    assert {s.uid for s in e.seances} == {s.uid for s in SEANCES_STANDARD}
    assert {"R_ECO002", "R_ABC006", "R_ABC007"} <= set(e.numeros)
    assert set(e.engagees) == {"ABC-2026-12"} and e.engagees["ABC-2026-12"].n_facture == "R_ABC006"
    assert e.erreur_coordonnees is None
    assert e.nda == NDA_FICTIF
    assert e.emetteur["nom"] == "FORMATEUR FICTIF ROSARI EI"
    assert (e.info["SIRET"], e.info["IBAN"], e.info["BIC"]) == (SIRET_FICTIF, IBAN_FICTIF, BIC_FICTIF)
    assert e.destinataires["ABC"]["nom"] == "ÉCOLE FICTIVE A"
    assert re.fullmatch(r"[0-9a-f]{64}", e.empreinte_agenda)
    assert entrees(p).empreinte_agenda == e.empreinte_agenda
    modifie = [s._replace(intitule="Statistiques M1 (salle 2)") if s.uid == "abc-1" else s for s in SEANCES_STANDARD]
    assert entrees(p, ics(modifie)).empreinte_agenda != e.empreinte_agenda


def test_t_prp_26_coordonnees_incompletes_sans_erreur_globale(tmp_path):
    """T-PRP-26 — SIRET absent : `erreur_coordonnees` le nomme, sans exception ; NDA absent :
    `nda=None`, sans exception."""
    sans_siret = installer(tmp_path / "sans_siret")
    retirer_info(sans_siret.master, "ROSARI", "SIRET")
    e = entrees(sans_siret)
    assert e.erreur_coordonnees and "SIRET" in e.erreur_coordonnees

    sans_nda = entrees(installer(tmp_path / "sans_nda", nda=None))
    assert sans_nda.nda is None and sans_nda.erreur_coordonnees is None


def test_t_prp_27_decisions_actives_et_exclure(tmp_path):
    """T-PRP-27 — `decisions_actives` : décisions du journal du mois, moins les décisions annulées
    (ni celles d'un autre mois) ; une décision `exclure` retire la séance, ni facturée ni à classer ;
    une décision annulée ne s'applique plus."""
    D = importlib.import_module("facturation.decisions")
    assert D.NATURES == ("classer", "exclure", "inclure", "bon_commande", "annuler")
    p = installer(tmp_path)
    a1, a2 = standard("abc-1"), standard("abc-2")
    ecrire_decision(p, "exclure", cle(a1), {"raison": "séance offerte"}, id_="DEC000000001", commentaire="séance offerte")
    ecrire_decision(p, "exclure", cle(a2), {"raison": "erreur de saisie"}, id_="DEC000000002", commentaire="erreur")
    ecrire_decision(p, "annuler", "DEC000000002", id_="DEC000000003")
    ecrire_decision(p, "exclure", cle(standard("eco-1")), {"raison": "autre mois"}, mois="2026-11", id_="DEC000000004")

    actives = D.decisions_actives(p, MOIS)
    ids = {d.id for d in actives}
    assert "DEC000000001" in ids
    assert not {"DEC000000002", "DEC000000004"} & ids
    d1 = next(d for d in actives if d.id == "DEC000000001")
    assert (d1.nature, d1.cible, d1.mois, d1.poste) == ("exclure", cle(a1), MOIS, POSTE_TEST)
    assert dict(d1.valeurs) == {"raison": "séance offerte"}
    assert isinstance(d1.horodatage, datetime) and d1.horodatage.tzinfo is not None
    assert list(D.decisions_actives(p, "2027-01")) == []

    prep = preparer(p)
    assert "abc-1" not in uids_factures(prep)
    assert "abc-1" not in {a.seance.uid for a in prep.a_classer}
    assert {"abc-2", "eco-1"} <= uids_factures(prep)


def test_t_prp_28_correspondances_mal_formees_erreur_globale(tmp_path):
    """T-PRP-28 — Une ligne de `Correspondances` mal formée (sans clé tarif) est une erreur globale :
    `ValueError` qui nomme l'onglet."""
    corr = correspondances_standard() + [correspondance("Atelier", "exact", "ABC", "Atelier", "M1", cle_tarif=" ")]
    p = installer(tmp_path, correspondances=corr)
    with pytest.raises(ValueError, match="Correspondances"):
        preparer(p)


def test_t_prp_29_destinataire_absent_ou_incomplet(tmp_path):
    """T-PRP-29 — Destinataire absent de `Coordonnées`, ou incomplet (colonne nommée) : brouillon
    bloqué ; les autres restent prêts."""
    regles = regles_standard() + [regle("NEW")]
    tarifs = tarifs_standard() + [tarif("NEW", "Droit", "M1", 50)]
    corr = correspondances_standard() + [correspondance("Droit M1", "exact", "NEW", "Droit", "M1")]
    seances = [*SEANCES_STANDARD, SeanceFictive("new-1", "2026-12-09", "09:00", "12:00", "Droit M1")]
    p = installer(tmp_path, regles=regles, tarifs=tarifs, correspondances=corr)
    changer_coordonnee(p.master, "Destinataire", "ECO", "Adresse L2", None)
    prep = preparer(p, ics(seances))

    new = brouillon(prep, "NEW-2026-12")
    assert new.statut == "bloqué" and new.destinataire is None
    assert re.search(r"destinataire", textes(new.blocages), re.IGNORECASE)
    assert "Coordonnées" in textes(new.blocages)
    eco = brouillon(prep, "ECO-2026-12")
    assert eco.statut == "bloqué" and "Adresse L2" in textes(eco.blocages)
    assert brouillon(prep, "ABC-2026-12").statut == "prêt"


def test_t_prp_29_heures_non_calculables(tmp_path):
    """T-PRP-29 — Heures non calculables (coefficient sans valeur) : blocage « heures non
    calculables : … (onglet Correspondances) », heures et montants à None."""
    corr = [{**c, "Mode heures": "coefficient", "Valeur": None} if c["École"] == "ECO" else c
            for c in correspondances_standard()]
    prep = preparer(installer(tmp_path, correspondances=corr))

    eco = brouillon(prep, "ECO-2026-12")
    assert eco.statut == "bloqué"
    bl = textes(eco.blocages)
    assert "heures non calculables" in bl and "Correspondances" in bl
    assert all(l.heures is None and l.montant is None for l in eco.lignes)
    assert eco.total_ht is None
    assert brouillon(prep, "ABC-2026-12").statut == "prêt"


def test_t_prp_30_spec_conforme_a_facture(tmp_path):
    """T-PRP-30 — `spec(n)` rend la spec de `facture.py` : fichier `ROSARI/ABC/ROSARI_ABC.xlsx`, onglet
    `onglet_facture(n, periode_label)`, dates en ISO, JSON sérialisable ; `charger_spec` et
    `controler_spec` l'acceptent."""
    from facturation import facture

    p = installer(tmp_path)
    abc = brouillon(preparer(p), "ABC-2026-12")
    spec = abc.spec("R_ABC001")

    assert {k: spec[k] for k in ("entite", "ecole", "ecole_label", "fichier_sortie", "onglet", "n_facture")} == {
        "entite": "ROSARI", "ecole": "ABC", "ecole_label": "ÉCOLE FICTIVE A",
        "fichier_sortie": "ROSARI/ABC/ROSARI_ABC.xlsx", "onglet": "R_ABC001 Déc 2026", "n_facture": "R_ABC001"}
    assert spec["masquer_jour"] is False
    assert not spec.get("ref_contrat") and not spec.get("ref_bon_commande")
    chemin = tmp_path / "spec.json"
    chemin.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
    relue = facture.charger_spec(chemin)
    assert facture.controler_spec(relue, prefixe=p.prefixe_facture, racine=p.racine) == (
        p.racine / "ROSARI" / "ABC" / "ROSARI_ABC.xlsx")
    assert facture.echeance_de_la_spec(relue) == date(2027, 2, 28)
    assert facture.construire_groupes(relue["lignes"]) == [
        ("M1", "Statistiques", [(date(2026, 12, d), 3.0, 60.0) for d in (3, 10, 22)])]


def test_t_prp_30_colonne_jour_masquee(tmp_path):
    """T-PRP-30 — `masquer_jour = (regle.colonne_jour == "masquée")`, reporté dans la spec."""
    prep = preparer(installer(tmp_path, regles=[regle("ABC", **{"Colonne Jour": "masquée"}), regle("ECO")]))
    abc = brouillon(prep, "ABC-2026-12")
    assert abc.masquer_jour is True and abc.spec("R_ABC001")["masquer_jour"] is True
    assert brouillon(prep, "ECO-2026-12").masquer_jour is False


# ──────────────────────────────────────────────────────────────────────
# T-PRP-31 à T-PRP-34 : différences, numéros courts, aperçus
# ──────────────────────────────────────────────────────────────────────

_ECO_SANS_TARIF = [t for t in tarifs_standard() if t["École"] != "ECO"]
_ABC_TAUX_62 = [tarif("ABC", "Statistiques", "M1", 62) if t["École"] == "ABC" else t for t in tarifs_standard()]
_ABC_FIN_DECALEE = [s._replace(fin="12:30") if s.uid == "abc-1" else s for s in SEANCES_STANDARD]
_ABC_AJOUT = [*SEANCES_STANDARD, SeanceFictive("abc-4", "2026-12-17", "09:00", "12:00", "Statistiques M1", "#ABC-M1-STATS")]


def _adresse_abc_modifiee(chemin_master):
    changer_coordonnee(chemin_master, "Destinataire", "ABC", "Adresse L1", "12 avenue Test")


CAS_DIFFERENCES = {
    "taux modifié": ({}, {"tarifs": _ABC_TAUX_62}, {("ABC-2026-12", "taux modifié"), ("ABC-2026-12", "total modifié")}),
    "heures modifiées": ({}, {"seances": _ABC_FIN_DECALEE}, {("ABC-2026-12", "heures modifiées")}),
    "séance ajoutée": ({}, {"seances": _ABC_AJOUT}, {("ABC-2026-12", "séance ajoutée"), ("ABC-2026-12", "total modifié")}),
    "coordonnées modifiées": ({}, {"retouche": _adresse_abc_modifiee}, {("ABC-2026-12", "coordonnées modifiées")}),
    "facture nouvelle": ({"seances": seances_sauf("eco-1", "eco-2")}, {}, {("ECO-2026-12", "facture nouvelle")}),
    "facture disparue": ({}, {"seances": seances_sauf("eco-1", "eco-2")}, {("ECO-2026-12", "facture disparue")}),
    "nouveau blocage": ({}, {"tarifs": _ECO_SANS_TARIF}, {("ECO-2026-12", "nouveau blocage")}),
    "blocage levé": ({"tarifs": _ECO_SANS_TARIF}, {}, {("ECO-2026-12", "blocage levé")}),
    "échéance modifiée": (
        {}, {"regles": [regle("ABC", **{"Délai de paiement": "30 jours"}), regle("ECO"), regle("FUT", "Plus tard")]},
        {("ABC-2026-12", "échéance modifiée")}),
    "contrat modifié": (
        {}, {"regles": [regle("ABC", **{"Délai de paiement": "45 jours fin de mois", "N° de contrat": "CT-0009"}),
                        regle("ECO"), regle("FUT", "Plus tard")]},
        {("ABC-2026-12", "contrat modifié")}),
}


def _preparation_du_cas(tmp_path, nom: str, cas: dict):
    cas = dict(cas)
    seances = cas.pop("seances", SEANCES_STANDARD)
    retouche = cas.pop("retouche", None)
    p = installer(tmp_path / nom, **cas)
    if retouche is not None:
        retouche(p.master)
    return preparer(p, ics(seances))


@pytest.mark.parametrize("nature", list(CAS_DIFFERENCES))
def test_t_prp_31_differences_par_nature(tmp_path, nature):
    """T-PRP-31 — `differences` apparie les brouillons par identifiant et les lignes par `cle_seance` ;
    chaque nature du contrat est produite par le changement qui la cause."""
    ancienne_cas, nouvelle_cas, attendues = CAS_DIFFERENCES[nature]
    ancienne = _preparation_du_cas(tmp_path, "ancienne", ancienne_cas)
    nouvelle = _preparation_du_cas(tmp_path, "nouvelle", nouvelle_cas)

    resultat = prp().differences(ancienne, nouvelle)
    assert attendues <= {(d.identifiant, d.nature) for d in resultat}, resultat
    assert all(isinstance(d.detail, str) and d.detail for d in resultat)


def test_t_prp_31_differences_sans_changement_ou_premiere_preparation(tmp_path):
    """T-PRP-31 — Ancienne préparation absente : `[]` ; même préparation : `[]` ; ordre stable."""
    P = prp()
    p = installer(tmp_path)
    prep = preparer(p)
    assert list(P.differences(None, prep)) == []
    assert list(P.differences(prep, preparer(p, maintenant=datetime(2026, 12, 29, 9, 0, tzinfo=PARIS)))) == []
    autre = _preparation_du_cas(tmp_path, "ajout", {"seances": _ABC_AJOUT, "tarifs": _ECO_SANS_TARIF})
    assert P.differences(prep, autre) == P.differences(prep, autre)


def test_t_prp_32_numeros_courts(tmp_path):
    """T-PRP-32 — `numeros_courts` : « F<n> » → identifiant de chaque brouillon, « S<n> » →
    `cle_seance`, chaque séance facturée ou déjà facturée une seule fois ; stable d'une préparation à
    l'autre sur les mêmes entrées."""
    P = prp()
    a2 = standard("abc-2")
    p = installer(tmp_path, sessions=[ligne_sessions(a2, ecole="ABC", n_facture="R_ABC001")])
    prep = preparer(p)
    courts = P.numeros_courts(prep)

    factures = {k: v for k, v in courts.items() if k.startswith("F")}
    seances = {k: v for k, v in courts.items() if k.startswith("S")}
    assert set(courts) == set(factures) | set(seances)
    assert set(factures) == {f"F{i}" for i in range(1, len(prep.brouillons) + 1)}
    assert set(factures.values()) == identifiants(prep)
    assert all(re.fullmatch(r"S[1-9][0-9]*", k) for k in seances)
    assert len(set(seances.values())) == len(seances)
    assert cles_facturees(prep) <= set(seances.values())
    assert cle(a2) in set(seances.values())
    assert P.numeros_courts(preparer(p, maintenant=datetime(2026, 12, 29, 9, 0, tzinfo=PARIS))) == courts


def test_t_prp_33_apercu_facture(tmp_path):
    """T-PRP-33 — `apercu_facture` : N° « attribué à l'émission (probable : R_ABC001) », onglet nommé à
    l'émission, période, SIRET, IBAN, BIC, NDA, marque `[tenue]` / `[à venir]` et numéro court de chaque
    ligne, état, différences, empreinte et commande de validation."""
    P = prp()
    abc = brouillon(preparer(installer(tmp_path)), "ABC-2026-12")
    difference = P.Difference(identifiant="ABC-2026-12", nature="taux modifié", detail="60 → 62 €/h")
    texte = P.apercu_facture(abc, differences=[difference], etat="à valider")

    for attendu in (
        "attribué à l'émission", "R_ABC001", "nommé à l'émission", "Déc 2026", "ÉCOLE FICTIVE A",
        SIRET_FICTIF, IBAN_FICTIF, BIC_FICTIF, NDA_FICTIF, "à valider", "taux modifié", "540.00",
        f"Empreinte : {abc.empreinte[:12]}", "facturation valider", f"--empreinte {abc.empreinte[:8]}",
    ):
        assert attendu in texte, attendu
    assert texte.count("[tenue]") == 2 and texte.count("[à venir]") == 1
    assert len(set(re.findall(r"\bS[0-9]+\b", texte))) >= len(abc.lignes)


def test_t_prp_34_apercu_global(tmp_path):
    """T-PRP-34 — `apercu_global` : en-tête « Mois : … — entité … — préparé le JJ/MM/AAAA à HH:MM »,
    « À FAIRE » (valider, classer, écoles sans règle), tableau des FACTURES avec les états, file à
    classer, HORS LOT (nombre de séances par école, sans montant), puis les sections annexes."""
    P = prp()
    prep = preparer(installer(tmp_path))
    annexe = ("FACTURES ÉCHUES NON PAYÉES : 1", ["R_ABC001 — 12 jours de retard"])
    texte = P.apercu_global(prep, etats={"ABC-2026-12": "validée", "ECO-2026-12": "à valider"}, sections=[annexe])

    assert texte.splitlines()[0].startswith("Mois : décembre 2026 — entité ROSARI — préparé le 20/12/2026 à 16:00")
    for titre in ("À FAIRE", "FACTURES", "FILE À CLASSER", "HORS LOT"):
        assert titre in texte, titre
    a_faire = bloc(texte, "À FAIRE", "FACTURES")
    assert re.search(r"valider", a_faire, re.IGNORECASE) and re.search(r"classer", a_faire, re.IGNORECASE)
    assert "SRG" in a_faire
    assert "ABC-2026-12" in texte and "ECO-2026-12" in texte and "validée" in texte
    assert "Réunion pédagogique" in texte
    hors_lot = texte[texte.rindex("HORS LOT"):texte.index(annexe[0])]
    assert "FUT" in hors_lot and "€" not in hors_lot
    assert "€" in texte
    assert texte.index(annexe[0]) > texte.rindex("HORS LOT") and annexe[1][0] in texte


# ──────────────────────────────────────────────────────────────────────
# T-PRP-35 à T-PRP-37 et T-PRP-45 à T-PRP-47 : avertissements, registre, pureté
# ──────────────────────────────────────────────────────────────────────


def test_t_prp_35_chevauchement_avertit_sans_bloquer(tmp_path):
    """T-PRP-35 — Séances de la même école qui se chevauchent sans être identiques : avertissement,
    pas de blocage."""
    chevauche = SeanceFictive("abc-x", "2026-12-03", "11:00", "13:00", "Statistiques M1", "#ABC-M1-STATS")
    abc = brouillon(preparer(installer(tmp_path), ics([*SEANCES_STANDARD, chevauche])), "ABC-2026-12")
    assert re.search(r"chevauch", textes(abc.avertissements), re.IGNORECASE)
    assert abc.statut == "prêt"
    assert not re.search(r"chevauch|doublon", textes(abc.blocages), re.IGNORECASE)


def test_t_prp_35_avertissements_seances_a_venir_et_creneaux_a_classer(tmp_path):
    """T-PRP-35 — « N séance(s) à venir : recontrôle avant envoi (F-23) » et « N créneau(x) à classer
    dans la période »."""
    prep = preparer(installer(tmp_path))
    avertissements = textes(brouillon(prep, "ABC-2026-12").avertissements)
    assert re.search(r"\b1 séance", avertissements) and "à venir" in avertissements
    assert any(re.search(r"créneau.*à classer", textes(b.avertissements)) for b in prep.brouillons)


def test_t_prp_35_ecole_sans_seance(tmp_path):
    """T-PRP-35 — École du lot sans séance dans sa période : listée dans `ecoles_sans_seance`, sans
    brouillon."""
    prep = preparer(installer(tmp_path), ics(seances_sauf("eco-1", "eco-2")))
    assert "ECO" in prep.ecoles_sans_seance
    assert identifiants(prep) == {"ABC-2026-12"}


def test_t_prp_36_periode_deja_facturee_forme_longue(tmp_path):
    """T-PRP-36 (💡 [proposé — à valider]) — La ligne `Factures` dont la période est écrite sous la
    forme longue de l'historique (« Décembre 2026 », § 15) bloque comme « Déc 2026 » : sinon la garde
    serait sans effet sur les factures faites à la main."""
    abc = brouillon(preparer(installer(tmp_path, factures=[_facture_registre("R_ABC004", "ABC", "Décembre 2026")])),
                    "ABC-2026-12")
    assert abc.statut == "bloqué"
    assert re.search(r"déjà facturée", textes(abc.blocages), re.IGNORECASE)


def test_t_prp_37_piece_abandonnee_ne_bloque_pas(tmp_path):
    """T-PRP-37 — Pièce abandonnée : le brouillon est régénéré, son N° n'est jamais réutilisé
    (`numero_probable` suivant, ADR029) et l'empreinte ne dépend pas du numéro probable. Garde : la
    même pièce non abandonnée l'engage (T-PRP-13)."""
    p = installer(tmp_path)
    sans_piece = brouillon(preparer(p), "ABC-2026-12")
    ecrire_piece(p, "R_ABC001", identifiant="ABC-2026-12", empreinte=sans_piece.empreinte, brouillon_=sans_piece,
                 abandonnee=True)
    prep = preparer(p)
    abc = brouillon(prep, "ABC-2026-12")
    assert not prep.deja_emises
    assert abc.numero_probable == "R_ABC002"
    assert abc.empreinte == sans_piece.empreinte


def test_t_prp_45_serie_complete_bloque(tmp_path):
    """T-PRP-45 — Série complète (`prochain_numero` lève `ValueError`) : blocage, numéro probable
    absent ; les autres écoles restent prêtes."""
    prep = preparer(installer(tmp_path, factures=[_facture_registre("R_ABC999", "ABC", "Sept 2026")]))
    abc = brouillon(prep, "ABC-2026-12")
    assert abc.statut == "bloqué"
    assert re.search(r"R_ABC.*complète", textes(abc.blocages))
    assert abc.numero_probable is None
    assert brouillon(prep, "ECO-2026-12").statut == "prêt"


def test_t_prp_46_numeros_manquants_avertissement(tmp_path):
    """T-PRP-46 — Numéro manquant de la série sans ligne au registre (`numeros_manquants`, ADR029) :
    avertissement qui le cite, sans blocage ; le numéro probable suit le plus grand."""
    factures = [_facture_registre("R_ABC001", "ABC", "Sept 2026"), _facture_registre("R_ABC003", "ABC", "Oct 2026")]
    abc = brouillon(preparer(installer(tmp_path, factures=factures)), "ABC-2026-12")
    assert "R_ABC002" in textes(abc.avertissements)
    assert abc.statut == "prêt"
    assert abc.numero_probable == "R_ABC004"


def test_t_prp_47_preparer_pure_et_deterministe(tmp_path):
    """T-PRP-47 — `preparer` est pure et déterministe : deux appels sur les mêmes entrées rendent les
    mêmes brouillons et le même aperçu, sans rien écrire sur le disque."""
    P = prp()
    p = installer(tmp_path)
    e = entrees(p)
    avant = arborescence(tmp_path)
    une = P.preparer(entite="ROSARI", mois=MOIS, entrees=e, prefixe="R_", maintenant=MAINTENANT)
    deux = P.preparer(entite="ROSARI", mois=MOIS, entrees=e, prefixe="R_", maintenant=MAINTENANT)
    assert une == deux
    assert P.apercu_global(une) == P.apercu_global(deux)
    assert arborescence(tmp_path) == avant


# ──────────────────────────────────────────────────────────────────────
# T-PRP-48 à T-PRP-60 : constats de revue de F-16a (rouges avant le correctif)
# ──────────────────────────────────────────────────────────────────────

_REGLE_ABC_CT0001 = regle("ABC", **{"N° de contrat": "CT-0001", "Une facture par contrat": "oui"})
_TARIFS_ABC_CT0001 = [tarif("ABC", "Statistiques", "M1", 60, "CT-0001") if t["École"] == "ABC" else t
                      for t in tarifs_standard()]


@pytest.mark.parametrize("cas", ["école passée à une facture par contrat", "suffixe haché par un contrat voisin"])
def test_t_prp_48_piece_engagee_protege_ses_seances_si_l_identifiant_change(tmp_path, cas):
    """T-PRP-48 (constat de revue, haut) — Une pièce « en émission » n'a pas encore écrit ses séances
    dans `Sessions`. Si l'identifiant de son brouillon change ensuite (l'école passe à « Une facture
    par contrat = oui » ; une règle « CT 0001 » rend haché le suffixe de « CT-0001 »), ses séances ne
    repartent jamais dans un nouveau brouillon : aucune `cle_seance` de la pièce dans les brouillons
    (§ 4 étape 4 : « jamais régénérés » ; sinon, double facturation). La pièce reste listée dans
    `deja_emises`. Garde : sans la pièce, ces séances sont facturées sous le nouvel identifiant."""
    if cas == "école passée à une facture par contrat":
        avant, identifiant = {}, "ABC-2026-12"
        apres = dict(regles=[_REGLE_ABC_CT0001, regle("ECO"), regle("FUT", "Plus tard")], tarifs=_TARIFS_ABC_CT0001)
    else:
        avant = dict(regles=[_REGLE_ABC_CT0001, regle("ECO"), regle("FUT", "Plus tard")], tarifs=_TARIFS_ABC_CT0001)
        identifiant = "ABC-2026-12-CT0001"
        voisin = regle("ABC", **{"N° de contrat": "CT 0001", "Une facture par contrat": "oui"})
        apres = dict(regles=[_REGLE_ABC_CT0001, voisin, regle("ECO"), regle("FUT", "Plus tard")],
                     tarifs=_TARIFS_ABC_CT0001)
    p = installer(tmp_path, **avant)
    engage = brouillon(preparer(p), identifiant)
    cles_piece = {l.cle_seance for l in engage.lignes}
    assert cles_piece == {cle(standard(u)) for u in ("abc-1", "abc-2", "abc-3")}
    ecrire_piece(p, "R_ABC001", identifiant=identifiant, empreinte=engage.empreinte, brouillon_=engage)

    sans_piece = preparer(installer(tmp_path / "garde", **apres))
    assert identifiant not in identifiants(sans_piece)
    assert cles_piece <= cles_facturees(sans_piece)

    master_standard(p.racine, **apres)
    prep = preparer(p)
    assert not cles_piece & cles_facturees(prep), [
        (b.identifiant, b.statut, sorted(_par_uid(b))) for b in prep.brouillons]
    assert "R_ABC001" in {n for _, n, _ in prep.deja_emises}


def _abc_deux_contrats(statut_ct2: str, **colonnes_ct2) -> dict:
    """`_contrats_abc` dont la ligne du contrat CT-0002 porte un autre « Statut lot » (§ 16)."""
    donnees = _contrats_abc()
    donnees["regles"][1] = regle("ABC", statut_ct2, **{"N° de contrat": "CT-0002", "Une facture par contrat": "oui",
                                                        **colonnes_ct2})
    return donnees


@pytest.mark.parametrize(
    "statut, colonnes", [("Hors facturation", {}), ("Plus tard", {"Rythme": "fin de module"})],
    ids=["contrat hors facturation", "contrat plus tard en fin de module"],
)
def test_t_prp_49_statut_lot_de_chaque_ligne_de_contrat(tmp_path, statut, colonnes):
    """T-PRP-49 (constat de revue, haut ; `💡 [proposé — à valider]`, le contrat se tait) — Depuis le
    § 16, chaque ligne de contrat porte son « Statut lot ». La séance d'un contrat « Hors facturation »
    ou « Plus tard » d'une école du lot n'est jamais dans un brouillon prêt (hors lot, ou blocage qui
    demande de trancher), sinon double facture si ce contrat est facturé à la main. Et la ligne hors
    lot en « fin de module » ne retient pas le contrat « Lot 1 » : CT-0001 est préparé."""
    p = installer(tmp_path, **_abc_deux_contrats(statut, **colonnes))
    prep = preparer(p, ics([standard("abc-1"), SEANCE_M2, standard("eco-1")]))

    assert "ABC" not in {e.ecole for e in prep.ecoles_non_preparees}, prep.ecoles_non_preparees
    assert "abc-1" in _par_uid(brouillon(prep, "ABC-2026-12-CT0001"))
    prets = [b.identifiant for b in prep.brouillons if b.statut == "prêt" and "abc-m2" in _par_uid(b)]
    assert prets == [], f"séance d'un contrat « {statut} » dans un brouillon prêt : {prets}"


@pytest.mark.parametrize(
    "cas, libelle",
    [("SIRET vide", "SIRET"), ("IBAN vide", "IBAN"), ("BIC vide", "BIC"),
     ("adresse de l'émetteur vide", "Adresse L1"), ("email de l'émetteur vide", "Email"),
     ("téléphone de l'émetteur vide", "Téléphone")],
)
def test_t_prp_50_valeur_d_une_mention_obligatoire_vide_bloque(tmp_path, cas, libelle):
    """T-PRP-50 (constat de revue, haut) — La ligne est présente mais sa valeur est vide (cellule
    réellement vidée) : SIRET, IBAN ou BIC de la ligne `Info`, adresse, email ou téléphone de
    l'émetteur (ADR030 §1 et §10). Tous les brouillons sont bloqués, par un blocage qui nomme
    l'onglet Coordonnées et la donnée manquante ; aucune exception."""
    p = installer(tmp_path)
    if libelle in ("SIRET", "IBAN", "BIC"):
        changer_info(p.master, "ROSARI", libelle, None)
    else:
        changer_coordonnee(p.master, "Émetteur", "ROSARI", libelle, None)
    prep = preparer(p)

    assert identifiants(prep) == {"ABC-2026-12", "ECO-2026-12"}
    for b in prep.brouillons:
        assert b.statut == "bloqué", (cas, b.identifiant, b.info, b.emetteur)
        assert any("Coordonnées" in x and libelle in x for x in b.blocages), b.blocages


def test_t_prp_51_ligne_sessions_illisible_ne_laisse_pas_refacturer_en_silence(tmp_path):
    """T-PRP-51 (constat de revue, moyen) — Une ligne `Sessions` illisible (Date_séance « début déc »)
    qui facture peut-être abc-1 ne se réduit pas à un avertissement noyé : le brouillon ABC est
    bloqué, ou « À FAIRE » nomme la ligne (onglet Sessions, ligne 2). Garde : la même ligne lisible
    retire abc-1 (« exacte »)."""
    a1 = standard("abc-1")
    lisible = ligne_sessions(a1, ecole="ABC", n_facture="R_ABC001", debut_="09:00", fin_="12:00")
    prep = preparer(installer(tmp_path, sessions=[{**lisible, "Date_séance": "début déc"}]))

    abc = brouillon(prep, "ABC-2026-12")
    a_faire = bloc(prp().apercu_global(prep), "À FAIRE", "FACTURES")
    assert abc.statut == "bloqué" or re.search(r"Sessions.*ligne 2", a_faire), (abc.blocages, a_faire)

    garde = preparer(installer(tmp_path / "lisible", sessions=[lisible]))
    assert cle(a1) in {d.cle_seance for d in garde.deja_facturees}


def test_t_prp_52_reference_de_tarif_differente_du_contrat_de_la_regle(tmp_path):
    """T-PRP-52 (constat de revue, moyen) — Sans « Une facture par contrat », la ligne de tarif reliée
    à CT-0002 sur une école dont la seule règle porte CT-0001 : blocage qui cite les deux références
    et Règles_facturation (la mention du contrat ne peut pas être fausse, ADR030 §3 ; même écart
    bloquant qu'en facturation par contrat). Garde : tarif relié à CT-0001 → prêt, contrat CT-0001."""
    regles = [regle("ABC", **{"Délai de paiement": "45 jours fin de mois", "N° de contrat": "CT-0001"}), regle("ECO"),
              regle("FUT", "Plus tard")]
    tarifs = [tarif("ABC", "Statistiques", "M1", 60, "CT-0002") if t["École"] == "ABC" else t for t in tarifs_standard()]
    prep = preparer(installer(tmp_path, regles=regles, tarifs=tarifs))

    abc = brouillon(prep, "ABC-2026-12")
    assert abc.statut == "bloqué", (abc.contrat, abc.blocages)
    assert any("CT-0001" in x and "CT-0002" in x and "Règles_facturation" in x for x in abc.blocages), abc.blocages
    assert brouillon(prep, "ECO-2026-12").statut == "prêt"

    garde = brouillon(preparer(installer(tmp_path / "garde", regles=regles, tarifs=_TARIFS_ABC_CT0001)), "ABC-2026-12")
    assert (garde.statut, garde.contrat) == ("prêt", "CT-0001")


def test_t_prp_53_cle_tarif_d_une_autre_ecole(tmp_path):
    """T-PRP-53 (constat de revue, moyen) — Une ligne de `Correspondances` de l'école ABC dont la clé
    tarif est celle d'ECO (« ECO|Économie|L3 ») ne facture pas ABC au taux d'ECO : brouillon ABC
    bloqué, blocage qui cite la clé et l'onglet Correspondances ; ECO reste prêt. Garde : avec la clé
    ABC, le brouillon est prêt (T-PRP-01)."""
    corr = [correspondance("#ABC-M1-STATS", "code", "ABC", "Statistiques", "M1", cle_tarif="ECO|Économie|L3")
            if c["École"] == "ABC" else c for c in correspondances_standard()]
    prep = preparer(installer(tmp_path, correspondances=corr))

    abc = brouillon(prep, "ABC-2026-12")
    assert abc.statut == "bloqué", [(l.uid, l.cle_tarif, l.taux) for l in abc.lignes]
    assert any("ECO|Économie|L3" in x and "Correspondances" in x for x in abc.blocages), abc.blocages
    assert brouillon(prep, "ECO-2026-12").statut == "prêt"


@pytest.mark.parametrize("mode", ["forfait", "coefficient"])
def test_t_prp_54_heures_nulles_bloquent(tmp_path, mode):
    """T-PRP-54 (constat de revue, moyen) — Une correspondance à 0 h (forfait ou coefficient 0, que
    `heures_facturables` accepte) ne donne pas un brouillon prêt à 0,00 € que `facture.charger_spec`
    refuserait (§ 18.3, heures ≤ 0) : blocage qui parle des heures et nomme Correspondances. Garde
    générale : la spec de chaque brouillon prêt passe `charger_spec` puis `controler_spec`."""
    from facturation import facture

    corr = [correspondance("#ABC-M1-STATS", "code", "ABC", "Statistiques", "M1", mode=mode, valeur=0)
            if c["École"] == "ABC" else c for c in correspondances_standard()]
    p = installer(tmp_path, correspondances=corr)
    prep = preparer(p)

    abc = brouillon(prep, "ABC-2026-12")
    assert abc.statut == "bloqué", [(l.uid, l.heures) for l in abc.lignes]
    bl = textes(abc.blocages)
    assert re.search(r"heures", bl, re.IGNORECASE) and "Correspondances" in bl, abc.blocages
    for b in prep.brouillons:
        if b.statut != "prêt":
            continue
        chemin = tmp_path / f"spec_{b.identifiant}.json"
        chemin.write_text(json.dumps(b.spec(b.numero_probable), ensure_ascii=False), encoding="utf-8")
        facture.controler_spec(facture.charger_spec(chemin), prefixe=p.prefixe_facture, racine=p.racine)


_ABC_INTITULE_MODIFIE = [s._replace(intitule="Statistiques M1 (salle 2)") if s.uid == "abc-1" else s
                         for s in SEANCES_STANDARD]
_ABC_MODULE_RENOMME = [
    correspondance("#ABC-M1-STATS", "code", "ABC", "Statistiques appliquées", "M1", cle_tarif="ABC|Statistiques|M1")
    if c["École"] == "ABC" else c for c in correspondances_standard()
]
_ABC_ECOLE_RENOMMEE = [{**e, "École": "ÉCOLE FICTIVE A RENOMMÉE"} if e["Code facturation"] == "ABC" else e
                       for e in ecoles_standard()]
CAS_EMPREINTE_SANS_DIFFERENCE = {
    "intitulé": ({"seances": _ABC_INTITULE_MODIFIE}, r"intitulé|salle 2"),
    "libellé de la ligne": ({"correspondances": _ABC_MODULE_RENOMME}, r"libellé|Statistiques appliquées"),
    "libellé de l'école": ({"ecoles": _ABC_ECOLE_RENOMMEE}, r"libellé|RENOMMÉE"),
}


@pytest.mark.parametrize("champ", list(CAS_EMPREINTE_SANS_DIFFERENCE))
def test_t_prp_55_tout_changement_d_empreinte_a_sa_difference(tmp_path, champ):
    """T-PRP-55 (constat de revue, moyen) — L'intitulé d'une séance, le libellé (Module) d'une ligne
    ou le libellé de l'école changent l'empreinte, donc font « à revalider » : `differences` rapporte
    alors au moins une différence du brouillon, qui nomme le champ (ou sa nouvelle valeur) ; sinon
    l'aperçu dit « CHANGEMENTS : aucun » et pousse à revalider sans relire."""
    cas, motif = CAS_EMPREINTE_SANS_DIFFERENCE[champ]
    ancienne = _preparation_du_cas(tmp_path, "ancienne", {})
    nouvelle = _preparation_du_cas(tmp_path, "nouvelle", cas)
    assert brouillon(ancienne, "ABC-2026-12").empreinte != brouillon(nouvelle, "ABC-2026-12").empreinte

    propres = [d for d in prp().differences(ancienne, nouvelle) if d.identifiant == "ABC-2026-12"]
    assert propres, f"empreinte de ABC-2026-12 changée ({champ}) sans aucune différence rapportée"
    assert re.search(motif, textes(f"{d.nature} : {d.detail}" for d in propres), re.IGNORECASE), propres


def test_t_prp_56_periode_deja_facturee_saisie_en_date(tmp_path):
    """T-PRP-56 (constat de revue, haut) — Excel en français convertit la saisie « Déc 2026 » de
    `Factures.Période` en date (01/12/2026) : la garde « période déjà facturée hors outil » la
    reconnaît et bloque ABC (Q-D6). Garde : la date du 01/11/2026 (ECO) ne bloque pas."""
    factures = [_facture_registre("R_ABC004", "ABC", datetime(2026, 12, 1)),
                _facture_registre("R_ECO004", "ECO", datetime(2026, 11, 1))]
    prep = preparer(installer(tmp_path, factures=factures))

    abc = brouillon(prep, "ABC-2026-12")
    assert abc.statut == "bloqué" and re.search(r"déjà facturée", textes(abc.blocages), re.IGNORECASE), abc.blocages
    assert not re.search(r"déjà facturée", textes(brouillon(prep, "ECO-2026-12").blocages), re.IGNORECASE)


@pytest.mark.parametrize("colonne", ["École", "Date_séance"])
def test_t_prp_57_ligne_sessions_sans_ecole_ou_sans_date(tmp_path, colonne):
    """T-PRP-57 (constat de revue, moyen) — Une ligne `Sessions` sans École (ou sans Date_séance), au
    N° et aux heures d'abc-1, n'est pas ignorée en silence par l'anti-doublon (ADR005) : abc-1 est
    retirée (`deja_facturees`), ou le brouillon ABC signale la ligne 2 de Sessions."""
    a1 = standard("abc-1")
    ligne = {**ligne_sessions(a1, ecole="ABC", n_facture="R_ABC001", debut_="09:00", fin_="12:00"), colonne: None}
    prep = preparer(installer(tmp_path, sessions=[ligne]))

    abc = brouillon(prep, "ABC-2026-12")
    signalee = re.search(r"Sessions.*ligne 2", textes([*abc.blocages, *abc.avertissements]))
    assert cle(a1) in {d.cle_seance for d in prep.deja_facturees} or signalee, (abc.blocages, abc.avertissements)


@pytest.mark.parametrize("cas", ["N° de contrat de la règle", "Réf. contrat du tarif"])
def test_t_prp_58_periode_deja_facturee_sans_contrat_dans_factures(tmp_path, cas):
    """T-PRP-58 (constat de revue, moyen) — Une facture faite à la main, inscrite dans `Factures` sans
    Contrat (cas des lignes historiques, Q-D6), pour la période d'un brouillon qui cite un contrat
    (N° de la règle, ou Réf. contrat commune des tarifs) : au moins un avertissement « période déjà
    facturée » (ou un blocage). `💡 [proposé — à valider]` : « vide égal vide » est tenu à la lettre
    pour le blocage ; un contrat vide dans `Factures` ne doit pas rendre la garde muette."""
    if cas == "N° de contrat de la règle":
        regles = [regle("ABC", **{"Délai de paiement": "45 jours fin de mois", "N° de contrat": "CT-0009"}),
                  regle("ECO"), regle("FUT", "Plus tard")]
        surcharges, identifiant, n, ecole = dict(regles=regles), "ABC-2026-12", "R_ABC004", "ABC"
    else:
        tarifs = [tarif("ECO", "Économie", "L3", "45 €/h", "CT-0003") if t["École"] == "ECO" else t
                  for t in tarifs_standard()]
        surcharges, identifiant, n, ecole = dict(tarifs=tarifs), "ECO-2026-12", "R_ECO004", "ECO"
    prep = preparer(installer(tmp_path, factures=[_facture_registre(n, ecole, "Déc 2026")], **surcharges))

    b = brouillon(prep, identifiant)
    assert b.contrat, "garde : le brouillon doit citer un contrat"
    assert re.search(r"déjà facturée", textes([*b.blocages, *b.avertissements]), re.IGNORECASE), (
        b.blocages, b.avertissements)


def test_t_prp_59_seance_exclue_non_reconnue_reste_visible(tmp_path):
    """T-PRP-59 (constat de revue, faible ; `💡 [proposé — à valider]`) — Une séance non reconnue
    exclue par décision (« Réunion pédagogique ») ne disparaît pas de tout l'aperçu global : seules
    les exclusions d'une école connue y étaient reprises (avertissement du brouillon). Garde : la
    décision s'applique (ni à classer ni facturée)."""
    p = installer(tmp_path)
    inconnu = standard("inconnu-1")
    ecrire_decision(p, "exclure", cle(inconnu), {"raison": "réunion interne"})
    prep = preparer(p)

    assert "inconnu-1" not in {a.seance.uid for a in prep.a_classer} | uids_factures(prep)
    assert "Réunion pédagogique" in prp().apercu_global(prep)


def test_t_prp_60_decision_classer_vers_une_ecole_inconnue(tmp_path):
    """T-PRP-60 (constat de revue, faible) — Une décision `classer` dont l'école n'existe plus dans
    Écoles ne fait pas échouer toute la préparation (« une erreur propre à une école ne retient
    jamais les autres ») : la séance revient dans la file à classer, avec une raison qui cite la
    décision ou l'école ; ABC et ECO restent prêts."""
    p = installer(tmp_path)
    inconnu = standard("inconnu-1")
    ecrire_decision(p, "classer", cle(inconnu), {
        "ecole": "ZZ9 Inconnue", "module": "Réunion", "promotion": "M1", "cle_tarif": "ZZ9|Réunion|M1",
        "mode_heures": "agenda", "valeur": "",
    })
    prep = preparer(p)

    assert {brouillon(prep, i).statut for i in ("ABC-2026-12", "ECO-2026-12")} == {"prêt"}
    raisons = {a.seance.uid: a.raison for a in prep.a_classer}
    assert "inconnu-1" in raisons and re.search(r"ZZ9|décision", raisons["inconnu-1"]), raisons
    assert "inconnu-1" not in uids_factures(prep)
