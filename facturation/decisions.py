"""decisions.py — Décisions ponctuelles d'Adrian sur un mois préparé (F-16a : type et lecture).

Une décision est un événement `decision` du journal du mois (`_sorties/<AAAA-MM>/journal.jsonl`,
schéma figé par `facturation.journal`). Elle vaut pour ce mois seulement ; la règle gardée, qui
vaut pour les mois suivants, s'écrit dans l'onglet `Correspondances` (F-16b, ADR028).

Natures (`NATURES`) et cible :
- `classer` : une séance (`cle_seance`) reconnue pour ce mois ; `valeurs` porte `ecole`, `module`,
  `promotion`, `cle_tarif`, `mode_heures` et `valeur`, en texte (`valeur` vide en mode agenda) ;
- `exclure` : une séance retirée de la préparation, ni facturée ni à classer ; `valeurs["raison"]` ;
- `inclure` : une séance « même jour sans heure » réintégrée (cible `cle_seance`), ou le blocage
  « période déjà facturée hors outil » levé (cible : identifiant de brouillon) ;
- `bon_commande` : la référence du bon de commande d'un brouillon, `valeurs["bon_commande"]` ;
- `annuler` : une décision précédente annulée (cible : son `id`).

L'écriture des décisions vient avec F-16b ; ce module ne fait que les lire.

Lectures retenues (`💡 [proposé — à valider]`) : une décision `annuler` n'est pas rendue parmi les
décisions actives (elle a produit son effet) ; une décision rendue deux fois sous le même `id`
compte une fois, avec son dernier contenu ; une nature inconnue est rendue telle quelle (la
préparation l'ignore) ; un événement d'une autre entité que le profil est refusé (`ValueError`,
§ 1.1) ; `Decision` porte aussi le `commentaire` de l'événement (champ facultatif, en dernier).
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from typing import TYPE_CHECKING

from facturation import chemins, journal

if TYPE_CHECKING:
    from facturation.config import Profil

NATURES = ("classer", "exclure", "inclure", "bon_commande", "annuler")
CLASSER, EXCLURE, INCLURE, BON_COMMANDE, ANNULER = NATURES


@dataclass(frozen=True)
class Decision:
    """Une décision du journal du mois.

    `cible` : une `cle_seance`, un identifiant de brouillon ou l'`id` d'une décision, selon la
    nature ; `valeurs` : les valeurs de la décision, en texte.
    """

    id: str
    nature: str
    cible: str
    valeurs: dict[str, str]
    mois: str
    horodatage: datetime
    poste: str
    commentaire: str = ""


def _texte(valeur) -> str:
    return "" if valeur is None else str(valeur)


def _decision(evenement: dict, chemin) -> Decision:
    """Une décision lue d'un événement ; `ValueError` qui nomme le journal si elle ne se lit pas."""
    try:
        horodatage = datetime.fromisoformat(str(evenement["horodatage"]))
    except (KeyError, ValueError) as exc:
        raise ValueError(
            f"Journal {chemin} : décision « {evenement.get('id')} » à l'horodatage illisible "
            f"({evenement.get('horodatage')!r})."
        ) from exc
    valeurs = evenement.get("valeurs") or {}
    if not isinstance(valeurs, dict):
        raise ValueError(f"Journal {chemin} : décision « {evenement.get('id')} » aux valeurs illisibles ({valeurs!r}).")
    return Decision(
        id=_texte(evenement.get("id")),
        nature=_texte(evenement.get("nature")),
        cible=_texte(evenement.get("cible")),
        valeurs={str(cle): _texte(valeur) for cle, valeur in valeurs.items()},
        mois=_texte(evenement.get("mois")),
        horodatage=horodatage,
        poste=_texte(evenement.get("poste")),
        commentaire=_texte(evenement.get("commentaire")),
    )


def decisions_actives(profil: "Profil", mois: str) -> list[Decision]:
    """Décisions du journal du mois, dans l'ordre du journal, moins les décisions annulées.

    Seules comptent les décisions dont le champ `mois` est `mois`. Journal absent → [].

    Raises:
        ValueError: mois invalide ; événement d'une autre entité que le profil ; décision illisible.
        fichiers.FichierAbime: ligne illisible au milieu du journal.
    """
    chemin = chemins.journal_mois(profil, mois)
    actives: dict[str, Decision] = {}
    for evenement in journal.lire_evenements(chemin, types=["decision"]):
        entite = evenement.get("entite")
        if entite is not None and entite != profil.entite:
            raise ValueError(
                f"Journal {chemin} : décision de l'entité « {entite} », alors que le profil est celui de "
                f"« {profil.entite} ». Les entités ne se mélangent jamais (ADR001)."
            )
        if evenement.get("mois") != mois:
            continue
        decision = _decision(evenement, chemin)
        if decision.nature == ANNULER:
            actives.pop(decision.cible, None)
            continue
        actives.pop(decision.id, None)
        actives[decision.id] = decision
    return list(actives.values())
