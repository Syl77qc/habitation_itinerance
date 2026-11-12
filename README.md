# Logement et itinérance au Québec

Infographie interactive qui explique, en un coup d'œil, comment le manque de logements mène des personnes à la rue, dans toutes les régions du Québec.

**La chaîne des causes**, de gauche à droite :
pas assez de logements construits → peu de logements libres → les loyers montent → évictions et loyers impayés → itinérance.

**Trois actions sur le schéma** :
1. **Construire** plus de logements, privés et sociaux;
2. **Prévenir les évictions** avant la perte du logement;
3. **Reloger tout de suite** (Logement d'abord), mais seulement s'il reste des logements libres.

Un bloc « Ce qui rend plus vulnérable » (pauvreté, santé mentale, dépendances, sortie de placement) entre dans l'étape des évictions : ces facteurs décident **qui** perd son logement en premier, alors que le manque de logements décide **combien** de personnes se retrouvent à la rue.

La flèche en pointillé (acheter coûte plus cher → moins de logements libres) est un lien observé par la SCHL, mais pas chiffré : quand les locataires ne peuvent plus acheter, la chaîne des déménagements ralentit et il se libère moins de logements. Le contrôle des loyers est présenté en deux fiches : le gel ou le plafond strict (ça ne règle pas le problème) et l'encadrement souple du TAL (jamais évalué).

Sous le schéma, les solutions sont classées en trois colonnes : *ça marche*, *ça marche à moitié*, *ça ne règle pas le problème*. Les étoiles (★ à ★★★) indiquent à quel point chaque constat est prouvé.

La colonne *ça marche* comprend aussi deux pistes que les acteurs ne proposent pas : repérer qui est vraiment à risque avant de verser l'aide d'urgence, et une aide souple après une violence conjugale. Sous les trois colonnes, une section « Ce que personne ne propose » présente quatre constats appuyés par des études, mais absents des demandes des acteurs : trois profils d'itinérance, le vieillissement, la location touristique de courte durée et la demande démographique.

Un clic sur une étape, une action ou une solution ouvre une fiche : l'idée en clair, les chiffres du Québec, ce que montrent les études (avec leurs références) et les sources.

Ouvrir `index.html` dans un navigateur. Aucune dépendance locale (police Google chargée en ligne, avec repli système).

Sources : études publiées (citées dans chaque fiche) et données publiques (SCHL, ISQ, MSSS, Statistique Canada, SHQ, APCIQ).

## Cartographie des parties prenantes

Page : [`parties-prenantes.html`](https://syl77qc.github.io/habitation_itinerance/parties-prenantes.html) · retour à l'[infographie](https://syl77qc.github.io/habitation_itinerance/).

Qui agit sur chaque maillon de la chaîne des causes, et ce que chacun a dit publiquement. Cinq vues : la carte des familles, les alliances par enjeu, les liens et fractures, un tableau filtrable des positions et une vue **Forum** : 28 propositions que les familles d'acteurs feront probablement au forum sur l'itinérance, comparées à la littérature scientifique et classées en zones (consensus possible, consensus fragile, à porter, à retravailler, avec le motif de chaque proposition à retravailler; chaque consensus possible porte aussi une étiquette de solidité : robuste, probable ou à confirmer) et regroupées en cinq blocs (construire, prévenir, reloger, gouverner, populations particulières) pour repérer les ententes partielles possibles. Ces propositions sont des inférences tirées de documents publics, pas des positions officielles. Chaque position est cotée de −2 à +2 et appuyée par une citation tirée d'un document public (mémoires de la consultation nationale santé mentale, itinérance et dépendance 2026, mémoires prébudgétaires 2026-2027, plateformes électorales et communiqués de la campagne électorale 2026). Dans l'infographie, les fiches des actions renvoient aux propositions du forum qui s'y rattachent.

Document de travail personnel : codage préliminaire, non validé. Il ne représente la position d'aucun ministère ni du gouvernement du Québec.

La source des données est le classeur `Parties_prenantes_habitation_itinerance.xlsx`. Après une modification, régénérer la page avec `python outils/generer_page.py` (requiert `openpyxl`).

## Révision des programmes de la SHQ

Page : [`revision-programmes/index-programmes-shq.html`](https://syl77qc.github.io/habitation_itinerance/revision-programmes/index-programmes-shq.html) · retour à l'[infographie](https://syl77qc.github.io/habitation_itinerance/).

Cartographies de processus des 14 programmes actifs de la SHQ (qui agit, à quelle étape, sur quel fondement), tirées des cadres normatifs publiés dans le [répertoire des documents servant à la prise de décision de la SHQ](https://www.quebec.ca/gouvernement/ministeres-organismes/societe-habitation/acces-a-linformation-a-la-societe-dhabitation-du-quebec/acces-information/documents-servant-a-la-prise-de-decision). Une matrice comparative (`simplification-programmes-shq.html`) propose des pistes de simplification reliées aux exigences actuelles, avec l'analyse complète (`analyse-simplification-shq.md`) et les données structurées (`matrice-simplification-shq.json`).

Les liens « PDF » renvoient aux documents officiels de la SHQ; la version consultée le 7 novembre 2026 et son empreinte figurent dans `sources-et-perimetre-shq.json`. La SHQ peut avoir remplacé un document depuis.

Document de travail personnel : analyse documentaire préliminaire, non validée. Ce n'est pas une publication de la SHQ et il ne représente la position d'aucun ministère ni du gouvernement du Québec.
