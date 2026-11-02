# Logement et itinérance au Québec

Infographie interactive qui explique, en un coup d'œil, comment le manque de logements mène des personnes à la rue, dans toutes les régions du Québec.

**La chaîne des causes**, de gauche à droite :
pas assez de logements construits → peu de logements libres → les loyers montent → évictions et loyers impayés → itinérance.

**Trois actions sur le schéma** :
1. **Construire** plus de logements, privés et sociaux;
2. **Prévenir les évictions** avant la perte du logement;
3. **Reloger tout de suite** (Logement d'abord), mais seulement s'il reste des logements libres.

Un bloc « Ce qui rend plus vulnérable » (pauvreté, santé mentale, dépendances, sortie de placement) entre dans l'étape des évictions : ces facteurs décident **qui** perd son logement en premier, alors que le manque de logements décide **combien** de personnes se retrouvent à la rue.

La flèche en pointillé (acheter coûte plus cher → moins de logements libres) est un lien souvent affirmé, mais pas démontré dans les études recensées.

Sous le schéma, les solutions sont classées en trois colonnes : *ça marche*, *ça marche à moitié*, *ça ne règle pas le problème*. Les étoiles (★ à ★★★) indiquent à quel point chaque constat est prouvé.

Un clic sur une étape, une action ou une solution ouvre une fiche : l'idée en clair, les chiffres du Québec, ce que montrent les études (avec leurs références) et les sources.

Ouvrir `index.html` dans un navigateur. Aucune dépendance locale (police Google chargée en ligne, avec repli système).

Sources : études publiées (citées dans chaque fiche) et données publiques (SCHL, ISQ, MSSS, Statistique Canada, SHQ, APCIQ).

## Cartographie des parties prenantes

Page : [`parties-prenantes.html`](https://syl77qc.github.io/habitation_itinerance/parties-prenantes.html) · retour à l'[infographie](https://syl77qc.github.io/habitation_itinerance/).

Qui agit sur chaque maillon de la chaîne des causes, et ce que chacun a dit publiquement. Quatre vues : la carte des familles, les alliances par enjeu, les liens et fractures, et un tableau filtrable des positions. Chaque position est cotée de −2 à +2 et appuyée par une citation tirée d'un document public (mémoires de la consultation nationale santé mentale, itinérance et dépendance 2026, mémoires prébudgétaires 2026-2027, communiqués).

Document de travail personnel : codage préliminaire, non validé. Il ne représente la position d'aucun ministère ni du gouvernement du Québec.

La source des données est le classeur `Parties_prenantes_habitation_itinerance.xlsx`. Après une modification, régénérer la page avec `python outils/generer_page.py` (requiert `openpyxl`).
