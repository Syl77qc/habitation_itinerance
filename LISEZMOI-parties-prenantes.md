# Parties prenantes — fichiers à ajouter au dépôt

À la racine du dépôt `habitation_itinerance` :

| Fichier | Rôle |
|---|---|
| `parties-prenantes.html` | La page web (carte, alliances, liens, tableau, forum). Générée, ne pas modifier à la main. |
| `Parties_prenantes_habitation_itinerance.xlsx` | La source des données. C'est ici qu'on corrige et qu'on valide. |
| `index.html` | L'infographie, avec les liens vers la carte des parties prenantes et, dans les fiches des actions, vers les propositions du forum. |
| `outils/generer_page.py` | Refait la page à partir du classeur. |
| `outils/gabarit_parties_prenantes.html` | Le modèle de la page utilisé par le script. |

## Mettre la page à jour après une modification du classeur

1. Enregistrer le classeur (il doit rester à la racine, avec le même nom).
2. Dans le dossier du dépôt : `python outils/generer_page.py` (il faut `pip install openpyxl` une première fois).
3. Pousser `parties-prenantes.html` et le classeur sur GitHub.

Le script vérifie que chaque position et chaque lien pointe vers une organisation de l'onglet « Parties prenantes ».

La vue « Forum » lit les onglets Propositions × familles (dont la colonne Bloc), Preuve et consensus, Blocs et Positions forum (détail). Les demandes de la campagne électorale 2026 sont codées dans l’onglet Campagne 2026, qui les compare aux positions d’avant la campagne; elles sont déjà intégrées à la matrice et aux sources de l’onglet Positions forum (détail). Les zones sont des formules : si le classeur a été modifié ailleurs qu'Excel, l'ouvrir et l'enregistrer dans Excel avant de lancer le script, pour que les valeurs soient recalculées.
