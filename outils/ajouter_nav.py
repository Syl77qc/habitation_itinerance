# -*- coding: utf-8 -*-
"""Ajoute (ou remet à jour) la barre de navigation commune en haut des pages du dépôt.

Usage, à la racine du dépôt : python outils/ajouter_nav.py

Pages statiques traitées ici : index.html, tableau-de-bord.html et revision-programmes/*.html.
(tableau-de-bord.html est la version publique du tableau de bord, produite par le dossier
« Tableau de bord » : copier Tableau_de_bord_habitation_public.html sous ce nom, puis relancer ce script.)
Les pages générées ont la barre dans leur source :
- parties-prenantes.html : outils/gabarit_parties_prenantes.html
- ecosysteme-habitation.html : outils/generer_ecosysteme.py
Pour modifier la barre partout, changer PAGES ci-dessous, relancer ce script,
puis reporter le même bloc dans le gabarit et dans le script de l'écosystème.

Le bloc est encadré par <!-- nav-site:debut --> et <!-- nav-site:fin -->;
relancer le script le remplace au lieu de le dupliquer.
"""
import re
import sys
from pathlib import Path

BASE = "https://syl77qc.github.io/habitation_itinerance/"
PAGES = [
    ("infographie", "Infographie", BASE),
    ("parties", "Parties prenantes", BASE + "parties-prenantes.html"),
    ("ecosysteme", "Écosystème", BASE + "ecosysteme-habitation.html"),
    ("revision", "Révision des programmes", BASE + "revision-programmes/index-programmes-shq.html"),
    ("tableau", "Tableau de bord", BASE + "tableau-de-bord.html"),
]

STYLE = (
    '<style>'
    '.site-nav{box-sizing:border-box;display:flex;flex-wrap:wrap;align-items:center;gap:4px 6px;padding:8px 16px;'
    'background:#172033;border-bottom:1px solid #2f3d55;font:14px/1.3 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;'
    'position:relative;z-index:50;margin:0 0 8px}'
    '.site-nav .site-nav-titre{color:#c9d3e0;font-weight:600;font-size:13px;margin-right:10px}'
    '.site-nav a.site-nav-lien{color:#e6edf3;text-decoration:none;padding:5px 11px;border-radius:999px;'
    'border:1px solid transparent;white-space:nowrap;font-weight:400}'
    '.site-nav a.site-nav-lien:hover{border-color:#4a6379}'
    '.site-nav a.site-nav-lien[aria-current="page"]{background:#e6edf3;color:#172033;font-weight:600}'
    '.site-nav a.site-nav-lien:focus-visible{outline:2px solid #7cc4f0;outline-offset:2px}'
    '@media print{.site-nav{display:none}}'
    '</style>'
)


def bloc(courant):
    liens = "".join(
        '<a class="site-nav-lien" href="%s"%s>%s</a>' % (url, ' aria-current="page"' if cle == courant else "", nom)
        for cle, nom, url in PAGES
    )
    return ('<!-- nav-site:debut -->\n' + STYLE + '\n<div class="site-nav" role="navigation" aria-label="Pages du projet">'
            '<span class="site-nav-titre">Logement et itinérance au Québec</span>' + liens + '</div>\n<!-- nav-site:fin -->\n')


MARQUES = re.compile(r"<!-- nav-site:debut -->.*?<!-- nav-site:fin -->\n?", re.S)


def poser(texte, courant):
    """Retire l'ancien bloc s'il existe, puis insère le nouveau juste après <body>."""
    texte = MARQUES.sub("", texte)
    m = re.search(r"<body[^>]*>\n?", texte)
    if not m:
        raise ValueError("balise <body> introuvable")
    return texte[:m.end()] + bloc(courant) + texte[m.end():]


def main():
    racine = Path(__file__).resolve().parent.parent
    cibles = [(racine / "index.html", "infographie"), (racine / "tableau-de-bord.html", "tableau")]
    cibles += [(p, "revision") for p in sorted((racine / "revision-programmes").glob("*.html"))]
    for chemin, courant in cibles:
        avant = chemin.read_text(encoding="utf-8")
        apres = poser(avant, courant)
        if apres != avant:
            chemin.write_text(apres, encoding="utf-8")
        print("barre posée :", chemin.relative_to(racine))


if __name__ == "__main__":
    sys.exit(main())
