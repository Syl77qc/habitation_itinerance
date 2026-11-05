"""Régénère parties-prenantes.html à partir du classeur Excel.

Utilisation (dans le dossier du dépôt) :
    python outils/generer_page.py

Le script lit Parties_prenantes_habitation_itinerance.xlsx (à la racine du dépôt),
insère les données dans outils/gabarit_parties_prenantes.html et écrit
parties-prenantes.html à la racine. Il faut Python 3 et openpyxl
(pip install openpyxl).
"""
import json, re, sys, datetime
from pathlib import Path
from openpyxl import load_workbook

RACINE = Path(__file__).resolve().parent.parent
CLASSEUR = RACINE / "Parties_prenantes_habitation_itinerance.xlsx"
GABARIT = Path(__file__).resolve().parent / "gabarit_parties_prenantes.html"
SORTIE = RACINE / "parties-prenantes.html"

PLACES = {"Transversal": "top", "Construire": "c1", "Logements libres et loyers": "c2", "Évictions": "c3",
          "Ce qui rend vulnérable": "c4", "Itinérance": "c5", "Reloger et logement social": "c6", "Autres acteurs": "autres"}


def lignes(ws):
    """Retourne les lignes d'un onglet sous forme de dictionnaires {en-tête: valeur}."""
    it = ws.iter_rows(values_only=True)
    entetes = [str(h).strip() if h is not None else "" for h in next(it)]
    for r in it:
        if r is None or all(v in (None, "") for v in r):
            continue
        yield dict(zip(entetes, r))


def texte(v):
    return "" if v is None else str(v).strip()


def lire_forum(wb):
    """Lit les onglets du forum (propositions probables, preuve, zones). Retourne None s'ils sont absents."""
    if not all(n in wb.sheetnames for n in ("Propositions × familles", "Preuve et consensus", "Positions forum (détail)")):
        return None
    ws = wb["Propositions × familles"]
    entetes = [texte(c.value) for c in ws[1]]
    ib = entetes.index("Bloc") if "Bloc" in entetes else len(entetes)
    fams = entetes[3:ib]
    props = {}
    for r in ws.iter_rows(min_row=2, values_only=True):
        pid = texte(r[0])
        if not re.fullmatch(r"P\d+", pid):
            continue
        props[pid] = dict(id=pid, t=texte(r[1]), th=texte(r[2]), codes=[texte(v) for v in r[3:3 + len(fams)]],
                          b=texte(r[ib]) if ib < len(r) else "")
    for r in lignes(wb["Preuve et consensus"]):
        pid = texte(r.get("ID"))
        if pid not in props:
            continue
        num = lambda k: int(r.get(k) or 0) if isinstance(r.get(k), (int, float)) else 0
        props[pid].update(ap=num("Appui (P+A)"), co=num("Conditionnel (C)"), op=num("Oppositions (O)"),
                          pr=texte(r.get("Niveau de preuve")), z=texte(r.get("Zone")), lit=texte(r.get("Ce que dit la littérature")),
                          mo=texte(r.get("Motif (à retravailler)")), pi=texte(r.get("Piste")))
        if not props[pid]["z"]:
            sys.exit(f"Zone vide pour {pid} : ouvrir le classeur dans Excel et l'enregistrer pour recalculer les formules.")
    det = []
    for r in lignes(wb["Positions forum (détail)"]):
        pid = texte(r.get("ID"))
        if pid in props:
            det.append(dict(p=pid, f=texte(r.get("Famille")), c=texte(r.get("Code")), k=texte(r.get("Confiance")),
                            s=texte(r.get("Sources (document, page)"))))
    themes = {}
    if "Grille des thèmes" in wb.sheetnames:
        for r in wb["Grille des thèmes"].iter_rows(min_row=2, values_only=True):
            t = texte(r[0])
            if t[:1] == "T" and t[1:2].isdigit():
                themes[t.split(" ")[0]] = t.split(" ", 1)[1]
    blocs = []
    if "Blocs" in wb.sheetnames:
        for r in wb["Blocs"].iter_rows(min_row=4, values_only=True):
            if re.fullmatch(r"B\d+", texte(r[0])):
                blocs.append(dict(code=texte(r[0]), nom=texte(r[1])))
    return dict(fams=fams, props=list(props.values()), det=det, themes=themes, blocs=blocs)


def main():
    if not CLASSEUR.exists():
        sys.exit(f"Classeur introuvable : {CLASSEUR}")
    wb = load_workbook(CLASSEUR, data_only=True)

    fams, nom2code = [], {}
    for r in lignes(wb["Familles"]):
        code, nom = texte(r.get("Code")), texte(r.get("Famille"))
        if not code.startswith("F"):
            continue  # ligne Total ou exemple
        place = texte(r.get("Place sur la carte web")) or "Autres acteurs"
        if place not in PLACES:
            sys.exit(f"Place inconnue pour {nom} : « {place} ». Valeurs permises : {', '.join(PLACES)}")
        fams.append(dict(code=code, nom=nom, sous=texte(r.get("Sous-familles")),
                         maillon=texte(r.get("Maillon(s) de la chaîne")), col=PLACES[place]))
        nom2code[nom] = code

    enj = [dict(code=texte(r.get("Code")), nom=texte(r.get("Enjeu")),
                aff=texte(r.get("Affirmation de référence (la cote mesure l'accord avec cette phrase)")),
                maillon=texte(r.get("Maillon")))
           for r in lignes(wb["Enjeux"]) if texte(r.get("Code")).startswith("E")]

    orgs, ids = [], set()
    for r in lignes(wb["Parties prenantes"]):
        i = texte(r.get("ID"))
        if not i:
            continue
        fam = nom2code.get(texte(r.get("Famille")))
        if not fam:
            sys.exit(f"Famille inconnue pour {i} : « {r.get('Famille')} »")
        orgs.append(dict(id=i, nom=texte(r.get("Organisation")), sig=texte(r.get("Sigle")), fam=fam,
                         sous=texte(r.get("Sous-famille")), reg=texte(r.get("Région")),
                         doc=texte(r.get("Source principale")) not in ("", "À documenter"), note=texte(r.get("Note"))))
        ids.add(i)

    pos = []
    for r in lignes(wb["Positions"]):
        o, e = texte(r.get("ID organisation")), texte(r.get("Code enjeu"))
        if not o:
            continue
        if o not in ids:
            sys.exit(f"Position pour une organisation absente de l'onglet Parties prenantes : {o}")
        try:
            c = int(r.get("Cote"))
        except (TypeError, ValueError):
            sys.exit(f"Cote invalide pour {o} · {e} : {r.get('Cote')}")
        pg = r.get("Page")
        pos.append(dict(o=o, e=e, c=c, k=texte(r.get("Confiance")), t=texte(r.get("Citation (textuelle)")),
                        pg=int(pg) if isinstance(pg, (int, float)) else texte(pg), d=texte(r.get("Document")),
                        u=texte(r.get("Lien")), n=texte(r.get("Note du codeur")), v=texte(r.get("Validation (à remplir)"))))

    links = []
    if "Liens observés" in wb.sheetnames:
        for r in lignes(wb["Liens observés"]):
            membres = [x.strip() for x in texte(r.get("Organisations (ID séparés par « ; »)")).split(";") if x.strip()]
            inconnus = [m for m in membres if m not in ids]
            if inconnus:
                sys.exit(f"Lien « {r.get('Lien')} » : ID inconnu(s) {inconnus}")
            links.append(dict(nom=texte(r.get("Lien")), orgs=membres, src=texte(r.get("Source"))))

    forum = lire_forum(wb)
    data = dict(fams=fams, orgs=orgs, pos=pos, enj=enj, links=links, forum=forum)
    js = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    date = datetime.date.today().strftime("%Y-%m-%d")
    html = GABARIT.read_text(encoding="utf-8").replace("__DATA__", js).replace("__DATE__", date)
    SORTIE.write_text(html, encoding="utf-8")
    print(f"Page écrite : {SORTIE.name} · {len(orgs)} parties prenantes · {len(pos)} positions · {len(links)} liens"
          + (f" · {len(forum['props'])} propositions du forum" if forum else ""))


if __name__ == "__main__":
    main()
