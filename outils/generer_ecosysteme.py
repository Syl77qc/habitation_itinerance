# -*- coding: utf-8 -*-
"""Génère ecosysteme-habitation.html à partir de Ecosysteme_Habitation_fonctions.xlsx.

Usage : python generer_ecosysteme.py (dans le dépôt : python outils/generer_ecosysteme.py)
Le classeur est cherché à côté du script, puis un dossier plus haut. La page est écrite à côté du classeur.
Nécessite openpyxl.
"""
import json, os, sys
from openpyxl import load_workbook

ICI = os.path.dirname(os.path.abspath(__file__))
NOM = "Ecosysteme_Habitation_fonctions.xlsx"
if len(sys.argv) > 1:
    SRC = sys.argv[1]
else:
    # Le classeur est à côté du script (dossier OneDrive) ou un niveau plus haut (dépôt GitHub, script dans outils/).
    SRC = next((c for c in (os.path.join(ICI, NOM), os.path.join(os.path.dirname(ICI), NOM)) if os.path.exists(c)), os.path.join(ICI, NOM))
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(SRC)), "ecosysteme-habitation.html")

wb = load_workbook(SRC, read_only=True)  # on lit les valeurs saisies, pas les formules

fonctions = []
for r in wb["Fonctions"].iter_rows(min_row=2, values_only=True):
    if r[0] and str(r[0]).startswith("F"):
        fonctions.append({"code": r[0], "nom": r[1], "condition": r[2], "couvre": r[3], "manque": r[4]})
codes = [f["code"] for f in fonctions]

orgs, familles = [], []
for r in wb["Classement"].iter_rows(min_row=2, values_only=True):
    if not r[0]:
        continue
    p = str(r[5] or "").strip()
    if p not in codes:
        sys.exit(f"Fonction principale inconnue pour {r[1]} : « {p} »")
    sec = [s.strip() for s in str(r[7] or "").replace(";", ",").split(",") if s.strip() in codes and s.strip() != p]
    orgs.append({"nom": r[1], "sigle": r[2] or "", "fam": r[3], "p": p, "s": sec, "role": r[8] or "Exerce",
                 "pouvoir": r[9] or "", "infl": r[10] or "", "pos": int(r[11] or 0), "note": r[12] or ""})
    if r[3] not in familles:
        familles.append(r[3])

acteurs = []
for r in wb["Acteurs principaux"].iter_rows(min_row=2, values_only=True):
    if r[0] and r[3]:
        if str(r[0]).strip() not in codes:
            sys.exit(f"Fonction inconnue pour {r[3]} : « {r[0]} »")
        acteurs.append({"f": str(r[0]).strip(), "volet": r[2] or "", "nom": r[3], "role": r[4] or "Exerce", "tient": r[5] or "",
                        "rep": str(r[6] or "Non"), "infl": r[7] or ""})

parcours, leviers = [], {}
if "Parcours" in wb.sheetnames:
    for r in wb["Parcours"].iter_rows(min_row=2, values_only=True):
        if r[0] and r[2] and r[4]:
            parcours.append({"projet": r[0], "n": int(r[1] or 0), "etape": r[2], "quoi": r[3] or "", "qui": r[4],
                             "acteur": r[5] or r[4], "type": r[6] or "Décide", "f": r[7] or "", "non": r[8] or "",
                             "seuil": r[9] or "", "rep": str(r[10] or ""), "src": r[11] or "",
                             "ecrit": str(r[12] or "").lower().startswith("écrit")})
            cle = f"{r[0]}|{int(r[1] or 0)}"
            if len(r) > 13 and r[13] and cle not in leviers:
                leviers[cle] = {"texte": r[13], "statut": str(r[14] or ""), "src": str(r[15] or "") if len(r) > 15 else ""}

liens = []
if "Liens entre fonctions" in wb.sheetnames:
    for r in wb["Liens entre fonctions"].iter_rows(min_row=2, values_only=True):
        if r[0] and r[2] and str(r[0]).strip() in codes and str(r[2]).strip() in codes:
            appui = str(r[7] or "").strip()
            liens.append({"de": str(r[0]).strip(), "vers": str(r[2]).strip(), "etiq": r[4] or "", "retour": str(r[5] or "").lower().startswith("retour"),
                          "expl": r[6] or "", "appui": "" if appui.lower().startswith("à documenter") else appui,
                          "vue": str(r[8] or "").lower().startswith("o"), "trace": str(r[9] or "")})

DATA = json.dumps({"fonctions": fonctions, "orgs": orgs, "familles": familles, "acteurs": acteurs, "parcours": parcours, "leviers": leviers, "liens": liens},
                  ensure_ascii=False).replace("</", "<\\/")

HTML = r"""<!DOCTYPE html>
<html lang="fr-CA">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Écosystème de l'Habitation</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Sans+Condensed:wght@500;600&display=swap" rel="stylesheet">
<style>
:root{
  /* Carte de système, thème sombre unique : chaîne de fonctions de gauche à droite, panneau de détail dessous */
  color-scheme: dark;
  --fond:#10171f; --carte:#17212c; --carte2:#1d2a37; --trait:#2f4254; --trait-fort:#4a6379;
  --texte:#e6edf3; --doux:#9fb1c1; --pale:#6f8496;
  --accent:#7cc4f0;
  --bon:#5fcf9a; --mince:#f0b95a; --trou:#f08a7a;
  --exerce:#7cc4f0; --represente:#c9a5f2; --eclaire:#9fb1c1;
  --sans:"IBM Plex Sans",system-ui,-apple-system,"Segoe UI",sans-serif;
  --titre:"IBM Plex Sans Condensed","IBM Plex Sans",system-ui,sans-serif;
  --mono:"IBM Plex Mono",ui-monospace,Consolas,monospace;
}
*{box-sizing:border-box}
html,body{margin:0}
body{background:var(--fond);color:var(--texte);font-family:var(--sans);font-size:15px;line-height:1.5}
.page{max-width:1400px;margin:0 auto;padding:24px 20px 56px}
h1{font-family:var(--titre);font-weight:600;font-size:30px;letter-spacing:.01em;margin:0 0 4px;text-wrap:balance}
h2{font-family:var(--titre);font-weight:600;font-size:19px;margin:0 0 8px;text-wrap:balance}
p{margin:0 0 10px;max-width:68ch}
.chapeau{color:var(--doux)}
.avis{display:inline-block;font-family:var(--mono);font-size:12px;color:var(--mince);border:1px dashed var(--mince);border-radius:4px;padding:2px 8px;margin-bottom:12px}
.num{font-family:var(--mono);font-variant-numeric:tabular-nums}
header{display:grid;gap:14px;grid-template-columns:minmax(0,1fr);margin-bottom:18px}
.resume{display:flex;flex-wrap:wrap;gap:10px}
.tuile{background:var(--carte);border:1px solid var(--trait);border-radius:6px;padding:8px 12px;min-width:150px}
.tuile b{display:block;font-family:var(--mono);font-size:22px;font-weight:600}
.tuile span{color:var(--doux);font-size:13px}
nav{display:flex;flex-wrap:wrap;gap:6px;border-bottom:1px solid var(--trait);margin-bottom:18px}
nav a{color:var(--doux);text-decoration:none;padding:8px 14px;border:1px solid transparent;border-bottom:none;border-radius:6px 6px 0 0;font-weight:500}
nav a[aria-current="true"]{color:var(--texte);background:var(--carte);border-color:var(--trait)}
nav a:focus-visible,button:focus-visible,input:focus-visible,select:focus-visible,.boite:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
section[hidden]{display:none}

/* ----- carte ----- */
.defile{overflow-x:auto;padding-bottom:6px}
.chaine{display:grid;grid-template-columns:repeat(6,minmax(170px,1fr));gap:14px 26px;min-width:1150px}
.boite{position:relative;background:var(--carte);border:1px solid var(--trait);border-radius:8px;padding:10px 12px 12px;text-align:left;color:inherit;font:inherit;cursor:pointer;display:flex;flex-direction:column;gap:6px}
.boite:hover{border-color:var(--trait-fort)}
.boite[aria-pressed="true"]{border-color:var(--accent);background:var(--carte2);box-shadow:0 0 0 1px var(--accent)}
.boite .code{font-family:var(--mono);font-size:12px;color:var(--pale)}
.boite .nom{font-family:var(--titre);font-weight:600;font-size:17px;line-height:1.2}
.boite .cond{color:var(--doux);font-size:13px;line-height:1.35;min-height:3.9em}
.boite .ligne{display:flex;align-items:baseline;justify-content:space-between;gap:8px}
.boite .gros{font-family:var(--mono);font-size:24px;font-weight:600}
.boite .petit{font-size:12px;color:var(--pale)}
.barre{display:flex;height:6px;border-radius:3px;overflow:hidden;background:#0c1218}
.barre i{display:block;height:100%}
.etat{font-family:var(--mono);font-size:11px;padding:1px 7px;border-radius:10px;border:1px solid currentColor;white-space:nowrap}
.etat.bon{color:var(--bon)} .etat.mince{color:var(--mince)} .etat.trou{color:var(--trou)}
.abs{font-size:12px;color:var(--pale);border:1px dashed var(--trait-fort);border-radius:4px;padding:1px 6px;align-self:flex-start}
/* liens de gauche à droite */
.chaine .lien::after{content:"";position:absolute;top:50%;right:-27px;width:26px;border-top:2px solid var(--trait-fort)}
.chaine .lien::before{content:"";position:absolute;top:calc(50% - 4px);right:-27px;border:5px solid transparent;border-left-color:var(--trait-fort);border-right:none}
.f7{grid-column:6}
.f7::after{content:"";position:absolute;left:50%;top:-15px;height:14px;border-left:2px solid var(--trait-fort)}
.f8{grid-column:1 / span 5;grid-row:2;align-self:start;border-style:dashed;flex-direction:row;flex-wrap:wrap;align-items:center;gap:6px 18px}
.f8 .cond{min-height:0;flex:1 1 300px}
.f8 .abs{align-self:center}
.retour{font-family:var(--mono);font-size:11px;color:var(--pale)}
.legende{display:flex;flex-wrap:wrap;gap:6px 18px;margin:12px 0 20px;color:var(--doux);font-size:13px}
.legende span{display:inline-flex;align-items:center;gap:6px}
.pastille{width:10px;height:10px;border-radius:50%;display:inline-block}
.tiret{width:22px;border-top:1px dashed var(--trait-fort);display:inline-block}

/* ----- panneau de détail ----- */
.detail{display:grid;grid-template-columns:minmax(0,5fr) minmax(0,7fr);gap:20px;background:var(--carte);border:1px solid var(--trait);border-radius:8px;padding:18px}
.detail h2 .code{font-family:var(--mono);font-size:13px;color:var(--pale);margin-right:8px}
dl.kv{display:grid;grid-template-columns:minmax(120px,max-content) minmax(0,1fr);gap:6px 16px;margin:10px 0 16px}
dl.kv dt{color:var(--pale);font-size:13px}
dl.kv dd{margin:0}
dl.kv dd.num{font-size:15px}
h3{font-size:12px;font-family:var(--mono);font-weight:500;letter-spacing:.06em;text-transform:uppercase;color:var(--pale);margin:14px 0 6px}
.fam-ligne{display:grid;grid-template-columns:minmax(0,1fr) 34px 90px;gap:8px;align-items:center;font-size:13px;padding:2px 0}
.fam-ligne .jauge{height:6px;background:#0c1218;border-radius:3px;overflow:hidden}
.fam-ligne .jauge i{display:block;height:100%;background:var(--accent)}
.liste{max-height:430px;overflow:auto;border:1px solid var(--trait);border-radius:6px}
.org{display:grid;grid-template-columns:12px minmax(0,1fr) auto;gap:8px;align-items:baseline;padding:5px 10px;border-bottom:1px solid #223140;font-size:14px}
.org:last-child{border-bottom:none}
.org small{color:var(--pale);display:block;font-size:12px}
.org .fam{font-size:12px;color:var(--pale);text-align:right}
.volet{margin-bottom:12px}
.act{display:grid;grid-template-columns:12px minmax(0,1fr) auto;gap:8px;align-items:baseline;padding:6px 10px;border:1px solid #223140;border-radius:6px;margin-bottom:4px;background:#141d27;font-size:14px}
.act.absent{border-style:dashed;border-color:var(--trait-fort)}
.act small{display:block;color:var(--pale);font-size:12.5px}
.chip{font-family:var(--mono);font-size:11px;padding:1px 7px;border-radius:10px;border:1px solid currentColor;white-space:nowrap}
.chip.oui{color:var(--bon)} .chip.partie{color:var(--mince)} .chip.non{color:var(--trou);border-style:dashed}
details.classeur{margin-top:16px;border-top:1px solid var(--trait);padding-top:10px}
details.classeur summary{cursor:pointer;color:var(--doux);font-size:14px}
details.classeur summary b{font-family:var(--mono);color:var(--texte)}
.puces{display:flex;flex-wrap:wrap;gap:6px}
.puce{font-size:13px;border:1px dashed var(--trait-fort);border-radius:4px;padding:2px 8px;color:var(--doux)}
.puce.strat{border-color:var(--mince);color:var(--texte)}
.onglets-role{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:8px}
button.filtre{background:transparent;border:1px solid var(--trait);color:var(--doux);border-radius:14px;padding:3px 10px;font:inherit;font-size:13px;cursor:pointer}
button.filtre[aria-pressed="true"]{border-color:var(--accent);color:var(--texte);background:var(--carte2)}

/* ----- tableaux ----- */
.tab-defile{overflow-x:auto;border:1px solid var(--trait);border-radius:8px}
table{border-collapse:collapse;width:100%;font-size:14px}
th,td{padding:6px 10px;border-bottom:1px solid #223140;text-align:left;vertical-align:top}
thead th{position:sticky;top:0;background:var(--carte2);font-weight:500;color:var(--doux);font-size:13px}
td.c,th.c{text-align:center}
td.cel{font-family:var(--mono);text-align:center;font-variant-numeric:tabular-nums}
td.cel small{color:var(--pale);display:block;font-size:11px}
tfoot td{font-weight:600;background:var(--carte)}
.outils{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:12px;align-items:end}
.outils label{display:grid;gap:3px;font-size:12px;color:var(--pale)}
input,select{background:var(--carte);border:1px solid var(--trait);color:var(--texte);border-radius:6px;padding:6px 9px;font:inherit;font-size:14px}
input{min-width:240px}
.compte{font-family:var(--mono);font-size:13px;color:var(--doux);margin-left:auto}
.lecture{background:var(--carte);border-left:3px solid var(--accent);padding:10px 14px;border-radius:0 6px 6px 0;margin:0 0 16px;max-width:none}
.lecture p{max-width:80ch}
.lecture p:last-child{margin-bottom:0}
.deux{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:20px;align-items:start}
.projet{background:var(--carte);border:1px solid var(--trait);border-radius:8px;padding:16px}
.projet .faits{display:flex;flex-wrap:wrap;gap:6px 18px;color:var(--doux);font-size:13px;margin:4px 0 14px}
.projet .faits b{font-family:var(--mono);color:var(--texte);font-weight:600}
.etape{display:grid;grid-template-columns:34px minmax(0,1fr);gap:0 10px;position:relative;padding-bottom:14px}
.etape::before{content:"";position:absolute;left:16px;top:30px;bottom:0;border-left:2px solid var(--trait)}
.etape:last-child::before{display:none}
.etape .no{font-family:var(--mono);font-weight:600;font-size:13px;width:34px;height:28px;border:1px solid var(--trait-fort);border-radius:6px;display:grid;place-items:center;background:var(--carte2)}
.etape h4{margin:2px 0 0;font-family:var(--titre);font-size:16px;font-weight:600}
.etape .quoi{color:var(--pale);font-size:13px;margin:0 0 6px}
.porte{display:grid;grid-template-columns:152px minmax(0,1fr);gap:8px;padding:5px 8px;border:1px solid #223140;border-radius:6px;margin-bottom:4px;font-size:14px;background:#141d27}
.porte.absent{border-style:dashed;border-color:var(--trait-fort)}
.porte small{display:block;color:var(--pale);font-size:12px}
.porte small.non{color:var(--doux)}
.tp{font-family:var(--mono);font-size:11px;padding:1px 6px;border-radius:4px;border:1px solid currentColor;white-space:nowrap;align-self:start;justify-self:start;font-weight:500}
.tp.decide{color:var(--trou)} .tp.fixe{color:var(--mince)} .tp.avis{color:var(--eclaire)} .tp.livre{color:var(--bon)}
.regle{font-family:var(--mono);font-size:11px;color:var(--accent);border:1px solid var(--accent);border-radius:4px;padding:0 4px;margin-left:6px;font-weight:500}
.fcode{font-family:var(--mono);font-size:11px;color:var(--pale);margin-left:6px}
@media (max-width:1000px){.deux{grid-template-columns:minmax(0,1fr)}}
@media (max-width:520px){.porte{grid-template-columns:minmax(0,1fr)}}
.choix-f{display:flex;flex-wrap:wrap;gap:6px;margin-bottom:16px}
.fiche-tete{display:grid;grid-template-columns:minmax(0,5fr) minmax(0,7fr);gap:20px;background:var(--carte);border:1px solid var(--trait);border-radius:8px;padding:16px;margin-bottom:16px}
.fiche-tete h2 .code{font-family:var(--mono);font-size:13px;color:var(--pale);margin-right:8px}
.puces-act{display:flex;flex-wrap:wrap;gap:6px}
.puce-act{font-size:13px;border:1px solid var(--trait);border-radius:4px;padding:2px 8px;background:#141d27}
.puce-act.absent{border-style:dashed;border-color:var(--trait-fort);color:var(--doux)}
.projet-f{background:var(--carte);border:1px solid var(--trait);border-radius:8px;padding:14px 16px;margin-bottom:16px}
.projet-f h3{margin-top:0}
.projet-f .faits{display:flex;flex-wrap:wrap;gap:6px 18px;color:var(--doux);font-size:13px;margin:0 0 12px}
.projet-f .faits b{font-family:var(--mono);color:var(--texte);font-weight:600}
.chemin{display:flex;align-items:flex-start;overflow-x:auto;padding-bottom:8px}
.col{flex:0 0 236px;display:flex;flex-direction:column;gap:6px}
.pas{min-height:78px;background:rgba(95,207,154,.10);border:1px solid var(--bon);border-radius:6px;padding:8px 10px}
.pas .titre{display:flex;gap:8px;align-items:baseline;font-family:var(--titre);font-weight:600;font-size:15px;line-height:1.25}
.pas .titre .num{font-family:var(--mono);font-size:12px;color:var(--bon)}
.pas small{display:block;color:var(--doux);font-size:12.5px;margin-top:2px}
.lien-pas{flex:0 0 60px;position:relative;height:78px}
.lien-pas::before{content:"";position:absolute;left:4px;right:8px;top:38px;border-top:2px solid var(--trait-fort)}
.lien-pas::after{content:"";position:absolute;right:3px;top:34px;border:5px solid transparent;border-left-color:var(--trait-fort);border-right:none}
.lien-pas.saut::before{border-top-style:dashed}
.lien-pas span{position:absolute;left:2px;right:6px;top:46px;text-align:center;font-family:var(--mono);font-size:10.5px;color:var(--pale);line-height:1.2}
.porte2{border:1px solid #223140;border-radius:6px;padding:6px 8px;background:#141d27;font-size:13.5px;display:grid;gap:3px}
.porte2.autre{opacity:.62;border-style:dashed}
.porte2 small{color:var(--pale);font-size:12px}
.porte2 small.non{color:var(--doux)}
.fcode.autre-ex{margin:0 4px 0 0}
.vide{color:var(--doux);font-size:14px;border:1px dashed var(--trait-fort);border-radius:6px;padding:10px 12px;margin:0}
.lienfiche{color:var(--accent);font-size:14px}
@media (max-width:900px){.fiche-tete{grid-template-columns:minmax(0,1fr)}}
.contexte{display:flex;flex-wrap:wrap;gap:4px 14px;align-items:baseline;border:1px dashed var(--trait-fort);border-radius:8px;padding:9px 14px;margin-bottom:14px;font-size:13px;color:var(--doux)}
.contexte b{color:var(--texte);font-weight:600}
.ens{display:grid;grid-template-columns:minmax(0,1fr) 340px;gap:20px;align-items:start}
.ens-tete{display:flex;justify-content:space-between;gap:12px;align-items:center;margin-bottom:6px}
.ens-tete p{margin:0}
.ens-defile{overflow-x:auto}
.reseau{position:relative;min-width:860px;padding:22px 56px 92px}
.noeuds{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));column-gap:92px;row-gap:74px;position:relative;z-index:2}
.noeud{background:var(--carte);border:1px solid var(--trait);border-radius:8px;padding:10px 12px;text-align:left;color:inherit;font:inherit;cursor:pointer;display:flex;flex-direction:column;gap:4px;min-height:116px}
.noeud:hover{border-color:var(--trait-fort)}
.noeud[aria-pressed="true"]{border-color:var(--accent);background:var(--carte2);box-shadow:0 0 0 1px var(--accent)}
.noeud:focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.noeud.dim{opacity:.35}
.noeud .code{font-family:var(--mono);font-size:12px;color:var(--pale)}
.noeud .nom{font-family:var(--titre);font-weight:600;font-size:17px;line-height:1.2}
.noeud .cond{color:var(--doux);font-size:13px;line-height:1.35}
.noeud .petit{font-family:var(--mono);font-size:11.5px;color:var(--pale);margin-top:auto}
.noeud[data-f="F1"]{grid-column:1;grid-row:1}.noeud[data-f="F2"]{grid-column:3;grid-row:1}
.noeud[data-f="F3"]{grid-column:1;grid-row:2}.noeud[data-f="F4"]{grid-column:2;grid-row:2}.noeud[data-f="F5"]{grid-column:3;grid-row:2}
.noeud[data-f="F8"]{grid-column:1;grid-row:3}.noeud[data-f="F7"]{grid-column:2;grid-row:3}.noeud[data-f="F6"]{grid-column:3;grid-row:3}
#liens-svg{position:absolute;left:0;top:0;width:100%;height:100%;z-index:1;overflow:visible;pointer-events:none}
.lien{fill:none;stroke:var(--trait-fort);stroke-width:1.6;stroke-linejoin:round}
.lien.retour{stroke-dasharray:5 5}
.lien.ad{opacity:.5}
.lien.actif{stroke:var(--accent);stroke-width:2.2}
.etiq{font:11px var(--sans);fill:var(--doux);paint-order:stroke;stroke:var(--fond);stroke-width:4px;stroke-linejoin:round}
.etiq.actif{fill:var(--accent)}
.trait-plein,.trait-pointille,.trait-pale{display:inline-block;width:26px;border-top:2px solid var(--trait-fort)}
.trait-pointille{border-top-style:dashed}
.trait-pale{opacity:.5}
.ens-panneau{background:var(--carte);border:1px solid var(--trait);border-radius:8px;padding:16px}
.ens-panneau h2 .code{font-family:var(--mono);font-size:13px;color:var(--pale);margin-right:8px}
.ens-panneau ol{margin:8px 0 0;padding-left:18px;font-size:14px;color:var(--doux)}
.ens-panneau ol li{margin-bottom:6px}
.ens-panneau ol b{color:var(--texte);font-weight:600}
.question{font-family:var(--titre);font-size:18px;border-top:1px solid var(--trait);padding-top:12px;margin-top:14px}
.lien-li{border-top:1px solid #223140;padding:8px 0;font-size:14px}
.lien-li .dir{display:block;font-family:var(--mono);font-size:11px;color:var(--pale)}
.lien-li button{background:none;border:0;color:var(--accent);font:inherit;font-weight:600;padding:0;cursor:pointer;text-decoration:underline;text-underline-offset:3px}
.lien-li small{display:block;color:var(--pale);font-size:12px;margin-top:3px}
.lien-li small.ad{color:var(--mince)}
.lienfiche-ens{display:inline-block;margin-top:10px;color:var(--accent);font-size:14px}
.levier{border-left:3px solid var(--accent);background:rgba(124,196,240,.08);padding:6px 8px;font-size:12.5px;line-height:1.4;color:var(--doux);border-radius:0 6px 6px 0}
.levier b{color:var(--texte);font-weight:600}
.levier .st{font-family:var(--mono);font-size:10.5px;border:1px solid currentColor;border-radius:8px;padding:0 5px;margin-left:4px;font-weight:500}
.levier .st.piste{color:var(--mince)} .levier .st.existant{color:var(--bon)}
.levier small{display:block;color:var(--pale);font-size:11.5px;margin-top:2px}
.etape .levier{margin:0 0 6px}
@media (max-width:1100px){.ens{grid-template-columns:minmax(0,1fr)}}
footer{margin-top:28px;color:var(--pale);font-size:13px;border-top:1px solid var(--trait);padding-top:14px}
footer p{max-width:90ch}
@media (max-width:900px){.detail{grid-template-columns:minmax(0,1fr)} h1{font-size:25px}}
@media (prefers-reduced-motion:no-preference){.boite{transition:border-color .12s,background .12s}}
</style>
</head>
<body>
<!-- nav-site:debut -->
<style>.site-nav{box-sizing:border-box;display:flex;flex-wrap:wrap;align-items:center;gap:4px 6px;padding:8px 16px;background:#172033;border-bottom:1px solid #2f3d55;font:14px/1.3 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;position:relative;z-index:50;margin:0 0 8px}.site-nav .site-nav-titre{color:#c9d3e0;font-weight:600;font-size:13px;margin-right:10px}.site-nav a.site-nav-lien{color:#e6edf3;text-decoration:none;padding:5px 11px;border-radius:999px;border:1px solid transparent;white-space:nowrap;font-weight:400}.site-nav a.site-nav-lien:hover{border-color:#4a6379}.site-nav a.site-nav-lien[aria-current="page"]{background:#e6edf3;color:#172033;font-weight:600}.site-nav a.site-nav-lien:focus-visible{outline:2px solid #7cc4f0;outline-offset:2px}@media print{.site-nav{display:none}}</style>
<div class="site-nav" role="navigation" aria-label="Pages du projet"><span class="site-nav-titre">Logement et itinérance au Québec</span><a class="site-nav-lien" href="https://syl77qc.github.io/habitation_itinerance/">Infographie</a><a class="site-nav-lien" href="https://syl77qc.github.io/habitation_itinerance/parties-prenantes.html">Parties prenantes</a><a class="site-nav-lien" href="https://syl77qc.github.io/habitation_itinerance/ecosysteme-habitation.html" aria-current="page">Écosystème</a><a class="site-nav-lien" href="https://syl77qc.github.io/habitation_itinerance/revision-programmes/index-programmes-shq.html">Révision des programmes</a></div>
<!-- nav-site:fin -->
<div class="page">
<header>
  <div>
    <span class="avis">Document de travail personnel · proposition à valider</span>
    <h1>L'écosystème de l'Habitation</h1>
    <p class="chapeau">On part des fonctions, puis on y place les principales parties prenantes. Chaque fonction tient une condition : un logement existe, est habité et le reste seulement si toutes les conditions sont réunies. Notre classeur des parties prenantes sert ensuite à voir qui nous manque.</p>
  </div>
  <div class="resume" id="resume"></div>
</header>

<nav aria-label="Vues">
  <a href="#ensemble" data-vue="ensemble">Vue d'ensemble</a>
  <a href="#carte" data-vue="carte">Carte par fonctions</a>
  <a href="#fonctions" data-vue="fonctions">Par fonction : acteurs et chemin critique</a>
  <a href="#parcours" data-vue="parcours">Parcours d'un projet</a>
  <a href="#familles" data-vue="familles">Notre classeur : familles × fonctions</a>
  <a href="#tableau" data-vue="tableau">Notre classeur : organisations</a>
</nav>

<section id="vue-ensemble">
  <div class="contexte"><b>Le contexte agit sur tout</b><span>Revenus et coût de la vie · Population et ménages · Territoires et infrastructures · Taux d'intérêt et conjoncture</span></div>
  <div class="ens">
    <div>
      <div class="ens-tete">
        <p class="chapeau">Choisissez une fonction pour voir ce dont elle dépend et ce qu'elle fait bouger.</p>
        <button type="button" id="ens-tout" class="filtre" aria-pressed="true">Voir l'ensemble</button>
      </div>
      <div class="ens-defile"><div class="reseau" id="reseau"><svg id="liens-svg" aria-hidden="true"></svg><div class="noeuds" id="noeuds"></div></div></div>
      <div class="legende">
        <span><i class="trait-plein"></i>Dépendance : règles, ressources, services</span>
        <span><i class="trait-pointille"></i>Retour qui fait évoluer le système</span>
        <span><i class="trait-pale"></i>Lien plausible, appui à documenter</span>
      </div>
      <p class="chapeau" style="font-size:13px">Les flèches montrent des dépendances, pas un ordre de passage. La vue d'ensemble montre les liens principaux; les autres apparaissent quand on choisit une fonction.</p>
    </div>
    <aside class="ens-panneau" id="ens-panneau" aria-live="polite"></aside>
  </div>
</section>

<section id="vue-carte" hidden>
  <div class="defile"><div class="chaine" id="chaine"></div></div>
  <div class="legende">
    <span><i class="pastille" style="background:var(--exerce)"></i>Exerce la fonction</span>
    <span><i class="pastille" style="background:var(--represente)"></i>Représente ceux qui l'exercent</span>
    <span><i class="pastille" style="background:var(--eclaire)"></i>Éclaire</span>
    <span><b class="etat bon">couvert</b>les trois quarts des acteurs principaux sont dans notre classeur</span>
    <span><b class="etat mince">mince</b>la moitié ou plus</span>
    <span><b class="etat trou">trou</b>moins de la moitié</span>
    <span><i class="tiret"></i>Acteur absent de notre classeur</span>
  </div>
  <div class="detail" id="detail"></div>
</section>

<section id="vue-fonctions" hidden>
  <div class="choix-f" id="fiche-choix" role="group" aria-label="Choisir une fonction"></div>
  <div id="fiche"></div>
  <div class="legende">
    <span><b class="tp decide">Décide</b>peut dire oui ou non</span>
    <span><b class="tp fixe">Fixe les conditions</b>pose la règle, le seuil ou le prix</span>
    <span><b class="tp avis">Donne un avis</b>recommande</span>
    <span><b class="tp livre">Livre</b>le projet dépend de sa capacité</span>
    <span><span class="fcode autre-ex">F6</span>Acteur d'une autre fonction, présent à la même porte</span>
    <span><b class="regle">§</b>Écrit dans une règle</span>
  </div>
</section>

<section id="vue-familles" hidden>
  <div class="lecture" id="lecture-familles"></div>
  <div class="tab-defile"><table id="matrice"></table></div>
</section>

<section id="vue-parcours" hidden>
  <div class="lecture" id="lecture-parcours"></div>
  <div class="legende">
    <span><b class="tp decide">Décide</b>peut dire oui ou non</span>
    <span><b class="tp fixe">Fixe les conditions</b>pose la règle, le seuil ou le prix</span>
    <span><b class="tp avis">Donne un avis</b>recommande</span>
    <span><b class="tp livre">Livre</b>le projet dépend de sa capacité</span>
    <span><i class="tiret"></i>Acteur absent du classeur des parties prenantes</span>
    <span><b class="regle">§</b>Écrit dans une règle, avec la section</span>
  </div>
  <div class="deux" id="liste-parcours"></div>
  <h2 style="margin-top:26px">Qui tient les portes</h2>
  <p class="chapeau">Nombre d'étapes où chaque acteur décide ou fixe les conditions, dans chacun des deux parcours.</p>
  <div class="tab-defile"><table id="portes"></table></div>
</section>

<section id="vue-tableau" hidden>
  <div class="outils">
    <label for="q">Recherche<input id="q" type="search" placeholder="Nom ou sigle"></label>
    <label for="sel-f">Fonction<select id="sel-f"></select></label>
    <label for="sel-fam">Famille<select id="sel-fam"></select></label>
    <label for="sel-r">Rôle<select id="sel-r"><option value="">Tous</option><option>Exerce</option><option>Représente</option><option>Éclaire</option></select></label>
    <span class="compte" id="compte"></span>
  </div>
  <div class="tab-defile" style="max-height:640px;overflow:auto"><table id="table-orgs"></table></div>
</section>

<footer>
  <p><b>Vue d'ensemble.</b> Les liens entre les fonctions s'inspirent de Clark (2013) : le résultat dépend des pièces et de leurs interactions. Chaque lien est appuyé par une étape de parcours ou une source; un lien pâle reste à documenter. Les leviers indiquent ce qui existe déjà (« existant ») ou ce qui mérite d'être étudié (« piste »).</p>
  <p><b>Comment lire.</b> La liste des acteurs principaux est une proposition à valider. « Principal » veut dire que la fonction ne s'exerce pas sans lui, ou qu'il parle pour ceux qui l'exercent. Le reclassement des organisations de notre classeur vient de leur nom et de leur famille, pas d'une lecture de leurs activités.</p>
  <p><b>Document de travail personnel.</b> Analyse préliminaire, non validée. Elle ne représente la position d'aucun ministère ni du gouvernement du Québec. Voir aussi l'<a href="https://syl77qc.github.io/habitation_itinerance/" style="color:var(--accent)">infographie sur le logement et l'itinérance</a> et la <a href="https://syl77qc.github.io/habitation_itinerance/parties-prenantes.html" style="color:var(--accent)">cartographie des parties prenantes</a>.</p>
  <p><b>Sources.</b> Classeur des parties prenantes, bâti sur des mémoires et documents publics. Ministère des Affaires municipales et de l'Habitation (2024), <i>Stratégie québécoise en habitation</i>. Clark, A. M. (2013), « What are the components of complex interventions in healthcare? », <i>Social Science &amp; Medicine</i>, 93, 185-193.</p>
</footer>
</div>

<script id="donnees" type="application/json">__DATA__</script>
<script>
(function(){
const D = JSON.parse(document.getElementById('donnees').textContent);
const F = D.fonctions, O = D.orgs, FAM = D.familles;
const LEV = D.leviers || {};
const levierHtml = (pr, n) => { const l = LEV[pr + '|' + n]; if (!l) return '';
  const st = (l.statut || '').toLowerCase().startsWith('piste') ? 'piste' : 'existant';
  return '<div class="levier"><b>Levier</b><span class="st ' + st + '">' + st + '</span> ' + esc(l.texte) + (l.src ? '<small>' + esc(l.src) + '</small>' : '') + '</div>'; };
const parCode = Object.fromEntries(F.map(f => [f.code, f]));
const coul = {Exerce:'var(--exerce)', 'Représente':'var(--represente)', 'Éclaire':'var(--eclaire)'};
const esc = s => String(s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const ACT = D.acteurs || [];
const chip = r => r === 'Oui' ? '<span class="chip oui">dans le classeur</span>' : r === 'En partie' ? '<span class="chip partie">en partie</span>' : '<span class="chip non">absent</span>';
F.forEach(f => {
  f.act = ACT.filter(a => a.f === f.code);
  f.orgs = O.filter(o => o.p === f.code);
  f.sec = O.filter(o => o.s.includes(f.code));
  f.n = {Exerce:0,'Représente':0,'Éclaire':0};
  f.act.forEach(a => f.n[a.role] = (f.n[a.role]||0) + 1);
  f.rep = {Oui:0,'En partie':0,Non:0};
  f.act.forEach(a => f.rep[a.rep] = (f.rep[a.rep]||0) + 1);
  f.couv = f.act.length ? (f.rep.Oui + 0.5 * f.rep['En partie']) / f.act.length : 0;
});
const etat = c => c >= 0.75 ? ['bon','couvert'] : c >= 0.5 ? ['mince','mince'] : ['trou','trou'];

/* résumé */
const nOui = ACT.filter(a => a.rep === 'Oui').length, nPart = ACT.filter(a => a.rep === 'En partie').length, nNon = ACT.filter(a => a.rep === 'Non').length;
const pire = F.slice().sort((a,b) => a.couv - b.couv)[0];
document.getElementById('resume').innerHTML = [
  [F.length, 'fonctions, donc ' + F.length + ' conditions'],
  [ACT.length, 'acteurs principaux'],
  [nOui, 'dans notre classeur'],
  [nPart, 'en partie seulement'],
  [nNon, 'absents de notre classeur'],
  [pire.rep.Non + ' sur ' + pire.act.length, 'absents dans « ' + pire.nom + ' »']
].map(t => '<div class="tuile"><b>' + esc(t[0]) + '</b><span>' + esc(t[1]) + '</span></div>').join('');

/* carte */
const chaine = document.getElementById('chaine');
const ordre = ['F1','F2','F3','F4','F5','F6','F8','F7'].filter(c => parCode[c]);
chaine.innerHTML = ordre.map(c => {
  const f = parCode[c], n = f.act.length, e = etat(f.couv), tot = Math.max(n,1);
  const lien = ['F1','F2','F3','F4','F5'].includes(c) ? ' lien' : '';
  const cls = c === 'F7' ? ' f7' : c === 'F8' ? ' f8' : '';
  const barre = '<div class="barre" title="Dans le classeur ' + f.rep.Oui + ', en partie ' + f.rep['En partie'] + ', absents ' + f.rep.Non + '">' +
    [['Oui','var(--bon)'],['En partie','var(--mince)'],['Non','var(--trou)']].map(r => '<i style="width:' + (100*f.rep[r[0]]/tot) + '%;background:' + r[1] + '"></i>').join('') + '</div>';
  if (c === 'F8') return '<button type="button" class="boite f8" data-f="F8" aria-pressed="false">' +
    '<span><span class="code">F8</span> <span class="nom">' + esc(f.nom) + '</span></span>' +
    '<span class="cond">' + esc(f.condition) + '</span>' +
    '<span class="gros">' + n + '</span><span class="petit">acteurs principaux</span><b class="etat ' + e[0] + '">' + e[1] + '</b>' +
    '<span class="abs">' + f.rep.Non + ' absents du classeur</span>' +
    '<span class="retour">← boucle de retour vers F1</span></button>';
  return '<button type="button" class="boite' + lien + cls + '" data-f="' + c + '" aria-pressed="false">' +
    '<span class="code">' + c + '</span><span class="nom">' + esc(f.nom) + '</span>' +
    '<span class="cond">' + esc(f.condition) + '</span>' +
    '<span class="ligne"><span><span class="gros">' + n + '</span> <span class="petit">acteurs</span></span><b class="etat ' + e[0] + '">' + e[1] + '</b></span>' +
    barre +
    '<span class="abs">' + f.rep.Non + ' absents du classeur</span></button>';
}).join('');

let choisie = 'F1';
function ligneOrg(o, second){
  return '<div class="org"><i class="pastille" style="background:' + coul[o.role] + '" title="' + esc(o.role) + '"></i>' +
    '<span>' + esc(o.nom) + (second ? ' <span class="petit" style="color:var(--pale)">(fonction principale : ' + esc(parCode[o.p].nom) + ')</span>' : '') +
    (o.note ? '<small>' + esc(o.note) + '</small>' : '') + '</span>' +
    '<span class="fam">' + esc(o.fam) + '</span></div>';
}
function detail(){
  const f = parCode[choisie];
  document.querySelectorAll('.boite').forEach(b => b.setAttribute('aria-pressed', b.dataset.f === choisie));
  const volets = [...new Set(f.act.map(a => a.volet))];
  const orgs = f.orgs.slice().sort((a,b) => a.nom.localeCompare(b.nom,'fr'));
  const sec = f.sec.slice().sort((a,b) => a.nom.localeCompare(b.nom,'fr'));
  document.getElementById('detail').innerHTML =
  '<div>' +
    '<h2><span class="code">' + f.code + '</span>' + esc(f.nom) + '</h2>' +
    '<p class="chapeau">' + esc(f.couvre) + '</p>' +
    '<dl class="kv">' +
      '<dt>Condition tenue</dt><dd>' + esc(f.condition) + '</dd>' +
      '<dt>Si elle manque</dt><dd>' + esc(f.manque) + '</dd>' +
      '<dt>Acteurs principaux</dt><dd class="num">' + f.act.length + '</dd>' +
      '<dt>Exercent</dt><dd class="num">' + f.n.Exerce + '</dd>' +
      '<dt>Représentent</dt><dd class="num">' + f.n['Représente'] + '</dd>' +
      '<dt>Éclairent</dt><dd class="num">' + f.n['Éclaire'] + '</dd>' +
    '</dl>' +
    '<h3>Présence dans notre classeur</h3>' +
    [['Oui','Dans le classeur','var(--bon)'],['En partie','En partie','var(--mince)'],['Non','Absents','var(--trou)']].map(x =>
      '<div class="fam-ligne"><span>' + x[1] + '</span><span class="num" style="text-align:right">' + f.rep[x[0]] + '</span><span class="jauge"><i style="width:' + (100*f.rep[x[0]]/Math.max(f.act.length,1)) + '%;background:' + x[2] + '"></i></span></div>').join('') +
    '<p style="margin:14px 0 0"><a class="lienfiche" href="#fonctions" data-f="' + f.code + '">Voir les acteurs et le chemin critique de cette fonction →</a></p>' +
    '<details class="classeur"><summary>Organisations de notre classeur rangées dans cette fonction : <b>' + orgs.length + '</b> en fonction principale, <b>' + sec.length + '</b> en second</summary>' +
      '<div class="liste" style="margin-top:8px;max-height:320px">' + (orgs.map(o => ligneOrg(o,false)).join('') + sec.map(o => ligneOrg(o,true)).join('') || '<div class="org"><span></span><span>Aucune organisation.</span><span></span></div>') + '</div></details>' +
  '</div>' +
  '<div>' +
    volets.map(v => '<div class="volet"><h3 style="margin-top:0">' + esc(v) + '</h3>' +
      f.act.filter(a => a.volet === v).map(a => '<div class="act' + (a.rep === 'Non' ? ' absent' : '') + '"><i class="pastille" style="background:' + coul[a.role] + '" title="' + esc(a.role) + '"></i>' +
        '<span>' + esc(a.nom) + '<small>' + esc(a.tient) + '</small></span>' + chip(a.rep) + '</div>').join('') + '</div>').join('') +
  '</div>';
}
chaine.addEventListener('click', e => { const b = e.target.closest('.boite'); if (b){ choisie = b.dataset.f; detail(); } });
detail();

/* matrice familles × fonctions */
(function(){
  const lignes = FAM.map(fam => {
    const os = O.filter(o => o.fam === fam);
    const cel = F.map(f => [os.filter(o => o.p === f.code).length, os.filter(o => o.s.includes(f.code)).length]);
    return {fam, tot: os.length, cel, touche: cel.filter(c => c[0] + c[1] > 0).length};
  });
  const maxCel = Math.max(...lignes.flatMap(l => l.cel.map(c => c[0])));
  let h = '<thead><tr><th>Famille</th>' + F.map(f => '<th class="c">' + f.code + '<br>' + esc(f.nom) + '</th>').join('') + '<th class="c">Total</th><th class="c">Fonctions touchées</th></tr></thead><tbody>';
  lignes.forEach(l => {
    h += '<tr><td>' + esc(l.fam) + '</td>' + l.cel.map(c => {
      const a = c[0] ? 0.10 + 0.55 * Math.sqrt(c[0] / maxCel) : 0;
      return '<td class="cel" style="background:rgba(124,196,240,' + a.toFixed(2) + ')">' + (c[0] || (c[1] ? '' : '·')) + (c[1] ? '<small>+' + c[1] + ' en second</small>' : '') + '</td>';
    }).join('') + '<td class="cel">' + l.tot + '</td><td class="cel">' + l.touche + ' / ' + F.length + '</td></tr>';
  });
  h += '</tbody><tfoot><tr><td>Total</td>' + F.map(f => '<td class="cel">' + f.orgs.length + '</td>').join('') + '<td class="cel">' + O.length + '</td><td></td></tr></tfoot>';
  document.getElementById('matrice').innerHTML = h;
  const large = lignes.slice().sort((a,b) => b.touche - a.touche || b.tot - a.tot).slice(0,3);
  const etroit = lignes.filter(l => l.touche === 1 && l.tot >= 2);
  document.getElementById('lecture-familles').innerHTML =
    '<p><b>Le grand chiffre</b> compte les organisations dont c\'est la fonction principale. <b>« En second »</b> compte celles qui touchent aussi cette fonction.</p>' +
    '<p>Familles les plus étendues : ' + large.map(l => esc(l.fam) + ' (<span class="num">' + l.touche + '</span> fonctions)').join(', ') + '. ' +
    (etroit.length ? 'Familles concentrées dans une seule fonction : ' + etroit.map(l => esc(l.fam)).join(', ') + '.' : '') + '</p>';
})();

/* tableau */
(function(){
  const sf = document.getElementById('sel-f'), sfam = document.getElementById('sel-fam'), sr = document.getElementById('sel-r'), q = document.getElementById('q');
  sf.innerHTML = '<option value="">Toutes</option>' + F.map(f => '<option value="' + f.code + '">' + f.code + ' ' + esc(f.nom) + '</option>').join('');
  sfam.innerHTML = '<option value="">Toutes</option>' + FAM.map(f => '<option>' + esc(f) + '</option>').join('');
  const t = document.getElementById('table-orgs');
  function rendre(){
    const mot = q.value.trim().toLowerCase();
    const l = O.filter(o => (!sf.value || o.p === sf.value || o.s.includes(sf.value)) && (!sfam.value || o.fam === sfam.value) && (!sr.value || o.role === sr.value) &&
      (!mot || (o.nom + ' ' + o.sigle).toLowerCase().includes(mot)));
    document.getElementById('compte').textContent = l.length + ' sur ' + O.length;
    t.innerHTML = '<thead><tr><th>Organisation</th><th>Famille</th><th>Fonction principale</th><th>Autres fonctions</th><th>Rôle</th><th>Type de pouvoir</th><th>Influence</th><th class="c">Positions codées</th></tr></thead><tbody>' +
      l.map(o => '<tr><td>' + esc(o.nom) + (o.note ? '<br><small style="color:var(--pale)">' + esc(o.note) + '</small>' : '') + '</td><td>' + esc(o.fam) + '</td><td>' + esc(parCode[o.p].nom) +
        '</td><td>' + o.s.map(c => esc(parCode[c].nom)).join(', ') + '</td><td style="white-space:nowrap"><i class="pastille" style="background:' + coul[o.role] + '"></i> ' + esc(o.role) +
        '</td><td>' + esc(o.pouvoir) + '</td><td>' + (o.infl ? esc(o.infl) : '<span style="color:var(--pale)">à coter</span>') + '</td><td class="cel">' + o.pos + '</td></tr>').join('') + '</tbody>';
  }
  [sf, sfam, sr].forEach(s => s.addEventListener('change', rendre));
  q.addEventListener('input', rendre);
  rendre();
})();

/* parcours d'un projet */
(function(){
  const P = D.parcours || [];
  const cls = {'Décide':'decide','Fixe les conditions':'fixe','Donne un avis':'avis','Livre':'livre'};
  const projets = [...new Set(P.map(p => p.projet))];
  const stats = {};
  document.getElementById('liste-parcours').innerHTML = projets.map(pr => {
    const l = P.filter(p => p.projet === pr);
    const nos = [...new Set(l.map(p => p.n))].sort((a,b) => a-b);
    const dec = l.filter(p => p.type === 'Décide');
    const acteursDec = new Set(dec.map(p => p.acteur));
    const absents = new Set(l.filter(p => p.rep === 'Non').map(p => p.acteur));
    const ecrits = l.filter(p => p.ecrit).length;
    stats[pr] = {etapes: nos.length, portes: dec.length, acteurs: acteursDec.size, lignes: l.length, absents: absents.size, tous: new Set(l.map(p => p.acteur)).size, ecrits};
    return '<div class="projet"><h2>' + esc(pr) + '</h2>' +
      '<div class="faits"><span><b>' + nos.length + '</b> étapes</span><span><b>' + dec.length + '</b> décisions oui ou non</span><span><b>' + acteursDec.size + '</b> acteurs qui décident</span><span><b>' + absents.size + '</b> sur ' + stats[pr].tous + ' acteurs absents du classeur</span></div>' +
      nos.map(n => {
        const e = l.filter(p => p.n === n);
        return '<div class="etape"><span class="no">' + n + '</span><div><h4>' + esc(e[0].etape) + '</h4><p class="quoi">' + esc(e[0].quoi) + '</p>' + levierHtml(pr, n) +
          e.map(p => '<div class="porte' + (p.rep === 'Non' ? ' absent' : '') + '"><b class="tp ' + (cls[p.type] || 'avis') + '">' + esc(p.type) + '</b><span>' + esc(p.qui) +
            (p.f && parCode[p.f] ? '<span class="fcode" title="' + esc(parCode[p.f].nom) + '">' + p.f + '</span>' : '') +
            (p.ecrit ? '<b class="regle" title="' + esc(p.src) + '">§</b>' : '') +
            (p.seuil ? '<small>' + esc(p.seuil) + (p.ecrit ? ' · ' + esc(p.src) : '') + '</small>' : (p.ecrit ? '<small>' + esc(p.src) + '</small>' : '')) +
            (p.non ? '<small class="non">Si c\'est non : ' + esc(p.non) + '</small>' : '') + '</span></div>').join('') +
          '</div></div>';
      }).join('') + '</div>';
  }).join('');

  /* qui tient les portes */
  const acteurs = [...new Set(P.map(p => p.acteur))].map(a => {
    const o = {a, rep: (P.find(p => p.acteur === a) || {}).rep, par: {}, tot: 0, dec: 0};
    projets.forEach(pr => {
      const l = P.filter(p => p.projet === pr && p.acteur === a && (p.type === 'Décide' || p.type === 'Fixe les conditions'));
      const d = new Set(l.filter(p => p.type === 'Décide').map(p => p.n)).size, f = new Set(l.map(p => p.n)).size;
      o.par[pr] = [d, f]; o.tot += f; o.dec += d;
    });
    return o;
  }).filter(o => o.tot > 0).sort((x,y) => y.tot - x.tot || y.dec - x.dec || x.a.localeCompare(y.a,'fr'));
  const max = Math.max(1, ...acteurs.flatMap(o => projets.map(pr => o.par[pr][1])));
  document.getElementById('portes').innerHTML = '<thead><tr><th>Acteur</th>' + projets.map(pr => '<th class="c">' + esc(pr) + '</th>').join('') + '<th class="c">Dans le classeur</th></tr></thead><tbody>' +
    acteurs.map(o => '<tr><td>' + esc(o.a) + '</td>' + projets.map(pr => { const c = o.par[pr];
      return '<td class="cel" style="background:rgba(124,196,240,' + (c[1] ? (0.08 + 0.5 * c[1] / max).toFixed(2) : 0) + ')">' + (c[1] ? c[1] + ' étape' + (c[1] > 1 ? 's' : '') + '<small>dont ' + c[0] + ' où il décide</small>' : '·') + '</td>'; }).join('') +
      '<td class="c">' + (o.rep === 'Non' ? '<span class="puce">non</span>' : esc(o.rep.toLowerCase())) + '</td></tr>').join('') + '</tbody>';

  if (projets.length){
    const communs = acteurs.filter(o => projets.every(pr => o.par[pr][1] > 0)).map(o => o.a);
    const premier = pr => acteurs.slice().sort((x,y) => y.par[pr][1] - x.par[pr][1] || y.par[pr][0] - x.par[pr][0])[0];
    document.getElementById('lecture-parcours').innerHTML =
      '<p><b>La chandelle, appliquée à deux projets.</b> Chaque étape est une porte. Le projet avance seulement si tous ceux qui la tiennent disent oui.</p>' +
      '<p>' + projets.map(pr => { const s = stats[pr], t = premier(pr); return esc(pr) + ' : <span class="num">' + s.etapes + '</span> étapes, <span class="num">' + s.portes + '</span> décisions oui ou non. L\'acteur le plus présent est « ' + esc(t.a) + ' » (<span class="num">' + t.par[pr][1] + '</span> étapes).'; }).join(' ') + '</p>' +
      (communs.length ? '<p>Acteurs qui tiennent une porte dans les deux parcours : ' + communs.map(esc).join(', ') + '.</p>' : '') +
      '<p>Les lignes marquées § viennent d\'un texte, avec la section. Les autres décrivent le cadre général et restent à valider avec une ville, un promoteur ou la SHQ.</p>';
  }
})();

/* fiches par fonction : acteurs et chemin critique */
(function(){
  const P = D.parcours || [];
  const projets = [...new Set(P.map(p => p.projet))];
  const cls = {'Décide':'decide','Fixe les conditions':'fixe','Donne un avis':'avis','Livre':'livre'};
  const ordreType = {'Décide':0,'Fixe les conditions':1,'Donne un avis':2,'Livre':3};
  let fsel = 'F3';
  const choix = document.getElementById('fiche-choix');
  choix.innerHTML = F.map(f => '<button type="button" class="filtre" data-f="' + f.code + '" aria-pressed="false">' + f.code + ' ' + esc(f.nom) + '</button>').join('');

  function porte(p){
    const autre = p.f !== fsel;
    return '<div class="porte2' + (autre ? ' autre' : '') + '"><span><b class="tp ' + (cls[p.type] || 'avis') + '">' + esc(p.type) + '</b>' +
      (autre && parCode[p.f] ? '<span class="fcode" title="' + esc(parCode[p.f].nom) + '">' + p.f + '</span>' : '') +
      (p.ecrit ? '<b class="regle" title="' + esc(p.src) + '">§</b>' : '') + '</span>' +
      '<span>' + esc(p.qui) + '</span>' +
      (p.seuil ? '<small>' + esc(p.seuil) + (p.ecrit ? ' · ' + esc(p.src) : '') + '</small>' : '') +
      (p.non && !autre ? '<small class="non">Si c\'est non : ' + esc(p.non) + '</small>' : '') + '</div>';
  }

  function chemin(pr){
    const tout = P.filter(p => p.projet === pr);
    const total = new Set(tout.map(p => p.n)).size;
    const nos = [...new Set(tout.filter(p => p.f === fsel).map(p => p.n))].sort((a,b) => a-b);
    const dec = new Set(tout.filter(p => p.f === fsel && p.type === 'Décide').map(p => p.n)).size;
    if (!nos.length) return {html: '<div class="projet-f"><h3>' + esc(pr) + '</h3><p class="vide">Aucune étape de ce parcours ne passe par cette fonction.</p></div>', nos, total, dec};
    let h = '';
    nos.forEach((n, i) => {
      if (i){
        const saut = n - nos[i-1] - 1;
        h += '<div class="lien-pas' + (saut ? ' saut' : '') + '">' + (saut ? '<span>+' + saut + ' étape' + (saut > 1 ? 's' : '') + '<br>ailleurs</span>' : '') + '</div>';
      }
      const l = tout.filter(p => p.n === n).sort((a,b) => (a.f !== fsel) - (b.f !== fsel) || ordreType[a.type] - ordreType[b.type]);
      h += '<div class="col"><div class="pas"><span class="titre"><span class="num">' + n + '</span>' + esc(l[0].etape) + '</span><small>' + esc(l[0].quoi) + '</small></div>' + levierHtml(pr, n) + l.map(porte).join('') + '</div>';
    });
    return {html: '<div class="projet-f"><h3>' + esc(pr) + '</h3><div class="faits"><span><b>' + nos.length + '</b> étapes sur ' + total + ' passent par cette fonction</span><span><b>' + dec + '</b> où un de ses acteurs décide</span></div><div class="chemin">' + h + '</div></div>', nos, total, dec};
  }

  function rendre(){
    const f = parCode[fsel];
    choix.querySelectorAll('button').forEach(b => b.setAttribute('aria-pressed', b.dataset.f === fsel));
    const tab = {};
    P.filter(p => p.f === fsel && (p.type === 'Décide' || p.type === 'Fixe les conditions')).forEach(p => {
      const t = tab[p.acteur] = tab[p.acteur] || {a: p.acteur, rep: p.rep, tout: new Set(), dec: new Set()};
      t.tout.add(p.projet + '|' + p.n);
      if (p.type === 'Décide') t.dec.add(p.projet + '|' + p.n);
    });
    const top = Object.values(tab).sort((x,y) => y.tout.size - x.tout.size || y.dec.size - x.dec.size || x.a.localeCompare(y.a,'fr'));
    const ch = projets.map(chemin);
    let lecture;
    if (ch.every(c => !c.nos.length)){
      lecture = 'Aucun des deux parcours ne fait intervenir cette fonction. C\'est un angle mort de nos parcours : elle agit autour des projets plutôt que dans leur chemin.';
    } else {
      lecture = projets.map((pr, i) => ch[i].nos.length ? esc(pr) + ' : <span class="num">' + ch[i].nos.length + '</span> étapes sur ' + ch[i].total + '.' : esc(pr) + ' : aucune étape.').join(' ');
      if (top.length) lecture += ' L\'acteur qui tient le plus de portes dans cette fonction : « ' + esc(top[0].a) + ' » (<span class="num">' + top[0].tout.size + '</span> étape' + (top[0].tout.size > 1 ? 's' : '') + ' dans les deux parcours, dont ' + top[0].dec.size + ' où il décide).';
    }
    document.getElementById('fiche').innerHTML =
      '<div class="fiche-tete"><div>' +
        '<h2><span class="code">' + f.code + '</span>' + esc(f.nom) + '</h2>' +
        '<dl class="kv"><dt>Condition tenue</dt><dd>' + esc(f.condition) + '</dd><dt>Si elle manque</dt><dd>' + esc(f.manque) + '</dd></dl>' +
        '<p class="chapeau" style="margin:0">' + lecture + '</p>' +
      '</div><div>' +
        '<h3 style="margin-top:0">Acteurs principaux</h3>' +
        '<div class="puces-act">' + f.act.map(a => '<span class="puce-act' + (a.rep === 'Non' ? ' absent' : '') + '" title="' + esc(a.tient) + '">' + esc(a.nom) + '</span>').join('') + '</div>' +
        (top.length ? '<h3>Qui tient les portes de cette fonction</h3><div class="liste" style="max-height:none">' + top.map(t =>
          '<div class="org"><i class="pastille" style="background:' + (t.dec.size ? 'var(--trou)' : 'var(--mince)') + '"></i><span>' + esc(t.a) + '<small>' + t.tout.size + ' étape' + (t.tout.size > 1 ? 's' : '') + ' dans les deux parcours, dont ' + t.dec.size + ' où il décide</small></span><span class="fam">' + (t.rep === 'Non' ? 'absent du classeur' : '') + '</span></div>').join('') + '</div>' : '') +
      '</div></div>' +
      ch.map(c => c.html).join('');
  }
  choix.addEventListener('click', e => { const b = e.target.closest('button'); if (b){ fsel = b.dataset.f; rendre(); } });
  document.getElementById('detail').addEventListener('click', e => { const a = e.target.closest('.lienfiche'); if (a){ fsel = a.dataset.f; rendre(); } });
  window.ouvrirFiche = c => { fsel = c; rendre(); };
  rendre();
})();

/* vue d'ensemble : les fonctions et leurs liens */
(function(){
  const L = D.liens || [];
  const P = D.parcours || [];
  const noeuds = document.getElementById('noeuds'), svg = document.getElementById('liens-svg'), reseau = document.getElementById('reseau');
  const panneau = document.getElementById('ens-panneau'), tout = document.getElementById('ens-tout');
  const portes = c => P.filter(p => p.f === c && (p.type === 'Décide' || p.type === 'Fixe les conditions')).length;
  noeuds.innerHTML = F.map(f => '<button type="button" class="noeud" data-f="' + f.code + '" aria-pressed="false"><span class="code">' + f.code + '</span><span class="nom">' + esc(f.nom) + '</span><span class="cond">' + esc(f.condition) + '</span><span class="petit">' + portes(f.code) + (portes(f.code) > 1 ? ' portes' : ' porte') + ' · ' + f.act.length + ' acteurs</span></button>').join('');
  let sel = null;

  function geo(){
    const b = reseau.getBoundingClientRect(), o = {};
    noeuds.querySelectorAll('.noeud').forEach(n => { const r = n.getBoundingClientRect(); o[n.dataset.f] = {x: r.left - b.left, y: r.top - b.top, w: r.width, h: r.height}; });
    const bas = Math.max(...Object.values(o).map(r => r.y + r.h));
    const gx = o.F4 && o.F3 ? o.F4.x - (o.F3.x + o.F3.w) : 80, gy = o.F3 && o.F1 ? o.F3.y - (o.F1.y + o.F1.h) : 70;
    return {o, W: b.width, H: b.height, bas, gx, gy};
  }
  function points(code, a, b, g){
    const parts = String(code || '').split(':'), t = parts[0], n = parts.slice(1).map(Number);
    const X = (r,f) => r.x + f*r.w, Y = (r,f) => r.y + f*r.h, B = r => r.y + r.h, D2 = r => r.x + r.w;
    switch (t){
      case 'bas-haut': { const ax = X(a,n[0]), bx = X(b,n[1]), ym = B(a) + n[2]*(b.y - B(a)); return Math.abs(ax-bx) < 1 ? [[ax,B(a)],[bx,b.y]] : [[ax,B(a)],[ax,ym],[bx,ym],[bx,b.y]]; }
      case 'haut-bas': { const ax = X(a,n[0]), bx = X(b,n[1]), ym = a.y - n[2]*(a.y - B(b)); return Math.abs(ax-bx) < 1 ? [[ax,a.y],[bx,B(b)]] : [[ax,a.y],[ax,ym],[bx,ym],[bx,B(b)]]; }
      case 'droite-gauche': { const y = Y(a,n[0]); return [[D2(a),y],[b.x,y]]; }
      case 'gauche-droite': { const y = Y(a,n[0]); return [[a.x,y],[D2(b),y]]; }
      case 'droite-haut': { const y = Y(a,n[0]), bx = X(b,n[1]); return [[D2(a),y],[bx,y],[bx,b.y]]; }
      case 'droite-couloir': { const y = Y(a,n[0]), xc = b.x - g.gx/2, yc = b.y - n[1]*g.gy, bx = X(b,n[2]); return [[D2(a),y],[xc,y],[xc,yc],[bx,yc],[bx,b.y]]; }
      case 'droite-ext': { const y = Y(a,n[0]), xe = g.W - n[1], yb = Y(b,n[2]); return [[D2(a),y],[xe,y],[xe,yb],[D2(b),yb]]; }
      case 'gauche-ext': { const y = Y(a,n[0]), xe = n[1], yb = Y(b,n[2]); return [[a.x,y],[xe,y],[xe,yb],[b.x,yb]]; }
      case 'bas-ext': { const ax = X(a,n[0]), yl = g.bas + n[1], bx = X(b,n[2]); return [[ax,B(a)],[ax,yl],[bx,yl],[bx,B(b)]]; }
      case 'gauche-ext-bas': { const y = Y(a,n[0]), xe = n[1], yl = g.bas + n[2], bx = X(b,n[3]); return [[a.x,y],[xe,y],[xe,yl],[bx,yl],[bx,B(b)]]; }
      case 'droite-ext-bas': { const y = Y(a,n[0]), xe = g.W - n[1], yl = g.bas + n[2], bx = X(b,n[3]); return [[D2(a),y],[xe,y],[xe,yl],[bx,yl],[bx,B(b)]]; }
      default: return [[a.x + a.w/2, a.y + a.h/2],[b.x + b.w/2, b.y + b.h/2]];
    }
  }
  function etiquette(pts, g, texte, actif){
    const i = Math.floor((pts.length - 1) / 2), p = pts[i], q = pts[i+1];
    const mx = (p[0] + q[0]) / 2, my = (p[1] + q[1]) / 2, cl = 'etiq' + (actif ? ' actif' : '');
    if (Math.abs(p[0] - q[0]) < 1){
      if (p[0] < 44) return '<text class="' + cl + '" x="' + (p[0]-3) + '" y="' + my + '" text-anchor="middle" transform="rotate(-90 ' + (p[0]-3) + ' ' + my + ')">' + esc(texte) + '</text>';
      if (p[0] > g.W - 44) return '<text class="' + cl + '" x="' + (p[0]+11) + '" y="' + my + '" text-anchor="middle" transform="rotate(-90 ' + (p[0]+11) + ' ' + my + ')">' + esc(texte) + '</text>';
      return '<text class="' + cl + '" x="' + (p[0]+6) + '" y="' + my + '" dominant-baseline="middle">' + esc(texte) + '</text>';
    }
    return '<text class="' + cl + '" x="' + mx + '" y="' + (my-6) + '" text-anchor="middle">' + esc(texte) + '</text>';
  }
  function dessiner(){
    if (!reseau.offsetWidth) return;
    const g = geo();
    svg.setAttribute('viewBox', '0 0 ' + g.W + ' ' + g.H);
    const montres = sel ? L.filter(l => l.de === sel || l.vers === sel) : L.filter(l => l.vue);
    let out = '<defs><marker id="pointe" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M1 1L9 5L1 9z" fill="context-stroke"/></marker></defs>', et = '';
    montres.forEach(l => {
      const a = g.o[l.de], b = g.o[l.vers];
      if (!a || !b) return;
      const pts = points(l.trace, a, b, g), actif = Boolean(sel);
      out += '<path class="lien' + (l.retour ? ' retour' : '') + (l.appui ? '' : ' ad') + (actif ? ' actif' : '') + '" d="M' + pts.map(p => p[0].toFixed(1) + ' ' + p[1].toFixed(1)).join(' L') + '" marker-end="url(#pointe)"/>';
      et += etiquette(pts, g, l.etiq, actif);
    });
    svg.innerHTML = out + et;
  }
  window.dessinerReseau = dessiner;

  function lienLi(l, sortant){
    const autre = sortant ? l.vers : l.de, f = parCode[autre];
    return '<div class="lien-li"><span class="dir">' + (sortant ? 'AGIT SUR →' : 'DÉPEND DE ←') + (l.retour ? ' · retour' : '') + '</span>' +
      '<button type="button" data-aller="' + autre + '">' + esc(f ? f.nom : autre) + '</button> · ' + esc(l.etiq) +
      '<small style="color:var(--doux)">' + esc(l.expl) + '</small>' +
      (l.appui ? '<small>Appui : ' + esc(l.appui) + '</small>' : '<small class="ad">Appui à documenter</small>') + '</div>';
  }
  function panneauRendu(){
    if (!sel){
      const nb = L.length, ret = L.filter(l => l.retour).length, ad = L.filter(l => !l.appui).length;
      panneau.innerHTML = '<h2>Un logement doit pouvoir être habité, et le rester</h2>' +
        '<p class="chapeau">Huit fonctions doivent se rencontrer. Chacune tient une condition; si une seule manque, le logement n\'existe pas ou ne reste pas accessible.</p>' +
        '<ol><li><b>Les conditions se rencontrent.</b> Règles, argent, sol et équipes rendent le projet possible.</li>' +
        '<li><b>Le parc dure.</b> Les logements sont détenus, habités, entretenus et rénovés.</li>' +
        '<li><b>Le système apprend.</b> Ce que vivent les ménages revient vers les connaissances, puis vers les règles.</li></ol>' +
        '<p class="chapeau" style="font-size:13px;margin-top:10px"><span class="num">' + nb + '</span> liens documentés, dont <span class="num">' + ret + '</span> retours. <span class="num">' + ad + '</span> lien' + (ad > 1 ? 's' : '') + ' encore sans appui.</p>' +
        '<p class="question">Qui peut agir sur la condition qui manque, et avec qui?</p>';
      return;
    }
    const f = parCode[sel], sort = L.filter(l => l.de === sel), entr = L.filter(l => l.vers === sel);
    panneau.innerHTML = '<h2><span class="code">' + f.code + '</span>' + esc(f.nom) + '</h2>' +
      '<dl class="kv"><dt>Condition tenue</dt><dd>' + esc(f.condition) + '</dd><dt>Si elle manque</dt><dd>' + esc(f.manque) + '</dd>' +
      '<dt>Portes</dt><dd class="num">' + portes(sel) + '</dd><dt>Acteurs principaux</dt><dd class="num">' + f.act.length + '</dd></dl>' +
      '<a class="lienfiche-ens" href="#fonctions" data-f="' + f.code + '">Voir ses acteurs, son chemin critique et ses leviers →</a>' +
      '<h3>Ses liens</h3>' + sort.map(l => lienLi(l, true)).join('') + entr.map(l => lienLi(l, false)).join('');
  }
  function maj(){
    const voisins = new Set(sel ? [sel, ...L.filter(l => l.de === sel || l.vers === sel).flatMap(l => [l.de, l.vers])] : F.map(f => f.code));
    noeuds.querySelectorAll('.noeud').forEach(n => { n.setAttribute('aria-pressed', n.dataset.f === sel); n.classList.toggle('dim', !voisins.has(n.dataset.f)); });
    tout.setAttribute('aria-pressed', !sel);
    panneauRendu();
    dessiner();
  }
  noeuds.addEventListener('click', e => { const b = e.target.closest('.noeud'); if (b){ sel = b.dataset.f === sel ? null : b.dataset.f; maj(); } });
  tout.addEventListener('click', () => { sel = null; maj(); });
  panneau.addEventListener('click', e => {
    const b = e.target.closest('[data-aller]');
    if (b){ sel = b.dataset.aller; maj(); return; }
    const a = e.target.closest('.lienfiche-ens');
    if (a && window.ouvrirFiche) window.ouvrirFiche(a.dataset.f);
  });
  if (window.ResizeObserver) new ResizeObserver(() => requestAnimationFrame(dessiner)).observe(reseau);
  window.addEventListener('resize', dessiner);
  maj();
})();

/* vues */
function vue(){
  const VUES = ['ensemble','carte','fonctions','familles','parcours','tableau'];
  const v = VUES.includes(location.hash.slice(1)) ? location.hash.slice(1) : 'ensemble';
  VUES.forEach(x => document.getElementById('vue-' + x).hidden = x !== v);
  document.querySelectorAll('nav a').forEach(a => a.setAttribute('aria-current', a.dataset.vue === v));
  if (v === 'ensemble' && window.dessinerReseau) requestAnimationFrame(window.dessinerReseau);
}
window.addEventListener('hashchange', vue); vue();
})();
</script>
</body>
</html>
"""

with open(OUT, "w", encoding="utf-8") as f:
    f.write(HTML.replace("__DATA__", DATA))
print("Page écrite :", OUT, "—", len(orgs), "organisations,", len(fonctions), "fonctions,", len(acteurs), "acteurs principaux,", len(parcours), "lignes de parcours,", len(leviers), "leviers,", len(liens), "liens")
