"""Road to Muscle Gain — version Streamlit fidèle au notebook Jupyter original.

Lancer : python -m streamlit run app.py
Données : fichier Excel synchronisé localement, sinon téléchargement Google Drive.

Le libellé « mois » du notebook original filtre les derniers BLOCS,
non des mois calendaires : cette sémantique est préservée ici.
"""
from __future__ import annotations

import io
import base64
import html
import math
from collections import Counter
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Patch, Rectangle
from matplotlib import font_manager
import numpy as np
import pandas as pd
import requests
import streamlit as st

# -----------------------------------------------------------------------------
# Config et palette identiques au notebook
# -----------------------------------------------------------------------------
FILE_ID = "1kDm-ua1YYQlHRQCHJDoPJaZFIOm0w83p"
EXCEL_NAME = "entrainement_blocs_donnees_brutes.xlsx"
FOND = "#142D24"
ORANGE = "#d95d00"
VERT = "#b5e222"
TEXTE = "#E8F0EB"
GRIS = "#C0C0C0"
POLICE_GRAPHIQUES = "Arial" if any(f.name == "Arial" for f in font_manager.fontManager.ttflist) else "DejaVu Sans"

PERIODES = ("Tout", "12 mois", "6 mois", "3 mois")
NB_BLOCS = {"Tout": None, "12 mois": 12, "6 mois": 6, "3 mois": 3}
CLES_COURBES = ("charges_realisees", "charges_prevues", "reps_realisees", "reps_prevues")
ONGLETS = ("Vue générale", "Pecs / épaules", "Dos / bras", "Jambes")

CSS = """
<style>
:root { --vert:#B5E222; --fond:#142D24; --orange:#D95D00; --espacement:12px; }
.stApp { background:#fff; color:#242d29; font-family:Arial,Helvetica,sans-serif; }
/* L'en-tête Streamlit est fixe : laisser assez de place aux 4 onglets ! */
.stApp .block-container { max-width:1220px; padding-top:4.2rem !important; padding-bottom:4rem; }
.stApp h1 { color:#D95D00 !important; font-size:clamp(2rem,3.4vw,3.1rem);
  font-weight:700; letter-spacing:-.03em; margin-top:0 !important; }
.stApp h2 { color:#D95D00 !important; font-size:clamp(1.65rem,2.5vw,2.35rem);
  font-weight:700; letter-spacing:-.03em; }
.stApp [data-testid="stVerticalBlock"] { gap:var(--espacement); }
/* Les onglets doivent rester visibles sous la barre supérieure, jamais masqués. */
.stApp [data-testid="stTabs"] [data-baseweb="tab-list"] {
  gap:0; display:flex !important; visibility:visible !important;
  position:relative; z-index:2; min-height:40px; margin-bottom:4px;
}
/* Apparence des onglets */
.stApp [data-testid="stTabs"] [data-baseweb="tab"] {
  background:#eee;
  border:1px solid #c5c5c5;
  color:#315246 !important;
  padding:.42rem .8rem;
  border-radius:0;
  font-family:Arial,sans-serif;
  font-size:15px;
}
/* Onglet sélectionné */
.stApp [data-testid="stTabs"] [data-baseweb="tab"][aria-selected="true"] {
  background:#fff !important;
  color:#D95D00 !important;
  border-top:3px solid #D95D00 !important;
}
/* Le texte suit la couleur de l'onglet */
.stApp [data-testid="stTabs"] [data-baseweb="tab"] p {
  color:inherit !important;
}
.stApp [data-testid="stTabs"] [data-baseweb="tab-highlight"] {
  background:transparent;
}
/* Période : une ligne compacte de quatre boutons répartis sur la largeur. */
[class*="st-key-periode_"] {
  width:65.2%; max-width:100%; margin:8px auto 0 auto;
  padding:12px 15px 10px; border-radius:18px; background:#142D24;
}
[class*="st-key-periode_"] [data-testid="stVerticalBlock"] { gap:7px !important; }
[class*="st-key-periode_"] [data-testid="stHorizontalBlock"] { gap:5px !important; }
[class*="st-key-periode_"] .period-title {
  color:#B5E222; font-weight:700; font-size:15px;
  margin:0 0 9px !important; line-height:1.3;
}
[class*="st-key-period_btn_"] button {
  background:#254438 !important; color:#E8F0EB !important;
  border:0 !important; border-radius:13px !important;
  height:34px !important; min-height:34px !important;
  padding:3px 5px !important; font-size:15px !important;
  font-family:Arial,Helvetica,sans-serif !important;
  width:100%; box-shadow:none !important;
}
[class*="st-key-period_btn_"][class*="_active"] button {
  background:#B5E222 !important; color:#142D24 !important;
}
/* Boutons de courbes : même fond #142D24 que les graphiques,
   symboles accolés au texte et groupe entier parfaitement centré. */
[class*="st-key-courbe_"] button,
[class*="st-key-courbe_"][class*="_active"] button,
[class*="st-key-courbe_"] button:hover {
  width:100% !important; height:35px !important; min-height:35px !important;
  background:#142D24 !important; color:#E8F0EB !important;
  border:1px solid #B5E222 !important; border-radius:13px !important;
  padding:4px 7px !important; font-size:15px !important;
  box-shadow:none !important; white-space:nowrap;
  display:flex !important; flex-direction:row !important;
  justify-content:center !important; align-items:center !important;
  gap:7px !important;
}
/* Streamlit insère un <p> (parfois entouré d'un <div>) dans le bouton.
   Il ne doit pas s'étirer et repousser le symbole au bord gauche. */
[class*="st-key-courbe_"] button > div,
[class*="st-key-courbe_"] button p {
  flex:0 0 auto !important; width:auto !important; max-width:max-content !important;
  margin:0 !important;
}
[class*="st-key-courbe_"] button::before {
  margin:0 !important; display:inline-block !important; flex:0 0 auto !important;
  line-height:1; font-size:16px;
}
[class*="st-key-courbe_"][class*="_charges_realisees_"] button::before {content:'●';color:#D95D00;}
[class*="st-key-courbe_"][class*="_charges_prevues_"] button::before {content:'┄';color:#E48E4C;}
[class*="st-key-courbe_"][class*="_reps_realisees_"] button::before {content:'●';color:#B5E222;}
[class*="st-key-courbe_"][class*="_reps_prevues_"] button::before {content:'┄';color:#C3E74E;}
/* Grille poids : les 3 cartes ont exactement la hauteur du graphique et
   les 12 px de gouttière sont identiques horizontalement/verticalement. */
.poids-grille { display:grid; grid-template-columns:minmax(0,1fr) minmax(0,2.7fr);
  gap:var(--espacement); width:100%; align-items:stretch; }
.poids-cartes { display:grid; grid-template-rows:repeat(3,minmax(0,1fr));
  gap:var(--espacement); min-width:0; }
.poids-grille img { display:block; width:100%; height:auto; align-self:stretch;
  border-radius:20px; min-width:0; }
.stat-card { background:#EDF2F0; padding:20px 23px; border-radius:12px;
  display:flex; flex-direction:column; justify-content:center;
  min-width:0; box-sizing:border-box; }
.stat-card .stat-label { color:#D95D00; font-size:18px; line-height:1.4; }
.stat-card .stat-valeur { color:#00441A; font-size:37px; line-height:1.15;
  font-weight:700; margin-top:3px; }
.stat-cartes-ligne { display:grid; grid-template-columns:repeat(3,minmax(0,1fr));
  gap:var(--espacement); margin-bottom:14px; }
.stat-cartes-ligne .stat-card { min-height:116px; }
.stApp [data-testid="stImage"] img { border-radius:20px; }
@media (max-width:750px) {
  .stApp .block-container { padding-top:4rem !important; }
  [class*="st-key-periode_"] { width:100%; padding:10px; }
  [class*="st-key-period_btn_"] button { font-size:12px !important; }
  .poids-grille { grid-template-columns:1fr; }
  .poids-cartes { grid-template-columns:repeat(3,minmax(0,1fr)); grid-template-rows:auto; }
  .stat-cartes-ligne { grid-template-columns:repeat(3,minmax(0,1fr)); }
  .stat-card { padding:13px 10px; }
  .stat-card .stat-label { font-size:13px; }
  .stat-card .stat-valeur { font-size:clamp(18px,4vw,26px); }
  [class*="st-key-courbe_"] button {font-size:11px;white-space:normal;height:43px;}
  [class*="st-key-courbe_"] button::before { margin-right:3px; }
}
@media (max-width:480px) {
  .poids-cartes, .stat-cartes-ligne { grid-template-columns:1fr; }
  .poids-cartes .stat-card { min-height:90px; }
}
</style>
"""


# -----------------------------------------------------------------------------
# Chargement Excel. Le répertoire Google Drive synchronisé a priorité en local.
# En ligne, Google Drive fournit le même fichier public.
# -----------------------------------------------------------------------------
@st.cache_data(ttl=900, show_spinner=False)
def _telecharger_excel() -> bytes:
    """Lit le classeur partagé ; plusieurs URLs évitent les pages HTML de Drive."""
    adresses = (
        f"https://drive.google.com/uc?export=download&id={FILE_ID}",
        f"https://drive.usercontent.google.com/download?id={FILE_ID}&export=download&confirm=t",
        f"https://docs.google.com/spreadsheets/d/{FILE_ID}/export?format=xlsx",
    )
    problemes = []
    for url in adresses:
        try:
            response = requests.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
            response.raise_for_status()
            if response.content.startswith(b"PK"):
                return response.content
            problemes.append("Google a renvoyé une page au lieu d'un fichier Excel")
        except requests.RequestException as erreur:
            problemes.append(type(erreur).__name__)
    raise RuntimeError(
        "Impossible de récupérer le classeur Excel public depuis Google Drive. "
        "Vérifie que le partage autorise toute personne disposant du lien "
        "(ou place le fichier XLSX à côté de app.py). "
        + ", ".join(problemes)
    )


@st.cache_data(ttl=900, show_spinner=False)
def _lire_classeur_local(chemin: str, mtime_ns: int) -> dict[str, pd.DataFrame]:
    # mtime_ns est volontaire : toute modification locale du fichier invalide le cache.
    with pd.ExcelFile(chemin) as classeur:
        return pd.read_excel(classeur, sheet_name=None)


@st.cache_data(ttl=900, show_spinner=False)
def _lire_classeur_distant(contenu: bytes) -> dict[str, pd.DataFrame]:
    with pd.ExcelFile(io.BytesIO(contenu)) as classeur:
        return pd.read_excel(classeur, sheet_name=None)


def charger_feuilles() -> dict[str, pd.DataFrame]:
    chemin = Path(__file__).resolve().parent / EXCEL_NAME
    if chemin.exists():
        return _lire_classeur_local(str(chemin), chemin.stat().st_mtime_ns)
    return _lire_classeur_distant(_telecharger_excel())


def valider_feuilles(feuilles: dict[str, pd.DataFrame]) -> None:
    obligations = {
        "releves_clean": ["bloc", "seance", "exercice", "serie_1"],
        "programme_clean": ["bloc", "seance", "exercice", "seance_type", "charge_prevue_clean"],
        "calendrier": ["bloc", "seance_type"],
        "seances": ["bloc", "semaine_cible", "duree_clean_min"],
        "pesees": ["bloc", "semaine", "poids_clean"],
    }
    for feuille, colonnes in obligations.items():
        if feuille not in feuilles:
            raise ValueError(f"La feuille Excel « {feuille} » est introuvable.")
        absentes = [col for col in colonnes if col not in feuilles[feuille].columns]
        if absentes:
            raise ValueError(f"Colonnes manquantes dans « {feuille} » : {', '.join(absentes)}")


# -----------------------------------------------------------------------------
# Calculs : formules et mesures identiques au notebook
# -----------------------------------------------------------------------------
def selectionner_colonnes(type_mesure: str) -> list[str]:
    return [f"serie_{i}_{type_mesure}_clean" for i in range(1, 5)]


def calculer_statistiques(donnees_series: pd.DataFrame):
    moyenne = donnees_series.apply(pd.to_numeric, errors="coerce").mean(axis=1)
    return moyenne, moyenne.min(), moyenne.max()


def moyenne_mesure(donnees: pd.DataFrame, type_mesure: str) -> pd.Series:
    colonnes = selectionner_colonnes(type_mesure)
    absentes = [c for c in colonnes if c not in donnees.columns]
    if absentes:
        # Les quatre colonnes existent dans le notebook. Un message explicite
        # évite de laisser une page blanche si un export Excel change de schéma.
        raise ValueError(f"Mesure {type_mesure} : colonnes manquantes : {', '.join(absentes)}")
    return calculer_statistiques(donnees[colonnes])[0]


def calculer_fidelite(valeur_prevue, valeur_realisee):
    if valeur_prevue == 0 or pd.isna(valeur_prevue) or pd.isna(valeur_realisee):
        return None
    return max(0, (1 - abs(valeur_prevue - valeur_realisee) / valeur_prevue) * 100)


def calculer_fidelite_moyenne(valeurs_prevues, valeurs_realisees):
    scores = [calculer_fidelite(p, r) for p, r in zip(valeurs_prevues, valeurs_realisees)]
    scores = [v for v in scores if v is not None and pd.notna(v)]
    return sum(scores) / len(scores) if scores else None


def calculer_progression(valeur_initiale, valeur_finale):
    if valeur_initiale == 0 or pd.isna(valeur_initiale) or pd.isna(valeur_finale):
        return None
    return (valeur_finale - valeur_initiale) / valeur_initiale * 100


def preparer_donnees(feuilles: dict[str, pd.DataFrame]):
    valider_feuilles(feuilles)
    releves = feuilles["releves_clean"].copy()
    releves["exercice"] = releves["exercice"].astype(str).str.lower().str.strip()
    realises = releves[releves["serie_1"].notna()].copy()
    programme = feuilles["programme_clean"].copy()
    programme["exercice"] = programme["exercice"].astype(str).str.lower().str.strip()
    type_par_exercice = programme.drop_duplicates("exercice").set_index("exercice")["seance_type"].to_dict()
    # .unique() respecte l'ordre de première apparition : identique au notebook.
    blocs_disponibles = realises["bloc"].dropna().unique().tolist()
    return releves, realises, programme, type_par_exercice, blocs_disponibles


def selectionner_periode(blocs_disponibles, periode: str):
    if periode not in NB_BLOCS:
        raise ValueError(f"Période inconnue : {periode}")
    nombre = NB_BLOCS[periode]
    return blocs_disponibles[-nombre:] if nombre is not None else blocs_disponibles


def filtrer_periode(releves, realises, blocs_disponibles, periode):
    blocs = selectionner_periode(blocs_disponibles, periode)
    return blocs, realises[realises["bloc"].isin(blocs)].copy(), releves[releves["bloc"].isin(blocs)].copy()


def statistiques_generales(feuilles, realises, prevus, blocs):
    faits = Counter(realises["exercice"].dropna().astype(str).str.lower().str.strip())
    attendus = Counter(prevus["exercice"].dropna().astype(str).str.lower().str.strip())
    total_faits = sum(faits.values())
    total_attendus = sum(attendus.values())
    taux = 100 * total_faits / total_attendus if total_attendus else np.nan
    calendrier = feuilles["calendrier"]
    calendrier = calendrier[calendrier["bloc"].isin(blocs)]
    types_seances = Counter(calendrier["seance_type"])
    nb_seances = sum(types_seances.get(t, 0) for t in ("A", "B", "C"))
    seances = feuilles["seances"]
    seances = seances[seances["bloc"].isin(blocs)]
    total_heures = pd.to_numeric(seances["duree_clean_min"], errors="coerce").sum() / 60
    pesees = feuilles["pesees"]
    pesees = pesees[pesees["bloc"].isin(blocs)].sort_values(["bloc", "semaine"]).copy()
    pesees["poids_clean"] = pd.to_numeric(pesees["poids_clean"], errors="coerce")
    pesees = pesees.dropna(subset=["poids_clean"])
    poids_initial = pesees["poids_clean"].iloc[0] if not pesees.empty else np.nan
    poids_actuel = pesees["poids_clean"].iloc[-1] if not pesees.empty else np.nan
    return dict(
        faits=faits, attendus=attendus, taux=taux, nb_seances=nb_seances,
        heures=total_heures, pesees=pesees, poids_initial=poids_initial,
        poids_actuel=poids_actuel, seances=seances,
    )


def records_exercices(realises):
    resultats = {}
    for nom, mesure in (("bench press", "charge"), ("back squat", "charge"), ("weighted pull-up", "reps")):
        donnees = realises[realises["exercice"].eq(nom)]
        moyenne = moyenne_mesure(donnees, mesure)
        resultats[nom] = moyenne.max() if not moyenne.empty else np.nan
    return resultats


def statistiques_progression(nom, realises, programme):
    donnees = realises[realises["exercice"].eq(nom)].sort_values(["bloc", "seance"]).copy()
    if donnees.empty:
        return None
    prevus = programme[programme["exercice"].eq(nom)].sort_values(["bloc", "seance"]).copy()
    # Les séances sont une clé unique dans le programme du notebook.
    # Si un import contient des doublons, ne pas multiplier les points du graphique.
    prevus = prevus.drop_duplicates(["bloc", "seance"], keep="first")
    comparaison = donnees.merge(prevus, on=["bloc", "seance"], how="left", suffixes=("_realise", "_prevu"))
    charge = moyenne_mesure(donnees, "charge").reset_index(drop=True)
    reps = moyenne_mesure(donnees, "reps").reset_index(drop=True)
    durees = moyenne_mesure(donnees, "duree_sec").reset_index(drop=True)
    charges_prevues = pd.to_numeric(comparaison["charge_prevue_clean"], errors="coerce").reset_index(drop=True)
    reps_prevues = moyenne_mesure(comparaison, "reps_prevues").reset_index(drop=True)
    durees_prevues = moyenne_mesure(comparaison, "duree_prevue").reset_index(drop=True)
    fidelites = (
        calculer_fidelite_moyenne(charges_prevues, charge),
        calculer_fidelite_moyenne(reps_prevues, reps),
        calculer_fidelite_moyenne(durees_prevues, durees),
    )
    valides = [v for v in fidelites if v is not None]
    respect = sum(valides) / len(valides) if valides else None
    sans_charge = nom in ("push-up", "dead bug")
    serie_progression = reps if sans_charge else charge
    progression = calculer_progression(serie_progression.iloc[0], serie_progression.iloc[-1])
    return dict(
        nom=nom, donnees=donnees, charge=charge, reps=reps, durees=durees,
        charges_prevues=charges_prevues, reps_prevues=reps_prevues,
        durees_prevues=durees_prevues, respect=respect, progression=progression,
        sans_charge=sans_charge,
    )


# -----------------------------------------------------------------------------
# Figures Matplotlib (couleurs, tailles et proportions du notebook)
# -----------------------------------------------------------------------------
def habiller_axes(fig, ax):
    fig.patch.set_facecolor(FOND)
    ax.set_facecolor(FOND)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_visible(False)
    ax.spines["bottom"].set_visible(True)
    ax.spines["bottom"].set_color(GRIS)
    ax.spines["bottom"].set_alpha(.20)
    ax.tick_params(axis="both", length=0, colors=TEXTE, labelsize=8)
    ax.grid(axis="y", color=GRIS, alpha=.20)
    ax.set_axisbelow(True)
    for tick in ax.get_xticklabels() + ax.get_yticklabels():
        tick.set_fontfamily(POLICE_GRAPHIQUES)


def figure_poids(pesees):
    if pesees.empty:
        return None
    blocs_semaine = [f"B{b} - S{s}" for b, s in zip(pesees["bloc"], pesees["semaine"])]
    nombres_blocs = pesees["bloc"].reset_index(drop=True)
    poids = pesees["poids_clean"].reset_index(drop=True)
    fig, ax = plt.subplots(figsize=(8, 4), dpi=160)
    x = np.arange(len(poids))
    ax.plot(x, poids, color=VERT, marker="o", markersize=4, markerfacecolor=ORANGE, markeredgewidth=0)
    pas = 4 if len(blocs_semaine) > 12 else 1
    ax.set_xticks(x[::pas], blocs_semaine[::pas])
    mini, maxi = poids.min(), poids.max()
    borne_basse = math.floor(mini * 2) / 2 - .5
    borne_haute = math.ceil(maxi * 2) / 2 + .5
    ax.set_ylim(borne_basse, borne_haute)
    ax.set_yticks(np.arange(borne_basse, borne_haute + .6, .5))
    for pos in range(1, len(nombres_blocs)):
        if nombres_blocs.iloc[pos] != nombres_blocs.iloc[pos - 1]:
            ax.axvline(pos - .5, color=GRIS, alpha=.3, linewidth=.7)
    habiller_axes(fig, ax)
    fig.tight_layout()
    fig.subplots_adjust(left=.08)
    return fig


def couleurs_frequences(faits):
    maximal = max(faits.values()) if faits else 0
    sortie = []
    for valeur in faits.values():
        if valeur == maximal:
            sortie.append(ORANGE)
        elif valeur >= maximal * .9:
            sortie.append("#ffcd3c")
        elif valeur >= maximal * .7:
            sortie.append(VERT)
        elif valeur >= maximal * .5:
            sortie.append("#74c69d")
        else:
            sortie.append("#1b4332")
    return sortie


def figure_frequences(faits, attendus, type_par_exercice):
    if not faits:
        return None
    noms = list(faits.keys())
    # Même ordre que le dictionnaire du notebook, pour les séparateurs y=5.5/11.5.
    fig, ax = plt.subplots(figsize=(10, 5), dpi=160)
    ax.barh(list(attendus.keys()), list(attendus.values()), color=GRIS, alpha=.2,
            label="Exercices prévus", height=.35)
    barres = ax.barh(noms, list(faits.values()), color=couleurs_frequences(faits),
                     label="Exercices faits", height=.35)
    maxi = max(faits.values())
    ax.set_xlabel("Fréquences", color=TEXTE)
    ax.bar_label(barres, padding=10, color=TEXTE, fontsize=8)
    ax.set_xlim(0, maxi + 3)
    for y in (5.5, 11.5):
        ax.axhline(y=y, color=GRIS, linewidth=.6, alpha=.2)
    for groupe, label, y in (("C", "Jambes", 2.5), ("A", "Dos/bras", 8.5), ("B", "Pecs/épaules", 14.5)):
        faits_groupe = sum(n for exo, n in faits.items() if type_par_exercice.get(exo) == groupe)
        prevus_groupe = sum(n for exo, n in attendus.items() if type_par_exercice.get(exo) == groupe)
        taux = 100 * faits_groupe / prevus_groupe if prevus_groupe else 0
        ax.text(maxi + .8, y, f"{label} - {taux:.1f} %", color=TEXTE, fontsize=8, fontweight="bold")
        ax.add_patch(Rectangle((maxi + .8, y - .5), 1.5, .12, facecolor="#315246", edgecolor="none"))
        ax.add_patch(Rectangle((maxi + .8, y - .5), 1.5 * min(taux / 100, 1), .12,
                               facecolor="#87f5cd", edgecolor="none"))
    ax.legend(handles=[Patch(facecolor=GRIS, alpha=.2, label="Exercices prévus"),
                       Patch(facecolor=ORANGE, label="Exercices faits")],
              fontsize=8, handlelength=1, handleheight=.7, frameon=False, labelcolor=TEXTE)
    for side in ("top", "right", "left", "bottom"):
        ax.spines[side].set_visible(False)
    ax.set_facecolor("none")
    fig.patch.set_facecolor(FOND)
    ax.tick_params(axis="both", length=0, colors=TEXTE, labelsize=9)
    ax.set_yticks(ax.get_yticks(), [s.capitalize() for s in list(dict.fromkeys(list(attendus) + noms))])
    fig.tight_layout()
    fig.subplots_adjust(left=.22)
    return fig


def figures_seances(seances):
    """Deux diagrammes : par bloc, puis par (bloc, semaine), comme Jupyter."""
    bloc_et_semaine = Counter(zip(seances["bloc"], seances["semaine_cible"]))
    total_bloc = Counter()
    for (bloc, _semaine), nombre in bloc_et_semaine.items():
        total_bloc[bloc] += nombre
    fig_bloc = _figure_barres_seances(
        list(total_bloc), list(total_bloc.values()), "bloc")
    labels = [f"B{b}-{s}" for b, s in bloc_et_semaine]
    fig_semaine = _figure_barres_seances(labels, list(bloc_et_semaine.values()), "semaine")
    return fig_bloc, fig_semaine


def _figure_barres_seances(labels, valeurs, type_graphique):
    if not valeurs:
        return None
    if type_graphique == "bloc":
        couleurs = [ORANGE if n >= 20 else "#ffcd3c" if n >= 18 else VERT if n >= 15
                    else "#74c69d" if n >= 10 else "#1b4332" for n in valeurs]
        axe_x, maxi = "Bloc d'entraînement", 30
    else:
        couleurs = [ORANGE if n >= 5 else "#ffcd3c" if n == 4 else VERT if n == 3
                    else "#74c69d" if n == 2 else "#1b4332" for n in valeurs]
        axe_x, maxi = "Bloc et semaine d'entraînement", 7
    fig, ax = plt.subplots(figsize=(10, 5), dpi=160)
    x = np.arange(len(valeurs))
    barres = ax.bar(x, valeurs, color=couleurs)
    ax.set_xlabel(axe_x, color=TEXTE)
    ax.set_ylabel("Nombre de séances", color=TEXTE)
    pas = 4 if type_graphique == "semaine" and len(labels) > 12 else 1
    ax.set_xticks(x[::pas], [str(label) for label in labels[::pas]])
    ax.bar_label(barres, padding=5, color=TEXTE, fontsize=8)
    ax.set_ylim(0, maxi)
    habiller_axes(fig, ax)
    fig.tight_layout()
    fig.subplots_adjust(left=.08)
    return fig


def _bornes_axe(ax, minimum, maximum):
    if pd.isna(minimum) or pd.isna(maximum):
        return
    bas, haut = minimum * .75, maximum * 1.25
    if bas == haut:
        bas, haut = (0, 1) if bas == 0 else (bas - 1, haut + 1)
    ax.set_ylim(bas, haut)


def figure_progression(stats, filtres):
    """Tracé directement transposé de graphique_progression() du notebook."""
    fig, ax1 = plt.subplots(figsize=(6.4, 4.8), dpi=160)
    ax2 = ax1.twinx()
    fig.patch.set_facecolor(FOND)
    ax1.patch.set_facecolor(FOND)
    ax2.patch.set_alpha(0)
    charge, reps, durees = stats["charge"], stats["reps"], stats["durees"]
    charges_prevues = stats["charges_prevues"]
    reps_prevues = stats["reps_prevues"]
    durees_prevues = stats["durees_prevues"]
    sans_charge = stats["sans_charge"]
    nb = len(stats["donnees"])
    x = np.arange(1, nb + 1)

    if charge.notna().any() and filtres["charges_realisees"] and not sans_charge:
        ax1.plot(x, charge, color=ORANGE, linewidth=2, marker="o", markersize=3.5, label="Charges moyennes")
    if charges_prevues.notna().any() and filtres["charges_prevues"] and not sans_charge:
        ax1.plot(x, charges_prevues, color="#e48e4c", linewidth=1.5, linestyle=(0, (2, 8)),
                 label="Charges prévues moyennes")
    if not reps.notna().any() and not (durees.notna().any() and charge.notna().any()):
        ax2.set_visible(False)
    if reps.notna().any() and filtres["reps_realisees"]:
        ax2.plot(x, reps, color=VERT, linewidth=1, marker="o", markersize=2.5, label="Reps moyennes")
    if reps_prevues.notna().any() and filtres["reps_prevues"]:
        ax2.plot(x, reps_prevues, color="#c3e74e", linewidth=.5, linestyle=(0, (2, 8)),
                 label="Reps prévues moyennes")
    if durees.notna().any() and not charge.notna().any():
        ax1.plot(x, durees, color="#74c69d", linewidth=1, marker="o", markersize=2,
                 label="Durées moyennes")
        ax1.plot(x, durees_prevues, color="#b7e4c7", linewidth=.5, linestyle=(0, (2, 8)),
                 label="Durées prévues moyennes")
    if durees.notna().any() and charge.notna().any():
        ax2.plot(x, durees, color="#74c69d", linewidth=1, marker="o", markersize=2,
                 label="Durées moyennes")
        ax2.plot(x, durees_prevues, color="#b7e4c7", linewidth=.5, linestyle=(0, (2, 8)),
                 label="Durées prévues moyennes")
    ax1.set_xlabel("Séances", color=GRIS)
    if reps.notna().any():
        ax2.set_ylabel("Reps moyennes", color=VERT)
    if charge.notna().any():
        _bornes_axe(ax1, charge.min(), charge.max())
    pas = 1 if nb <= 20 else max(1, round(nb / 9))
    ax1.set_xticks(x[::pas])
    if reps.notna().any():
        _bornes_axe(ax2, reps.min(), reps.max())
    if durees.notna().any() and not charge.notna().any():
        _bornes_axe(ax1, durees.min(), durees.max())
        ax1.set_ylabel("Durées moyennes (sec)", color="#74c69d")
    else:
        ax1.set_ylabel("Charges moyennes (kg)", color=ORANGE)
    if durees.notna().any() and charge.notna().any():
        _bornes_axe(ax2, durees.min(), durees.max())
        ax2.set_ylabel("Durées moyennes (sec)", color="#74c69d")
    blocs = stats["donnees"]["bloc"].reset_index(drop=True)
    changements = blocs != blocs.shift(1)
    changements.iloc[0] = False
    for pos, changement in enumerate(changements):
        if changement:
            ax1.axvline(pos + .5, color=GRIS, alpha=.3, linewidth=.7)
    ax1.set_title(stats["nom"].capitalize(), fontsize=18, pad=30, color=GRIS)
    blocs_uniques = blocs.unique()
    nb_cumules = 0
    for bloc in blocs_uniques:
        n = (blocs == bloc).sum()
        centre = (nb_cumules + 1 + nb_cumules + n) / 2
        try:
            mise_en_avant = (len(blocs_uniques) <= 6 or bloc == blocs_uniques[0]
                              or (bloc % 6 == 0 and (blocs_uniques[-1] % 6 >= 3 or blocs_uniques[-1] % 6 == 0))
                              or bloc == blocs_uniques[-1])
        except TypeError:
            mise_en_avant = True
        if mise_en_avant:
            ax1.text(centre, .76, f"B{bloc}", transform=ax1.get_xaxis_transform(),
                     ha="center", color=GRIS)
        nb_cumules += n
    for ax in (ax1, ax2):
        ax.tick_params(length=0, colors=GRIS, labelsize=9)
        for cote in ("top", "right", "left", "bottom"):
            ax.spines[cote].set_visible(False)
    ax1.grid(axis="y", color=GRIS, alpha=.10, linewidth=.5)
    respect = stats["respect"]
    progression = stats["progression"]
    texte_respect = f"Respect\n{respect:.1f} %" if respect is not None else "Respect\nN/D"
    texte_progression = f"Progression\n{progression:.1f} %" if progression is not None else "Progression\nN/D"
    for pos, texte in ((.40, texte_respect), (.60, texte_progression)):
        ax1.text(pos, .98, texte, transform=ax1.transAxes, ha="center", fontsize=10,
                 color=GRIS, linespacing=1.2,
                 bbox=dict(boxstyle="round,pad=0.6", facecolor="#23483A", edgecolor="#416B56", linewidth=1))
    return fig


# -----------------------------------------------------------------------------
# Interface Streamlit
# -----------------------------------------------------------------------------
def afficher_figure(fig):
    if fig is None:
        st.info("Pas de données pour cette période.")
        return
    st.pyplot(fig, use_container_width=True)
    plt.close(fig)



def carte(titre, valeur):
    """Carte HTML réutilisable, sans marge Streamlit entre les éléments."""
    return (
        '<div class="stat-card">'
        f'<div class="stat-label">{html.escape(str(titre))}</div>'
        f'<div class="stat-valeur">{html.escape(str(valeur))}</div>'
        '</div>'
    )


def ligne_cartes(cartes):
    st.markdown('<div class="stat-cartes-ligne">' + ''.join(cartes) + '</div>',
                unsafe_allow_html=True)


def image_figure_base64(fig):
    """PNG Matplotlib affiché dans une grille HTML : hauteur partagée avec les cartes."""
    if fig is None:
        return None
    flux = io.BytesIO()
    fig.savefig(flux, format='png', dpi=160, facecolor=fig.get_facecolor())
    plt.close(fig)
    return base64.b64encode(flux.getvalue()).decode('ascii')


def formater(valeur, unite=""):
    return f"{valeur:.1f}{unite}" if pd.notna(valeur) else "N/D"



def _choisir_periode(nouvelle_periode):
    """Un seul état, partagé par les 4 onglets."""
    st.session_state['periode_commune'] = nouvelle_periode


def bandeau_periode(indice):
    with st.container(key=f'periode_{indice}'):
        st.markdown('<p class="period-title">PÉRIODE D\'ANALYSE</p>', unsafe_allow_html=True)
        colonnes = st.columns(4, gap='small')
        for numero, (col, periode) in enumerate(zip(colonnes, PERIODES)):
            actif = st.session_state['periode_commune'] == periode
            with col:
                with st.container(key=f'period_btn_{indice}_{numero}_{"active" if actif else "inactive"}'):
                    st.button(periode, key=f'period_click_{indice}_{numero}',
                              use_container_width=True, on_click=_choisir_periode, args=(periode,))


def _basculer_courbe(cle):
    st.session_state[f"courbe_{cle}"] = not st.session_state[f"courbe_{cle}"]


def boutons_courbes(type_seance):
    etiquettes = (("charges_realisees", "Charges réalisées"),
                  ("charges_prevues", "Charges prévues"),
                  ("reps_realisees", "Reps réalisées"),
                  ("reps_prevues", "Reps prévues"))
    colonnes = st.columns(4, gap="small")
    for col, (cle, texte) in zip(colonnes, etiquettes):
        actif = st.session_state[f"courbe_{cle}"]
        # L'état est partagé par les trois onglets, comme les ToggleButtons Jupyter.
        with col:
            with st.container(key=f"courbe_{type_seance}_{cle}_{'active' if actif else 'inactive'}"):
                st.button(texte, key=f"bouton_{type_seance}_{cle}",
                          use_container_width=True, on_click=_basculer_courbe, args=(cle,))



def vue_generale(feuilles, realises, prevus, blocs, types):
    stats = statistiques_generales(feuilles, realises, prevus, blocs)
    st.title("Vue d'ensemble : programme prise de masse")
    evolution = stats['poids_actuel'] - stats['poids_initial']
    cartes_poids = ''.join((
        carte('Poids initial', formater(stats['poids_initial'], ' kg')),
        carte('Poids actuel', formater(stats['poids_actuel'], ' kg')),
        carte('Poids évolution', f'{evolution:+.1f} kg' if pd.notna(evolution) else 'N/D'),
    ))
    poids_png = image_figure_base64(figure_poids(stats['pesees']))
    if poids_png is not None:
        st.markdown(
            '<div class="poids-grille"><div class="poids-cartes">' + cartes_poids + '</div>'
            + f'<img alt="Évolution du poids" src="data:image/png;base64,{poids_png}">'
            + '</div>', unsafe_allow_html=True,
        )
    else:
        st.markdown('<div class="poids-cartes">' + cartes_poids + '</div>', unsafe_allow_html=True)
        st.info('Pas de données de poids pour cette période.')

    st.header('Suivi des entraînements')
    ligne_cartes([
        carte('Exercices accomplis', formater(stats['taux'], ' %')),
        carte('Séances réalisées', str(stats['nb_seances'])),
        carte('Heures cumulées', formater(stats['heures'], ' h')),
    ])
    afficher_figure(figure_frequences(stats['faits'], stats['attendus'], types))
    fig_bloc, fig_semaine = figures_seances(stats['seances'])
    afficher_figure(fig_bloc)
    afficher_figure(fig_semaine)
    st.header('Records sur 3 exercices phares')
    records = records_exercices(realises)
    ligne_cartes([
        carte('PR Développé couché', formater(records['bench press'], ' kg')),
        carte('PR Back squat', formater(records['back squat'], ' kg')),
        carte('PR Traction lestée 10kg', formater(records['weighted pull-up'], ' reps')),
    ])


def vue_exercices(type_seance, titre, realises, programme, types):
    st.title(titre)
    boutons_courbes(type_seance)
    filtres = {cle: st.session_state[f"courbe_{cle}"] for cle in CLES_COURBES}
    # Ordre original : première apparition dans les relevés, non tri alphabétique.
    exercices = [e for e in realises["exercice"].drop_duplicates()
                 if types.get(e) == type_seance]
    if not exercices:
        st.info("Pas d'exercice réalisé dans cette période.")
        return
    for position in range(0, len(exercices), 2):
        cols = st.columns(2, gap="small")
        for col, exo in zip(cols, exercices[position:position + 2]):
            with col:
                stats = statistiques_progression(exo, realises, programme)
                if stats is not None:
                    afficher_figure(figure_progression(stats, filtres))


def initialiser_etat():
    if "periode_commune" not in st.session_state:
        st.session_state["periode_commune"] = "Tout"
    for cle in CLES_COURBES:
        st.session_state.setdefault(f"courbe_{cle}", True)


def main():
    st.set_page_config(page_title="Road to Muscle Gain", layout="wide", initial_sidebar_state="collapsed")
    st.markdown(CSS, unsafe_allow_html=True)
    plt.rcParams["font.family"] = [POLICE_GRAPHIQUES]
    initialiser_etat()
    # Les onglets sont créés AVANT le téléchargement : jamais une page entièrement vide.
    onglets = st.tabs(list(ONGLETS))
    try:
        with st.spinner("Chargement du fichier Excel…"):
            feuilles = charger_feuilles()
            releves, realises, programme, types, blocs_disponibles = preparer_donnees(feuilles)
        if not blocs_disponibles:
            st.warning("Aucune séance réalisée dans le fichier Excel.")
            return
        blocs, realises_filtres, prevus_filtres = filtrer_periode(
            releves, realises, blocs_disponibles, st.session_state["periode_commune"])
        descriptions = (("B", "Progression des séances Pecs / épaules"),
                        ("A", "Progression des séances Dos / bras"),
                        ("C", "Progression des séances Jambes"))
        with onglets[0]:
            bandeau_periode(0)
            vue_generale(feuilles, realises_filtres, prevus_filtres, blocs, types)
        for i, (type_seance, titre) in enumerate(descriptions, start=1):
            with onglets[i]:
                bandeau_periode(i)
                vue_exercices(type_seance, titre, realises_filtres, programme, types)
    except Exception as erreur:
        st.error("Impossible d'afficher le dashboard : " + str(erreur))
        st.exception(erreur)


if __name__ == "__main__":
    main()
