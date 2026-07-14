# -*- coding: utf-8 -*-
"""
Module de logique métier : lecture du fichier de mouvements de stock,
calcul des indicateurs de pilotage et détection des anomalies.

Conçu pour le format réel transmis par le service (export ERP à plat, une ligne par
mouvement, colonne "Type de flux de mouvement" identifiant Sorties/Entrées/
Transfères/Ajustement) — ex. fichier "Mvts_02_2026.xlsx".

Ce module ne dépend d'aucun framework web : il est importé tel quel par le backend
Flask (app.py) et peut aussi être utilisé en ligne de commande / notebook.
"""

import re
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd


# ============================================================================
# Détection souple des colonnes (au cas où les libellés varient légèrement
# d'un mois à l'autre dans l'export ERP)
# ============================================================================

def _normaliser(texte):
    t = unicodedata.normalize("NFD", str(texte))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", t).strip().lower()


CANDIDATS_COLONNES = {
    "date": ["date du mouvement", "datem", "date mouvement"],
    "article": ["article"],
    "designation": ["version", "designation", "libelle"],
    "categorie": ["descr_section", "categorie", "section"],
    "quantite": ["qte", "quantite"],
    "pmp": ["pmp"],
    "montant": ["mnt", "montant", "valeur du mouvement"],
    "magasin": ["magasin"],
    "type_flux": ["type de flux de mouvement", "type de flux"],
    "type_detail": ["type"],
    "origine": ["type d'origine du mouvement", "origine du mouvement", "reference"],
}

# Colonnes à exclure des correspondances "partielles" trop larges (ex: "Type" seul
# ne doit pas capter "Type de flux de mouvement" ni "Type d'origine du mouvement").
COLONNES_EXACTES_UNIQUEMENT = {"type_detail"}


def _detecter_colonne(colonnes, candidats, exact_only=False):
    norm = {c: _normaliser(c) for c in colonnes}
    for cand in candidats:
        for original, n in norm.items():
            if n == cand:
                return original
    if not exact_only:
        for cand in candidats:
            for original, n in norm.items():
                if cand in n:
                    return original
    return None


def detecter_mapping(colonnes):
    mapping = {}
    for champ, candidats in CANDIDATS_COLONNES.items():
        mapping[champ] = _detecter_colonne(
            colonnes, candidats, exact_only=(champ in COLONNES_EXACTES_UNIQUEMENT)
        )
    return mapping


def choisir_feuille(chemin_fichier):
    """Choisit, parmi les feuilles du classeur, celle qui contient une colonne de
    type de flux de mouvement (la feuille de données ; les autres feuilles éventuelles
    — sommaires, feuilles vides — sont ignorées)."""
    xl = pd.ExcelFile(chemin_fichier)
    for nom in xl.sheet_names:
        entete = pd.read_excel(chemin_fichier, sheet_name=nom, nrows=0).columns
        mapping = detecter_mapping(entete)
        if mapping["type_flux"] and mapping["article"] and mapping["quantite"]:
            return nom, mapping
    raise ValueError(
        "Aucune feuille du classeur ne contient les colonnes attendues "
        "(Article, Qté, Type de flux de mouvement, ...)."
    )


# ============================================================================
# Parsing des dates ERP
# ============================================================================

MOIS_FR = {
    "JAN": 1, "FEV": 2, "MAR": 3, "AVR": 4, "MAI": 5, "JUN": 6,
    "JUI": 7, "JUL": 7, "AOU": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
}


def parser_date_erp(valeur):
    if pd.isna(valeur):
        return pd.NaT
    if isinstance(valeur, (pd.Timestamp, datetime)):
        return pd.Timestamp(valeur)
    texte = str(valeur).strip().upper()
    m = re.match(r"^(\d{1,2})-([A-Z]{3})-(\d{4})(?:\s+(\d{1,2}):(\d{2}):(\d{2}))?", texte)
    if m:
        jour, mois_abr, annee = int(m.group(1)), m.group(2), int(m.group(3))
        mois = MOIS_FR.get(mois_abr)
        if mois:
            h, mi, s = int(m.group(4) or 0), int(m.group(5) or 0), int(m.group(6) or 0)
            return pd.Timestamp(annee, mois, jour, h, mi, s)
    return pd.to_datetime(valeur, errors="coerce", dayfirst=True)


# ============================================================================
# Chargement et normalisation
# ============================================================================

def charger_mouvements(chemin_fichier):
    """Charge le classeur, détecte la feuille et les colonnes pertinentes, et retourne
    un DataFrame normalisé (schéma stable quel que soit le fichier source) ainsi que
    la liste des avertissements rencontrés."""
    avertissements = []
    nom_feuille, mapping = choisir_feuille(chemin_fichier)
    df = pd.read_excel(chemin_fichier, sheet_name=nom_feuille)
    df.columns = [str(c).strip() for c in df.columns]
    # re-détecter sur les colonnes réelles (au cas où les strip changent la correspondance)
    mapping = detecter_mapping(df.columns)

    manquants_requis = [c for c in ("article", "quantite", "type_flux") if not mapping[c]]
    if manquants_requis:
        raise ValueError(f"Colonnes requises introuvables : {manquants_requis}")

    d = pd.DataFrame()
    d["Article"] = df[mapping["article"]].astype(str).str.strip()
    d["Désignation"] = df[mapping["designation"]].astype(str).str.strip() if mapping["designation"] else ""
    d["Catégorie"] = df[mapping["categorie"]].astype(str).str.strip() if mapping["categorie"] else "(non classé)"
    d["Qté"] = pd.to_numeric(df[mapping["quantite"]], errors="coerce").fillna(0.0)
    d["PMP"] = pd.to_numeric(df[mapping["pmp"]], errors="coerce") if mapping["pmp"] else np.nan
    d["Montant"] = pd.to_numeric(df[mapping["montant"]], errors="coerce").fillna(0.0) if mapping["montant"] else 0.0
    d["Magasin"] = df[mapping["magasin"]].astype(str).str.strip().replace("nan", "(sans magasin)") if mapping["magasin"] else "(sans magasin)"
    d["Magasin"] = d["Magasin"].replace("", "(sans magasin)")
    d["TypeFlux"] = df[mapping["type_flux"]].astype(str).str.strip()
    d["NatureConso"] = df[mapping["type_detail"]].astype(str).str.strip() if mapping["type_detail"] else ""
    d["DateM"] = df[mapping["date"]].apply(parser_date_erp) if mapping["date"] else pd.NaT

    # normalisation des libellés de flux (ex: "Transfères"/"Transferts" -> "Transferts")
    def _norm_flux(v):
        n = _normaliser(v)
        if "sortie" in n:
            return "Sorties"
        if "entr" in n:
            return "Entrées"
        if "transf" in n:
            return "Transferts"
        if "ajust" in n:
            return "Ajustement"
        return v or "(non défini)"

    d["TypeFlux"] = d["TypeFlux"].apply(_norm_flux)
    d["AbsQte"] = d["Qté"].abs()
    d = d.dropna(subset=["Article"])
    d = d[d["Article"].str.len() > 0].reset_index(drop=True)

    if d["DateM"].isna().all():
        avertissements.append("Aucune date de mouvement exploitable : la tendance temporelle ne peut pas être calculée.")
    if mapping["montant"] is None:
        avertissements.append("Colonne de montant non trouvée : les indicateurs de valeur seront à 0.")

    return d, avertissements


# ============================================================================
# Indicateurs de pilotage
# ============================================================================

def kpis_generaux(df):
    dates = df["DateM"].dropna()
    return {
        "nb_mouvements": int(len(df)),
        "nb_articles": int(df["Article"].nunique()),
        "nb_magasins": int(df["Magasin"].nunique()),
        "valeur_totale": float(df["Montant"].abs().sum()),
        "periode_debut": dates.min().strftime("%d/%m/%Y") if len(dates) else None,
        "periode_fin": dates.max().strftime("%d/%m/%Y") if len(dates) else None,
    }


def repartition_par_type_flux(df):
    s = df["TypeFlux"].value_counts()
    return [{"type": k, "nb": int(v)} for k, v in s.items()]


def activite_par_magasin(df):
    g = df.groupby("Magasin").agg(nb=("Magasin", "count"), valeur=("Montant", lambda x: float(x.abs().sum())))
    g = g.sort_values("nb", ascending=False).reset_index()
    return g.to_dict("records")


def top_articles_quantite(df, n=10):
    sorties = df[df["TypeFlux"] == "Sorties"]
    g = (sorties.groupby(["Article", "Désignation"])["Qté"].apply(lambda x: float(x.abs().sum()))
         .reset_index(name="qte").sort_values("qte", ascending=False).head(n))
    return g.to_dict("records")


def top_articles_valeur(df, n=10):
    sorties = df[df["TypeFlux"] == "Sorties"]
    g = (sorties.groupby(["Article", "Désignation"])["Montant"].apply(lambda x: float(x.abs().sum()))
         .reset_index(name="valeur").sort_values("valeur", ascending=False).head(n))
    return g.to_dict("records")


def categories_plus_couteuses(df, n=10):
    g = (df.groupby("Catégorie")["Montant"].apply(lambda x: float(x.abs().sum()))
         .reset_index(name="valeur").sort_values("valeur", ascending=False).head(n))
    return g.to_dict("records")


def tendance_mensuelle(df):
    """Tendance par semaine calendaire (le mot 'mensuelle' du cahier des charges
    correspond, sur un export d'un seul mois, à un suivi infra-mensuel par semaine)."""
    d = df.dropna(subset=["DateM"]).copy()
    if d.empty:
        return []
    d["Semaine"] = d["DateM"].dt.to_period("W").apply(lambda p: p.start_time.strftime("%d/%m"))
    d["MontantSigne"] = d["Montant"].abs() * np.sign(d["Qté"])
    g = d.groupby("Semaine").agg(nb=("DateM", "count"), valeur_nette=("MontantSigne", "sum")).reset_index()
    g["ordre"] = pd.to_datetime(g["Semaine"] + "/2026", format="%d/%m/%Y", errors="coerce")
    g = g.sort_values("ordre").drop(columns="ordre")
    return g.to_dict("records")


def calculer_indicateurs(df):
    return {
        "kpis": kpis_generaux(df),
        "repartition_flux": repartition_par_type_flux(df),
        "activite_magasins": activite_par_magasin(df),
        "top_articles_quantite": top_articles_quantite(df),
        "top_articles_valeur": top_articles_valeur(df),
        "categories_couteuses": categories_plus_couteuses(df),
        "tendance": tendance_mensuelle(df),
    }


# ============================================================================
# Détection des anomalies (6 règles + rupture de stock)
# ============================================================================

def _lignes_dict(df):
    """Sérialise un sous-ensemble de lignes pour la réponse JSON / l'export."""
    cols = ["DateM", "Article", "Désignation", "Magasin", "TypeFlux", "Qté", "Montant"]
    out = df[cols].copy()
    out["DateM"] = out["DateM"].apply(lambda d: d.strftime("%d/%m/%Y %H:%M") if pd.notna(d) else "")
    out = out.rename(columns={"DateM": "Date", "TypeFlux": "Flux"})
    return out.round({"Qté": 2, "Montant": 2}).to_dict("records")


def calculer_stock_cumule(df):
    """Stock théorique cumulé par article, base 0 en début de période (aucun stock
    d'ouverture n'est fourni dans ce fichier)."""
    d = df.sort_values(["Article", "DateM"], na_position="first").copy()
    d["StockCumule"] = d.groupby("Article")["Qté"].cumsum()
    d["StockAvant"] = d["StockCumule"] - d["Qté"]
    return d


def detect_stock_negatif(df_stock):
    return df_stock[df_stock["StockCumule"] < 0]


def detect_rupture_stock(df_stock):
    """Mouvements après lesquels le stock cumulé de l'article retombe exactement à 0
    (rupture / stock devenu vide), alors qu'il était positif juste avant."""
    return df_stock[(df_stock["StockCumule"] == 0) & (df_stock["StockAvant"] > 0)]


def detect_sortie_superieure_stock(df_stock):
    sorties = df_stock[df_stock["TypeFlux"] == "Sorties"]
    return sorties[sorties["Qté"].abs() > sorties["StockAvant"]]


def detect_consommation_excessive(df, seuil_z=3.0):
    sorties = df[df["TypeFlux"] == "Sorties"].copy()
    stats = sorties.groupby("Article")["AbsQte"].agg(["mean", "std", "count"])
    stats = stats[stats["count"] >= 3]
    sorties = sorties.merge(stats, on="Article", how="inner")
    sorties["z"] = (sorties["AbsQte"] - sorties["mean"]) / sorties["std"].replace(0, np.nan)
    return sorties[sorties["z"] > seuil_z].sort_values("z", ascending=False)


def detect_rotation_faible(df):
    compte = df.groupby("Article").size()
    return df[df["Article"].isin(compte[compte == 1].index)]


def detect_ajustements_frequents(df, seuil=2):
    aj = df[df["TypeFlux"] == "Ajustement"]
    compte = aj.groupby("Article").size()
    return aj[aj["Article"].isin(compte[compte >= seuil].index)]


def detect_sorties_sans_entree(df):
    types_par_article = df.groupby("Article")["TypeFlux"].apply(set)
    articles = [a for a, t in types_par_article.items() if "Sorties" in t and "Entrées" not in t]
    return df[df["Article"].isin(articles)]


RULE_DESCRIPTIONS = {
    "stock_negatif": "Stock théorique cumulé (base 0 en début de période) devenu négatif après un mouvement.",
    "rupture_stock": "Stock d'un article tombé exactement à zéro après un mouvement de sortie — rupture potentielle.",
    "sortie_superieure_stock": "Sortie dont la quantité dépasse le stock disponible juste avant le mouvement.",
    "consommation_excessive": "Sortie dont la quantité s'écarte fortement (score-z) de la consommation habituelle de l'article.",
    "rotation_faible": "Article ne présentant qu'un seul mouvement sur toute la période analysée.",
    "ajustements_frequents": "Article ayant subi plusieurs ajustements de coût moyen pondéré sur la période.",
    "sorties_sans_entree": "Article uniquement mouvementé en sortie, sans aucune entrée enregistrée sur la période.",
}


def calculer_anomalies(df, seuil_z=3.0, seuil_ajust=2):
    df_stock = calculer_stock_cumule(df)
    resultats = {
        "stock_negatif": detect_stock_negatif(df_stock),
        "rupture_stock": detect_rupture_stock(df_stock),
        "sortie_superieure_stock": detect_sortie_superieure_stock(df_stock),
        "consommation_excessive": detect_consommation_excessive(df, seuil_z),
        "rotation_faible": detect_rotation_faible(df),
        "ajustements_frequents": detect_ajustements_frequents(df, seuil_ajust),
        "sorties_sans_entree": detect_sorties_sans_entree(df),
    }
    serialise = {}
    for cle, sous_df in resultats.items():
        serialise[cle] = {
            "libelle": {
                "stock_negatif": "Stock cumulé négatif",
                "rupture_stock": "Stock devenu vide (rupture)",
                "sortie_superieure_stock": "Sortie > stock disponible",
                "consommation_excessive": "Consommation excessive",
                "rotation_faible": "Rotation très faible",
                "ajustements_frequents": "Ajustements fréquents",
                "sorties_sans_entree": "Sorties sans entrée",
            }[cle],
            "description": RULE_DESCRIPTIONS[cle],
            "nb": int(len(sous_df)),
            "lignes": _lignes_dict(sous_df.head(500)),
        }
    return serialise, resultats


def analyser_fichier(chemin_fichier, seuil_z=3.0, seuil_ajust=2):
    """Point d'entrée unique utilisé par le backend Flask."""
    df, avertissements = charger_mouvements(chemin_fichier)
    indicateurs = calculer_indicateurs(df)
    anomalies_json, anomalies_df = calculer_anomalies(df, seuil_z, seuil_ajust)
    return {
        "avertissements": avertissements,
        "indicateurs": indicateurs,
        "anomalies": anomalies_json,
    }, df, anomalies_df
