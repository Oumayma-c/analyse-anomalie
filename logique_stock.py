# -*- coding: utf-8 -*-
"""
Module de logique métier : lecture du fichier de mouvements de stock,
calcul des indicateurs de pilotage et détection des anomalies.

Conçu pour être universel :
1. Détecte si le fichier contient des onglets séparés de mouvements (Entrées, Sorties, Ajustements, Transferts) et les fusionne.
2. Sinon, scanne intelligemment tous les onglets du fichier pour trouver celui qui contient les colonnes nécessaires (Article et Quantité), évitant les onglets vides ou d'introduction.
"""

import re
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd


# ============================================================================
# Détection et Normalisation des Textes
# ============================================================================

def _normaliser(texte):
    t = unicodedata.normalize("NFD", str(texte))
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", t).strip().lower()


CANDIDATS_COLONNES = {
    "date": ["date", "date du mouvement", "datem", "date mouvement", "date_mvt", "date mvt"],
    "article": ["article", "code", "code article", "ref", "reference", "item", "code_article"],
    "designation": ["version", "designation", "libelle", "nom", "description", "libelle article", "des_article"],
    "categorie": ["descr_section", "categorie", "section", "famille", "groupe", "description_categorie"],
    "quantite": ["qte", "quantite", "quantites", "qty", "qté", "nombre", "volume"],
    "pmp": ["pmp", "prix moyen", "prix unitaire", "pu", "paf"],
    "montant": ["mnt", "montant", "valeur du mouvement", "valeur", "total", "montant total", "valeur mvt"],
    "magasin": ["magasin", "depot", "emplacement", "site", "stock", "mag"],
    "type_flux": ["type de flux de mouvement", "type de flux", "type flux", "flux", "mouvement", "type_mvt", "sens", "type"],
    "type_detail": ["type"],
    "origine": ["type d'origine du mouvement", "origine du mouvement", "reference", "origine"],
}

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
# Chargement intelligent multi-onglets & mono-onglet universel
# ============================================================================

def charger_mouvements(chemin_fichier):
    avertissements = []
    xl = pd.ExcelFile(chemin_fichier)
    toutes_feuilles = xl.sheet_names

    mappage_feuilles = {
        "entree": "Entrées",
        "sortie": "Sorties",
        "ajust": "Ajustement",
        "transf": "Transferts"
    }

    dfs_a_combiner = []
    onglets_mouvements_trouves = False

    for nom_f in toutes_feuilles:
        nom_normalise = _normaliser(nom_f)
        type_flux_detecte = None
        for cle, label in mappage_feuilles.items():
            if cle in nom_normalise:
                type_flux_detecte = label
                break
        
        if type_flux_detecte:
            df_temp = pd.read_excel(chemin_fichier, sheet_name=nom_f)
            df_temp.columns = [str(c).strip() for c in df_temp.columns]
            mapping = detecter_mapping(df_temp.columns)

            if mapping["article"] and mapping["quantite"]:
                onglets_mouvements_trouves = True
                d_sub = pd.DataFrame()
                d_sub["Article"] = df_temp[mapping["article"]].astype(str).str.strip()
                d_sub["Désignation"] = df_temp[mapping["designation"]].astype(str).str.strip() if mapping["designation"] else ""
                d_sub["Catégorie"] = df_temp[mapping["categorie"]].astype(str).str.strip() if mapping["categorie"] else "(non classé)"
                
                qtes_brutes = pd.to_numeric(df_temp[mapping["quantite"]], errors="coerce").fillna(0.0)
                if type_flux_detecte == "Sorties":
                    d_sub["Qté"] = -qtes_brutes.abs()
                else:
                    d_sub["Qté"] = qtes_brutes.abs()

                d_sub["PMP"] = pd.to_numeric(df_temp[mapping["pmp"]], errors="coerce") if mapping["pmp"] else np.nan
                d_sub["Montant"] = pd.to_numeric(df_temp[mapping["montant"]], errors="coerce").fillna(0.0) if mapping["montant"] else 0.0
                d_sub["Magasin"] = df_temp[mapping["magasin"]].astype(str).str.strip().replace("nan", "(sans magasin)") if mapping["magasin"] else "(sans magasin)"
                d_sub["Magasin"] = d_sub["Magasin"].replace("", "(sans magasin)")
                
                d_sub["TypeFlux"] = type_flux_detecte
                d_sub["NatureConso"] = df_temp[mapping["type_detail"]].astype(str).str.strip() if mapping["type_detail"] else ""
                d_sub["DateM"] = df_temp[mapping["date"]].apply(parser_date_erp) if mapping["date"] else pd.NaT
                d_sub["AbsQte"] = d_sub["Qté"].abs()

                dfs_a_combiner.append(d_sub)

    if not onglets_mouvements_trouves:
        feuille_trouvee = None
        mapping_trouve = None
        
        for nom_f in toutes_feuilles:
            try:
                df_test = pd.read_excel(chemin_fichier, sheet_name=nom_f, nrows=5)
                df_test.columns = [str(c).strip() for c in df_test.columns]
                m = detecter_mapping(df_test.columns)
                if m["article"] and m["quantite"]:
                    feuille_trouvee = nom_f
                    mapping_trouve = m
                    break
            except Exception:
                continue

        if feuille_trouvee is None:
            raise ValueError(
                "Aucun onglet contenant les colonnes requises (Article et Quantité) n'a été trouvé dans ce fichier."
            )

        df_temp = pd.read_excel(chemin_fichier, sheet_name=feuille_trouvee)
        df_temp.columns = [str(c).strip() for c in df_temp.columns]
        mapping = mapping_trouve

        df_final = pd.DataFrame()
        df_final["Article"] = df_temp[mapping["article"]].astype(str).str.strip()
        df_final["Désignation"] = df_temp[mapping["designation"]].astype(str).str.strip() if mapping["designation"] else ""
        df_final["Catégorie"] = df_temp[mapping["categorie"]].astype(str).str.strip() if mapping["categorie"] else "(non classé)"
        df_final["Qté"] = pd.to_numeric(df_temp[mapping["quantite"]], errors="coerce").fillna(0.0)
        df_final["PMP"] = pd.to_numeric(df_temp[mapping["pmp"]], errors="coerce") if mapping["pmp"] else np.nan
        df_final["Montant"] = pd.to_numeric(df_temp[mapping["montant"]], errors="coerce").fillna(0.0) if mapping["montant"] else 0.0
        df_final["Magasin"] = df_temp[mapping["magasin"]].astype(str).str.strip().replace("nan", "(sans magasin)") if mapping["magasin"] else "(sans magasin)"
        df_final["Magasin"] = df_final["Magasin"].replace("", "(sans magasin)")
        df_final["TypeFlux"] = df_temp[mapping["type_flux"]].astype(str).str.strip() if mapping["type_flux"] else "Mouvement"
        df_final["NatureConso"] = df_temp[mapping["type_detail"]].astype(str).str.strip() if mapping["type_detail"] else ""
        df_final["DateM"] = df_temp[mapping["date"]].apply(parser_date_erp) if mapping["date"] else pd.NaT
        df_final["AbsQte"] = df_final["Qté"].abs()
    else:
        df_final = pd.concat(dfs_a_combiner, ignore_index=True)

    df_final = df_final.dropna(subset=["Article"])
    df_final = df_final[df_final["Article"].str.len() > 0].reset_index(drop=True)

    if df_final["DateM"].isna().all():
        avertissements.append("Aucune date de mouvement exploitable : la tendance temporelle ne peut pas être calculée.")

    return df_final, avertissements


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
    df_temp = df.copy()
    df_temp["AbsMontant"] = df_temp["Montant"].abs()
    g = df_temp.groupby("Magasin").agg(
        nb=("Magasin", "count"), 
        valeur=("AbsMontant", "sum")
    )
    g = g.sort_values("nb", ascending=False).reset_index()
    return g.to_dict("records")


def top_articles_quantite(df, n=10):
    sorties = df[df["TypeFlux"] == "Sorties"].copy()
    if sorties.empty:
        sorties = df.copy()
    sorties["AbsQte"] = sorties["Qté"].abs()
    g = (sorties.groupby(["Article", "Désignation"])["AbsQte"].sum()
         .reset_index(name="qte").sort_values("qte", ascending=False).head(n))
    return g.to_dict("records")


def top_articles_valeur(df, n=10):
    sorties = df[df["TypeFlux"] == "Sorties"].copy()
    if sorties.empty:
        sorties = df.copy()
    sorties["AbsMontant"] = sorties["Montant"].abs()
    g = (sorties.groupby(["Article", "Désignation"])["AbsMontant"].sum()
         .reset_index(name="valeur").sort_values("valeur", ascending=False).head(n))
    return g.to_dict("records")


def categories_plus_couteuses(df, n=10):
    df_temp = df.copy()
    df_temp["AbsMontant"] = df_temp["Montant"].abs()
    g = (df_temp.groupby("Catégorie")["AbsMontant"].sum()
         .reset_index(name="valeur").sort_values("valeur", ascending=False).head(n))
    return g.to_dict("records")


def evolution_categories_mensuelle(df, n=10):
    df_temp = df.dropna(subset=["DateM"]).copy()
    if df_temp.empty:
        return []
    
    # Création de l'axe temporel standardisé YYYY-MM
    df_temp["Mois"] = df_temp["DateM"].dt.strftime("%Y-%m")
    df_temp["AbsMontant"] = df_temp["Montant"].abs()
    
    # Isolation des familles principales pour ne pas saturer l'affichage
    top_cats = df_temp.groupby("Catégorie")["AbsMontant"].sum().nlargest(n).index
    df_filtered = df_temp[df_temp["Catégorie"].isin(top_cats)]
    
    # Agrégation croisée
    g = df_filtered.groupby(["Mois", "Catégorie"])["AbsMontant"].sum().reset_index()
    
    # Label propre pour l'affichage final (ex: "05/2026")
    g["Mois_Affichage"] = g["Mois"].apply(lambda x: datetime.strptime(x, "%Y-%m").strftime("%m/%Y"))
    g = g.sort_values("Mois")
    
    return g.to_dict("records")


def tendance_mensuelle(df):
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
        "evolution_categories": evolution_categories_mensuelle(df), # Inclus pour l'onglet dynamique
    }


# ============================================================================
# Détection des anomalies (6 règles + rupture de stock)
# ============================================================================

def _lignes_dict(df):
    cols = ["DateM", "Article", "Désignation", "Magasin", "TypeFlux", "Qté", "Montant"]
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    out = df[cols].copy()
    out["DateM"] = out["DateM"].apply(lambda d: d.strftime("%d/%m/%Y %H:%M") if pd.notna(d) and hasattr(d, "strftime") else "")
    out = out.rename(columns={"DateM": "Date", "TypeFlux": "Flux"})
    return out.round({"Qté": 2, "Montant": 2}).to_dict("records")


def calculer_stock_cumule(df):
    d = df.sort_values(["Article", "DateM"], na_position="first").copy()
    d["StockCumule"] = d.groupby("Article")["Qté"].cumsum()
    d["StockAvant"] = d["StockCumule"] - d["Qté"]
    return d


def detect_stock_negatif(df_stock):
    return df_stock[df_stock["StockCumule"] < 0]


def detect_rupture_stock(df_stock):
    return df_stock[(df_stock["StockCumule"] == 0) & (df_stock["StockAvant"] > 0)]


def detect_sortie_superieure_stock(df_stock):
    sorties = df_stock[df_stock["TypeFlux"] == "Sorties"]
    return sorties[sorties["Qté"].abs() > sorties["StockAvant"]]


def detect_consommation_excessive(df, seuil_z=3.0):
    sorties = df[df["TypeFlux"] == "Sorties"].copy()
    if sorties.empty:
        return pd.DataFrame(columns=df.columns)
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
    df, avertissements = charger_mouvements(chemin_fichier)
    indicateurs = calculer_indicateurs(df)
    anomalies_json, anomalies_df = calculer_anomalies(df, seuil_z, seuil_ajust)
    return {
        "avertissements": avertissements,
        "indicateurs": indicateurs,
        "anomalies": anomalies_json,
    }, df, anomalies_df
