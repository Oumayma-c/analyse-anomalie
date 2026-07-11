import pandas as pd
import numpy as np

def analyser_fichier(chemin_ou_flux_excel, col_article_override=None, col_flux_override=None):
    """
    Moteur algorithmique de diagnostic logistique.
    Détecte automatiquement l'en-tête, mappe les colonnes opérationnelles
    et calcule les indicateurs d'anomalies de stock.
    Accepte un chemin de fichier (str) ou un objet en mémoire (BytesIO).
    """
    # 1. Lecture intelligente du fichier Excel (recherche du bon en-tête)
    try:
        df_test = pd.read_excel(chemin_ou_flux_excel, nrows=10)
    except Exception as e:
        raise ValueError(f"Erreur lors de la première lecture du fichier : {e}")
    
    meilleur_header = 0
    max_colonnes_valides = 0
    for i in range(min(5, len(df_test))):
        try:
            test_df = pd.read_excel(chemin_ou_flux_excel, header=i, nrows=5)
            cols_valides = sum(1 for c in test_df.columns if not str(c).startswith('Unnamed:'))
            if cols_valides > max_colonnes_valides:
                max_colonnes_valides = cols_valides
                meilleur_header = i
        except:
            continue

    df = pd.read_excel(chemin_ou_flux_excel, header=meilleur_header)
    
    # Nettoyage initial des en-têtes de colonnes
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(how='all', axis=1)
    df = df.loc[:, ~df.columns.str.contains('^Unnamed:')]
    
    df_propre = df.copy()
    
    # 2. Mappage intelligent et élargi des colonnes
    dictionnaire_colonnes = {
        "Date": ["date", "date mouvement", "dates", "jour", "période", "créé le", "le", "timestamp", "annee", "mois", "dt"],
        "Article": ["article", "code article", "ref", "reference", "référence", "id", "code_art", "code", "matériel", "sku", "art", "no_article", "item", "item_code", "produit", "code produit", "designation_code"],
        "Designation": ["designation", "désignation", "libelle", "nom", "produit", "item", "description", "libellé", "nom produit", "libelle article", "nom article"],
        "Magasin": ["magasin", "depot", "dépôt", "emplacement", "site", "zone", "entrepot", "entrepôt", "whse", "loc", "localisation", "stock_loc"],
        "Flux": ["flux", "mouvement", "type flux", "sens", "quantite", "quantité", "volume", "qty", "qte", "qté", "solde", "nombre", "mvt", "quantités", "stk", "stock", "entree", "sortie", "unite", "unités", "qte_mvt", "nb"]
    }
    
    colonnes_trouvees = {}
    
    # Priorité 1 : Recherche par correspondance
    for cle, variantes in dictionnaire_colonnes.items():
        for col_reel in df_propre.columns:
            col_lower = str(col_reel).lower().strip()
            # Test exact ou inclusion
            if col_lower in variantes or any(v == col_lower for v in variantes) or any(v in col_lower for v in variantes):
                # Éviter de prendre une colonne de prix si ce n'est pas le seul choix
                if not any(p in col_lower for p in ["prix", "montant", "cout", "coût", "ttc", "ht", "valeur"]):
                    colonnes_trouvees[cle] = col_reel
                    break

    # Remplacement manuel si sélectionné dans l'interface
    if col_article_override and col_article_override in df_propre.columns:
        colonnes_trouvees["Article"] = col_article_override
    if col_flux_override and col_flux_override in df_propre.columns:
        colonnes_trouvees["Flux"] = col_flux_override

    # 3. Moteur de Détection des Alertes Logistiques
    alertes = {
        "negatifs": [],
        "vides": [],
        "excessifs": [],
        "faible_rotation": [],
        "sorties_sans_entree": []
    }
    
    col_art = colonnes_trouvees.get("Article")
    col_flux = colonnes_trouvees.get("Flux")
    col_des = colonnes_trouvees.get("Designation", col_art)
    
    if col_art and col_flux and col_art in df_propre.columns and col_flux in df_propre.columns:
        # Conversion forcée en numérique
        df_propre[col_flux] = pd.to_numeric(df_propre[col_flux], errors='coerce').fillna(0)
        
        # Analyse du comportement historique par article
        for art, sub_df in df_propre.groupby(col_art):
            if pd.isna(art) or str(art).strip() == "":
                continue
                
            designation_article = str(sub_df[col_des].iloc[0]) if col_des in sub_df.columns else "Non spécifiée"
            
            total_flux = sub_df[col_flux].sum()
            entrees = sub_df[sub_df[col_flux] > 0][col_flux].sum()
            sorties = abs(sub_df[sub_df[col_flux] < 0][col_flux].sum())
            nb_mouvements = len(sub_df)
            
            # A. Stock cumulé négatif
            if total_flux < 0:
                alertes["negatifs"].append({
                    "Code Article": str(art), 
                    "Désignation": designation_article, 
                    "Alerte": f"Niveau critique : {total_flux}"
                })
                
            # B. Stock devenu vide
            elif total_flux == 0 and nb_mouvements > 0:
                alertes["vides"].append({
                    "Code Article": str(art), 
                    "Désignation": designation_article, 
                    "Suivi": "Rupture complète : Stock actuellement épuisé (0)"
                })
                
            # C. Sorties sans entrée
            if sorties > 0 and entrees == 0:
                alertes["sorties_sans_entree"].append({
                    "Code Article": str(art), 
                    "Désignation": designation_article, 
                    "Anomalie": f"Déstockage anormal : {sorties} unités sorties sans aucune entrée"
                })
                
            # D. Sorties / Consommations excessives
            sorties_individuelles = sub_df[sub_df[col_flux] < 0][col_flux].abs()
            if not sorties_individuelles.empty:
                moyenne_sorties = sorties_individuelles.mean()
                max_sortie = sorties_individuelles.max()
                if max_sortie > (moyenne_sorties * 3) and len(sorties_individuelles) > 1:
                    alertes["excessifs"].append({
                        "Code Article": str(art), 
                        "Désignation": designation_article, 
                        "Vigilance": f"Surconsommation : Pic à -{max_sortie} (Moyenne : -{round(moyenne_sorties, 1)})"
                    })
            
            # E. Rotation très faible
            if nb_mouvements <= 1 and total_flux > 0:
                alertes["faible_rotation"].append({
                    "Code Article": str(art), 
                    "Désignation": designation_article, 
                    "Optimisation": f"Stock dormant : {total_flux} unités avec 1 seul mouvement"
                })

    # 4. Préparation des statistiques pour l'onglet Activités
    analyses_textes = []
    analyses_numeriques = []
    
    if col_art and col_art in df_propre.columns:
        repart = df_propre[col_art].value_counts().reset_index()
        repart.columns = ['Valeur', 'Nombre'] if len(repart.columns) == 2 else ['Valeur', 'Nombre'][:len(repart.columns)]
        
        if 'Valeur' in repart.columns and not repart.empty:
            repart['Valeur'] = repart['Valeur'].astype(str)
            analyses_textes.append({
                "colonne": f"Mouvements par {col_art}", 
                "donnees": repart.head(10).to_dict(orient='records')
            })
        
    if col_art and col_flux and col_art in df_propre.columns and col_flux in df_propre.columns:
        try:
            groupe = df_propre.groupby(col_art)[col_flux].sum().reset_index()
            groupe.columns = ['Valeur', 'Volume']
            groupe = groupe.sort_values(by='Volume', ascending=False)
            groupe['Valeur'] = groupe['Valeur'].astype(str)
            
            if 'Valeur' in groupe.columns and not groupe.empty:
                analyses_numeriques.append({
                    "colonne": f"Volumes cumulés finaux par {col_art} (Top 10)",
                    "donnees": groupe.head(10).to_dict(orient='records')
                })
        except:
            pass

    resultats_json = {
        "metriques": {
            "total_lignes": len(df_propre),
            "total_colonnes": len(df_propre.columns)
        },
        "analyses_textes": analyses_textes,
        "analyses_numeriques": analyses_numeriques,
        "colonnes_cibles": colonnes_trouvees,
        "alertes_logistiques": alertes
    }
    
    return resultats_json, df_propre
