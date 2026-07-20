import streamlit as st
import pandas as pd
import logique_stock as ls  # Importation de votre module de logique métier
import os
import plotly.express as px  # Pour les diagrammes circulaires et avancés

# 1. Configuration initiale de la page Streamlit
st.set_page_config(page_title="Dashboard Stocks CMG", layout="wide")

st.title(" Système d'Analyse Universel des Stocks ")
st.markdown("### *Outil d'aide à la décision par Chargement de Fichier (Anonymisé)*")
st.write("Ce site web utilise votre moteur algorithmique pour nettoyer, analyser et détecter les anomalies de n'importe quel export de mouvements de stock.")

st.markdown("---")

# 2. 📥 Zone de téléchargement du fichier Excel
fichier_uploade = st.file_uploader(
    label="Veuillez glisser-déposer ou sélectionner le fichier Excel des mouvements (.xlsx)", 
    type=["xlsx"]
)

if fichier_uploade is not None:
    nom_fichier_temporaire = "temp_analyse_stock.xlsx"
    
    try:
        # Sauvegarde temporaire du fichier chargé pour traitement
        with open(nom_fichier_temporaire, "wb") as f:
            f.write(fichier_uploade.getvalue())

        # Exécution de votre algorithme métier
        resultats_json, df_propre, anomalies_df = ls.analyser_fichier(nom_fichier_temporaire)
        
        # Extraction des indicateurs calculés
        kpis = resultats_json["indicateurs"]["kpis"]
        repart_flux = pd.DataFrame(resultats_json["indicateurs"]["repartition_flux"])
        categories = pd.DataFrame(resultats_json["indicateurs"]["categories_couteuses"])
        anomalies = resultats_json["anomalies"]

        # Affichage des KPIs Généraux
        st.success(" Analyse réussie ! Voici les indicateurs généraux du fichier importé :")
        col1, col2, col3 = st.columns(3)
        col1.metric(" Nombre total de mouvements", f"{kpis['nb_mouvements']:,}")
        col2.metric(" Articles différents mouvementés", f"{kpis['nb_articles']}")
        col3.metric(" Nombre de magasins actifs", f"{kpis['nb_magasins']}")

        st.markdown("---")

        # 3. Création des Onglets principaux du Dashboard
        onglet1, onglet2, onglet3 = st.tabs([
            " Activité & Volumes de Flux", 
            " Structure Analytique (Pareto/ABC)", 
            " Alertes & Anomalies Logistiques"
        ])

        # --- ONGLET 1 : ACTIVITÉ & VOLUMES DE FLUX ---
        with onglet1:
            st.subheader("Répartition des Mouvements par Type de Flux (Poids Relatif)")
            st.write("Cette section analyse la répartition des volumes de transactions (Entrées, Sorties, Ajustements) enregistrées dans le fichier pour comprendre la dynamique des flux.")
            
            if not repart_flux.empty:
                repart_flux["pourcentage"] = (repart_flux["nb"] / repart_flux["nb"].sum()) * 100
                
                g1_col1, g1_col2 = st.columns(2)
                
                with g1_col1:
                    st.markdown("**Diagramme en barres (Poids en %)**")
                    fig_bar_flux = px.bar(
                        repart_flux, x="type", y="pourcentage", text="pourcentage",
                        color="type",
                        labels={"type": "Type de Flux", "pourcentage": "Proportion (%)"}
                    )
                    fig_bar_flux.update_traces(texttemplate='%{text:.1f}%', textposition='outside')
                    st.plotly_chart(fig_bar_flux, use_container_width=True)
                    st.caption("Description : Ce graphique montre la part relative de chaque flux. Une dominance des sorties indique une forte consommation, tandis qu'un taux élevé d'ajustements peut révéler des écarts d'inventaire.")
                
                with g1_col2:
                    st.markdown("**Diagramme circulaire (Répartition globale)**")
                    fig_pie_flux = px.pie(
                        repart_flux, names="type", values="nb",
                        labels={"type": "Type de Flux", "nb": "Nombre de mouvements"},
                        hole=0.1
                    )
                    fig_pie_flux.update_traces(textinfo='percent+value')
                    st.plotly_chart(fig_pie_flux, use_container_width=True)
                    st.caption("Description : Visualisation en secteurs du volume brut des transactions par flux permettant d'isoler instantanément la catégorie dominante.")
            
            if kpis['periode_debut'] and kpis['periode_fin']:
                st.info(f"📅 Période couverte par ce fichier : du {kpis['periode_debut']} au {kpis['periode_fin']}")

        # --- ONGLET 2 : STRUCTURE ANALYTIQUE (PARETO / ABC / TEMPOREL) ---
        with onglet2:
            st.subheader("🔮 Évolution Mensuelle des Mouvements par Catégorie")
            st.write("Cette vue dynamique décompose les flux financiers bruts mensuels par famille d'articles, mettant en valeur les tendances logistiques récurrentes.")
            
            if "evolution_categories" in resultats_json["indicateurs"]:
                df_evo = pd.DataFrame(resultats_json["indicateurs"]["evolution_categories"])
                
                if not df_evo.empty:
                    # Graphique dynamique : Barres groupées par mois
                    fig_evo = px.bar(
                        df_evo, 
                        x="Mois_Affichage", 
                        y="AbsMontant", 
                        color="Catégorie",
                        barmode="group",
                        labels={"Mois_Affichage": "Mois d'analyse", "AbsMontant": "Montant Cumulé (Brut)"},
                        template="plotly_dark"
                    )
                    fig_evo.update_layout(xaxis_tickangle=0, legend_title_text="Catégories")
                    st.plotly_chart(fig_evo, use_container_width=True)
                    st.caption("Description : L'axe horizontal liste la chronologie des mois. Pour chaque mois, la comparaison des barres permet d'identifier l'évolution de l'empreinte financière des catégories.")
                else:
                    st.info("Aucune donnée temporelle valide trouvée pour décomposer les catégories.")
            
            st.markdown("---")
            
            # --- TABLEAU DE SYNTHÈSE ANALYTIQUE (STYLE EXCEL) ---
            st.markdown("### 📋 Tableau de Synthèse Analytique (Classement Général)")
            if not categories.empty:
                col_cat = "Catégorie" if "Catégorie" in categories.columns else "Categorie"
                df_tableau = categories.head(10).copy()[[col_cat, "valeur"]]
                df_tableau["valeur"] = df_tableau["valeur"].apply(lambda x: f"{x:,.0f}")
                
                total_general = categories["valeur"].sum()
                ligne_total = pd.DataFrame([{col_cat: "Total général", "valeur": f"{total_general:,.0f}"}])
                df_tableau = pd.concat([df_tableau, ligne_total], ignore_index=True)
                
                st.table(df_tableau.set_index(col_cat))
            else:
                st.info("Aucune donnée de montant trouvée pour classer les catégories.")

        # --- ONGLET 3 : ALERTES & ANOMALIES LOGISTIQUES ---
        with onglet3:
            st.subheader("Détection Automatique des Dysfonctionnements de Stock")
            st.write("Règles métier appliquées pour identifier les risques d'exploitation.")
            
            correspondance_noms = {
                "Stock cumulé négatif": "sorties",
                "stock_negatif": "sorties",
                "stock_cumule_negatif": "sorties",
                
                "Stock devenu vide (rupture)": "stock non mouvementé",
                "stock_vide": "stock non mouvementé",
                "stock_devenu_vide": "stock non mouvementé",
                "rupture": "stock non mouvementé",
                
                "Sortie > stock disponible": "SSF",
                "sortie_superieure": "SSF",
                "sortie_stock_disponible": "SSF",
                
                "Consommation excessive": "articles mouvementées",
                "consommation_excessive": "articles mouvementées",
                
                "Rotation très faible": "mouvt faible",
                "rotation_faible": "mouvt faible",
                "rotation_tres_faible": "mouvt faible"
            }
            
            aucun_incident = True
            for cle_anomalie, data in anomalies.items():
                if data["nb"] > 0:
                    aucun_incident = False
                    
                    libelle_original = data.get('libelle', '')
                    
                    if cle_anomalie in correspondance_noms:
                        nouveau_nom = correspondance_noms[cle_anomalie]
                    elif libelle_original in correspondance_noms:
                        nouveau_nom = correspondance_noms[libelle_original]
                    else:
                        nouveau_nom = libelle_original if libelle_original else cle_anomalie
                    
                    with st.expander(f" {nouveau_nom} : {data['nb']} incident(s) détecté(s)"):
                        st.error(f"**Description de la règle :** {data['description']}")
                        st.write("Échantillon des lignes affectées (anonymisé) :")
                        
                        df_ano_visu = pd.DataFrame(data["lignes"])
                        if not df_ano_visu.empty:
                            colonnes_visu = [c for c in ["Date", "Article", "Désignation", "Designation", "Magasin", "Flux"] if c in df_ano_visu.columns]
                            st.dataframe(df_ano_visu[colonnes_visu])
            
            if aucun_incident:
                st.success(" Félicitations ! Aucune anomalie logistique n'a été détectée dans ce fichier.")

    except Exception as e:
        st.error(f" Une erreur s'est produite lors de l'analyse de ce fichier : {e}")

else:
    st.info(" En attente d'un fichier. Veuillez glisser-déposer un fichier Excel ci-dessus pour lancer l'analyse.")
