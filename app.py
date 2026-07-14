import streamlit as st
import pandas as pd
import logique_stock as ls  # Importation de ton module logique
import os

# Configuration du site web
st.set_page_config(page_title="Dashboard Stocks CMG", layout="wide")

st.title("📊 Systeme d'Analyse Universel des Stocks - CMG")
st.markdown("### *Outil d'aide a la decision par Chargement de Fichier (Anonymise)*")
st.write("Ce site web utilise ton moteur algorithmique pour nettoyer, analyser et detecter les anomalies de n'importe quel export de mouvements de stock.")

st.markdown("---")

# 📥 Zone de téléchargement du fichier Excel
fichier_uploade = st.file_uploader(
    label="Veuillez glisser-deposer ou selectionner le fichier Excel des mouvements (.xlsx)", 
    type=["xlsx"]
)

if fichier_uploade is not None:
    # Fichier tampon dans le dossier courant pour éviter les conflits de droits Windows
    nom_fichier_temporaire = "temp_analyse_stock.xlsx"
    
    try:
        # Écriture du fichier sur le disque (écrase l'ancien s'il existe)
        with open(nom_fichier_temporaire, "wb") as f:
            f.write(fichier_uploade.getvalue())

        # Appel de ton module personnalisé
        resultats_json, df_propre, anomalies_df = ls.analyser_fichier(nom_fichier_temporaire)
        
        # --- NOTE : On ne supprime pas le fichier ici pour éviter le WinError 32 de Windows ---
        
        # Récupération des indicateurs calculés dynamiquement
        kpis = resultats_json["indicateurs"]["kpis"]
        repart_flux = pd.DataFrame(resultats_json["indicateurs"]["repartition_flux"])
        categories = pd.DataFrame(resultats_json["indicateurs"]["categories_couteuses"])
        anomalies = resultats_json["anomalies"]

        # 1. Affichage des KPIs Généraux du fichier chargé
        st.success("✅ Analyse reussie ! Voici les indicateurs generaux du fichier importé :")
        col1, col2, col3 = st.columns(3)
        col1.metric("📦 Nombre total de mouvements", f"{kpis['nb_mouvements']:,}")
        col2.metric("🔢 Articles differents mouvementes", f"{kpis['nb_articles']}")
        col3.metric("🏠 Nombre de magasins actifs", f"{kpis['nb_magasins']}")

        st.markdown("---")

        # 2. Onglets de Navigation
        onglet1, onglet2, onglet3 = st.tabs([
            "📈 Activite & Volumes de Flux", 
            "🔮 Structure Analytique (Pareto/ABC)", 
            "⚠️ Alertes & Anomalies Logistiques"
        ])

        with onglet1:
            st.subheader("Repartition des Mouvements par Type de Flux (Poids Relatif)")
            st.write("Analyse des volumes de transactions (Entrees, Sorties, Ajustements) enregistres dans ce fichier.")
            
            if not repart_flux.empty:
                repart_flux["pourcentage"] = (repart_flux["nb"] / repart_flux["nb"].sum()) * 100
                st.bar_chart(data=repart_flux, x="type", y="pourcentage", color="#4AP0A1")
            
            if kpis['periode_debut'] and kpis['periode_fin']:
                st.info(f"📅 Periode couverte par ce fichier : du {kpis['periode_debut']} au {kpis['periode_fin']}")

        with onglet2:
            st.subheader("Classement Proportionnel des Categories (Consommations / Immobilisations)")
            st.write("Structure des categories les plus sollicitees financierement dans le fichier importé (en %).")
            
            if not categories.empty:
                total_val_cat = categories["valeur"].sum()
                categories["Part_Relative_Pourcent"] = (categories["valeur"] / total_val_cat) * 100
                
                col_cat = "Catégorie" if "Catégorie" in categories.columns else "Categorie"
                st.bar_chart(data=categories.head(5), x=col_cat, y="Part_Relative_Pourcent", color="#FF4B4B")
            else:
                st.info("Aucune donnee de montant trouvee pour classer les categories.")

        with onglet3:
            st.subheader("Detection Automatique des Dysfonctionnements de Stock")
            st.write("Regles metier appliquees pour identifier les risques d'exploitation.")
            
            # Affichage dynamique des alertes selon le contenu du fichier
            aucun_incident = True
            for cle_anomalie, data in anomalies.items():
                if data["nb"] > 0:
                    aucun_incident = False
                    with st.expander(f"🔴 {data['libelle']} : {data['nb']} incident(s) detecte(s)"):
                        st.error(f"**Description de la regle :** {data['description']}")
                        st.write("Echantillon des lignes affectees (anonymise) :")
                        
                        df_ano_visu = pd.DataFrame(data["lignes"])
                        if not df_ano_visu.empty:
                            colonnes_visu = [c for c in ["Date", "Article", "Désignation", "Designation", "Magasin", "Flux"] if c in df_ano_visu.columns]
                            st.dataframe(df_ano_visu[colonnes_visu])
            
            if aucun_incident:
                st.success("🎉 Félicitations ! Aucune anomalie logistique n'a été détectée dans ce fichier.")

    except Exception as e:
        st.error(f"❌ Une erreur s'est produite lors de l'analyse de ce fichier : {e}")

else:
    st.info("👋 En attente d'un fichier. Veuillez glisser-deposer un fichier Excel ci-dessus pour lancer l'analyse.")
