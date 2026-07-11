import streamlit as st
import pandas as pd
import logique_stock as ls

# Configuration globale de la page web
st.set_page_config(page_title="Suivi de Stock Analytique", layout="wide")

st.title("🚀 Système de Diagnostic Logistique Universel")
st.write("Déposez votre fichier Excel de mouvements de stock pour générer automatiquement le rapport.")
st.markdown("---")

# Zone de téléchargement du fichier Excel
fichier_uploade = st.file_uploader(
    label="Glissez-déposez votre tableau Excel ici (.xlsx)", 
    type=["xlsx"]
)

if fichier_uploade is not None:
    try:
        # Déclenchement du moteur d'analyse en transmettant directement le flux mémoire
        resultats, df_propre = ls.analyser_fichier(fichier_uploade)
        
        st.success("✅ Fichier traité avec succès !")
        
        # 1. Métriques Générales de ton tableau
        col1, col2 = st.columns(2)
        col1.metric("📋 Lignes de flux détectées", f"{resultats['metriques']['total_lignes']:,}")
        col2.metric("🔢 Colonnes analysées", f"{resultats['metriques']['total_colonnes']}")

        st.markdown("---")

        # 2. Les 3 Onglets Structurés
        onglet1, onglet2, onglet3 = st.tabs([
            "📈 Activités", 
            "📋 Structure", 
            "⚠️ Alerte"
        ])

        # ================= 1. ONGLET ACTIVITÉS =================
        with onglet1:
            st.subheader("Analyse des Activités & Volumes Graphiques")
            
            # Affichage des graphiques pour les colonnes de texte (répartitions)
            if resultats.get("analyses_textes"):
                for index, analyse in enumerate(resultats["analyses_textes"]):
                    df_graph = pd.DataFrame(analyse["donnees"])
                    if not df_graph.empty and "Valeur" in df_graph.columns:
                        st.write(f"**{analyse['colonne']}**")
                        st.bar_chart(data=df_graph, x="Valeur", y="Nombre", color="#4AF0A1")
                    else:
                        st.warning(f"📊 Données insuffisantes pour afficher : {analyse['colonne']}")
            
            # Affichage des graphiques pour les colonnes numériques
            if resultats.get("analyses_numeriques"):
                st.markdown("---")
                st.subheader("Synthèse Quantitative des Flux")
                for index, analyse in enumerate(resultats["analyses_numeriques"]):
                    df_num = pd.DataFrame(analyse["donnees"])
                    if not df_num.empty and "Valeur" in df_num.columns:
                        st.write(f"**{analyse['colonne']}**")
                        st.bar_chart(data=df_num, x="Valeur", y="Volume", color="#FF4B4B")
                    else:
                        st.warning(f"📊 Données quantitatives insuffisantes pour afficher : {analyse['colonne']}")
            
            if not resultats.get("analyses_textes") and not resultats.get("analyses_numeriques"):
                st.info("💡 Aucune donnée d'activité graphique n'a pu être construite. Modifiez l'onglet Structure pour ajuster le mappage.")

        # ================= 2. ONGLET STRUCTURE =================
        with onglet2:
            st.subheader("Structure Analytique Filtrée")
            
            # BLOC DE DIAGNOSTIC POUR COMPRENDRE CE QUE LE SCRIPT A RECONNU
            st.info("🔍 **Diagnostic de détection automatique des colonnes :**")
            cibles = resultats.get("colonnes_cibles", {})
            
            # Présentation sous forme de petites métriques de succès
            diag_cols = st.columns(5)
            types_attendus = ["Date", "Article", "Designation", "Magasin", "Flux"]
            for idx, t in enumerate(types_attendus):
                nom_trouve = cibles.get(t)
                if nom_trouve:
                    diag_cols[idx].success(f"**{t}** trouvé :\n`{nom_trouve}`")
                else:
                    diag_cols[idx].error(f"**{t}**\n❌ Non détecté")
            
            st.markdown("---")
            st.write("#### Aperçu des données opérationnelles identifiées :")
            
            # Extraction des colonnes correspondantes trouvées par la logique
            colonnes_a_afficher = [nom_reel for nom_cle, nom_reel in cibles.items() if nom_cle in types_attendus]
            
            # Filtre manuel de secours anti-prix
            mots_prix = ["prix", "montant", "valeur", "unitaire", "cout", "coût", "mnt", "ttc", "ht", "price", "amount"]
            colonnes_a_afficher = [c for c in colonnes_a_afficher if not any(mot in str(c).lower() for mot in mots_prix)]
            
            if colonnes_a_afficher:
                st.dataframe(df_propre[colonnes_a_afficher].head(100), use_container_width=True)
            else:
                st.warning("⚠️ Les colonnes logistiques clés (Article / Flux) n'ont pas pu être toutes identifiées automatiquement.")
                st.write("Affichage global sécurisé du tableau complet (sans prix) :")
                cols_sans_prix = [c for c in df_propre.columns if not any(mot in str(c).lower() for mot in mots_prix)]
                st.dataframe(df_propre[cols_sans_prix].head(100), use_container_width=True)

        # ================= 3. ONGLET ALERTE =================
        with onglet3:
            st.subheader("Centre d'Alerte et de Diagnostic des Risques Logistiques")
            
            al = resultats.get("alertes_logistiques", {
                "negatifs": [], "vides": [], "excessifs": [], "faible_rotation": [], "sorties_sans_entree": []
            })
            
            # Compteur global d'anomalies
            total_anomalies = len(al["negatifs"]) + len(al["vides"]) + len(al["sorties_sans_entree"]) + len(al["excessifs"]) + len(al["faible_rotation"])
            
            if not cibles.get("Article") or not cibles.get("Flux"):
                st.error("🛑 Impossible de calculer les alertes logistiques car les colonnes fondamentales '**Article**' et '**Flux**' n'ont pas été détectées dans votre fichier Excel.")
            elif total_anomalies == 0:
                st.success("✅ Aucun comportement critique ou anomalie de stock détectés.")
            else:
                st.warning(f"⚠️ {total_anomalies} anomalies ou points de vigilance détectés sur vos flux.")
                
                # 1. Stocks Cumulés Négatifs
                if al["negatifs"]:
                    with st.expander("🛑 ALERTES : Stock cumulé négatif", expanded=True):
                        st.dataframe(pd.DataFrame(al["negatifs"]), use_container_width=True)
                        
                # 2. Stocks devenus Vides
                if al["vides"]:
                    with st.expander("📭 SUIVI : Stock devenu vide", expanded=False):
                        st.dataframe(pd.DataFrame(al["vides"]), use_container_width=True)
                        
                # 3. Sorties sans Entrées
                if al["sorties_sans_entree"]:
                    with st.expander("🚨 ANOMALIE : Sorties sans entrée", expanded=True):
                        st.dataframe(pd.DataFrame(al["sorties_sans_entree"]), use_container_width=True)

                # 4. Consommations Excessives
                if al["excessifs"]:
                    with st.expander("📈 VIGILANCE : Sorties consommations excessives", expanded=False):
                        st.dataframe(pd.DataFrame(al["excessifs"]), use_container_width=True)
                        
                # 5. Rotation très Faible
                if al["faible_rotation"]:
                    with st.expander("⏳ OPTIMISATION : Rotation très faible", expanded=False):
                        st.dataframe(pd.DataFrame(al["faible_rotation"]), use_container_width=True)

    except Exception as e:
        st.error(f"❌ Impossible d'analyser ce fichier automatiquement : {e}")
else:
    st.info("👋 Module en veille. Veuillez importer un tableau Excel pour initialiser les onglets.")
   
