import streamlit as st
import pandas as pd
#import seaborn as sns
import matplotlib.pyplot as plt
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from docx import Document
from docx.shared import Inches

# Initialisation
st.set_page_config(page_title="Estimation charge de travail", layout="wide")
st.title("🕒 Estimation de la charge de travail")

# Mode d'analyse
mode = st.sidebar.radio("Mode d'analyse", ["Responsables Projet", "Agents Techniques"])

# === MODULE POUR LES RESPONSABLES PROJET ===
if mode == "Responsables Projet":
    mode_rp = st.sidebar.radio("Mode RP", ["Staff unique"])

    # Info sur les seuils de fréquence mensuelle
    with st.expander("ℹ️ Aide - Seuils typiques de fréquence mensuelle"):
        st.markdown("""
        - **1.0** = tâche mensuelle (1x/mois)
        - **0.25** = tâche trimestrielle (1x/3 mois)
        - **0.5** = tâche bimestrielle (1x/2 mois)
        - **0.166** = tâche semestrielle (1x/6 mois)
        - **0.083** = tâche annuelle (1x/12 mois)
        - **2.0** = tâche bimensuelle (2x/mois)
        - **4.0** = tâche hebdomadaire (1x/semaine)
        - **20.0** = tâche quotidienne (5j/semaine en moyenne)
        """)

    # Paramètre global : heures totales de travail/mois
    heures_max_mensuelles = st.sidebar.number_input("Heures de travail max/personne/mois", min_value=1, value=160)
    temps_adm = st.sidebar.number_input("Temps administratif mensuel (en heures)", min_value=0.0, value=4.0, step=0.5)
    temps_forma = st.sidebar.number_input("Temps de formation continue / renforcement (en heures)", min_value=0.0, value=2.0, step=0.5)
    marge_imprevue = st.sidebar.slider("% de marge pour imprévus", min_value=0, max_value=30, value=5)

    if mode_rp == "Équipe":
        pass
    else:
        st.sidebar.header("🔍 Staff unique - Paramètres")
        staff_file = st.sidebar.file_uploader("📅 Charger un fichier avec les projets par RP (matricule, nb_projets) : ", type=["xlsx"])

        charges_par_agent = {}  # Initialize charges_par_agent here to avoid scope issues
        
        if staff_file:
            agent_data = pd.read_excel(staff_file)

            rp_data = pd.DataFrame({
                "Tâche": [
                    "Proposition des plan de remédiation ou de mitigation",
                    "Suivi du processus de validation des plans de remédiation",
                    "Vérification et validation des planification des BR",
                    "Organisation des missions terrain de vérification de la qualité d'intervention",
                    "Suivi des actions de remédiation auprès des bénéficiaires",
                    "Elaboration de budget",
                    "Suivi budgétaire",
                    "Participer aux réunion trimestrielle de coordination ICI Geneve",
                    "Organiser et participer à la reunion mensuelle avec le partenaire",
                    "Organiser et participer aux missions conjointes de suivis des activités sur le terrain",
                    "Participer au développement d'outils et d'approche",
                    "Elaborer des TDRs pour ateliers de formation",
                    "Organiser l'atelier annuelle de renforcement",
                    "Analyse des programmes de suivi des coachs",
                    "Rapport mensuel de suivi des Coachs",
                    "Travail avec les finances",
                    "Travail avec la logistique",
                    "Travail avec la communication",
                    "Travail avec l'équipe travail forcé",
                    "Travail avec l'équipe formation",
                    "Conception des outils d'enquête",
                    "Présentation des résultats d'enquête",
                    "Rapport narratif d'avancement",
                    "Liste nominative des enfants travailleurs remédiés",
                    "Rapport situationnel mensuel",
                    "Compte rendu réunions avec le partenaire",
                    "Templates du partenaire",
                    "Point de suivi SSRTE",
                    "KPI mensuel"
                ],
                "Durée (heures)": [1.5, 1, 1, 2, 3, 1.25, 1, 0.5, 0.1, 0.5, 0.1, 0, 0.5, 0.5, 0.5, 0.5, 0.05, 0.05, 0.05, 0.05, 0.05, 0.1, 2, 2, 1, 0.1, 0.1, 0.1, 0.5],
                "Fréquence mensuelle": [1, 4, 1, 1, 4, 1, 4, 0.33, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 4, 1]
            })

            st.subheader("📝 Liste des Tâches (RP)")
            with st.expander("ℹ️ Section - Valeurs des paramètres de calcul"):
                edited_df = st.data_editor(rp_data, num_rows="dynamic", use_container_width=True)        

            results = []

            for _, row in agent_data.iterrows():
                agent_id = row["Matricule"]
                nb_projets = row["Nombre de projets"]
                df = edited_df.copy()
                df["Temps par projet (h)"] = df["Durée (heures)"] * df["Fréquence mensuelle"]
                df["Temps total (tous projets)"] = df["Temps par projet (h)"] * nb_projets
                
                total_heur = df["Temps total (tous projets)"].sum()
                total_heur += temps_adm + temps_forma
                total_heures = total_heur * (1 + marge_imprevue / 100)
                charge_pct = (total_heures / heures_max_mensuelles) * 100

                charges_par_agent[agent_id] = df.copy()

                if charge_pct <= 40:
                    statut = "✅ Charge faible. Il reste de la marge pour ajouter d'autres responsabilités."
                    nb_requi = 1
                elif charge_pct <= 85:
                    statut = "➡️ Charge modérée. Suivi recommandé si d'autres tâches sont attendues."
                    nb_requi = 1
                elif charge_pct <= 100:
                    statut = "🔶 Charge élevée. Un ajustement organisationnel peut être envisagé."
                    nb_requi = 1
                else:
                    statut = "⚠️ Surcharge détectée. Il est conseillé d'envisager le recrutement de responsable(s) projet(s) supplémentaire(s)."
                    nb_requi = int(total_heures / heures_max_mensuelles + 0.99)

                results.append({
                    "Matricule": agent_id,
                    "Nombre total d'heures de travail/mois": heures_max_mensuelles,
                    "Nombre de projets": nb_projets,
                    "Heures totales": total_heures,
                    "% de charge": round(charge_pct, 1),
                    "Statut": statut,
                    "Nombre de RP requis": nb_requi
                })

            result_df = pd.DataFrame(results)
            st.subheader("📋 Résultats par agent")
            with st.expander("ℹ️ Section - Résultats par agent"):
                st.dataframe(result_df)

                output = BytesIO()
                result_df.to_excel(output, index=False, engine="openpyxl")
                st.download_button(
                    label="📀 Télécharger les résultats (Excel)",
                    data=output.getvalue(),
                    file_name="charges_par_agent.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.subheader("📊 Temps total par tâche")
            with st.expander("ℹ️ Section - Temps total par tâche"):
                # Only display chart options if we have data
                if charges_par_agent:
                    agent_ids = ["Tous"] + list(charges_par_agent.keys())
                    selected_agent = st.selectbox("Sélectionner un agent pour voir le graphique des tâches :", agent_ids)
    
                    if selected_agent == "Tous":
                        df_concat = pd.concat(charges_par_agent.values(), ignore_index=True)
                        total_per_task = df_concat.groupby("Tâche")["Temps total (tous projets)"].sum().reset_index()
                        total_per_task = total_per_task.sort_values("Temps total (tous projets)")
                        title = "Histogramme du temps total par tâche - Tous les agents"
                    else:
                        selected_df = charges_par_agent[selected_agent]
                        total_per_task = selected_df.groupby("Tâche")["Temps total (tous projets)"].sum().reset_index()
                        total_per_task = total_per_task.sort_values("Temps total (tous projets)")
                        title = f"Histogramme du temps total par tâche - Agent {selected_agent}"
    
                    # Préparation des données avec les valeurs formatées directement
                    total_per_task["text_values"] = total_per_task["Temps total (tous projets)"].apply(lambda x: f"{x:.1f}")
                    
                    # Inside the chart generation section
                    # Création du graphique interactif avec Plotly
                    fig = px.bar(
                        total_per_task,
                        x="Temps total (tous projets)",
                        y="Tâche",
                        orientation='h',
                        title=title,
                        labels={"Temps total (tous projets)": "Heures totales", "Tâche": ""},
                        text="text_values"  # Utiliser la colonne formatée
                    )

                    # Personnalisation du graphique avec étiquettes à l'intérieur
                    fig.update_traces(
                        textposition='inside',       # Position des étiquettes à l'intérieur des barres
                        textfont=dict(
                            size=14,
                            color='white'            # Couleur blanche pour contraster avec les barres
                        )
                    )

                    # Personnalisation du graphique
                    fig.update_layout(
                        height=800,
                        width=900,
                        yaxis={'categoryorder':'total ascending'},
                        margin=dict(l=10, r=10, t=50, b=10),
                        hoverlabel=dict(bgcolor="white", font_size=14),
                        title={
                            'text': title,
                            'x': 0.5,        # Centre le titre horizontalement
                            'xanchor': 'center',  # Point d'ancrage du titre
                            'yanchor': 'top'      # Position verticale du titre
                        }
                    )
                    
                    # Affichage du graphique
                    st.plotly_chart(fig, use_container_width=True)

# === MODULE POUR LES AGENTS TECHNIQUES ===

elif mode == "Agents Techniques":
    mode_at = st.sidebar.radio("Mode AT", ["Staff unique"])

    # Info sur les seuils de fréquence mensuelle
    with st.expander("ℹ️ Aide - Seuils typiques de fréquence mensuelle"):
        st.markdown("""
        - **1.0** = tâche mensuelle (1x/mois)
        - **0.25** = tâche trimestrielle (1x/3 mois)
        - **0.5** = tâche bimestrielle (1x/2 mois)
        - **0.166** = tâche semestrielle (1x/6 mois)
        - **0.083** = tâche annuelle (1x/12 mois)
        - **2.0** = tâche bimensuelle (2x/mois)
        - **4.0** = tâche hebdomadaire (1x/semaine)
        - **20.0** = tâche quotidienne (5j/semaine en moyenne)
        """)

    # Paramètre global : temps de déplacement moyen & heures totales de travail/mois
    heures_max_mensuelles = st.sidebar.number_input("Heures de travail max/personne/mois", min_value=1, value=160)
    temps_deplacement = st.sidebar.number_input("Temps moyen de déplacement (aller-retour, en heures)", min_value=0.0, value=1.0, step=0.1)
    temps_admin = st.sidebar.number_input("Temps administratif mensuel (en heures)", min_value=0.0, value=4.0, step=0.5)
    temps_formation = st.sidebar.number_input("Temps de formation continue / renforcement (en heures)", min_value=0.0, value=2.0, step=0.5)
    marge_imprevus = st.sidebar.slider("% de marge pour imprévus", min_value=0, max_value=30, value=5)

    if mode_at == "Équipe":
        pass

    else:
        st.sidebar.header("🔍 Staff unique - Paramètres")
        staff_file = st.sidebar.file_uploader("📥 Charger un fichier avec les projets par agent (matricule, nb_projets, coopératives, structures, agents op)", type=["xlsx"])
        uploaded_file = st.sidebar.file_uploader("📥 Charger le fichier des tâches associées", type=["xlsx"])
        # heures_max_mensuelles = st.sidebar.number_input("Heures de travail max/mois pour chaque AT", min_value=1, value=160)

        charges_par_at = {}  # Initialize charges_par_at here to avoid scope issues

        if staff_file and uploaded_file:
            agent_data = pd.read_excel(staff_file)
            task_df = pd.read_excel(uploaded_file)

            results = []

            for _, row in agent_data.iterrows():
                agent_id = row["Matricule"]
                nb_projet = row["Nombre de projets"]
                nb_coop = row["Nombre de coopératives"]
                nb_struc = row["Nombre de structures"]
                nb_agents_ope = row["Nombre d'agents opérationnels"]

                df = task_df.copy()
                multiplicateurs = {
                    # "projet": nb_projets,
                    "coopérative": nb_coop,
                    "structure": nb_struc,
                    "agent_op": nb_agents_ope,
                    "unique": 1
                }
                df["Temps total (heures)"] = (
                    (df["Durée (heures)"] + temps_deplacement) *
                    df["Fréquence mensuelle"] *
                    df["Facteur"].map(multiplicateurs)
                )

                base_heures = df["Temps total (heures)"].sum()
                base_heures += temps_admin + temps_formation
                total_heure = base_heures * (1 + marge_imprevus / 100)
                charge_pct = (total_heure / heures_max_mensuelles) * 100

                charges_par_at[agent_id] = df.copy()

                if charge_pct <= 40:
                    statut = "✅ Charge faible. Il reste de la marge pour ajouter d'autres missions."
                    nb_requis = 1
                elif charge_pct <= 85:
                    statut = "➡️ Charge modérée. Un suivi peut être utile."
                    nb_requis = 1
                elif charge_pct <= 100:
                    statut = "🔶 Charge élevée. Réévaluation possible."
                    nb_requis = 1
                else:
                    statut = "⚠️ Surcharge détectée. Recrutement conseillé : AT supplémentaire(s)."
                    nb_requis = int(total_heure / heures_max_mensuelles + 0.99)

                results.append({
                    "Matricule": agent_id,
                    "Nombre total d'heures de travail/mois": heures_max_mensuelles,
                    "Nombre de projets": nb_projet,
                    "Nombre de coopératives": nb_coop,
                    "Nombre de structures": nb_struc,
                    "Nombre d'agents opérationnels": nb_agents_ope,
                    "Heures totales": total_heure,
                    "% de charge": round(charge_pct, 1),
                    "Statut": statut,
                    "Nombre d'AT requis": nb_requis
                })

            result_df = pd.DataFrame(results)
            st.subheader("📋 Résultats par agent technique")
            with st.expander("ℹ️ Section - Résultats par agent technique"):
                st.dataframe(result_df)
    
                output = BytesIO()
                result_df.to_excel(output, index=False, engine="openpyxl")
                st.download_button(
                    label="💾 Télécharger les résultats (Excel)",
                    data=output.getvalue(),
                    file_name="Charges_par_AT.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            st.subheader("📊 Temps total par tâche")
            with st.expander("ℹ️ Section - Temps total par tâche"):
                # Only display chart options if we have data
                if charges_par_at:
                    agent_ids = ["Tous"] + list(charges_par_at.keys())
                    selected_at = st.selectbox("Sélectionner un agent pour voir le graphique des tâches :", agent_ids)
    
                    # Dans la section de génération du graphique, remplacez le code Matplotlib par:
                    if selected_at == "Tous":
                        df_concat = pd.concat(charges_par_at.values(), ignore_index=True)
                        total_per_task = df_concat.groupby("Tâche")["Temps total (heures)"].sum().reset_index()
                        total_per_task = total_per_task.sort_values("Temps total (heures)")
                        title = "Histogramme du temps total par tâche - Tous les agents"
                    else:
                        selected_df = charges_par_at[selected_at]
                        total_per_task = selected_df.groupby("Tâche")["Temps total (heures)"].sum().reset_index()
                        total_per_task = total_per_task.sort_values("Temps total (heures)")
                        title = f"Histogramme du temps total par tâche - Agent {selected_at}"
                    
                    # Préparation des données avec les valeurs formatées directement
                    total_per_task["text_values"] = total_per_task["Temps total (heures)"].apply(lambda x: f"{x:.1f}")

                    # Création du graphique interactif avec Plotly
                    fig = px.bar(
                        total_per_task,
                        x="Temps total (heures)",
                        y="Tâche",
                        orientation='h',
                        title=title,
                        labels={"Temps total (heures)": "Heures totales", "Tâche": ""},
                        text="text_values"  # Utiliser la colonne formatée
                    )

                    # Personnalisation du graphique avec étiquettes à l'intérieur
                    fig.update_traces(
                        textposition='inside',       # Position des étiquettes à l'intérieur des barres
                        textfont=dict(
                            size=14,
                            color='white'            # Couleur blanche pour contraster avec les barres
                        )
                    )

                    # Personnalisation du graphique
                    fig.update_layout(
                        height=800,
                        width=900,
                        yaxis={'categoryorder':'total ascending'},
                        margin=dict(l=10, r=10, t=50, b=10),
                        hoverlabel=dict(bgcolor="white", font_size=14),
                        title={
                            'text': title,
                            'x': 0.5,        # Centre le titre horizontalement
                            'xanchor': 'center',  # Point d'ancrage du titre
                            'yanchor': 'top'      # Position verticale du titre
                        }
                    )
                    
                    # Affichage du graphique
                    st.plotly_chart(fig, use_container_width=True)
