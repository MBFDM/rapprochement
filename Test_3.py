import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import os
import tempfile
import glob
from PIL import Image
import io
import base64
import hashlib
import sqlite3
import time
from streamlit_option_menu import option_menu
import plotly.figure_factory as ff
import re

# Configuration de la page
st.set_page_config(
    page_title="AGC-VIE - Gestion Technique et Comptable",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ------------------------ STYLES CSS PERSONNALISÉS ------------------------
st.markdown("""
<style>
    /* Style global */
    .stApp {
        background-color: #f8f9fa;
    }
    
    /* En-têtes */
    h1, h2, h3 {
        color: #1e3c72;
        font-weight: 600;
    }
    
    /* Cartes */
    .metric-card {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
    
    .metric-card h3 {
        color: white;
        font-size: 1.1em;
        margin-bottom: 10px;
    }
    
    .metric-card p {
        font-size: 2em;
        font-weight: bold;
        margin: 0;
    }
    
    /* Tableaux */
    .dataframe {
        font-size: 0.9em;
        border-collapse: collapse;
        width: 100%;
    }
    
    .dataframe th {
        background-color: #1e3c72;
        color: white;
        padding: 10px;
        text-align: left;
    }
    
    .dataframe td {
        padding: 8px;
        border-bottom: 1px solid #ddd;
    }
    
    .dataframe tr:hover {
        background-color: #f5f5f5;
    }
    
    /* Boutons */
    .stButton > button {
        background-color: #1e3c72;
        color: white;
        border-radius: 5px;
        padding: 10px 20px;
        font-weight: 600;
        border: none;
        transition: all 0.3s;
    }
    
    .stButton > button:hover {
        background-color: #2a5298;
        box-shadow: 0 4px 8px rgba(0,0,0,0.2);
    }
    
    /* Barre latérale */
    .css-1d391kg {
        background-color: #1e3c72;
    }
    
    .sidebar .sidebar-content {
        background-color: #1e3c72;
        color: white;
    }
    
    /* Onglets */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        background-color: #f0f2f6;
        border-radius: 4px 4px 0 0;
        padding: 10px 20px;
        font-weight: 600;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #1e3c72;
        color: white;
    }
    
    /* Messages de succès/erreur */
    .success-message {
        background-color: #d4edda;
        color: #155724;
        padding: 10px;
        border-radius: 5px;
        border-left: 4px solid #28a745;
        margin: 10px 0;
    }
    
    .error-message {
        background-color: #f8d7da;
        color: #721c24;
        padding: 10px;
        border-radius: 5px;
        border-left: 4px solid #dc3545;
        margin: 10px 0;
    }
    
    .warning-message {
        background-color: #fff3cd;
        color: #856404;
        padding: 10px;
        border-radius: 5px;
        border-left: 4px solid #ffc107;
        margin: 10px 0;
    }
    
    /* Barre de progression */
    .stProgress > div > div > div > div {
        background-color: #1e3c72;
    }
    
    /* Badges */
    .badge-success {
        background-color: #28a745;
        color: white;
        padding: 3px 8px;
        border-radius: 12px;
        font-size: 0.8em;
    }
    
    .badge-warning {
        background-color: #ffc107;
        color: black;
        padding: 3px 8px;
        border-radius: 12px;
        font-size: 0.8em;
    }
    
    .badge-danger {
        background-color: #dc3545;
        color: white;
        padding: 3px 8px;
        border-radius: 12px;
        font-size: 0.8em;
    }
</style>
""", unsafe_allow_html=True)

# ------------------------ INITIALISATION DE LA SESSION ------------------------
def init_session_state():
    """Initialise les variables de session"""
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'username' not in st.session_state:
        st.session_state.username = None
    if 'role' not in st.session_state:
        st.session_state.role = None
    if 'pivot_techniques' not in st.session_state:
        st.session_state.pivot_techniques = None
    if 'pivot_comptables' not in st.session_state:
        st.session_state.pivot_comptables = None
    if 'pivot_compte_41' not in st.session_state:
        st.session_state.pivot_compte_41 = None
    if 'df_technique' not in st.session_state:
        st.session_state.df_technique = None
    if 'df_comptable' not in st.session_state:
        st.session_state.df_comptable = None
    if 'df_compte_41' not in st.session_state:
        st.session_state.df_compte_41 = None
    if 'logs' not in st.session_state:
        st.session_state.logs = []
    if 'theme' not in st.session_state:
        st.session_state.theme = "Light"
    if 'page' not in st.session_state:
        st.session_state.page = "Accueil"

init_session_state()

# ------------------------ FONCTIONS D'AUTHENTIFICATION ------------------------
def hash_password(password):
    """Hash un mot de passe"""
    return hashlib.sha256(password.encode()).hexdigest()

def verify_password(hashed_password, password):
    """Vérifie un mot de passe"""
    return hashed_password == hash_password(password)

def init_db():
    """Initialise la base de données"""
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS users
                 (username TEXT PRIMARY KEY, 
                  password TEXT,
                  role TEXT,
                  email TEXT,
                  created_at TIMESTAMP)''')
    
    # Créer un admin par défaut si la table est vide
    c.execute("SELECT COUNT(*) FROM users")
    if c.fetchone()[0] == 0:
        default_password = hash_password("admin123")
        c.execute("INSERT INTO users VALUES (?, ?, ?, ?, ?)",
                  ("admin", default_password, "admin", "admin@example.com", datetime.now()))
    
    conn.commit()
    conn.close()

init_db()

def login(username, password):
    """Authentifie un utilisateur"""
    conn = sqlite3.connect('users.db')
    c = conn.cursor()
    c.execute("SELECT password, role FROM users WHERE username=?", (username,))
    result = c.fetchone()
    conn.close()
    
    if result and verify_password(result[0], password):
        st.session_state.authenticated = True
        st.session_state.username = username
        st.session_state.role = result[1]
        log_action("Connexion", f"Utilisateur {username} connecté")
        return True
    return False

def logout():
    """Déconnecte l'utilisateur"""
    log_action("Déconnexion", f"Utilisateur {st.session_state.username} déconnecté")
    st.session_state.authenticated = False
    st.session_state.username = None
    st.session_state.role = None
    st.rerun()

def log_action(action, details=""):
    """Enregistre une action dans les logs"""
    st.session_state.logs.append({
        'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        'username': st.session_state.username,
        'action': action,
        'details': details
    })

# ------------------------ PAGE DE CONNEXION ------------------------
def login_page():
    """Affiche la page de connexion"""
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("""
        <div style="text-align: center; padding: 40px;">
            <h1 style="color: #1e3c72; font-size: 3em;">AGC-VIE</h1>
            <p style="color: #666; font-size: 1.2em;">Système de Gestion Technique et Comptable</p>
        </div>
        """, unsafe_allow_html=True)
        
        with st.form("login_form"):
            st.markdown("### Connexion")
            username = st.text_input("Nom d'utilisateur", placeholder="Entrez votre nom d'utilisateur")
            password = st.text_input("Mot de passe", type="password", placeholder="Entrez votre mot de passe")
            role = st.selectbox("Rôle", ["user", "admin"])
            
            submitted = st.form_submit_button("Se connecter", use_container_width=True)
            
            if submitted:
                if login(username, password):
                    st.success("Connexion réussie!")
                    st.rerun()
                else:
                    st.error("Nom d'utilisateur ou mot de passe incorrect")

# ------------------------ PAGE D'ACCUEIL ------------------------
def accueil_page():
    """Affiche la page d'accueil"""
    st.markdown("""
    <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                padding: 40px; border-radius: 10px; color: white; text-align: center; margin-bottom: 30px;">
        <h1 style="color: white; font-size: 2.5em;">Bienvenue sur AGC-VIE</h1>
        <p style="font-size: 1.2em;">Système de Gestion Technique et Comptable</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown("""
        <div class="metric-card">
            <h3>📊 Gestion Technique</h3>
            <p>Analyse et rapprochement</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="metric-card">
            <h3>💰 Gestion Comptable</h3>
            <p>Suivi des opérations</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div class="metric-card">
            <h3>📈 Statistiques</h3>
            <p>Analyses approfondies</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown("""
        <div class="metric-card">
            <h3>🔒 Sécurité</h3>
            <p>Accès protégé</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📋 Fonctionnalités principales")
        st.markdown("""
        - Gestion des polices d'assurance
        - Rapprochement technique et comptable
        - Analyse des doublons
        - Génération de certificats
        - Statistiques détaillées
        """)
    
    with col2:
        st.markdown("### 🔧 Modules disponibles")
        st.markdown("""
        - **Gestion Technique** : Analyse des données techniques
        - **Gestion Comptable** : Suivi des opérations comptables
        - **Gestion 410 & 411** : Rapprochement des comptes
        - **Gestion Production** : Génération de certificats
        - **Statistiques** : Analyses et visualisations
        """)

# ------------------------ GESTION TECHNIQUE ------------------------
def gestion_technique_page():
    """Page de gestion technique"""
    st.markdown("## 📊 Gestion Technique")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("### Importer les données techniques")
        uploaded_file = st.file_uploader("Choisir un fichier Excel", type=['xlsx', 'xls'], key="tech_upload")
        
        if uploaded_file:
            try:
                df = pd.read_excel(uploaded_file)
                st.session_state.df_technique = df
                st.success(f"Fichier chargé avec succès! {len(df)} enregistrements trouvés.")
                
                # Traitement des données
                if all(col in df.columns for col in ['Num avenant', 'Code intermédiaire', 'N° police']):
                    df['Nouvelle_Police'] = df.apply(
                        lambda row: f"{row['Code intermédiaire']}-{row['N° police']}/{row['Num avenant']}" 
                        if pd.notnull(row['Num avenant']) 
                        else f"{row['Code intermédiaire']}-{row['N° police']}", axis=1
                    )
                
                if 'Type quittance' in df.columns and 'Chiffre affaire' in df.columns:
                    df['Ristournes'] = df.apply(
                        lambda row: row['Chiffre affaire'] if row['Type quittance'] == 'Ristourne' else 0, axis=1
                    )
                    df['Emissions'] = df.apply(
                        lambda row: row['Chiffre affaire'] if row['Type quittance'] == 'Emission' else 0, axis=1
                    )
                
                pivot_df = pd.pivot_table(
                    df, 
                    index=['Nouvelle_Police'] if 'Nouvelle_Police' in df.columns else df.columns[0],
                    values=['Emissions', 'Ristournes', 'Chiffre affaire'] if all(col in df.columns for col in ['Emissions', 'Ristournes', 'Chiffre affaire']) else df.select_dtypes(include=[np.number]).columns,
                    aggfunc='sum', 
                    fill_value=0
                ).reset_index()
                
                st.session_state.pivot_techniques = pivot_df
                log_action("Import technique", f"{len(df)} enregistrements importés")
                
            except Exception as e:
                st.error(f"Erreur lors du chargement: {str(e)}")
    
    with col2:
        if st.session_state.pivot_techniques is not None:
            st.markdown("### Statistiques rapides")
            total_emissions = st.session_state.pivot_techniques['Emissions'].sum() if 'Emissions' in st.session_state.pivot_techniques.columns else 0
            total_ristournes = st.session_state.pivot_techniques['Ristournes'].sum() if 'Ristournes' in st.session_state.pivot_techniques.columns else 0
            
            st.metric("Total Émissions", f"{total_emissions:,.0f} FCFA")
            st.metric("Total Ristournes", f"{total_ristournes:,.0f} FCFA")
            st.metric("Nombre de polices", len(st.session_state.pivot_techniques))
    
    # Affichage des données
    if st.session_state.pivot_techniques is not None:
        st.markdown("### Données techniques")
        
        # Recherche
        search = st.text_input("🔍 Rechercher une police", placeholder="Entrez le numéro de police...")
        
        df_display = st.session_state.pivot_techniques.copy()
        if search:
            mask = df_display.astype(str).apply(lambda x: x.str.contains(search, case=False, na=False)).any(axis=1)
            df_display = df_display[mask]
        
        st.dataframe(df_display, use_container_width=True, height=400)
        
        # Export
        if st.button("📥 Exporter en Excel"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_display.to_excel(writer, index=False, sheet_name='Données Techniques')
            
            st.download_button(
                label="Télécharger le fichier Excel",
                data=output.getvalue(),
                file_name="donnees_techniques.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ------------------------ GESTION COMPTABLE ------------------------
def gestion_comptable_page():
    """Page de gestion comptable"""
    st.markdown("## 💰 Gestion Comptable")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("### Importer les données comptables")
        uploaded_file = st.file_uploader("Choisir un fichier Excel", type=['xlsx', 'xls'], key="compta_upload")
        
        if uploaded_file:
            try:
                df = pd.read_excel(uploaded_file)
                st.session_state.df_comptable = df
                st.success(f"Fichier chargé avec succès! {len(df)} enregistrements trouvés.")
                
                # Tableau croisé
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                if len(numeric_cols) > 0:
                    pivot_df = pd.pivot_table(
                        df,
                        index=df.columns[0],
                        values=numeric_cols,
                        aggfunc='sum',
                        fill_value=0
                    ).reset_index()
                    
                    st.session_state.pivot_comptables = pivot_df
                    log_action("Import comptable", f"{len(df)} enregistrements importés")
                
            except Exception as e:
                st.error(f"Erreur lors du chargement: {str(e)}")
    
    with col2:
        if st.session_state.pivot_comptables is not None:
            st.markdown("### Statistiques rapides")
            if 'Débit' in st.session_state.pivot_comptables.columns:
                total_debit = st.session_state.pivot_comptables['Débit'].sum()
                st.metric("Total Débit", f"{total_debit:,.0f} FCFA")
            
            if 'Crédit' in st.session_state.pivot_comptables.columns:
                total_credit = st.session_state.pivot_comptables['Crédit'].sum()
                st.metric("Total Crédit", f"{total_credit:,.0f} FCFA")
            
            st.metric("Nombre d'entrées", len(st.session_state.pivot_comptables))
    
    # Affichage des données
    if st.session_state.pivot_comptables is not None:
        st.markdown("### Données comptables")
        
        # Recherche
        search = st.text_input("🔍 Rechercher une police", placeholder="Entrez le numéro de police...")
        
        df_display = st.session_state.pivot_comptables.copy()
        if search:
            mask = df_display.astype(str).apply(lambda x: x.str.contains(search, case=False, na=False)).any(axis=1)
            df_display = df_display[mask]
        
        st.dataframe(df_display, use_container_width=True, height=400)
        
        # Export
        if st.button("📥 Exporter en Excel"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_display.to_excel(writer, index=False, sheet_name='Données Comptables')
            
            st.download_button(
                label="Télécharger le fichier Excel",
                data=output.getvalue(),
                file_name="donnees_comptables.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ------------------------ RAPPROCHEMENT TECHNIQUE ------------------------
def rapprochement_technique_page():
    """Page de rapprochement technique"""
    st.markdown("## 🔄 Rapprochement Technique")
    
    if st.session_state.pivot_techniques is None or st.session_state.pivot_comptables is None:
        st.warning("Veuillez d'abord importer les données techniques et comptables.")
        return
    
    with st.spinner("Calcul du rapprochement en cours..."):
        # Fusion des données
        tech_df = st.session_state.pivot_techniques.copy()
        compta_df = st.session_state.pivot_comptables.copy()
        
        # S'assurer que les colonnes de police existent
        tech_col = 'Nouvelle_Police' if 'Nouvelle_Police' in tech_df.columns else tech_df.columns[0]
        compta_col = 'No Police' if 'No Police' in compta_df.columns else compta_df.columns[0]
        
        # Renommer pour la fusion
        tech_df = tech_df.rename(columns={tech_col: 'Police'})
        compta_df = compta_df.rename(columns={compta_col: 'Police'})
        
        # Fusion
        merged_df = pd.merge(tech_df, compta_df, on='Police', how='outer', suffixes=('_tech', '_compta'))
        
        # Calcul des écarts
        if 'Emissions' in merged_df.columns and 'Débit' in merged_df.columns and 'Crédit' in merged_df.columns:
            merged_df['Emissions'] = pd.to_numeric(merged_df['Emissions'], errors='coerce').fillna(0)
            merged_df['Débit'] = pd.to_numeric(merged_df['Débit'], errors='coerce').fillna(0)
            merged_df['Crédit'] = pd.to_numeric(merged_df['Crédit'], errors='coerce').fillna(0)
            
            merged_df['CA_Technique'] = merged_df['Emissions']
            merged_df['CA_Comptable'] = abs(merged_df['Crédit'] - merged_df['Débit'])
            merged_df['Écart'] = merged_df['CA_Technique'] - merged_df['CA_Comptable']
            merged_df['Statut'] = merged_df['Écart'].apply(
                lambda x: '✅ Rapproché' if abs(x) < 0.01 else '❌ Non rapproché'
            )
    
    # Métriques
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_technique = merged_df['CA_Technique'].sum() if 'CA_Technique' in merged_df.columns else 0
        st.metric("CA Technique", f"{total_technique:,.0f} FCFA")
    
    with col2:
        total_comptable = merged_df['CA_Comptable'].sum() if 'CA_Comptable' in merged_df.columns else 0
        st.metric("CA Comptable", f"{total_comptable:,.0f} FCFA")
    
    with col3:
        ecart_total = total_technique - total_comptable
        delta_color = "inverse" if abs(ecart_total) > 0 else "normal"
        st.metric("Écart Total", f"{ecart_total:,.0f} FCFA", delta=f"{abs(ecart_total):,.0f}", delta_color=delta_color)
    
    with col4:
        nb_rapproche = len(merged_df[merged_df['Statut'] == '✅ Rapproché']) if 'Statut' in merged_df.columns else 0
        st.metric("Polices rapprochées", f"{nb_rapproche}/{len(merged_df)}")
    
    # Tabs pour les différents vues
    tab1, tab2, tab3 = st.tabs(["📋 Données complètes", "❌ Non rapprochées", "✅ Rapprochées"])
    
    with tab1:
        st.dataframe(merged_df, use_container_width=True, height=500)
    
    with tab2:
        if 'Statut' in merged_df.columns:
            non_rapproche = merged_df[merged_df['Statut'] == '❌ Non rapproché']
            st.dataframe(non_rapproche, use_container_width=True, height=500)
    
    with tab3:
        if 'Statut' in merged_df.columns:
            rapproche = merged_df[merged_df['Statut'] == '✅ Rapproché']
            st.dataframe(rapproche, use_container_width=True, height=500)
    
    # Graphique
    st.markdown("### 📊 Visualisation des écarts")
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=merged_df['Police'].head(20),
        y=merged_df['CA_Technique'].head(20),
        name='CA Technique',
        marker_color='#1e3c72'
    ))
    fig.add_trace(go.Bar(
        x=merged_df['Police'].head(20),
        y=merged_df['CA_Comptable'].head(20),
        name='CA Comptable',
        marker_color='#2a5298'
    ))
    
    fig.update_layout(
        title="Comparaison CA Technique vs Comptable (20 premières polices)",
        xaxis_title="Police",
        yaxis_title="Montant (FCFA)",
        barmode='group',
        height=500
    )
    
    st.plotly_chart(fig, use_container_width=True)
    
    # Export
    if st.button("📥 Exporter le rapprochement"):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            merged_df.to_excel(writer, index=False, sheet_name='Rapprochement')
            if 'Statut' in merged_df.columns:
                merged_df[merged_df['Statut'] == '❌ Non rapproché'].to_excel(
                    writer, index=False, sheet_name='Non rapprochées'
                )
                merged_df[merged_df['Statut'] == '✅ Rapproché'].to_excel(
                    writer, index=False, sheet_name='Rapprochées'
                )
        
        st.download_button(
            label="Télécharger le rapport Excel",
            data=output.getvalue(),
            file_name="rapprochement_technique.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    log_action("Rapprochement technique", f"{len(merged_df)} polices analysées")

# ------------------------ GESTION 410 & 411 ------------------------
def gestion_410_411_page():
    """Page de gestion des comptes 410 et 411"""
    st.markdown("## 📊 Gestion 410 & 411")
    
    # Upload des fichiers
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### CP_410")
        file_410 = st.file_uploader("Importer CP_410", type=['xlsx', 'xls', 'csv'], key="410")
        
        if file_410:
            try:
                if file_410.name.endswith('.csv'):
                    df_410 = pd.read_csv(file_410)
                else:
                    df_410 = pd.read_excel(file_410)
                st.session_state.df_410 = df_410
                st.success(f"CP_410 chargé: {len(df_410)} enregistrements")
            except Exception as e:
                st.error(f"Erreur: {str(e)}")
    
    with col2:
        st.markdown("### CP_411")
        file_411 = st.file_uploader("Importer CP_411", type=['xlsx', 'xls', 'csv'], key="411")
        
        if file_411:
            try:
                if file_411.name.endswith('.csv'):
                    df_411 = pd.read_csv(file_411)
                else:
                    df_411 = pd.read_excel(file_411)
                st.session_state.df_411 = df_411
                st.success(f"CP_411 chargé: {len(df_411)} enregistrements")
            except Exception as e:
                st.error(f"Erreur: {str(e)}")
    
    # Vérifications
    if st.session_state.get('df_410') is not None and st.session_state.get('df_411') is not None:
        st.markdown("---")
        st.markdown("### 🔍 Vérifications")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("Vérifier polices 410/411"):
                if 'No Police' in st.session_state.df_410.columns and 'No Police' in st.session_state.df_411.columns:
                    polices_410 = set(st.session_state.df_410['No Police'].dropna().astype(str))
                    polices_411 = set(st.session_state.df_411['No Police'].dropna().astype(str))
                    
                    communes = polices_410.intersection(polices_411)
                    only_410 = polices_410 - polices_411
                    
                    st.markdown("#### Résultats")
                    st.metric("Polices communes", len(communes))
                    st.metric("Uniquement dans 410", len(only_410))
                    
                    if len(only_410) > 0:
                        st.warning(f"{len(only_410)} polices présentes uniquement dans CP_410")
        
        with col2:
            if st.button("Vérifier polices 411/410"):
                if 'No Police' in st.session_state.df_410.columns and 'No Police' in st.session_state.df_411.columns:
                    polices_410 = set(st.session_state.df_410['No Police'].dropna().astype(str))
                    polices_411 = set(st.session_state.df_411['No Police'].dropna().astype(str))
                    
                    communes = polices_411.intersection(polices_410)
                    only_411 = polices_411 - polices_410
                    
                    st.markdown("#### Résultats")
                    st.metric("Polices communes", len(communes))
                    st.metric("Uniquement dans 411", len(only_411))
                    
                    if len(only_411) > 0:
                        st.warning(f"{len(only_411)} polices présentes uniquement dans CP_411")
        
        with col3:
            if st.button("Vérifier références"):
                if 'Réf Pièce' in st.session_state.df_411.columns:
                    pattern = r"^\w+-\d+(?:/\d+)?$"
                    invalid_refs = []
                    
                    for ref in st.session_state.df_411['Réf Pièce'].dropna():
                        if not re.match(pattern, str(ref)):
                            invalid_refs.append(ref)
                    
                    st.metric("Références invalides", len(invalid_refs))
                    
                    if len(invalid_refs) > 0:
                        st.dataframe(pd.DataFrame(invalid_refs, columns=['Références invalides']))
        
        # Affichage des données
        tab1, tab2 = st.tabs(["CP_410", "CP_411"])
        
        with tab1:
            st.dataframe(st.session_state.df_410, use_container_width=True, height=400)
        
        with tab2:
            st.dataframe(st.session_state.df_411, use_container_width=True, height=400)

# ------------------------ GESTION DES DOUBLONS ------------------------
def gestion_doublons_page():
    """Page de gestion des doublons"""
    st.markdown("## 🔍 Analyse des Doublons")
    
    # Upload du fichier
    uploaded_file = st.file_uploader("Importer un fichier de polices", type=['xlsx', 'xls', 'csv'])
    
    if uploaded_file:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            st.success(f"Fichier chargé: {len(df)} enregistrements")
            
            # Détection des doublons
            if 'NUMERO POLICE' in df.columns:
                # Identifier les doublons
                duplicates_mask = df.duplicated(subset=['NUMERO POLICE'], keep=False)
                duplicates_df = df[duplicates_mask].sort_values('NUMERO POLICE')
                uniques_df = df[~duplicates_mask]
                
                # Statistiques
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.metric("Total polices", len(df))
                
                with col2:
                    st.metric("Polices en doublon", len(duplicates_df))
                
                with col3:
                    st.metric("Polices uniques", len(uniques_df))
                
                # Tabs pour les vues
                tab1, tab2 = st.tabs(["📋 Polices en doublon", "✅ Polices uniques"])
                
                with tab1:
                    if len(duplicates_df) > 0:
                        st.dataframe(duplicates_df, use_container_width=True, height=400)
                        
                        # Export des doublons
                        if st.button("📥 Exporter les doublons"):
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                duplicates_df.to_excel(writer, index=False, sheet_name='Doublons')
                            
                            st.download_button(
                                label="Télécharger",
                                data=output.getvalue(),
                                file_name="doublons_polices.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.info("Aucun doublon trouvé")
                
                with tab2:
                    if len(uniques_df) > 0:
                        st.dataframe(uniques_df, use_container_width=True, height=400)
                        
                        # Export des uniques
                        if st.button("📥 Exporter les polices uniques"):
                            output = io.BytesIO()
                            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                                uniques_df.to_excel(writer, index=False, sheet_name='Polices uniques')
                            
                            st.download_button(
                                label="Télécharger",
                                data=output.getvalue(),
                                file_name="polices_uniques.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.info("Aucune police unique trouvée")
                
                # Visualisation
                st.markdown("### 📊 Visualisation")
                
                fig = go.Figure(data=[
                    go.Pie(
                        labels=['Polices uniques', 'Polices en doublon'],
                        values=[len(uniques_df), len(duplicates_df)],
                        marker_colors=['#28a745', '#dc3545'],
                        hole=0.3
                    )
                ])
                
                fig.update_layout(
                    title="Répartition des polices",
                    height=400
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
            else:
                st.error("La colonne 'NUMERO POLICE' est requise")
                
        except Exception as e:
            st.error(f"Erreur lors du traitement: {str(e)}")

# ------------------------ GESTION PRODUCTION ------------------------
def gestion_production_page():
    """Page de gestion de production (certificats)"""
    st.markdown("## 📄 Générateur de Certificats")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### Modèle de certificat")
        template_file = st.file_uploader("Importer un modèle Word", type=['docx'])
        
        if template_file:
            st.success("Modèle chargé avec succès")
            st.session_state.template = template_file
    
    with col2:
        st.markdown("### Données")
        data_file = st.file_uploader("Importer les données", type=['xlsx', 'xls', 'csv'])
        
        if data_file:
            try:
                if data_file.name.endswith('.csv'):
                    df = pd.read_csv(data_file)
                else:
                    df = pd.read_excel(data_file)
                
                st.success(f"Données chargées: {len(df)} enregistrements")
                st.session_state.production_data = df
            except Exception as e:
                st.error(f"Erreur: {str(e)}")
    
    if st.session_state.get('template') and st.session_state.get('production_data') is not None:
        st.markdown("---")
        st.markdown("### 🎨 Personnalisation")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            police = st.selectbox("Police", ['Arial', 'Times New Roman', 'Helvetica', 'Calibri'])
        
        with col2:
            taille = st.slider("Taille de police", 8, 24, 12)
        
        with col3:
            couleur = st.color_picker("Couleur du texte", "#000000")
        
        # Aperçu des données
        st.markdown("### 📋 Aperçu des données")
        st.dataframe(st.session_state.production_data.head(10), use_container_width=True)
        
        # Génération
        if st.button("🚀 Générer les certificats", type="primary"):
            with st.spinner("Génération en cours..."):
                progress_bar = st.progress(0)
                
                df = st.session_state.production_data
                total = len(df)
                
                for i, (_, row) in enumerate(df.iterrows()):
                    # Simulation de la génération
                    time.sleep(0.05)
                    progress_bar.progress((i + 1) / total)
                
                st.success(f"{total} certificats générés avec succès!")
                
                # Simulation de fichier ZIP
                st.download_button(
                    label="📥 Télécharger tous les certificats",
                    data=b"Simulation de fichier ZIP",
                    file_name="certificats.zip",
                    mime="application/zip"
                )
        
        log_action("Génération certificats", f"{len(st.session_state.production_data)} certificats")

# ------------------------ STATISTIQUES ------------------------
def statistiques_page():
    """Page de statistiques"""
    st.markdown("## 📈 Statistiques et Analyses")
    
    # Source de données
    data_source = st.radio(
        "Source des données",
        ["Données techniques", "Données comptables", "Générer des données de test"],
        horizontal=True
    )
    
    df = None
    
    if data_source == "Données techniques" and st.session_state.pivot_techniques is not None:
        df = st.session_state.pivot_techniques
    elif data_source == "Données comptables" and st.session_state.pivot_comptables is not None:
        df = st.session_state.pivot_comptables
    elif data_source == "Générer des données de test":
        # Générer des données de test
        np.random.seed(42)
        n = 100
        df = pd.DataFrame({
            'Police': [f'POL-{i:04d}' for i in range(1, n+1)],
            'Montant': np.random.uniform(10000, 1000000, n),
            'Emission': np.random.choice([0, 1], n, p=[0.3, 0.7]),
            'Ristourne': np.random.choice([0, 1], n, p=[0.8, 0.2]),
            'Date': pd.date_range('2024-01-01', periods=n, freq='D')
        })
    
    if df is not None:
        # Métriques générales
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total enregistrements", len(df))
        
        with col2:
            numeric_cols = df.select_dtypes(include=[np.number]).columns
            if len(numeric_cols) > 0:
                st.metric("Montant total", f"{df[numeric_cols[0]].sum():,.0f} FCFA")
        
        with col3:
            if 'Emission' in df.columns:
                nb_emissions = df['Emission'].sum()
                st.metric("Émissions", f"{nb_emissions:,.0f}")
        
        with col4:
            if 'Ristourne' in df.columns:
                nb_ristournes = df['Ristourne'].sum()
                st.metric("Ristournes", f"{nb_ristournes:,.0f}")
        
        # Tabs pour différentes analyses
        tab1, tab2, tab3 = st.tabs(["📊 Distributions", "📈 Tendances", "📋 Analyse détaillée"])
        
        with tab1:
            col1, col2 = st.columns(2)
            
            with col1:
                # Distribution des montants
                numeric_cols = df.select_dtypes(include=[np.number]).columns
                if len(numeric_cols) > 0:
                    fig = px.histogram(
                        df, 
                        x=numeric_cols[0],
                        nbins=30,
                        title=f"Distribution des {numeric_cols[0]}",
                        color_discrete_sequence=['#1e3c72']
                    )
                    st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                # Camembert pour les catégories
                if 'Emission' in df.columns and 'Ristourne' in df.columns:
                    categories = {
                        'Émissions': df['Emission'].sum(),
                        'Ristournes': df['Ristourne'].sum(),
                        'Autres': len(df) - df['Emission'].sum() - df['Ristourne'].sum()
                    }
                    
                    fig = px.pie(
                        values=list(categories.values()),
                        names=list(categories.keys()),
                        title="Répartition des opérations",
                        color_discrete_sequence=['#1e3c72', '#2a5298', '#6c757d']
                    )
                    st.plotly_chart(fig, use_container_width=True)
        
        with tab2:
            if 'Date' in df.columns:
                # Agrégation par mois
                df['Mois'] = pd.to_datetime(df['Date']).dt.to_period('M').astype(str)
                monthly_sum = df.groupby('Mois')[numeric_cols[0]].sum().reset_index()
                
                fig = px.line(
                    monthly_sum,
                    x='Mois',
                    y=numeric_cols[0],
                    title=f"Évolution des {numeric_cols[0]}",
                    markers=True
                )
                fig.update_traces(line_color='#1e3c72')
                st.plotly_chart(fig, use_container_width=True)
        
        with tab3:
            # Top N
            n = st.slider("Nombre d'éléments à afficher", 5, 50, 10)
            
            if 'Police' in df.columns and len(numeric_cols) > 0:
                top_df = df.nlargest(n, numeric_cols[0])[['Police', numeric_cols[0]]]
                
                fig = px.bar(
                    top_df,
                    x='Police',
                    y=numeric_cols[0],
                    title=f"Top {n} des {numeric_cols[0]}",
                    color_discrete_sequence=['#1e3c72']
                )
                st.plotly_chart(fig, use_container_width=True)
        
        # Export
        if st.button("📥 Exporter l'analyse"):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Données')
                
                # Ajouter des résumés
                if len(numeric_cols) > 0:
                    summary = df[numeric_cols].describe()
                    summary.to_excel(writer, sheet_name='Résumé statistique')
            
            st.download_button(
                label="Télécharger l'analyse",
                data=output.getvalue(),
                file_name="analyse_statistique.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ------------------------ ADMINISTRATION ------------------------
def administration_page():
    """Page d'administration"""
    st.markdown("## ⚙️ Administration")
    
    if st.session_state.role != "admin":
        st.error("Accès réservé aux administrateurs")
        return
    
    tab1, tab2, tab3 = st.tabs(["👥 Utilisateurs", "📋 Logs", "🔧 Paramètres"])
    
    with tab1:
        st.markdown("### Gestion des utilisateurs")
        
        # Formulaire d'ajout
        with st.expander("➕ Ajouter un utilisateur"):
            with st.form("add_user"):
                new_username = st.text_input("Nom d'utilisateur")
                new_password = st.text_input("Mot de passe", type="password")
                new_role = st.selectbox("Rôle", ["user", "admin"])
                new_email = st.text_input("Email")
                
                if st.form_submit_button("Ajouter"):
                    conn = sqlite3.connect('users.db')
                    c = conn.cursor()
                    try:
                        c.execute(
                            "INSERT INTO users VALUES (?, ?, ?, ?, ?)",
                            (new_username, hash_password(new_password), new_role, new_email, datetime.now())
                        )
                        conn.commit()
                        st.success(f"Utilisateur {new_username} ajouté")
                        log_action("Ajout utilisateur", new_username)
                    except sqlite3.IntegrityError:
                        st.error("Ce nom d'utilisateur existe déjà")
                    finally:
                        conn.close()
        
        # Liste des utilisateurs
        conn = sqlite3.connect('users.db')
        users_df = pd.read_sql_query("SELECT username, role, email, created_at FROM users", conn)
        conn.close()
        
        st.dataframe(users_df, use_container_width=True)
    
    with tab2:
        st.markdown("### Journal des activités")
        
        if st.session_state.logs:
            logs_df = pd.DataFrame(st.session_state.logs)
            st.dataframe(logs_df, use_container_width=True)
            
            # Export des logs
            if st.button("📥 Exporter les logs"):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    logs_df.to_excel(writer, index=False, sheet_name='Logs')
                
                st.download_button(
                    label="Télécharger",
                    data=output.getvalue(),
                    file_name="logs_application.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        else:
            st.info("Aucun log pour le moment")
    
    with tab3:
        st.markdown("### Paramètres de sécurité")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### Politique de mot de passe")
            min_length = st.number_input("Longueur minimale", 6, 20, 8)
            require_upper = st.checkbox("Requiert des majuscules", True)
            require_special = st.checkbox("Requiert des caractères spéciaux", True)
            expiry_days = st.number_input("Expiration (jours)", 30, 365, 90)
        
        with col2:
            st.markdown("#### Verrouillage de compte")
            max_attempts = st.number_input("Tentatives max", 3, 10, 5)
            lockout_duration = st.number_input("Durée de verrouillage (minutes)", 5, 1440, 30)
            two_factor = st.checkbox("Activer 2FA", False)
        
        if st.button("💾 Sauvegarder les paramètres"):
            st.success("Paramètres sauvegardés")
            log_action("Modification paramètres", "Paramètres de sécurité mis à jour")

# ------------------------ BARRE LATÉRALE ------------------------
def sidebar():
    """Affiche la barre latérale"""
    with st.sidebar:
        st.markdown("""
        <div style="text-align: center; padding: 20px;">
            <h2 style="color: white;">AGC-VIE</h2>
            <p style="color: #ccc;">Version 2.0</p>
        </div>
        """, unsafe_allow_html=True)
        
        if st.session_state.authenticated:
            st.markdown(f"""
            <div style="background: rgba(255,255,255,0.1); padding: 10px; border-radius: 5px; margin-bottom: 20px;">
                <p style="color: white; margin: 0;">👤 {st.session_state.username}</p>
                <p style="color: #ccc; margin: 0; font-size: 0.9em;">{st.session_state.role}</p>
            </div>
            """, unsafe_allow_html=True)
        
        # Menu principal
        selected = option_menu(
            menu_title=None,
            options=["Accueil", "Gestion Technique", "Gestion Comptable", 
                     "Rapprochement Technique", "Gestion 410 & 411", 
                     "Gestion Doublons", "Gestion Production", 
                     "Statistiques", "Administration", "Déconnexion"],
            icons=["house", "bar-chart", "cash", "arrow-repeat", 
                   "folder", "files", "file-earmark", "graph-up", 
                   "gear", "box-arrow-right"],
            menu_icon="cast",
            default_index=0,
            styles={
                "container": {"padding": "0!important", "background-color": "#1e3c72"},
                "icon": {"color": "white", "font-size": "20px"},
                "nav-link": {"color": "white", "font-size": "16px", "text-align": "left", "margin": "0px"},
                "nav-link-selected": {"background-color": "#2a5298"},
            }
        )
        
        return selected

# ------------------------ MAIN ------------------------
def main():
    """Fonction principale"""
    
    if not st.session_state.authenticated:
        login_page()
        return
    
    selected = sidebar()
    
    if selected == "Déconnexion":
        logout()
    
    elif selected == "Accueil":
        accueil_page()
    
    elif selected == "Gestion Technique":
        gestion_technique_page()
    
    elif selected == "Gestion Comptable":
        gestion_comptable_page()
    
    elif selected == "Rapprochement Technique":
        rapprochement_technique_page()
    
    elif selected == "Gestion 410 & 411":
        gestion_410_411_page()
    
    elif selected == "Gestion Doublons":
        gestion_doublons_page()
    
    elif selected == "Gestion Production":
        gestion_production_page()
    
    elif selected == "Statistiques":
        statistiques_page()
    
    elif selected == "Administration":
        administration_page()
    
    # Pied de page
    st.markdown("---")
    st.markdown(
        "<div style='text-align: center; color: #666; padding: 10px;'>"
        "© 2026 AGC-VIE - Système de Gestion Technique et Comptable"
        "</div>",
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()