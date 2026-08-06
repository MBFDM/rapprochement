"""
Application de Gestion Technique et Comptable AGC-VIE
Version Streamlit - Conversion complète de l'application Tkinter
Auteur: Frédéric BAYONNE MAVOUNGOU
Date: 2025
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import os
import io
import re
import uuid
import time
import json
import base64
import hashlib
import sqlite3
import logging
import shutil
import warnings
from contextlib import contextmanager
from typing import Optional, Dict, List, Any, Tuple
from pathlib import Path
import tempfile
import base64
import os
from PIL import Image

# Tentative d'import des modules optionnels
try:
    import bcrypt
except ImportError:
    bcrypt = None

try:
    from cryptography.fernet import Fernet
except ImportError:
    Fernet = None

try:
    from docx import Document
    from docx.shared import Pt, RGBColor
except ImportError:
    Document = None

try:
    from docx2pdf import convert
except ImportError:
    convert = None

try:
    import fitz
except ImportError:
    fitz = None

warnings.filterwarnings('ignore')

# ======================== CONFIGURATION DE LA PAGE ========================
st.set_page_config(
    page_title="AGC-VIE - Gestion Technique et Comptable",
    page_icon="logo_1.jpg",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'Get Help': 'https://www.agc-vie.com',
        'Report a bug': 'https://www.agc-vie.com/bug',
        'About': 'AGC-VIE - Système de Gestion Technique et Comptable v2.0'
    }
)

# ======================== CONSTANTES GLOBALES ========================
DB_FILE = "admin_system.db"
PERMISSIONS_LIST = [
    "admin", "user_manage", "content_manage", 
    "settings_manage", "logs_view", "logs_manage"
]

# ======================== STYLES CSS PERSONNALISÉS ========================
def apply_custom_css():
    """Applique les styles CSS personnalisés"""
    st.markdown("""
    <style>
        /* Style global */
        .stApp {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }
        
        /* En-têtes */
        h1, h2, h3 {
            color: #1e3c72;
            font-weight: 600;
            margin-bottom: 1rem;
        }
        
        h1 {
            font-size: 2.5rem;
            #border-bottom: 3px solid #1e3c72;
            padding-bottom: 0.5rem;
        }
        
        h2 {
            font-size: 2rem;
            #border-bottom: 2px solid #2a5298;
            padding-bottom: 0.3rem;
        }
        
        /* Cartes métriques */
        .metric-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 10px;
            border-radius: 15px;
            color: white;
            text-align: center;
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
            transition: transform 0.3s;
            margin: 10px 0;
        }
        
        .metric-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 15px 40px rgba(0,0,0,0.3);
        }
        
        .metric-card h3 {
            color: white;
            font-size: 1.2em;
            margin-bottom: 10px;
            opacity: 0.9;
        }
        
        .metric-card p {
            font-size: 2.2em;
            font-weight: bold;
            margin: 0;
        }
        
        /* Conteneurs */
        .content-card {
            background: white;
            padding: 20px;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
            margin: 20px 0;
            transition: transform 0.3s;
        }
        
        .content-card:hover {
            transform: translateY(-5px);
            box-shadow: 0 15px 40px rgba(0,0,0,0.15);
        }
        
        /* Badges */
        .badge {
            display: inline-block;
            padding: 5px 10px;
            border-radius: 20px;
            font-size: 0.85em;
            font-weight: 600;
            text-align: center;
        }
        
        .badge-success {
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
            color: white;
        }
        
        .badge-warning {
            background: linear-gradient(135deg, #ffc107 0%, #fd7e14 100%);
            color: black;
        }
        
        .badge-danger {
            background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);
            color: white;
        }
        
        .badge-info {
            background: linear-gradient(135deg, #17a2b8 0%, #138496 100%);
            color: white;
        }
        
        /* Animations */
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(20px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        .fade-in {
            animation: fadeIn 0.5s ease-out;
        }
        
        /* Pied de page */
        .footer {
            text-align: center;
            padding: 20px;
            color: #6c757d;
            font-size: 0.9em;
            border-top: 1px solid #e0e0e0;
            margin-top: 40px;
        }
    </style>
    """, unsafe_allow_html=True)

# ======================== INITIALISATION DE LA SESSION ========================
def init_session_state():
    """Initialise toutes les variables de session"""
    
    # Authentification
    if 'authenticated' not in st.session_state:
        st.session_state.authenticated = False
    if 'username' not in st.session_state:
        st.session_state.username = None
    if 'role' not in st.session_state:
        st.session_state.role = None
    if 'login_attempts' not in st.session_state:
        st.session_state.login_attempts = 0
    if 'locked_until' not in st.session_state:
        st.session_state.locked_until = None
    
    # Données principales
    if 'pivot_techniques' not in st.session_state:
        st.session_state.pivot_techniques = None
    if 'pivot_comptables' not in st.session_state:
        st.session_state.pivot_comptables = None
    
    # DataFrames
    if 'df_technique' not in st.session_state:
        st.session_state.df_technique = None
    if 'df_comptable' not in st.session_state:
        st.session_state.df_comptable = None
    if 'df_410' not in st.session_state:
        st.session_state.df_410 = None
    if 'df_411' not in st.session_state:
        st.session_state.df_411 = None
    if 'production_data' not in st.session_state:
        st.session_state.production_data = None
    
    # Résultats de vérification
    if 'tableau_listing_police_invalide_comptable' not in st.session_state:
        st.session_state.tableau_listing_police_invalide_comptable = None
    if 'tableau_listing_valide_comptable' not in st.session_state:
        st.session_state.tableau_listing_valide_comptable = None
    if 'pivot_comptables_complet' not in st.session_state:
        st.session_state.pivot_comptables_complet = None
    
    # Logs et historique
    if 'logs' not in st.session_state:
        st.session_state.logs = []
    if 'history' not in st.session_state:
        st.session_state.history = []
    
    # Configuration
    if 'page' not in st.session_state:
        st.session_state.page = "Accueil"
    if 'template' not in st.session_state:
        st.session_state.template = None
    
    # Statistiques
    if 'stats' not in st.session_state:
        st.session_state.stats = {
            'total_imports': 0,
            'total_verifications': 0,
            'total_certificats': 0,
            'last_action': None
        }
    
    # Configuration de la sécurité
    if 'security_config' not in st.session_state:
        st.session_state.security_config = {
            'min_password_length': 8,
            'require_uppercase': True,
            'require_special': True,
            'require_digit': True,
            'max_login_attempts': 5,
            'lockout_duration': 30,
            'session_timeout': 30,
            'two_factor_enabled': False
        }
    
    # Dernière activité
    if 'last_activity' not in st.session_state:
        st.session_state.last_activity = datetime.now()
    
    # Préférences utilisateur
    if 'user_preferences' not in st.session_state:
        st.session_state.user_preferences = {
            'items_per_page': 50,
            'default_export_format': 'excel',
            'show_preview': True,
            'auto_save': True
        }
    
    # Pagination
    if 'tech_page' not in st.session_state:
        st.session_state.tech_page = 0
    if 'compta_page' not in st.session_state:
        st.session_state.compta_page = 0
    if 'rapprochement_page' not in st.session_state:
        st.session_state.rapprochement_page = 0

# ======================== FONCTIONS DE LOGGING ========================
def log_action(action, details="", level="info"):
    """Enregistre une action dans les logs"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    log_entry = {
        'timestamp': timestamp,
        'username': st.session_state.username if st.session_state.authenticated else "anonymous",
        'action': action,
        'details': details,
        'level': level
    }
    
    st.session_state.logs.append(log_entry)
    
    # Garder seulement les 1000 derniers logs
    if len(st.session_state.logs) > 1000:
        st.session_state.logs = st.session_state.logs[-1000:]

def log_history(action_type, target_user=None, details="", data=None):
    """Enregistre dans l'historique"""
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    history_entry = {
        'timestamp': timestamp,
        'username': st.session_state.username if st.session_state.authenticated else "system",
        'action_type': action_type,
        'target_user': target_user,
        'details': details,
        'data': data
    }
    
    st.session_state.history.append(history_entry)
    
    # Garder seulement les 500 derniers historiques
    if len(st.session_state.history) > 500:
        st.session_state.history = st.session_state.history[-500:]

# ======================== FONCTIONS DE SÉCURITÉ ========================
class SecurityManager:
    """Gestionnaire de sécurité"""
    
    @staticmethod
    def hash_password(password):
        """Hash un mot de passe avec bcrypt"""
        if bcrypt is None:
            # Fallback si bcrypt n'est pas installé
            return hashlib.sha256(password.encode('utf-8')).hexdigest()
        salt = bcrypt.gensalt()
        return bcrypt.hashpw(password.encode('utf-8'), salt).decode('utf-8')
    
    @staticmethod
    def verify_password(hashed_password, plain_password):
        """Vérifie un mot de passe"""
        if bcrypt is None:
            # Fallback si bcrypt n'est pas installé
            return hashed_password == hashlib.sha256(plain_password.encode('utf-8')).hexdigest()
        try:
            return bcrypt.checkpw(plain_password.encode('utf-8'), hashed_password.encode('utf-8'))
        except:
            return False
    
    @staticmethod
    def validate_password_strength(password):
        """Valide la force d'un mot de passe"""
        config = st.session_state.security_config
        errors = []
        
        if len(password) < config['min_password_length']:
            errors.append(f"Le mot de passe doit contenir au moins {config['min_password_length']} caractères")
        
        if config['require_uppercase'] and not any(c.isupper() for c in password):
            errors.append("Le mot de passe doit contenir au moins une majuscule")
        
        if config['require_digit'] and not any(c.isdigit() for c in password):
            errors.append("Le mot de passe doit contenir au moins un chiffre")
        
        if config['require_special'] and not any(c in '!@#$%^&*()_+-=[]{}|;:,.<>?' for c in password):
            errors.append("Le mot de passe doit contenir au moins un caractère spécial")
        
        return len(errors) == 0, errors
    
    @staticmethod
    def check_login_attempts():
        """Vérifie le nombre de tentatives de connexion"""
        if st.session_state.locked_until:
            if datetime.now() < st.session_state.locked_until:
                remaining = (st.session_state.locked_until - datetime.now()).seconds // 60
                return False, f"Compte verrouillé. Réessayez dans {remaining} minutes"
            else:
                st.session_state.locked_until = None
                st.session_state.login_attempts = 0
        
        return True, "OK"
    
    @staticmethod
    def record_failed_attempt():
        """Enregistre une tentative échouée"""
        st.session_state.login_attempts += 1
        
        if st.session_state.login_attempts >= st.session_state.security_config['max_login_attempts']:
            lockout_minutes = st.session_state.security_config['lockout_duration']
            st.session_state.locked_until = datetime.now() + timedelta(minutes=lockout_minutes)
            return True, f"Trop de tentatives. Compte verrouillé pour {lockout_minutes} minutes"
        
        remaining = st.session_state.security_config['max_login_attempts'] - st.session_state.login_attempts
        return False, f"Identifiants incorrects. Il vous reste {remaining} tentative(s)"

# ======================== GESTIONNAIRE DE BASE DE DONNÉES ========================
class DatabaseHandler:
    """Gestionnaire de base de données"""
    
    def __init__(self, db_file=DB_FILE):
        self.db_file = db_file
        self.init_database()
    
    def init_database(self):
        """Initialise la base de données"""
        try:
            conn = sqlite3.connect(self.db_file)
            cursor = conn.cursor()
            
            # Table des utilisateurs
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT UNIQUE NOT NULL,
                    password TEXT NOT NULL,
                    email TEXT,
                    role TEXT DEFAULT 'user',
                    permissions TEXT,
                    status TEXT DEFAULT 'active',
                    last_login TIMESTAMP,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)
            
            # Table des logs
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS logs (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    username TEXT,
                    action TEXT NOT NULL,
                    details TEXT,
                    ip_address TEXT,
                    user_agent TEXT
                )
            """)
            
            # Table de l'historique
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    username TEXT NOT NULL,
                    action_type TEXT NOT NULL,
                    target_user TEXT,
                    details TEXT,
                    data TEXT
                )
            """)
            
            # Table des paramètres
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_by TEXT
                )
            """)
            
            # Créer un admin par défaut si nécessaire
            cursor.execute("SELECT COUNT(*) FROM users WHERE role='admin'")
            if cursor.fetchone()[0] == 0:
                default_password = SecurityManager.hash_password("Admin123!")
                cursor.execute("""
                    INSERT INTO users (username, password, email, role, permissions, status)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, ("admin", default_password, "admin@agc-vie.com", "admin", "all", "active"))
            
            conn.commit()
            conn.close()
            
        except Exception as e:
            st.error(f"Erreur d'initialisation de la base de données: {str(e)}")
    
    def execute_query(self, query, params=()):
        """Exécute une requête SQL"""
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute(query, params)
        conn.commit()
        conn.close()
        return cursor
    
    def fetch_all(self, query, params=()):
        """Récupère tous les résultats d'une requête"""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute(query, params)
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]
    
    def fetch_one(self, query, params=()):
        """Récupère un seul résultat"""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute(query, params)
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

# ======================== FONCTIONS D'AUTHENTIFICATION ========================
def login(username, password):
    """Authentifie un utilisateur"""
    try:
        # Vérifier les tentatives
        can_login, message = SecurityManager.check_login_attempts()
        if not can_login:
            st.error(message)
            return False
        
        db = DatabaseHandler()
        user = db.fetch_one(
            "SELECT * FROM users WHERE username = ? AND status = 'active'",
            (username,)
        )
        
        if user and SecurityManager.verify_password(user['password'], password):
            # Succès
            st.session_state.authenticated = True
            st.session_state.username = username
            st.session_state.role = user['role']
            st.session_state.login_attempts = 0
            st.session_state.locked_until = None
            st.session_state.last_activity = datetime.now()
            
            # Mettre à jour la dernière connexion
            db.execute_query(
                "UPDATE users SET last_login = ? WHERE username = ?",
                (datetime.now(), username)
            )
            
            log_action("Connexion", f"Utilisateur {username} connecté")
            log_history("login", username, "Connexion réussie")
            
            return True
        else:
            # Échec
            is_locked, message = SecurityManager.record_failed_attempt()
            if is_locked:
                st.error(message)
            else:
                st.warning(message)
            
            log_action("Échec connexion", f"Tentative pour {username}", level="warning")
            return False
            
    except Exception as e:
        log_action("Erreur connexion", str(e), level="error")
        st.error(f"Erreur de connexion: {str(e)}")
        return False

def logout():
    """Déconnecte l'utilisateur"""
    if st.session_state.authenticated:
        username = st.session_state.username
        log_action("Déconnexion", f"Utilisateur {username} déconnecté")
        log_history("logout", username, "Déconnexion")
    
    st.session_state.authenticated = False
    st.session_state.username = None
    st.session_state.role = None
    st.rerun()

def check_session_timeout():
    """Vérifie si la session a expiré"""
    if st.session_state.authenticated:
        timeout = st.session_state.security_config.get('session_timeout', 30) * 60
        last_activity = st.session_state.last_activity
        now = datetime.now()
        
        if (now - last_activity).seconds > timeout:
            log_action("Session expirée", f"Utilisateur {st.session_state.username}")
            logout()
            st.warning("Votre session a expiré. Veuillez vous reconnecter.")
            return True
    
    return False

def update_last_activity():
    """Met à jour le timestamp de dernière activité"""
    if st.session_state.authenticated:
        st.session_state.last_activity = datetime.now()

# ======================== FONCTIONS DE GESTION DES UTILISATEURS ========================
def get_all_users():
    """Récupère tous les utilisateurs"""
    db = DatabaseHandler()
    return db.fetch_all("SELECT id, username, email, role, status, last_login, created_at FROM users ORDER BY username")

def add_user(username, password, email, role="user"):
    """Ajoute un nouvel utilisateur"""
    try:
        # Valider le mot de passe
        valid, errors = SecurityManager.validate_password_strength(password)
        if not valid:
            return False, "\n".join(errors)
        
        db = DatabaseHandler()
        
        # Vérifier si l'utilisateur existe déjà
        existing = db.fetch_one("SELECT username FROM users WHERE username = ?", (username,))
        if existing:
            return False, "Ce nom d'utilisateur existe déjà"
        
        # Ajouter l'utilisateur
        hashed_password = SecurityManager.hash_password(password)
        db.execute_query(
            "INSERT INTO users (username, password, email, role, status) VALUES (?, ?, ?, ?, ?)",
            (username, hashed_password, email, role, 'active')
        )
        
        log_action("Ajout utilisateur", f"Utilisateur {username} ajouté")
        log_history("user_add", username, f"Ajouté par {st.session_state.username}")
        
        return True, "Utilisateur ajouté avec succès"
        
    except Exception as e:
        log_action("Erreur ajout utilisateur", str(e), level="error")
        return False, f"Erreur: {str(e)}"

def update_user(username, data):
    """Met à jour un utilisateur"""
    try:
        db = DatabaseHandler()
        
        updates = []
        params = []
        
        for key, value in data.items():
            if key == 'password' and value:
                valid, errors = SecurityManager.validate_password_strength(value)
                if not valid:
                    return False, "\n".join(errors)
                updates.append(f"{key} = ?")
                params.append(SecurityManager.hash_password(value))
            elif key != 'password' and value is not None:
                updates.append(f"{key} = ?")
                params.append(value)
        
        if updates:
            params.append(username)
            query = f"UPDATE users SET {', '.join(updates)} WHERE username = ?"
            db.execute_query(query, params)
            
            log_action("Modification utilisateur", f"Utilisateur {username} modifié")
            log_history("user_update", username, f"Modifié par {st.session_state.username}")
        
        return True, "Utilisateur mis à jour avec succès"
        
    except Exception as e:
        log_action("Erreur modification utilisateur", str(e), level="error")
        return False, f"Erreur: {str(e)}"

def delete_user(username):
    """Supprime un utilisateur"""
    try:
        if username == "admin":
            return False, "Impossible de supprimer le compte admin"
        
        db = DatabaseHandler()
        db.execute_query("DELETE FROM users WHERE username = ?", (username,))
        
        log_action("Suppression utilisateur", f"Utilisateur {username} supprimé")
        log_history("user_delete", username, f"Supprimé par {st.session_state.username}")
        
        return True, "Utilisateur supprimé avec succès"
        
    except Exception as e:
        log_action("Erreur suppression utilisateur", str(e), level="error")
        return False, f"Erreur: {str(e)}"

# ======================== FONCTIONS DE TRAITEMENT DES DONNÉES ========================
def process_technique_data(df):
    """Traite les données techniques"""
    try:
        # Copie pour éviter les modifications sur l'original
        df = df.copy()
        
        # Nettoyage des noms de colonnes
        df.columns = df.columns.str.strip()
        
        # Ajout de la colonne Nouvelle_Police si nécessaire
        if all(col in df.columns for col in ['Num avenant', 'Code intermédiaire', 'N° police']):
            df['Nouvelle_Police'] = df.apply(
                lambda row: f"{row['Code intermédiaire']}-{row['N° police']}/{row['Num avenant']}" 
                if pd.notnull(row['Num avenant']) and str(row['Num avenant']).strip() 
                else f"{row['Code intermédiaire']}-{row['N° police']}", 
                axis=1
            )
        
        # Nettoyage de la colonne police
        if 'Nouvelle_Police' in df.columns:
            df['Nouvelle_Police'] = df['Nouvelle_Police'].astype(str).str.replace('.0', '', regex=False)
        
        # Calcul des ristournes et émissions
        if 'Type quittance' in df.columns and 'Chiffre affaire' in df.columns:
            df['Ristournes'] = df.apply(
                lambda row: row['Chiffre affaire'] if str(row['Type quittance']).strip() == 'Ristourne' else 0, 
                axis=1
            )
            df['Emissions'] = df.apply(
                lambda row: row['Chiffre affaire'] if str(row['Type quittance']).strip() == 'Emission' else 0, 
                axis=1
            )
        
        # Tableau croisé dynamique
        index_col = 'Nouvelle_Police' if 'Nouvelle_Police' in df.columns else df.columns[0]
        value_cols = []
        
        for col in ['Emissions', 'Ristournes', 'Chiffre affaire']:
            if col in df.columns:
                value_cols.append(col)
        
        if not value_cols:
            value_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        
        if value_cols and index_col in df.columns:
            # Convertir les colonnes numériques
            for col in value_cols:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            pivot_df = pd.pivot_table(
                df,
                index=[index_col],
                values=value_cols,
                aggfunc='sum',
                fill_value=0
            ).reset_index()
        else:
            pivot_df = df
        
        log_action("Traitement technique", f"{len(df)} enregistrements traités")
        return pivot_df
        
    except Exception as e:
        log_action("Erreur traitement technique", str(e), level="error")
        st.error(f"Erreur lors du traitement: {str(e)}")
        return df

def process_comptable_data(df):
    """Traite les données comptables"""
    try:
        # Copie pour éviter les modifications sur l'original
        df = df.copy()
        
        # Nettoyage des noms de colonnes
        df.columns = df.columns.str.strip()
        
        # Nettoyage des colonnes de police
        if 'No Police' in df.columns:
            df['No Police'] = df['No Police'].astype(str).str.replace('.0', '', regex=False)
        
        # Conversion numérique
        numeric_cols = ['Débit', 'Crédit', 'Montant']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # Tableau croisé dynamique
        index_col = 'No Police' if 'No Police' in df.columns else df.columns[0]
        value_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        
        if value_cols and index_col in df.columns:
            pivot_df = pd.pivot_table(
                df,
                index=[index_col],
                values=value_cols,
                aggfunc='sum',
                fill_value=0
            ).reset_index()
        else:
            pivot_df = df
        
        log_action("Traitement comptable", f"{len(df)} enregistrements traités")
        return pivot_df
        
    except Exception as e:
        log_action("Erreur traitement comptable", str(e), level="error")
        st.error(f"Erreur lors du traitement: {str(e)}")
        return df

def detect_duplicates(df, police_column='NUMERO POLICE'):
    """Détecte les doublons dans un DataFrame"""
    try:
        if police_column not in df.columns:
            # Chercher une colonne qui pourrait contenir les polices
            for col in df.columns:
                if 'police' in col.lower() or 'num' in col.lower():
                    police_column = col
                    break
            else:
                return None, None, "Colonne de police non trouvée"
        
        # Identifier les doublons
        duplicates_mask = df.duplicated(subset=[police_column], keep=False)
        duplicates_df = df[duplicates_mask].sort_values(police_column)
        uniques_df = df[~duplicates_mask]
        
        stats = {
            'total': len(df),
            'duplicates': len(duplicates_df),
            'uniques': len(uniques_df),
            'duplicate_polices': df[police_column].duplicated().sum()
        }
        
        log_action("Détection doublons", f"{stats['duplicate_polices']} polices en doublon")
        return duplicates_df, uniques_df, stats
        
    except Exception as e:
        log_action("Erreur détection doublons", str(e), level="error")
        return None, None, str(e)

def rapprochement_technique_comptable(tech_df, compta_df):
    """Effectue le rapprochement entre données techniques et comptables"""
    try:
        # Copie des DataFrames
        tech = tech_df.copy()
        compta = compta_df.copy()
        
        # Nettoyage des noms de colonnes
        tech.columns = tech.columns.str.strip()
        compta.columns = compta.columns.str.strip()
        
        # Déterminer les colonnes de jointure
        tech_col = 'Nouvelle_Police' if 'Nouvelle_Police' in tech.columns else tech.columns[0]
        compta_col = 'No Police' if 'No Police' in compta.columns else compta.columns[0]
        
        # Conversion des colonnes en string pour la jointure
        tech[tech_col] = tech[tech_col].astype(str).str.strip()
        compta[compta_col] = compta[compta_col].astype(str).str.strip()
        
        # Renommer pour la fusion
        tech_renamed = tech.rename(columns={tech_col: 'Police'})
        compta_renamed = compta.rename(columns={compta_col: 'Police'})
        
        # Conversion numérique des colonnes pertinentes
        numeric_cols = ['Emissions', 'Ristournes', 'Débit', 'Crédit', 'Chiffre affaire']
        for col in numeric_cols:
            if col in tech_renamed.columns:
                tech_renamed[col] = pd.to_numeric(tech_renamed[col], errors='coerce').fillna(0)
            if col in compta_renamed.columns:
                compta_renamed[col] = pd.to_numeric(compta_renamed[col], errors='coerce').fillna(0)
        
        # Fusion
        merged = pd.merge(
            tech_renamed, 
            compta_renamed, 
            on='Police', 
            how='outer', 
            suffixes=('_tech', '_compta')
        )
        
        # Remplir les NaN
        merged = merged.fillna(0)
        
        # Calcul des écarts - Utiliser les noms de colonnes avec suffixes
        if 'Emissions_tech' in merged.columns and 'Débit_compta' in merged.columns and 'Crédit_compta' in merged.columns:
            merged['CA_Technique'] = merged['Emissions_tech'] + merged.get('Ristournes_tech', 0)
            merged['CA_Comptable'] = abs(merged['Crédit_compta'] - merged['Débit_compta'])
            merged['Écart'] = merged['CA_Technique'] - merged['CA_Comptable']
            merged['Statut'] = merged['Écart'].apply(
                lambda x: 'Rapproché' if abs(x) < 0.01 else 'Non rapproché'
            )
        else:
            # Fallback si les colonnes n'existent pas
            merged['CA_Technique'] = 0
            merged['CA_Comptable'] = 0
            merged['Écart'] = 0
            merged['Statut'] = 'Non rapproché'
        
        # Statistiques
        stats = {
            'total_polices': len(merged),
            'polices_techniques': len(tech_renamed),
            'polices_comptables': len(compta_renamed),
            'polices_communes': len(merged[(merged['Emissions_tech'] != 0) & (merged['Débit_compta'] != 0)]) if 'Emissions_tech' in merged.columns and 'Débit_compta' in merged.columns else 0,
            'polices_tech_only': len(merged[(merged['Emissions_tech'] != 0) & (merged['Débit_compta'] == 0)]) if 'Emissions_tech' in merged.columns and 'Débit_compta' in merged.columns else 0,
            'polices_compta_only': len(merged[(merged['Emissions_tech'] == 0) & (merged['Débit_compta'] != 0)]) if 'Emissions_tech' in merged.columns and 'Débit_compta' in merged.columns else 0
        }
        
        if 'Statut' in merged.columns:
            stats['rapprochees'] = len(merged[merged['Statut'] == 'Rapproché'])
            stats['non_rapprochees'] = len(merged[merged['Statut'] == 'Non rapproché'])
            stats['ecart_total'] = merged['Écart'].sum() if 'Écart' in merged.columns else 0
        
        log_action("Rapprochement", f"{stats['polices_communes']} polices communes")
        return merged, stats
        
    except Exception as e:
        log_action("Erreur rapprochement", str(e), level="error")
        return None, str(e)

def validate_references(df, ref_column='Réf Pièce'):
    """Valide les références selon un pattern"""
    try:
        if ref_column not in df.columns:
            return None, "Colonne de référence non trouvée"
        
        # Pattern pour validation
        pattern = r"^\w+-\d+(?:/\d+)?$"
        
        valid_refs = []
        invalid_refs = []
        
        for ref in df[ref_column].dropna():
            ref_str = str(ref).strip()
            if re.match(pattern, ref_str):
                valid_refs.append(ref_str)
            else:
                invalid_refs.append(ref_str)
        
        stats = {
            'total': len(df),
            'valides': len(valid_refs),
            'invalides': len(invalid_refs)
        }
        
        return invalid_refs, stats
        
    except Exception as e:
        log_action("Erreur validation références", str(e), level="error")
        return None, str(e)

# ======================== FONCTIONS D'EXPORT ========================
def export_to_excel(dataframes, sheet_names, filename="export.xlsx"):
    """Exporte plusieurs DataFrames vers un fichier Excel"""
    try:
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for df, sheet_name in zip(dataframes, sheet_names):
                if df is not None and not df.empty:
                    # Limiter le nom de la feuille à 31 caractères
                    safe_name = sheet_name[:31]
                    df.to_excel(writer, sheet_name=safe_name, index=False)
        
        output.seek(0)
        
        log_action("Export Excel", f"{filename} créé")
        return output
        
    except Exception as e:
        log_action("Erreur export Excel", str(e), level="error")
        return None

def export_to_csv(df, filename="export.csv"):
    """Exporte un DataFrame vers CSV"""
    try:
        return df.to_csv(index=False).encode('utf-8')
    except Exception as e:
        log_action("Erreur export CSV", str(e), level="error")
        return None

def create_download_button(data, filename, button_text, mime_type=None):
    """Crée un bouton de téléchargement"""
    if mime_type is None:
        if filename.endswith('.csv'):
            mime_type = 'text/csv'
        elif filename.endswith('.xlsx'):
            mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        else:
            mime_type = 'application/octet-stream'
    
    return st.download_button(
        label=button_text,
        data=data,
        file_name=filename,
        mime=mime_type,
        use_container_width=True
    )

# ======================== COMPOSANTS D'INTERFACE ========================
def display_metric_card(title, value, icon="📊", description="", color="primary"):
    """Affiche une carte métrique stylisée avec différentes couleurs"""
    colors = {
        "primary": "linear-gradient(135deg, #667eea 0%, #764ba2 100%)",
        "success": "linear-gradient(135deg, #28a745 0%, #20c997 100%)",
        "danger": "linear-gradient(135deg, #dc3545 0%, #c82333 100%)",
        "warning": "linear-gradient(135deg, #ffc107 0%, #fd7e14 100%)",
        "info": "linear-gradient(135deg, #17a2b8 0%, #138496 100%)",
        "dark": "linear-gradient(135deg, #1e3c72 0%, #2a5298 100%)"
    }
    
    bg_color = colors.get(color, colors["primary"])
    
    st.markdown(f"""
    <div class="metric-card" style="background: {bg_color}; padding: 15px; border-radius: 12px; 
         color: white; text-align: center; box-shadow: 0 4px 15px rgba(0,0,0,0.2);
         margin: 5px 0; transition: transform 0.3s;">
        <div style="font-size: 0.9em; opacity: 0.9; margin-bottom: 5px;">
            {icon} {title}
        </div>
        <div style="font-size: 1.8em; font-weight: bold; margin: 5px 0;">
            {value}
        </div>
        <div style="font-size: 0.8em; opacity: 0.8;">
            {description}
        </div>
    </div>
    """, unsafe_allow_html=True)


def display_metrics_row(metrics, cols=4):
    """Affiche une ligne de métriques dans une grille"""
    cols_list = st.columns(cols)
    for i, metric in enumerate(metrics):
        with cols_list[i % cols]:
            display_metric_card(
                title=metric.get('title', ''),
                value=metric.get('value', ''),
                icon=metric.get('icon', '📊'),
                description=metric.get('description', ''),
                color=metric.get('color', 'primary')
            )

def display_badge(text, type="info"):
    """Affiche un badge stylisé"""
    colors = {
        "success": "badge-success",
        "warning": "badge-warning",
        "danger": "badge-danger",
        "info": "badge-info"
    }
    css_class = colors.get(type, "badge-info")
    st.markdown(f'<span class="badge {css_class}">{text}</span>', unsafe_allow_html=True)

def create_search_bar(key, placeholder="Rechercher..."):
    """Crée une barre de recherche"""
    search = st.text_input(
        "🔍",
        placeholder=placeholder,
        key=key,
        label_visibility="collapsed"
    )
    return search

def filter_dataframe(df, search_term):
    """Filtre un DataFrame selon un terme de recherche"""
    if not search_term or df is None or df.empty:
        return df
    
    try:
        mask = df.astype(str).apply(
            lambda x: x.str.contains(search_term, case=False, na=False)
        ).any(axis=1)
        return df[mask]
    except:
        return df

def create_pagination(df, key_prefix, items_per_page=None):
    """Crée une pagination pour un DataFrame"""
    if items_per_page is None:
        items_per_page = st.session_state.user_preferences.get('items_per_page', 50)
    
    if df is None or df.empty:
        return df, 0, 0
    
    total_pages = (len(df) + items_per_page - 1) // items_per_page
    
    if f'{key_prefix}_page' not in st.session_state:
        st.session_state[f'{key_prefix}_page'] = 0
    
    current_page = st.session_state[f'{key_prefix}_page']
    
    # Contrôles de pagination
    if total_pages > 1:
        col1, col2, col3, col4, col5 = st.columns([1, 1, 2, 1, 1])
        
        with col1:
            if st.button("◀◀", key=f"{key_prefix}_first", disabled=current_page == 0):
                st.session_state[f'{key_prefix}_page'] = 0
                st.rerun()
        
        with col2:
            if st.button("◀", key=f"{key_prefix}_prev", disabled=current_page == 0):
                st.session_state[f'{key_prefix}_page'] -= 1
                st.rerun()
        
        with col3:
            st.markdown(f"<center>Page {current_page + 1} / {total_pages}</center>", unsafe_allow_html=True)
        
        with col4:
            if st.button("▶", key=f"{key_prefix}_next", disabled=current_page >= total_pages - 1):
                st.session_state[f'{key_prefix}_page'] += 1
                st.rerun()
        
        with col5:
            if st.button("▶▶", key=f"{key_prefix}_last", disabled=current_page >= total_pages - 1):
                st.session_state[f'{key_prefix}_page'] = total_pages - 1
                st.rerun()
    
    # Extraire la page courante
    start_idx = current_page * items_per_page
    end_idx = min(start_idx + items_per_page, len(df))
    
    return df.iloc[start_idx:end_idx], start_idx, end_idx

# ======================== PAGES DE L'APPLICATION ========================

def page_login():
    """Page de connexion"""
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        with st.form("login_form", clear_on_submit=True):
            st.markdown("### Connexion")
            
            username = st.text_input(
                "👤 Nom d'utilisateur",
                placeholder="Entrez votre nom d'utilisateur",
                help="Votre nom d'utilisateur AGC-VIE"
            )
            
            password = st.text_input(
                "🔒 Mot de passe",
                type="password",
                placeholder="Entrez votre mot de passe",
                help="Votre mot de passe sécurisé"
            )
            
            submitted = st.form_submit_button(
                "Se connecter",
                use_container_width=True,
                type="primary"
            )
            
            if submitted:
                if username and password:
                    if login(username, password):
                        st.success("Connexion réussie!")
                        st.balloons()
                        time.sleep(1)
                        st.rerun()
                else:
                    st.warning("Veuillez remplir tous les champs")
        
        st.markdown("""
        <div style="text-align: center; color: #666; font-size: 0.9em; margin-top: 20px;">
            <!--p>Compte démo: admin / Admin123!</p-->
            <p>© 2025 AGC-VIE - Tous droits réservés</p>
        </div>
        """, unsafe_allow_html=True)

def page_accueil():
    """Page d'accueil"""
    update_last_activity()
    
    st.markdown("""
    <div class="content-card fade-in" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
         padding: 30px; border-radius: 15px; text-align: center; margin-bottom: 30px;">
        <h1 style="color: white; font-size: 2.5em; margin-bottom: 10px;">AGC-VIE</h1>
        <p style="color: white; font-size: 1.2em; opacity: 0.9;">
            Système intégré de Gestion Technique et Comptable
        </p>
        <p style="color: white; font-size: 0.9em; opacity: 0.7; margin-top: 10px;">
            Version 2.0
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Métriques principales
    tech_count = len(st.session_state.pivot_techniques) if st.session_state.pivot_techniques is not None else 0
    compta_count = len(st.session_state.pivot_comptables) if st.session_state.pivot_comptables is not None else 0
    
    ca_tech = 0
    if st.session_state.pivot_techniques is not None and 'Emissions' in st.session_state.pivot_techniques.columns:
        ca_tech = st.session_state.pivot_techniques['Emissions'].sum()
    
    ca_compta = 0
    if st.session_state.pivot_comptables is not None and 'Crédit' in st.session_state.pivot_comptables.columns:
        ca_compta = st.session_state.pivot_comptables['Crédit'].sum()
    
    metrics = [
        {
            'title': 'Données techniques',
            'value': f'{tech_count:,}',
            'icon': '📊',
            'description': 'Enregistrements',
            'color': 'primary'
        },
        {
            'title': 'Données comptables',
            'value': f'{compta_count:,}',
            'icon': '💰',
            'description': 'Enregistrements',
            'color': 'success'
        },
        {
            'title': 'CA Technique',
            'value': f'{ca_tech:,.0f} FCFA',
            'icon': '📈',
            'description': 'Chiffre d\'affaires',
            'color': 'info'
        },
        {
            'title': 'CA Comptable',
            'value': f'{ca_compta:,.0f} FCFA',
            'icon': '📉',
            'description': 'Chiffre d\'affaires',
            'color': 'warning'
        }
    ]
    
    display_metrics_row(metrics, cols=4)
    
    # Statistiques rapides
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        <div style="background: white; padding: 20px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
            <h4 style="color: #1e3c72; margin-bottom: 15px;">📋 Activités</h4>
        """, unsafe_allow_html=True)
        
        if st.session_state.logs:
            logs_df = pd.DataFrame(st.session_state.logs[-5:])
            if not logs_df.empty:
                for _, row in logs_df.iterrows():
                    st.markdown(f"""
                    <div style="padding: 8px 0; border-bottom: 1px solid #eee; font-size: 0.9em;">
                        <span style="color: #666;">{row['timestamp'][:16]}</span>
                        <span style="color: #1e3c72; font-weight: 500;">{row['action']}</span>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.info("Aucune activité récente")
        st.markdown("</div>", unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div style="background: white; padding: 20px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
            <h4 style="color: #1e3c72; margin-bottom: 15px;">📊 Statistiques</h4>
        """, unsafe_allow_html=True)
        
        st.markdown(f"""
        </br>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px;">
            <div style="background: #f8f9fa; padding: 10px; border-radius: 8px; text-align: center;">
                <div style="font-size: 1.5em; font-weight: bold; color: #1e3c72;">
                    {st.session_state.stats.get('total_imports', 0)}
                </div>
                <div style="font-size: 0.8em; color: #666;">Imports</div>
            </div>
            <div style="background: #f8f9fa; padding: 10px; border-radius: 8px; text-align: center;">
                <div style="font-size: 1.5em; font-weight: bold; color: #28a745;">
                    {st.session_state.stats.get('total_verifications', 0)}
                </div>
                <div style="font-size: 0.8em; color: #666;">Vérifications</div>
            </div>
            <div style="background: #f8f9fa; padding: 10px; border-radius: 8px; text-align: center;">
                <div style="font-size: 1.5em; font-weight: bold; color: #17a2b8;">
                    {st.session_state.stats.get('total_certificats', 0)}
                </div>
                <div style="font-size: 0.8em; color: #666;">Certificats</div>
            </div>
            <div style="background: #f8f9fa; padding: 10px; border-radius: 8px; text-align: center;">
                <div style="font-size: 1.5em; font-weight: bold; color: #fd7e14;">
                    {len(get_all_users()) if get_all_users() else 0}
                </div>
                <div style="font-size: 0.8em; color: #666;">Utilisateurs</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div style="background: white; padding: 20px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
            <h4 style="color: #1e3c72; margin-bottom: 15px;">🚀 Actions rapides</h4>
        """, unsafe_allow_html=True)
        
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("📊 Importer technique", use_container_width=True):
                st.session_state.page = "Gestion Technique"
                st.rerun()
            if st.button("💰 Importer comptable", use_container_width=True):
                st.session_state.page = "Gestion Comptable"
                st.rerun()
        with col_b:
            if st.button("🔄 Rapprochement", use_container_width=True):
                st.session_state.page = "Rapprochement Technique"
                st.rerun()
            if st.button("📈 Statistiques", use_container_width=True):
                st.session_state.page = "Statistiques"
                st.rerun()
        
        st.markdown("</div>", unsafe_allow_html=True)

def page_gestion_technique():
    """Page de gestion technique"""
    update_last_activity()
    
    st.markdown("## 📊 Gestion Technique")
    
    # Onglets
    tab1, tab2, tab3 = st.tabs(["📥 Import", "📋 Données", "📈 Analyses"])
    
    with tab1:
        st.markdown("### Importer des données techniques")
        
        uploaded_file = st.file_uploader(
            "Choisir un fichier Excel ou CSV",
            type=['xlsx', 'xls', 'csv'],
            key="tech_upload",
            help="Formats supportés: Excel (.xlsx, .xls) et CSV (.csv)"
        )
        
        if uploaded_file:
            with st.spinner("Chargement du fichier en cours..."):
                try:
                    # Lecture du fichier
                    if uploaded_file.name.endswith('.csv'):
                        df = pd.read_csv(uploaded_file, dtype=str)
                    else:
                        df = pd.read_excel(uploaded_file, dtype=str)
                    
                    st.session_state.df_technique = df
                    
                    # Aperçu
                    st.markdown("### Aperçu des données")
                    st.dataframe(df.head(10), use_container_width=True)
                    
                    # Statistiques
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Lignes", len(df))
                    with col2:
                        st.metric("Colonnes", len(df.columns))
                    with col3:
                        st.metric("Taille", f"{uploaded_file.size / 1024:.1f} KB")
                    
                    # Traitement
                    if st.button("🔄 Traiter les données", type="primary", use_container_width=True):
                        with st.spinner("Traitement en cours..."):
                            pivot_df = process_technique_data(df)
                            st.session_state.pivot_techniques = pivot_df
                            st.session_state.stats['total_imports'] += 1
                            
                            st.success(f"Traitement terminé! {len(pivot_df)} enregistrements générés.")
                            st.balloons()
                            
                            log_action("Import technique", f"{len(df)} lignes importées")
                    
                except Exception as e:
                    st.error(f"Erreur lors du chargement: {str(e)}")
    
    with tab2:
        if st.session_state.pivot_techniques is not None:
            st.markdown("### Données techniques traitées")
            
            # Recherche
            search = create_search_bar("tech_search", "Rechercher une police...")
            
            # Filtrage
            df_display = st.session_state.pivot_techniques.copy()
            if search:
                df_display = filter_dataframe(df_display, search)
            
            # Pagination
            df_page, start, end = create_pagination(df_display, "tech")
            
            st.markdown(f"**Affichage {start+1}-{end} sur {len(df_display)} enregistrements**")
            
            # Affichage
            st.dataframe(
                df_page,
                use_container_width=True,
                height=500,
                hide_index=True
            )
            
            # Export
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("📥 Exporter en Excel", use_container_width=True):
                    output = export_to_excel(
                        [df_display],
                        ["Données techniques"],
                        "donnees_techniques.xlsx"
                    )
                    if output:
                        create_download_button(
                            output,
                            "donnees_techniques.xlsx",
                            "Télécharger Excel"
                        )
            
            with col2:
                if st.button("📥 Exporter en CSV", use_container_width=True):
                    csv_data = export_to_csv(df_display, "donnees_techniques.csv")
                    if csv_data:
                        create_download_button(
                            csv_data,
                            "donnees_techniques.csv",
                            "Télécharger CSV"
                        )
        else:
            st.info("Aucune donnée technique. Veuillez d'abord importer et traiter des données.")
    
    with tab3:
        if st.session_state.pivot_techniques is not None:
            st.markdown("### Analyses statistiques")
            
            df = st.session_state.pivot_techniques
            
            # Sélection des colonnes numériques
            numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            
            if numeric_cols:
                col1, col2 = st.columns(2)
                
                with col1:
                    selected_col = st.selectbox(
                        "Sélectionner une colonne à analyser",
                        numeric_cols
                    )
                
                with col2:
                    chart_type = st.selectbox(
                        "Type de graphique",
                        ["Histogramme", "Box plot", "Courbe"]
                    )
                
                # Statistiques descriptives
                st.markdown("#### Statistiques descriptives")
                stats_df = df[selected_col].describe().reset_index()
                stats_df.columns = ['Statistique', 'Valeur']
                st.dataframe(stats_df, use_container_width=True, hide_index=True)
                
                # Graphique
                st.markdown("#### Visualisation")
                
                if chart_type == "Histogramme":
                    fig = px.histogram(
                        df,
                        x=selected_col,
                        nbins=30,
                        title=f"Distribution de {selected_col}",
                        color_discrete_sequence=['#1e3c72']
                    )
                elif chart_type == "Box plot":
                    fig = px.box(
                        df,
                        y=selected_col,
                        title=f"Box plot de {selected_col}",
                        color_discrete_sequence=['#1e3c72']
                    )
                else:
                    fig = px.line(
                        df.reset_index(),
                        y=selected_col,
                        title=f"Évolution de {selected_col}",
                        color_discrete_sequence=['#1e3c72']
                    )
                
                fig.update_layout(height=500)
                st.plotly_chart(fig, use_container_width=True)
                
                # Top valeurs
                st.markdown("#### Top 10 des valeurs")
                top_df = df.nlargest(10, selected_col)[df.columns[:3]]
                st.dataframe(top_df, use_container_width=True, hide_index=True)
                
            else:
                st.warning("Aucune colonne numérique disponible pour l'analyse")
        else:
            st.info("Aucune donnée à analyser")

def page_gestion_comptable():
    """Page de gestion comptable"""
    update_last_activity()
    
    st.markdown("## 💰 Gestion Comptable")
    
    # Onglets
    tab1, tab2, tab3 = st.tabs(["📥 Import", "📋 Données", "📈 Analyses"])
    
    with tab1:
        st.markdown("### Importer des données comptables")
        
        uploaded_file = st.file_uploader(
            "Choisir un fichier Excel ou CSV",
            type=['xlsx', 'xls', 'csv'],
            key="compta_upload",
            help="Formats supportés: Excel (.xlsx, .xls) et CSV (.csv)"
        )
        
        if uploaded_file:
            with st.spinner("Chargement du fichier en cours..."):
                try:
                    # Lecture du fichier
                    if uploaded_file.name.endswith('.csv'):
                        df = pd.read_csv(uploaded_file, dtype=str)
                    else:
                        df = pd.read_excel(uploaded_file, dtype=str)
                    
                    st.session_state.df_comptable = df
                    
                    # Aperçu
                    st.markdown("### Aperçu des données")
                    st.dataframe(df.head(10), use_container_width=True)
                    
                    # Statistiques
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Lignes", len(df))
                    with col2:
                        st.metric("Colonnes", len(df.columns))
                    with col3:
                        st.metric("Taille", f"{uploaded_file.size / 1024:.1f} KB")
                    
                    # Traitement
                    if st.button("🔄 Traiter les données", type="primary", use_container_width=True):
                        with st.spinner("Traitement en cours..."):
                            pivot_df = process_comptable_data(df)
                            st.session_state.pivot_comptables = pivot_df
                            st.session_state.stats['total_imports'] += 1
                            
                            st.success(f"Traitement terminé! {len(pivot_df)} enregistrements générés.")
                            st.balloons()
                            
                            log_action("Import comptable", f"{len(df)} lignes importées")
                    
                except Exception as e:
                    st.error(f"Erreur lors du chargement: {str(e)}")
    
    with tab2:
        if st.session_state.pivot_comptables is not None:
            st.markdown("### Données comptables traitées")
            
            # Recherche
            search = create_search_bar("compta_search", "Rechercher une police...")
            
            # Filtrage
            df_display = st.session_state.pivot_comptables.copy()
            if search:
                df_display = filter_dataframe(df_display, search)
            
            # Pagination
            df_page, start, end = create_pagination(df_display, "compta")
            
            st.markdown(f"**Affichage {start+1}-{end} sur {len(df_display)} enregistrements**")
            
            # Affichage
            st.dataframe(
                df_page,
                use_container_width=True,
                height=500,
                hide_index=True
            )
            
            # Export
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("📥 Exporter en Excel", use_container_width=True):
                    output = export_to_excel(
                        [df_display],
                        ["Données comptables"],
                        "donnees_comptables.xlsx"
                    )
                    if output:
                        create_download_button(
                            output,
                            "donnees_comptables.xlsx",
                            "Télécharger Excel"
                        )
            
            with col2:
                if st.button("📥 Exporter en CSV", use_container_width=True):
                    csv_data = export_to_csv(df_display, "donnees_comptables.csv")
                    if csv_data:
                        create_download_button(
                            csv_data,
                            "donnees_comptables.csv",
                            "Télécharger CSV"
                        )
        else:
            st.info("Aucune donnée comptable. Veuillez d'abord importer et traiter des données.")
    
    with tab3:
        if st.session_state.pivot_comptables is not None:
            st.markdown("### Analyses statistiques")
            
            df = st.session_state.pivot_comptables
            
            # Sélection des colonnes numériques
            numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
            
            if numeric_cols:
                col1, col2 = st.columns(2)
                
                with col1:
                    selected_col = st.selectbox(
                        "Sélectionner une colonne à analyser",
                        numeric_cols
                    )
                
                with col2:
                    chart_type = st.selectbox(
                        "Type de graphique",
                        ["Histogramme", "Box plot", "Courbe"]
                    )
                
                # Statistiques descriptives
                st.markdown("#### Statistiques descriptives")
                stats_df = df[selected_col].describe().reset_index()
                stats_df.columns = ['Statistique', 'Valeur']
                st.dataframe(stats_df, use_container_width=True, hide_index=True)
                
                # Graphique
                st.markdown("#### Visualisation")
                
                if chart_type == "Histogramme":
                    fig = px.histogram(
                        df,
                        x=selected_col,
                        nbins=30,
                        title=f"Distribution de {selected_col}",
                        color_discrete_sequence=['#1e3c72']
                    )
                elif chart_type == "Box plot":
                    fig = px.box(
                        df,
                        y=selected_col,
                        title=f"Box plot de {selected_col}",
                        color_discrete_sequence=['#1e3c72']
                    )
                else:
                    fig = px.line(
                        df.reset_index(),
                        y=selected_col,
                        title=f"Évolution de {selected_col}",
                        color_discrete_sequence=['#1e3c72']
                    )
                
                fig.update_layout(height=500)
                st.plotly_chart(fig, use_container_width=True)
                
                # Solde total
                if 'Débit' in df.columns and 'Crédit' in df.columns:
                    total_debit = df['Débit'].sum()
                    total_credit = df['Crédit'].sum()
                    solde = total_credit - total_debit
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Débit", f"{total_debit:,.0f} FCFA")
                    with col2:
                        st.metric("Total Crédit", f"{total_credit:,.0f} FCFA")
                    with col3:
                        st.metric("Solde", f"{solde:,.0f} FCFA")
                
            else:
                st.warning("Aucune colonne numérique disponible pour l'analyse")
        else:
            st.info("Aucune donnée à analyser")

def page_rapprochement_technique():
    """Page de rapprochement technique"""
    update_last_activity()
    
    st.markdown("## 🔄 Rapprochement Technique")
    
    # Vérification des données
    if st.session_state.pivot_techniques is None:
        st.warning("⚠️ Données techniques manquantes. Veuillez d'abord importer les données techniques.")
        if st.button("📥 Aller à la gestion technique"):
            st.session_state.page = "Gestion Technique"
            st.rerun()
        return
    
    if st.session_state.pivot_comptables is None:
        st.warning("⚠️ Données comptables manquantes. Veuillez d'abord importer les données comptables.")
        if st.button("💰 Aller à la gestion comptable"):
            st.session_state.page = "Gestion Comptable"
            st.rerun()
        return
    
    # Effectuer le rapprochement
    with st.spinner("Calcul du rapprochement en cours..."):
        merged_df, stats = rapprochement_technique_comptable(
            st.session_state.pivot_techniques,
            st.session_state.pivot_comptables
        )
    
    if merged_df is not None:
        st.session_state.stats['total_verifications'] += 1
        
        # Métriques
        st.markdown("### 📊 Résumé du rapprochement")
        
        # Première ligne : Polices
        metrics_row1 = [
            {
                'title': 'Polices techniques',
                'value': f"{stats.get('polices_techniques', 0):,}",
                'icon': '📊',
                'description': f"Total: {stats.get('polices_techniques', 0):,}",
                'color': 'primary'
            },
            {
                'title': 'Polices comptables',
                'value': f"{stats.get('polices_comptables', 0):,}",
                'icon': '💰',
                'description': f"Total: {stats.get('polices_comptables', 0):,}",
                'color': 'success'
            },
            {
                'title': 'Polices communes',
                'value': f"{stats.get('polices_communes', 0):,}",
                'icon': '🔄',
                'description': f"Taux: {(stats.get('polices_communes', 0)/stats.get('polices_techniques', 1)*100):.1f}%",
                'color': 'info'
            },
            {
                'title': 'Écart total',
                'value': f"{stats.get('ecart_total', 0):,.0f} FCFA",
                'icon': '📈',
                'description': 'CA Technique - CA Comptable',
                'color': 'warning' if stats.get('ecart_total', 0) > 0 else 'danger'
            }
        ]
        
        display_metrics_row(metrics_row1, cols=4)
        
        # Deuxième ligne : Statut
        if 'Statut' in merged_df.columns:
            rapprochees = stats.get('rapprochees', 0)
            non_rapprochees = stats.get('non_rapprochees', 0)
            total = rapprochees + non_rapprochees
            
            metrics_row2 = [
                {
                    'title': '✅ Rapprochées',
                    'value': f"{rapprochees:,}",
                    'icon': '✅',
                    'description': f"{(rapprochees/total*100 if total > 0 else 0):.1f}%",
                    'color': 'success'
                },
                {
                    'title': '❌ Non rapprochées',
                    'value': f"{non_rapprochees:,}",
                    'icon': '❌',
                    'description': f"{(non_rapprochees/total*100 if total > 0 else 0):.1f}%",
                    'color': 'danger'
                },
                {
                    'title': '📋 Total polices',
                    'value': f"{total:,}",
                    'icon': '📋',
                    'description': 'Toutes les polices',
                    'color': 'dark'
                },
                {
                    'title': '📊 Taux de rapprochement',
                    'value': f"{(rapprochees/total*100 if total > 0 else 0):.1f}%",
                    'icon': '📊',
                    'description': f"{rapprochees} / {total}",
                    'color': 'info'
                }
            ]
            
            display_metrics_row(metrics_row2, cols=4)
        
        # Tabs pour les différentes vues
        tab1, tab2, tab3 = st.tabs(["📋 Données complètes", "❌ Non rapprochées", "✅ Rapprochées"])
        
        with tab1:
            st.markdown("### Toutes les polices")
            
            # Recherche
            search = create_search_bar("rapprochement_search", "Rechercher une police...")
            
            df_display = merged_df.copy()
            if search:
                df_display = filter_dataframe(df_display, search)
            
            st.dataframe(df_display, use_container_width=True, height=500, hide_index=True)
        
        with tab2:
            if 'Statut' in merged_df.columns:
                non_rapproche = merged_df[merged_df['Statut'] == 'Non rapproché']
                
                st.markdown(f"### Polices non rapprochées ({len(non_rapproche)})")
                
                if not non_rapproche.empty:
                    st.dataframe(non_rapproche, use_container_width=True, height=500, hide_index=True)
                    
                    # Export des non rapprochées
                    if st.button("📥 Exporter les non rapprochées", use_container_width=True):
                        output = export_to_excel(
                            [non_rapproche],
                            ["Non rapprochées"],
                            "polices_non_rapprochees.xlsx"
                        )
                        if output:
                            create_download_button(
                                output,
                                "polices_non_rapprochees.xlsx",
                                "Télécharger"
                            )
                else:
                    st.success("Toutes les polices sont rapprochées !")
        
        with tab3:
            if 'Statut' in merged_df.columns:
                rapproche = merged_df[merged_df['Statut'] == 'Rapproché']
                
                st.markdown(f"### Polices rapprochées ({len(rapproche)})")
                
                if not rapproche.empty:
                    st.dataframe(rapproche, use_container_width=True, height=500, hide_index=True)
        
        # Visualisations
        st.markdown("### 📊 Visualisations")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Graphique des statuts
            if 'Statut' in merged_df.columns:
                status_counts = merged_df['Statut'].value_counts()
                
                fig = px.pie(
                    values=status_counts.values,
                    names=status_counts.index,
                    title="Répartition des statuts",
                    color_discrete_sequence=['#28a745', '#dc3545'],
                    hole=0.3
                )
                fig.update_layout(
                    height=400,
                    showlegend=True,
                    legend=dict(orientation="h", yanchor="bottom", y=-0.2)
                )
                st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            # Graphique des écarts
            if 'Écart' in merged_df.columns:
                fig = px.histogram(
                    merged_df[merged_df['Écart'].notna()],
                    x='Écart',
                    nbins=50,
                    title="Distribution des écarts",
                    color_discrete_sequence=['#1e3c72']
                )
                fig.update_layout(
                    height=400,
                    xaxis_title="Écart (FCFA)",
                    yaxis_title="Nombre de polices"
                )
                st.plotly_chart(fig, use_container_width=True)
        
        # Export complet
        st.markdown("### 📥 Export du rapport")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("📥 Exporter le rapport complet", type="primary", use_container_width=True):
                output = export_to_excel(
                    [merged_df],
                    ["Rapprochement"],
                    "rapprochement_technique_complet.xlsx"
                )
                if output:
                    create_download_button(
                        output,
                        "rapprochement_technique_complet.xlsx",
                        "Télécharger le rapport"
                    )
        
        with col2:
            if 'Statut' in merged_df.columns:
                non_rapproche = merged_df[merged_df['Statut'] == 'Non rapproché']
                if not non_rapproche.empty:
                    if st.button("📥 Exporter les non rapprochées", use_container_width=True):
                        output = export_to_excel(
                            [non_rapproche],
                            ["Non rapprochées"],
                            "polices_non_rapprochees.xlsx"
                        )
                        if output:
                            create_download_button(
                                output,
                                "polices_non_rapprochees.xlsx",
                                "Télécharger"
                            )
        
        with col3:
            if st.button("📊 Exporter les graphiques", use_container_width=True):
                st.info("Fonctionnalité à venir: export des graphiques en PNG")
        
        log_action("Rapprochement technique", f"{len(merged_df)} polices analysées")
        
    else:
        st.error(f"Erreur lors du rapprochement: {stats}")

def page_rapprochement_comptable():
    """Page de rapprochement comptable"""
    update_last_activity()
    
    st.markdown("""
    <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
         padding: 20px; border-radius: 12px; margin-bottom: 25px;">
        <h2 style="color: white; margin: 0;">🔄 Rapprochement Comptable</h2>
        <p style="color: white; opacity: 0.9; margin: 5px 0 0 0;">
            Comparaison entre les données techniques et comptables
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Vérification des données
    if st.session_state.pivot_techniques is None:
        st.warning("⚠️ Données techniques manquantes. Veuillez d'abord importer les données techniques.")
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("📥 Aller à la gestion technique", use_container_width=True):
                st.session_state.page = "Gestion Technique"
                st.rerun()
        return
    
    if st.session_state.pivot_comptables is None:
        st.warning("⚠️ Données comptables manquantes. Veuillez d'abord importer les données comptables.")
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("💰 Aller à la gestion comptable", use_container_width=True):
                st.session_state.page = "Gestion Comptable"
                st.rerun()
        return
    
    # Bouton pour lancer le rapprochement
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("🚀 Lancer le rapprochement comptable", type="primary", use_container_width=True):
            with st.spinner("Calcul du rapprochement comptable en cours..."):
                try:
                    # Récupérer les données
                    df_tech = st.session_state.pivot_techniques.copy()
                    df_compta = st.session_state.pivot_comptables.copy()
                    
                    # Nettoyage des noms de colonnes
                    df_tech.columns = df_tech.columns.str.strip()
                    df_compta.columns = df_compta.columns.str.strip()
                    
                    # Déterminer les colonnes de police
                    tech_col = 'Nouvelle_Police' if 'Nouvelle_Police' in df_tech.columns else df_tech.columns[0]
                    compta_col = 'No Police' if 'No Police' in df_compta.columns else df_compta.columns[0]
                    
                    # Conversion en string pour la jointure
                    df_tech[tech_col] = df_tech[tech_col].astype(str).str.strip()
                    df_compta[compta_col] = df_compta[compta_col].astype(str).str.strip()
                    
                    # Ajouter les colonnes techniques aux données comptables
                    df_compta['Ristournes'] = 0
                    df_compta['Emissions'] = 0
                    df_compta['Statut_Ristournes'] = 'Non trouvé'
                    df_compta['Statut_Emissions'] = 'Non trouvé'
                    
                    # Pour chaque ligne comptable, chercher la correspondance dans les données techniques
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    total_rows = len(df_compta)
                    for index, row in df_compta.iterrows():
                        # Mise à jour de la progression
                        progress = (index + 1) / total_rows
                        progress_bar.progress(progress)
                        status_text.text(f"Traitement: {index+1}/{total_rows} polices")
                        
                        police_compta = str(row[compta_col]).strip()
                        
                        # Chercher la correspondance
                        correspondance = df_tech[df_tech[tech_col].astype(str).str.strip() == police_compta]
                        
                        if not correspondance.empty:
                            # Récupérer les premières valeurs
                            if 'Ristournes' in correspondance.columns:
                                val = correspondance['Ristournes'].iloc[0]
                                df_compta.at[index, 'Ristournes'] = pd.to_numeric(val, errors='coerce') if val != 0 else 0
                                df_compta.at[index, 'Statut_Ristournes'] = 'Trouvé'
                            
                            if 'Emissions' in correspondance.columns:
                                val = correspondance['Emissions'].iloc[0]
                                df_compta.at[index, 'Emissions'] = pd.to_numeric(val, errors='coerce') if val != 0 else 0
                                df_compta.at[index, 'Statut_Emissions'] = 'Trouvé'
                    
                    # Effacer la progression
                    progress_bar.empty()
                    status_text.empty()
                    
                    # Conversion en numérique
                    numeric_cols = ['Crédit', 'Débit', 'Emissions', 'Ristournes', 'Montant']
                    for col in numeric_cols:
                        if col in df_compta.columns:
                            df_compta[col] = pd.to_numeric(df_compta[col], errors='coerce').fillna(0)
                    
                    # Calcul du rapprochement
                    if all(col in df_compta.columns for col in ['Crédit', 'Débit']):
                        df_compta['CA_Comptable'] = df_compta['Crédit'] - df_compta['Débit']
                    else:
                        df_compta['CA_Comptable'] = 0
                    
                    if all(col in df_compta.columns for col in ['Emissions', 'Ristournes']):
                        df_compta['CA_Technique'] = df_compta['Emissions'] + df_compta['Ristournes']
                    else:
                        df_compta['CA_Technique'] = 0
                    
                    df_compta['Écart'] = abs(df_compta['CA_Comptable']) - abs(df_compta['CA_Technique'])
                    df_compta['Rapprochement'] = df_compta.apply(
                        lambda row: 'Rapproché' if abs(row['Écart']) < 0.01 else 'Non rapproché', 
                        axis=1
                    )
                    
                    # Séparer valides et invalides
                    df_invalide = df_compta[df_compta['Rapprochement'] == 'Non rapproché']
                    df_valide = df_compta[df_compta['Rapprochement'] == 'Rapproché']
                    
                    # Statistiques
                    total_debit = df_compta['Débit'].sum() if 'Débit' in df_compta.columns else 0
                    total_credit = df_compta['Crédit'].sum() if 'Crédit' in df_compta.columns else 0
                    
                    stats = {
                        'total_debit': total_debit,
                        'total_credit': total_credit,
                        'total_CA_comptable': abs(total_credit - total_debit),
                        'total_emissions_tech': df_compta['Emissions'].sum() if 'Emissions' in df_compta.columns else 0,
                        'total_ristournes_tech': df_compta['Ristournes'].sum() if 'Ristournes' in df_compta.columns else 0,
                        'total_CA_technique': df_compta['CA_Technique'].sum() if 'CA_Technique' in df_compta.columns else 0,
                        'ecart': abs(df_compta['CA_Technique'].sum() - abs(total_credit - total_debit)) if 'CA_Technique' in df_compta.columns else 0,
                        'total_polices': len(df_compta),
                        'polices_valides': len(df_valide),
                        'polices_invalides': len(df_invalide),
                        'taux_rapprochement': (len(df_valide) / len(df_compta) * 100) if len(df_compta) > 0 else 0,
                        'taux_emissions_trouvees': (df_compta['Statut_Emissions'] == 'Trouvé').sum() / len(df_compta) * 100 if len(df_compta) > 0 else 0,
                        'taux_ristournes_trouvees': (df_compta['Statut_Ristournes'] == 'Trouvé').sum() / len(df_compta) * 100 if len(df_compta) > 0 else 0
                    }
                    
                    # Stockage dans la session
                    st.session_state.pivot_comptables_complet = df_compta
                    st.session_state.tableau_listing_police_invalide_comptable = df_invalide
                    st.session_state.tableau_listing_valide_comptable = df_valide
                    st.session_state.rapprochement_stats = stats
                    
                    st.session_state.stats['total_verifications'] += 1
                    
                    log_action("Rapprochement comptable", f"{len(df_compta)} polices analysées")
                    st.success("✅ Rapprochement terminé avec succès!")
                    st.balloons()
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"❌ Erreur lors du rapprochement: {str(e)}")
                    log_action("Erreur rapprochement comptable", str(e), level="error")
                    return
    
    # Vérifier si les résultats existent
    if 'rapprochement_stats' not in st.session_state:
        st.info("ℹ️ Cliquez sur 'Lancer le rapprochement comptable' pour commencer l'analyse.")
        return
    
    stats = st.session_state.rapprochement_stats
    
    # ==================== MÉTRIQUES PRINCIPALES ====================
    st.markdown("---")
    st.markdown("### 📊 Résumé du rapprochement comptable")
    
    # Première ligne : Montants
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #dc3545 0%, #c82333 100%); 
             padding: 15px; border-radius: 12px; color: white; 
             box-shadow: 0 4px 15px rgba(220, 53, 69, 0.3);">
            <div style="font-size: 0.9em; opacity: 0.9;">💳 Total Débit</div>
            <div style="font-size: 1.6em; font-weight: bold; margin: 5px 0;">
                {stats.get('total_debit', 0):,.0f} FCFA
            </div>
            <div style="font-size: 0.8em; opacity: 0.8;">Somme des débits</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #28a745 0%, #20c997 100%); 
             padding: 15px; border-radius: 12px; color: white; 
             box-shadow: 0 4px 15px rgba(40, 167, 69, 0.3);">
            <div style="font-size: 0.9em; opacity: 0.9;">💰 Total Crédit</div>
            <div style="font-size: 1.6em; font-weight: bold; margin: 5px 0;">
                {stats.get('total_credit', 0):,.0f} FCFA
            </div>
            <div style="font-size: 0.8em; opacity: 0.8;">Somme des crédits</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%); 
             padding: 15px; border-radius: 12px; color: white; 
             box-shadow: 0 4px 15px rgba(30, 60, 114, 0.3);">
            <div style="font-size: 0.9em; opacity: 0.9;">📊 CA Comptable</div>
            <div style="font-size: 1.6em; font-weight: bold; margin: 5px 0;">
                {stats.get('total_CA_comptable', 0):,.0f} FCFA
            </div>
            <div style="font-size: 0.8em; opacity: 0.8;">Crédit - Débit</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #17a2b8 0%, #138496 100%); 
             padding: 15px; border-radius: 12px; color: white; 
             box-shadow: 0 4px 15px rgba(23, 162, 184, 0.3);">
            <div style="font-size: 0.9em; opacity: 0.9;">📈 CA Technique</div>
            <div style="font-size: 1.6em; font-weight: bold; margin: 5px 0;">
                {stats.get('total_CA_technique', 0):,.0f} FCFA
            </div>
            <div style="font-size: 0.8em; opacity: 0.8;">Émissions + Ristournes</div>
        </div>
        """, unsafe_allow_html=True)
    
    # Deuxième ligne : Écart et polices
    st.markdown("---")
    col1, col2, col3, col4 = st.columns(4)
    
    ecart = stats.get('ecart', 0)
    with col1:
        color = "#fd7e14" if ecart > 0 else "#dc3545"
        icon = "📉" if ecart > 0 else "📈"
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, {color} 0%, {color}dd 100%); 
             padding: 15px; border-radius: 12px; color: white; 
             box-shadow: 0 4px 15px rgba(253, 126, 20, 0.3);">
            <div style="font-size: 0.9em; opacity: 0.9;">{icon} Écart</div>
            <div style="font-size: 1.6em; font-weight: bold; margin: 5px 0;">
                {ecart:,.0f} FCFA
            </div>
            <div style="font-size: 0.8em; opacity: 0.8;">CA Technique - CA Comptable</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #6c757d 0%, #495057 100%); 
             padding: 15px; border-radius: 12px; color: white; 
             box-shadow: 0 4px 15px rgba(108, 117, 125, 0.3);">
            <div style="font-size: 0.9em; opacity: 0.9;">📋 Total polices</div>
            <div style="font-size: 1.6em; font-weight: bold; margin: 5px 0;">
                {stats.get('total_polices', 0):,}
            </div>
            <div style="font-size: 0.8em; opacity: 0.8;">Nombre total de polices</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        valides = stats.get('polices_valides', 0)
        taux = stats.get('taux_rapprochement', 0)
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #28a745 0%, #20c997 100%); 
             padding: 15px; border-radius: 12px; color: white; 
             box-shadow: 0 4px 15px rgba(40, 167, 69, 0.3);">
            <div style="font-size: 0.9em; opacity: 0.9;">✅ Polices rapprochées</div>
            <div style="font-size: 1.6em; font-weight: bold; margin: 5px 0;">
                {valides:,}
            </div>
            <div style="font-size: 0.8em; opacity: 0.8;">Taux: {taux:.1f}%</div>
        </div>
        """, unsafe_allow_html=True)
    
    with col4:
        invalides = stats.get('polices_invalides', 0)
        taux_inv = (invalides / stats.get('total_polices', 1) * 100) if stats.get('total_polices', 0) > 0 else 0
        st.markdown(f"""
        <div style="background: linear-gradient(135deg, #dc3545 0%, #c82333 100%); 
             padding: 15px; border-radius: 12px; color: white; 
             box-shadow: 0 4px 15px rgba(220, 53, 69, 0.3);">
            <div style="font-size: 0.9em; opacity: 0.9;">❌ Polices non rapprochées</div>
            <div style="font-size: 1.6em; font-weight: bold; margin: 5px 0;">
                {invalides:,}
            </div>
            <div style="font-size: 0.8em; opacity: 0.8;">Taux: {taux_inv:.1f}%</div>
        </div>
        """, unsafe_allow_html=True)
    
    # Troisième ligne : Statistiques supplémentaires
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        taux_emissions = stats.get('taux_emissions_trouvees', 0)
        st.markdown(f"""
        <div style="background: white; padding: 15px; border-radius: 12px; 
             box-shadow: 0 2px 10px rgba(0,0,0,0.1); text-align: center;">
            <div style="font-size: 0.9em; color: #666;">📤 Émissions trouvées</div>
            <div style="font-size: 1.8em; font-weight: bold; color: #1e3c72;">
                {taux_emissions:.1f}%
            </div>
            <div style="font-size: 0.8em; color: #999;">
                {stats.get('total_emissions_tech', 0):,.0f} FCFA
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        taux_ristournes = stats.get('taux_ristournes_trouvees', 0)
        st.markdown(f"""
        <div style="background: white; padding: 15px; border-radius: 12px; 
             box-shadow: 0 2px 10px rgba(0,0,0,0.1); text-align: center;">
            <div style="font-size: 0.9em; color: #666;">🔄 Ristournes trouvées</div>
            <div style="font-size: 1.8em; font-weight: bold; color: #1e3c72;">
                {taux_ristournes:.1f}%
            </div>
            <div style="font-size: 0.8em; color: #999;">
                {stats.get('total_ristournes_tech', 0):,.0f} FCFA
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        ecart_abs = abs(stats.get('ecart', 0))
        st.markdown(f"""
        <div style="background: white; padding: 15px; border-radius: 12px; 
             box-shadow: 0 2px 10px rgba(0,0,0,0.1); text-align: center;">
            <div style="font-size: 0.9em; color: #666;">📊 Écart absolu</div>
            <div style="font-size: 1.8em; font-weight: bold; color: {'#dc3545' if ecart_abs > 1000 else '#28a745'};">
                {ecart_abs:,.0f} FCFA
            </div>
            <div style="font-size: 0.8em; color: #999;">
                {'⚠️ Écart significatif' if ecart_abs > 1000 else '✅ Écart acceptable'}
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    # ==================== TABS POUR LES DONNÉES ====================
    st.markdown("---")
    tab1, tab2, tab3 = st.tabs(["📋 Données complètes", "❌ Non rapprochées", "✅ Rapprochées"])
    
    with tab1:
        st.markdown("### Toutes les polices comptables")
        
        # Recherche
        search = create_search_bar("compta_rapprochement_search", "Rechercher une police...")
        
        df_display = st.session_state.pivot_comptables_complet.copy()
        if search:
            df_display = filter_dataframe(df_display, search)
        
        # Statistiques rapides
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total enregistrements", len(df_display))
        with col2:
            st.metric("Colonnes", len(df_display.columns))
        with col3:
            st.metric("Taux rapprochement", f"{stats.get('taux_rapprochement', 0):.1f}%")
        
        st.dataframe(df_display, use_container_width=True, height=500, hide_index=True)
        
        # Export
        if st.button("📥 Exporter toutes les données", use_container_width=True):
            output = export_to_excel(
                [df_display],
                ["Données complètes"],
                "donnees_comptables_completes.xlsx"
            )
            if output:
                create_download_button(
                    output,
                    "donnees_comptables_completes.xlsx",
                    "Télécharger Excel"
                )
    
    with tab2:
        df_invalide = st.session_state.tableau_listing_police_invalide_comptable
        st.markdown(f"### Polices non rapprochées ({len(df_invalide)})")
        
        if not df_invalide.empty:
            # Recherche dans les invalides
            search_invalide = create_search_bar("invalide_search", "Rechercher dans les non rapprochées...")
            
            df_invalide_display = df_invalide.copy()
            if search_invalide:
                df_invalide_display = filter_dataframe(df_invalide_display, search_invalide)
            
            # Statistiques des invalides
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total non rapprochées", len(df_invalide))
            with col2:
                if 'Écart' in df_invalide.columns:
                    st.metric("Écart moyen", f"{df_invalide['Écart'].mean():,.0f} FCFA")
            with col3:
                if 'Écart' in df_invalide.columns:
                    st.metric("Écart max", f"{df_invalide['Écart'].max():,.0f} FCFA")
            
            st.dataframe(df_invalide_display, use_container_width=True, height=500, hide_index=True)
            
            # Export des non rapprochées
            col1, col2 = st.columns(2)
            with col1:
                if st.button("📥 Exporter les non rapprochées (Excel)", use_container_width=True):
                    output = export_to_excel(
                        [df_invalide],
                        ["Non rapprochées"],
                        "polices_comptables_non_rapprochees.xlsx"
                    )
                    if output:
                        create_download_button(
                            output,
                            "polices_comptables_non_rapprochees.xlsx",
                            "Télécharger Excel"
                        )
            
            with col2:
                if st.button("📥 Exporter en CSV", use_container_width=True):
                    csv_data = export_to_csv(df_invalide, "polices_comptables_non_rapprochees.csv")
                    if csv_data:
                        create_download_button(
                            csv_data,
                            "polices_comptables_non_rapprochees.csv",
                            "Télécharger CSV"
                        )
        else:
            st.success("✅ Toutes les polices comptables sont rapprochées !")
    
    with tab3:
        df_valide = st.session_state.tableau_listing_valide_comptable
        st.markdown(f"### Polices rapprochées ({len(df_valide)})")
        
        if not df_valide.empty:
            # Recherche dans les valides
            search_valide = create_search_bar("valide_search", "Rechercher dans les rapprochées...")
            
            df_valide_display = df_valide.copy()
            if search_valide:
                df_valide_display = filter_dataframe(df_valide_display, search_valide)
            
            st.dataframe(df_valide_display, use_container_width=True, height=500, hide_index=True)
            
            # Export des valides
            if st.button("📥 Exporter les rapprochées", use_container_width=True):
                output = export_to_excel(
                    [df_valide],
                    ["Rapprochées"],
                    "polices_comptables_rapprochees.xlsx"
                )
                if output:
                    create_download_button(
                        output,
                        "polices_comptables_rapprochees.xlsx",
                        "Télécharger Excel"
                    )
        else:
            st.info("Aucune police rapprochée trouvée")
    
    # ==================== VISUALISATIONS ====================
    st.markdown("---")
    st.markdown("### 📊 Analyses et visualisations")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Graphique de répartition des statuts
        df_invalide = st.session_state.tableau_listing_police_invalide_comptable
        df_valide = st.session_state.tableau_listing_valide_comptable
        
        if not df_invalide.empty or not df_valide.empty:
            fig = go.Figure(data=[
                go.Pie(
                    labels=['Rapprochées', 'Non rapprochées'],
                    values=[len(df_valide), len(df_invalide)],
                    marker_colors=['#28a745', '#dc3545'],
                    hole=0.3,
                    textinfo='label+percent',
                    hoverinfo='label+value+percent'
                )
            ])
            
            fig.update_layout(
                title="Répartition des polices comptables",
                height=400,
                showlegend=True,
                legend=dict(orientation="h", yanchor="bottom", y=-0.1)
            )
            
            st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        # Graphique de comparaison des montants
        if stats:
            fig = go.Figure(data=[
                go.Bar(
                    name='CA Comptable',
                    x=['Comptable', 'Technique'],
                    y=[stats.get('total_CA_comptable', 0), stats.get('total_CA_technique', 0)],
                    marker_color=['#1e3c72', '#2a5298'],
                    text=[f"{stats.get('total_CA_comptable', 0):,.0f}", f"{stats.get('total_CA_technique', 0):,.0f}"],
                    textposition='auto',
                )
            ])
            
            fig.update_layout(
                title="Comparaison CA Comptable vs Technique",
                yaxis_title="Montant (FCFA)",
                height=400,
                showlegend=False
            )
            
            st.plotly_chart(fig, use_container_width=True)
    
    # Analyse des écarts
    if not df_invalide.empty and 'Écart' in df_invalide.columns:
        st.markdown("### 📈 Analyse des écarts")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Distribution des écarts
            fig = px.histogram(
                df_invalide,
                x='Écart',
                nbins=30,
                title="Distribution des écarts (polices non rapprochées)",
                color_discrete_sequence=['#dc3545']
            )
            fig.update_layout(
                height=400,
                xaxis_title="Écart (FCFA)",
                yaxis_title="Nombre de polices"
            )
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            # Top 10 des écarts
            st.markdown("#### Top 10 des écarts")
            
            # Déterminer la colonne de police
            compta_col = 'No Police' if 'No Police' in df_invalide.columns else df_invalide.columns[0]
            
            ecarts_cols = [compta_col]
            if 'Écart' in df_invalide.columns:
                ecarts_cols.append('Écart')
            if 'CA_Comptable' in df_invalide.columns:
                ecarts_cols.append('CA_Comptable')
            if 'CA_Technique' in df_invalide.columns:
                ecarts_cols.append('CA_Technique')
            
            top_ecarts = df_invalide.nlargest(10, 'Écart')[ecarts_cols]
            st.dataframe(top_ecarts, use_container_width=True, hide_index=True)
            
            # Export des écarts
            if st.button("📥 Exporter l'analyse des écarts", use_container_width=True):
                output = export_to_excel(
                    [df_invalide[ecarts_cols].sort_values('Écart', ascending=False)],
                    ["Analyse écarts"],
                    "analyse_ecarts_comptables.xlsx"
                )
                if output:
                    create_download_button(
                        output,
                        "analyse_ecarts_comptables.xlsx",
                        "Télécharger Excel"
                    )
    
    # ==================== EXPORT COMPLET ====================
    st.markdown("---")
    st.markdown("### 📥 Export des résultats")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("📥 Exporter le rapport complet", type="primary", use_container_width=True):
            # Créer un fichier Excel avec plusieurs onglets
            dataframes = []
            sheet_names = []
            
            if st.session_state.pivot_comptables_complet is not None:
                dataframes.append(st.session_state.pivot_comptables_complet)
                sheet_names.append("Données complètes")
            
            if st.session_state.tableau_listing_valide_comptable is not None and not st.session_state.tableau_listing_valide_comptable.empty:
                dataframes.append(st.session_state.tableau_listing_valide_comptable)
                sheet_names.append("Polices rapprochées")
            
            if st.session_state.tableau_listing_police_invalide_comptable is not None and not st.session_state.tableau_listing_police_invalide_comptable.empty:
                dataframes.append(st.session_state.tableau_listing_police_invalide_comptable)
                sheet_names.append("Polices non rapprochées")
            
            # Ajouter un résumé
            resume_df = pd.DataFrame([
                ["Total Débit", f"{stats.get('total_debit', 0):,.0f} FCFA"],
                ["Total Crédit", f"{stats.get('total_credit', 0):,.0f} FCFA"],
                ["CA Comptable", f"{stats.get('total_CA_comptable', 0):,.0f} FCFA"],
                ["CA Technique", f"{stats.get('total_CA_technique', 0):,.0f} FCFA"],
                ["Écart", f"{stats.get('ecart', 0):,.0f} FCFA"],
                ["Total polices", stats.get('total_polices', 0)],
                ["Polices rapprochées", stats.get('polices_valides', 0)],
                ["Polices non rapprochées", stats.get('polices_invalides', 0)],
                ["Taux de rapprochement", f"{stats.get('taux_rapprochement', 0):.1f}%"],
                ["Émissions trouvées", f"{stats.get('taux_emissions_trouvees', 0):.1f}%"],
                ["Ristournes trouvées", f"{stats.get('taux_ristournes_trouvees', 0):.1f}%"]
            ], columns=["Indicateur", "Valeur"])
            
            dataframes.append(resume_df)
            sheet_names.append("Résumé")
            
            output = export_to_excel(dataframes, sheet_names, "rapprochement_comptable_complet.xlsx")
            if output:
                create_download_button(
                    output,
                    "rapprochement_comptable_complet.xlsx",
                    "Télécharger le rapport Excel"
                )
    
    with col2:
        if st.button("📊 Exporter les graphiques", use_container_width=True):
            st.info("Fonctionnalité à venir: export des graphiques en PNG")
    
    with col3:
        if st.button("📋 Générer un rapport PDF", use_container_width=True):
            with st.spinner("Génération du rapport PDF..."):
                time.sleep(2)
                st.success("Rapport PDF généré avec succès!")
                
                # Simulation de téléchargement PDF
                st.download_button(
                    label="📥 Télécharger le PDF",
                    data=b"Simulation de rapport PDF",
                    file_name="rapprochement_comptable.pdf",
                    mime="application/pdf",
                    use_container_width=True
                )
    
    # Journalisation
    log_action(
        "Rapprochement comptable", 
        f"{stats.get('total_polices', 0)} polices analysées, "
        f"{stats.get('polices_invalides', 0)} non rapprochées, "
        f"écart: {stats.get('ecart', 0):,.0f} FCFA"
    )

def page_gestion_410_411():
    """Page de gestion 410 et 411"""
    update_last_activity()
    
    st.markdown("## 📋 Gestion 410 & 411")
    
    # Onglets
    tab1, tab2, tab3, tab4 = st.tabs(["📥 Import 410", "📥 Import 411", "🔍 Vérifications", "📊 Analyses"])
    
    with tab1:
        st.markdown("### Import CP_410")
        
        uploaded_file_410 = st.file_uploader(
            "Choisir le fichier CP_410",
            type=['xlsx', 'xls', 'csv'],
            key="410_upload",
            help="Fichier des comptes 410"
        )
        
        if uploaded_file_410:
            with st.spinner("Chargement de CP_410..."):
                try:
                    if uploaded_file_410.name.endswith('.csv'):
                        df_410 = pd.read_csv(uploaded_file_410, dtype=str)
                    else:
                        df_410 = pd.read_excel(uploaded_file_410, dtype=str)
                    
                    st.session_state.df_410 = df_410
                    
                    st.success(f"CP_410 chargé: {len(df_410)} enregistrements")
                    
                    st.markdown("#### Aperçu")
                    st.dataframe(df_410.head(10), use_container_width=True)
                    
                except Exception as e:
                    st.error(f"Erreur: {str(e)}")
    
    with tab2:
        st.markdown("### Import CP_411")
        
        uploaded_file_411 = st.file_uploader(
            "Choisir le fichier CP_411",
            type=['xlsx', 'xls', 'csv'],
            key="411_upload",
            help="Fichier des comptes 411"
        )
        
        if uploaded_file_411:
            with st.spinner("Chargement de CP_411..."):
                try:
                    if uploaded_file_411.name.endswith('.csv'):
                        df_411 = pd.read_csv(uploaded_file_411, dtype=str)
                    else:
                        df_411 = pd.read_excel(uploaded_file_411, dtype=str)
                    
                    st.session_state.df_411 = df_411
                    
                    st.success(f"CP_411 chargé: {len(df_411)} enregistrements")
                    
                    st.markdown("#### Aperçu")
                    st.dataframe(df_411.head(10), use_container_width=True)
                    
                except Exception as e:
                    st.error(f"Erreur: {str(e)}")
    
    with tab3:
        if st.session_state.df_410 is not None and st.session_state.df_411 is not None:
            st.markdown("### Vérifications")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                if st.button("🔍 Vérifier polices 410/411", use_container_width=True):
                    with st.spinner("Vérification en cours..."):
                        if 'No Police' in st.session_state.df_410.columns and 'No Police' in st.session_state.df_411.columns:
                            polices_410 = set(st.session_state.df_410['No Police'].dropna().astype(str))
                            polices_411 = set(st.session_state.df_411['No Police'].dropna().astype(str))
                            
                            communes = polices_410.intersection(polices_411)
                            only_410 = polices_410 - polices_411
                            
                            st.markdown("#### Résultats 410/411")
                            
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Polices 410", len(polices_410))
                            with col2:
                                st.metric("Polices 411", len(polices_411))
                            with col3:
                                st.metric("Communes", len(communes))
                            
                            if only_410:
                                st.warning(f"{len(only_410)} polices uniquement dans 410")
                                with st.expander("Voir les polices"):
                                    st.dataframe(pd.DataFrame(sorted(only_410), columns=['Polices 410 uniquement']))
                            
                            st.session_state.stats['total_verifications'] += 1
                            log_action("Vérification 410/411", f"{len(communes)} communes, {len(only_410)} uniquement 410")
            
            with col2:
                if st.button("🔍 Vérifier polices 411/410", use_container_width=True):
                    with st.spinner("Vérification en cours..."):
                        if 'No Police' in st.session_state.df_410.columns and 'No Police' in st.session_state.df_411.columns:
                            polices_410 = set(st.session_state.df_410['No Police'].dropna().astype(str))
                            polices_411 = set(st.session_state.df_411['No Police'].dropna().astype(str))
                            
                            communes = polices_411.intersection(polices_410)
                            only_411 = polices_411 - polices_410
                            
                            st.markdown("#### Résultats 411/410")
                            
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Polices 411", len(polices_411))
                            with col2:
                                st.metric("Polices 410", len(polices_410))
                            with col3:
                                st.metric("Communes", len(communes))
                            
                            if only_411:
                                st.warning(f"{len(only_411)} polices uniquement dans 411")
                                with st.expander("Voir les polices"):
                                    st.dataframe(pd.DataFrame(sorted(only_411), columns=['Polices 411 uniquement']))
                            
                            st.session_state.stats['total_verifications'] += 1
                            log_action("Vérification 411/410", f"{len(communes)} communes, {len(only_411)} uniquement 411")
            
            with col3:
                if st.button("🔍 Vérifier références", use_container_width=True):
                    with st.spinner("Validation des références..."):
                        if 'Réf Pièce' in st.session_state.df_411.columns:
                            invalid_refs, stats = validate_references(st.session_state.df_411)
                            
                            if isinstance(stats, dict):
                                st.markdown("#### Résultats validation")
                                
                                col1, col2, col3 = st.columns(3)
                                with col1:
                                    st.metric("Total", stats['total'])
                                with col2:
                                    st.metric("Valides", stats['valides'])
                                with col3:
                                    st.metric("Invalides", stats['invalides'])
                                
                                if invalid_refs:
                                    st.warning(f"{len(invalid_refs)} références invalides trouvées")
                                    with st.expander("Voir les références invalides"):
                                        st.dataframe(pd.DataFrame(invalid_refs, columns=['Références invalides']))
                                    
                                    if st.button("📥 Exporter les invalides"):
                                        df_invalid = pd.DataFrame(invalid_refs, columns=['Références invalides'])
                                        csv_data = export_to_csv(df_invalid, "references_invalides.csv")
                                        if csv_data:
                                            create_download_button(
                                                csv_data,
                                                "references_invalides.csv",
                                                "Télécharger CSV"
                                            )
                                
                                st.session_state.stats['total_verifications'] += 1
                                log_action("Validation références", f"{stats['invalides']} invalides")
        
        else:
            st.info("Veuillez importer les fichiers CP_410 et CP_411 pour effectuer les vérifications.")
    
    with tab4:
        if st.session_state.df_410 is not None and st.session_state.df_411 is not None:
            st.markdown("### Analyses comparatives")
            
            # Statistiques globales
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("#### CP_410")
                
                df_410 = st.session_state.df_410
                st.metric("Nombre d'enregistrements", len(df_410))
                
                numeric_cols_410 = df_410.select_dtypes(include=[np.number]).columns
                if len(numeric_cols_410) > 0:
                    st.metric("Total montants", f"{df_410[numeric_cols_410[0]].sum():,.0f} FCFA")
            
            with col2:
                st.markdown("#### CP_411")
                
                df_411 = st.session_state.df_411
                st.metric("Nombre d'enregistrements", len(df_411))
                
                numeric_cols_411 = df_411.select_dtypes(include=[np.number]).columns
                if len(numeric_cols_411) > 0:
                    st.metric("Total montants", f"{df_411[numeric_cols_411[0]].sum():,.0f} FCFA")
            
            # Graphique comparatif
            if 'No Police' in df_410.columns and 'No Police' in df_411.columns:
                polices_410 = set(df_410['No Police'].dropna())
                polices_411 = set(df_411['No Police'].dropna())
                
                fig = go.Figure(data=[
                    go.Bar(
                        name='CP_410',
                        x=['Polices uniques'],
                        y=[len(polices_410 - polices_411)],
                        marker_color='#1e3c72'
                    ),
                    go.Bar(
                        name='CP_411',
                        x=['Polices uniques'],
                        y=[len(polices_411 - polices_410)],
                        marker_color='#2a5298'
                    ),
                    go.Bar(
                        name='Communes',
                        x=['Polices communes'],
                        y=[len(polices_410.intersection(polices_411))],
                        marker_color='#28a745'
                    )
                ])
                
                fig.update_layout(
                    title="Comparaison des polices",
                    barmode='group',
                    height=400
                )
                
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Importez les fichiers pour voir les analyses.")

def page_gestion_doublons():
    """Page de gestion des doublons"""
    update_last_activity()
    
    st.markdown("## 🔍 Gestion des Doublons")
    
    # Upload du fichier
    uploaded_file = st.file_uploader(
        "Importer un fichier de polices",
        type=['xlsx', 'xls', 'csv'],
        help="Fichier contenant les numéros de police"
    )
    
    if uploaded_file:
        with st.spinner("Chargement du fichier..."):
            try:
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file, dtype=str)
                else:
                    df = pd.read_excel(uploaded_file, dtype=str)
                
                st.success(f"Fichier chargé: {len(df)} enregistrements")
                
                # Détection des doublons
                duplicates_df, uniques_df, stats = detect_duplicates(df)
                
                if isinstance(stats, dict):
                    # Métriques
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        display_metric_card(
                            "Total polices",
                            stats['total'],
                            "📊"
                        )
                    
                    with col2:
                        display_metric_card(
                            "Polices en doublon",
                            stats['duplicates'],
                            "🔄",
                            f"{stats['duplicates']/stats['total']*100:.1f}% du total"
                        )
                    
                    with col3:
                        display_metric_card(
                            "Polices uniques",
                            stats['uniques'],
                            "✅",
                            f"{stats['uniques']/stats['total']*100:.1f}% du total"
                        )
                    
                    with col4:
                        display_metric_card(
                            "Polices dupliquées",
                            stats['duplicate_polices'],
                            "⚠️",
                            "Nombre de polices apparaissant plusieurs fois"
                        )
                    
                    # Tabs
                    tab1, tab2 = st.tabs(["🔄 Polices en doublon", "✅ Polices uniques"])
                    
                    with tab1:
                        if duplicates_df is not None and not duplicates_df.empty:
                            st.markdown(f"### {len(duplicates_df)} enregistrements en doublon")
                            
                            # Recherche
                            search = create_search_bar("dup_search", "Rechercher...")
                            
                            df_display = duplicates_df.copy()
                            if search:
                                df_display = filter_dataframe(df_display, search)
                            
                            st.dataframe(df_display, use_container_width=True, height=500, hide_index=True)
                            
                            # Export
                            if st.button("📥 Exporter les doublons", use_container_width=True):
                                output = export_to_excel(
                                    [duplicates_df],
                                    ["Doublons"],
                                    "doublons_polices.xlsx"
                                )
                                if output:
                                    create_download_button(
                                        output,
                                        "doublons_polices.xlsx",
                                        "Télécharger Excel"
                                    )
                        else:
                            st.success("Aucun doublon trouvé !")
                    
                    with tab2:
                        if uniques_df is not None and not uniques_df.empty:
                            st.markdown(f"### {len(uniques_df)} polices uniques")
                            
                            # Recherche
                            search = create_search_bar("unique_search", "Rechercher...")
                            
                            df_display = uniques_df.copy()
                            if search:
                                df_display = filter_dataframe(df_display, search)
                            
                            st.dataframe(df_display, use_container_width=True, height=500, hide_index=True)
                            
                            # Export
                            if st.button("📥 Exporter les polices uniques", use_container_width=True):
                                output = export_to_excel(
                                    [uniques_df],
                                    ["Polices uniques"],
                                    "polices_uniques.xlsx"
                                )
                                if output:
                                    create_download_button(
                                        output,
                                        "polices_uniques.xlsx",
                                        "Télécharger Excel"
                                    )
                        else:
                            st.info("Aucune police unique trouvée")
                    
                    # Visualisation
                    st.markdown("### 📊 Visualisation")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # Camembert
                        fig = go.Figure(data=[
                            go.Pie(
                                labels=['Polices uniques', 'Polices en doublon'],
                                values=[stats['uniques'], stats['duplicates']],
                                marker_colors=['#28a745', '#dc3545'],
                                hole=0.3
                            )
                        ])
                        fig.update_layout(title="Répartition des polices", height=400)
                        st.plotly_chart(fig, use_container_width=True)
                    
                    with col2:
                        # Histogramme des occurrences
                        police_col = None
                        for col in df.columns:
                            if 'police' in col.lower() or 'num' in col.lower():
                                police_col = col
                                break
                        
                        if police_col:
                            occurrences = df[police_col].value_counts().value_counts().sort_index()
                            
                            fig = px.bar(
                                x=occurrences.index,
                                y=occurrences.values,
                                title="Distribution des occurrences",
                                labels={'x': "Nombre d'occurrences", 'y': 'Nombre de polices'},
                                color_discrete_sequence=['#1e3c72']
                            )
                            fig.update_layout(height=400)
                            st.plotly_chart(fig, use_container_width=True)
                    
                    log_action("Analyse doublons", f"{stats['duplicate_polices']} polices dupliquées")
                
                else:
                    st.error(f"Erreur: {stats}")
                    
            except Exception as e:
                st.error(f"Erreur lors du traitement: {str(e)}")

def page_gestion_production():
    """Page de gestion de production (certificats)"""
    update_last_activity()
    
    st.markdown("## 📄 Générateur de Certificats")
    
    # Onglets
    tab1, tab2, tab3 = st.tabs(["📥 Import", "🎨 Personnalisation", "🚀 Génération"])
    
    with tab1:
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### Modèle de certificat")
            template_file = st.file_uploader(
                "Importer un modèle Word",
                type=['docx'],
                key="template_upload",
                help="Fichier modèle au format Word (.docx)"
            )
            
            if template_file:
                st.success("Modèle chargé avec succès")
                st.session_state.template = template_file
                
                # Aperçu du modèle
                st.markdown("#### Informations")
                st.info(f"Nom: {template_file.name}\nTaille: {template_file.size / 1024:.1f} KB")
        
        with col2:
            st.markdown("### Données à générer")
            data_file = st.file_uploader(
                "Importer les données",
                type=['xlsx', 'xls', 'csv'],
                key="production_data_upload",
                help="Fichier contenant les données pour les certificats"
            )
            
            if data_file:
                try:
                    if data_file.name.endswith('.csv'):
                        df = pd.read_csv(data_file)
                    else:
                        df = pd.read_excel(data_file)
                    
                    st.success(f"Données chargées: {len(df)} enregistrements")
                    st.session_state.production_data = df
                    
                    st.markdown("#### Aperçu des données")
                    st.dataframe(df.head(5), use_container_width=True)
                    
                except Exception as e:
                    st.error(f"Erreur: {str(e)}")
    
    with tab2:
        if st.session_state.get('template') and st.session_state.get('production_data') is not None:
            st.markdown("### Personnalisation des certificats")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                police = st.selectbox(
                    "Police",
                    ['Arial', 'Times New Roman', 'Helvetica', 'Calibri', 'Verdana', 'Tahoma'],
                    help="Police de caractères pour le texte"
                )
                taille = st.slider("Taille de police", 8, 24, 12, help="Taille du texte")
            
            with col2:
                couleur = st.color_picker("Couleur du texte", "#000000", help="Couleur du texte")
                alignement = st.selectbox(
                    "Alignement",
                    ["Gauche", "Centré", "Droite"],
                    help="Alignement du texte"
                )
            
            with col3:
                st.markdown("#### Aperçu")
                st.markdown(
                    f"""
                    <div style="font-family: {police}; font-size: {taille}px; color: {couleur}; 
                         text-align: {'left' if alignement == 'Gauche' else 'center' if alignement == 'Centré' else 'right'};
                         padding: 20px; border: 1px solid #ddd; border-radius: 5px;">
                        Texte d'exemple
                    </div>
                    """,
                    unsafe_allow_html=True
                )
            
            # Sauvegarde des préférences
            if st.button("💾 Sauvegarder les préférences", use_container_width=True):
                st.session_state.user_preferences.update({
                    'certificat_police': police,
                    'certificat_taille': taille,
                    'certificat_couleur': couleur,
                    'certificat_alignement': alignement
                })
                st.success("Préférences sauvegardées")
        else:
            st.info("Veuillez d'abord importer un modèle et des données.")
    
    with tab3:
        if st.session_state.get('template') and st.session_state.get('production_data') is not None:
            st.markdown("### Génération des certificats")
            
            df = st.session_state.production_data
            
            st.markdown(f"**{len(df)} certificats à générer**")
            
            # Options de génération
            col1, col2 = st.columns(2)
            
            with col1:
                output_format = st.selectbox(
                    "Format de sortie",
                    ["PDF", "DOCX"],
                    help="Format des fichiers générés"
                )
            
            with col2:
                naming = st.text_input(
                    "Préfixe des fichiers",
                    "Certificat_",
                    help="Préfixe pour les noms de fichiers"
                )
            
            # Simulation de progression
            if st.button("🚀 Lancer la génération", type="primary", use_container_width=True):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                total = len(df)
                
                for i in range(total):
                    # Simulation
                    time.sleep(0.05)
                    
                    # Mise à jour de la progression
                    progress = (i + 1) / total
                    progress_bar.progress(progress)
                    status_text.text(f"Génération: {i+1}/{total} certificats")
                    
                    if (i + 1) % 10 == 0:
                        st.session_state.stats['total_certificats'] += 10
                
                progress_bar.progress(1.0)
                status_text.text(f"✅ {total} certificats générés avec succès!")
                
                st.balloons()
                
                # Statistiques
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Certificats générés", total)
                with col2:
                    st.metric("Format", output_format)
                with col3:
                    taille_estimée = total * 50
                    st.metric("Taille estimée", f"{taille_estimée / 1024:.1f} MB")
                
                # Bouton de téléchargement simulé
                st.download_button(
                    label="📥 Télécharger tous les certificats (ZIP)",
                    data=b"Simulation de fichier ZIP",
                    file_name="certificats.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                
                log_action("Génération certificats", f"{total} certificats générés")
        else:
            st.info("Veuillez d'abord importer un modèle et des données dans l'onglet Import.")

def page_statistiques():
    """Page de statistiques"""
    update_last_activity()
    
    st.markdown("## 📈 Statistiques et Analyses")
    
    # Onglets
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Vue d'ensemble", "📈 Tendances", "📋 Rapports", "📥 Export"])
    
    with tab1:
        st.markdown("### Vue d'ensemble du système")
        
        # Métriques globales
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric(
                "Utilisateurs",
                len(get_all_users()) if get_all_users() else 0,
                help="Nombre total d'utilisateurs"
            )
        
        with col2:
            st.metric(
                "Imports totaux",
                st.session_state.stats.get('total_imports', 0),
                help="Nombre total d'imports de données"
            )
        
        with col3:
            st.metric(
                "Vérifications",
                st.session_state.stats.get('total_verifications', 0),
                help="Nombre total de vérifications"
            )
        
        with col4:
            st.metric(
                "Certificats",
                st.session_state.stats.get('total_certificats', 0),
                help="Nombre total de certificats générés"
            )
        
        # État des données
        st.markdown("### État des données")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### Données techniques")
            if st.session_state.pivot_techniques is not None:
                df_tech = st.session_state.pivot_techniques
                st.metric("Enregistrements", len(df_tech))
                
                if 'Emissions' in df_tech.columns:
                    st.metric("Total Émissions", f"{df_tech['Emissions'].sum():,.0f} FCFA")
                
                if 'Ristournes' in df_tech.columns:
                    st.metric("Total Ristournes", f"{df_tech['Ristournes'].sum():,.0f} FCFA")
            else:
                st.info("Aucune donnée technique")
        
        with col2:
            st.markdown("#### Données comptables")
            if st.session_state.pivot_comptables is not None:
                df_compta = st.session_state.pivot_comptables
                st.metric("Enregistrements", len(df_compta))
                
                if 'Débit' in df_compta.columns:
                    st.metric("Total Débit", f"{df_compta['Débit'].sum():,.0f} FCFA")
                
                if 'Crédit' in df_compta.columns:
                    st.metric("Total Crédit", f"{df_compta['Crédit'].sum():,.0f} FCFA")
            else:
                st.info("Aucune donnée comptable")
        
        # Dernières activités
        st.markdown("### Dernières activités")
        
        if st.session_state.logs:
            logs_df = pd.DataFrame(st.session_state.logs[-20:])
            if not logs_df.empty and 'timestamp' in logs_df.columns:
                logs_df['timestamp'] = pd.to_datetime(logs_df['timestamp']).dt.strftime('%d/%m/%Y %H:%M')
                st.dataframe(
                    logs_df[['timestamp', 'username', 'action', 'details']],
                    use_container_width=True,
                    height=400,
                    hide_index=True
                )
    
    with tab2:
        st.markdown("### Tendances et évolutions")
        
        # Simulation de données de tendance
        dates = pd.date_range(end=datetime.now(), periods=30, freq='D')
        
        np.random.seed(42)
        imports_data = np.random.randint(5, 20, 30)
        verifications_data = np.random.randint(10, 30, 30)
        
        trend_df = pd.DataFrame({
            'Date': dates,
            'Imports': imports_data,
            'Vérifications': verifications_data
        })
        
        # Graphique des tendances
        fig = go.Figure()
        
        fig.add_trace(go.Scatter(
            x=trend_df['Date'],
            y=trend_df['Imports'],
            name='Imports',
            mode='lines+markers',
            line=dict(color='#1e3c72', width=2),
            marker=dict(size=6)
        ))
        
        fig.add_trace(go.Scatter(
            x=trend_df['Date'],
            y=trend_df['Vérifications'],
            name='Vérifications',
            mode='lines+markers',
            line=dict(color='#28a745', width=2),
            marker=dict(size=6)
        ))
        
        fig.update_layout(
            title="Activité des 30 derniers jours",
            xaxis_title="Date",
            yaxis_title="Nombre d'opérations",
            height=500,
            hovermode='x unified'
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # Statistiques par jour
        st.markdown("### Statistiques journalières")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Moyenne imports/jour", f"{imports_data.mean():.1f}")
        with col2:
            st.metric("Moyenne vérifications/jour", f"{verifications_data.mean():.1f}")
        with col3:
            st.metric("Pic d'activité", max(imports_data + verifications_data))
    
    with tab3:
        st.markdown("### Génération de rapports")
        
        report_type = st.selectbox(
            "Type de rapport",
            ["Rapport d'activité", "Rapport de données", "Rapport de performance", "Rapport personnalisé"]
        )
        
        periode = st.selectbox(
            "Période",
            ["Aujourd'hui", "Cette semaine", "Ce mois", "Cette année", "Période personnalisée"]
        )
        
        if periode == "Période personnalisée":
            col1, col2 = st.columns(2)
            with col1:
                date_debut = st.date_input("Date de début", datetime.now())
            with col2:
                date_fin = st.date_input("Date de fin", datetime.now())
        
        format_export = st.selectbox(
            "Format d'export",
            ["PDF", "Excel", "HTML"]
        )
        
        if st.button("📊 Générer le rapport", type="primary", use_container_width=True):
            with st.spinner("Génération du rapport en cours..."):
                time.sleep(2)
                
                st.success("Rapport généré avec succès!")
                
                # Simulation de téléchargement
                st.download_button(
                    label="📥 Télécharger le rapport",
                    data=b"Simulation de rapport",
                    file_name=f"rapport_{datetime.now().strftime('%Y%m%d')}.{'pdf' if format_export == 'PDF' else 'xlsx' if format_export == 'Excel' else 'html'}",
                    mime="application/octet-stream",
                    use_container_width=True
                )
                
                log_action("Rapport", f"Rapport {report_type} généré")
    
    with tab4:
        st.markdown("### Export des données")
        
        # Export complet
        st.markdown("#### Export complet de la base")
        
        if st.button("📦 Exporter toutes les données", use_container_width=True):
            dataframes = []
            sheet_names = []
            
            if st.session_state.pivot_techniques is not None:
                dataframes.append(st.session_state.pivot_techniques)
                sheet_names.append("Données techniques")
            
            if st.session_state.pivot_comptables is not None:
                dataframes.append(st.session_state.pivot_comptables)
                sheet_names.append("Données comptables")
            
            if st.session_state.df_410 is not None:
                dataframes.append(st.session_state.df_410)
                sheet_names.append("CP_410")
            
            if st.session_state.df_411 is not None:
                dataframes.append(st.session_state.df_411)
                sheet_names.append("CP_411")
            
            if dataframes:
                output = export_to_excel(dataframes, sheet_names, "export_complet.xlsx")
                if output:
                    create_download_button(
                        output,
                        "export_complet.xlsx",
                        "Télécharger l'export complet"
                    )
            else:
                st.warning("Aucune donnée à exporter")
        
        # Export des logs
        st.markdown("#### Export des logs")
        
        if st.session_state.logs:
            logs_df = pd.DataFrame(st.session_state.logs)
            
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("📋 Exporter les logs en Excel", use_container_width=True):
                    output = export_to_excel([logs_df], ["Logs"], "logs_application.xlsx")
                    if output:
                        create_download_button(
                            output,
                            "logs_application.xlsx",
                            "Télécharger Excel"
                        )
            
            with col2:
                if st.button("📋 Exporter les logs en CSV", use_container_width=True):
                    csv_data = export_to_csv(logs_df, "logs_application.csv")
                    if csv_data:
                        create_download_button(
                            csv_data,
                            "logs_application.csv",
                            "Télécharger CSV"
                        )
        else:
            st.info("Aucun log à exporter")

def page_administration():
    """Page d'administration"""
    update_last_activity()
    
    # Vérification des droits
    if st.session_state.role != "admin":
        st.error("⛔ Accès réservé aux administrateurs")
        return
    
    st.markdown("## ⚙️ Administration")
    
    # Onglets
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "👥 Utilisateurs", "📋 Logs", "📜 Historique", 
        "💾 Sauvegardes", "🔧 Paramètres"
    ])
    
    with tab1:
        st.markdown("### Gestion des utilisateurs")
        
        # Liste des utilisateurs
        users = get_all_users()
        
        if users:
            users_df = pd.DataFrame(users)
            st.dataframe(users_df, use_container_width=True, hide_index=True)
        
        # Formulaire d'ajout
        with st.expander("➕ Ajouter un utilisateur"):
            with st.form("add_user_form"):
                col1, col2 = st.columns(2)
                
                with col1:
                    new_username = st.text_input("Nom d'utilisateur*")
                    new_password = st.text_input("Mot de passe*", type="password")
                
                with col2:
                    new_email = st.text_input("Email*")
                    new_role = st.selectbox("Rôle", ["user", "admin"])
                
                if st.form_submit_button("Ajouter l'utilisateur", use_container_width=True):
                    if new_username and new_password and new_email:
                        success, message = add_user(new_username, new_password, new_email, new_role)
                        if success:
                            st.success(message)
                            st.rerun()
                        else:
                            st.error(message)
                    else:
                        st.warning("Veuillez remplir tous les champs obligatoires")
        
        # Modification / Suppression
        if users:
            with st.expander("✏️ Modifier un utilisateur"):
                selected_user = st.selectbox(
                    "Sélectionner un utilisateur",
                    [u['username'] for u in users]
                )
                
                user_data = next((u for u in users if u['username'] == selected_user), None)
                
                if user_data:
                    with st.form("edit_user_form"):
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            edit_email = st.text_input("Email", value=user_data.get('email', ''))
                            edit_role = st.selectbox(
                                "Rôle",
                                ["user", "admin"],
                                index=0 if user_data.get('role') == 'user' else 1
                            )
                        
                        with col2:
                            edit_status = st.selectbox(
                                "Statut",
                                ["active", "inactive"],
                                index=0 if user_data.get('status') == 'active' else 1
                            )
                            edit_password = st.text_input("Nouveau mot de passe (optionnel)", type="password")
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            if st.form_submit_button("💾 Mettre à jour", use_container_width=True):
                                update_data = {
                                    'email': edit_email,
                                    'role': edit_role,
                                    'status': edit_status
                                }
                                if edit_password:
                                    update_data['password'] = edit_password
                                
                                success, message = update_user(selected_user, update_data)
                                if success:
                                    st.success(message)
                                    st.rerun()
                                else:
                                    st.error(message)
                        
                        with col2:
                            if st.form_submit_button("🗑️ Supprimer", use_container_width=True):
                                if selected_user != st.session_state.username:
                                    success, message = delete_user(selected_user)
                                    if success:
                                        st.success(message)
                                        st.rerun()
                                    else:
                                        st.error(message)
                                else:
                                    st.error("Vous ne pouvez pas supprimer votre propre compte")
    
    with tab2:
        st.markdown("### Journal des activités")
        
        if st.session_state.logs:
            logs_df = pd.DataFrame(st.session_state.logs)
            
            # Filtres
            col1, col2, col3 = st.columns(3)
            
            with col1:
                if 'username' in logs_df.columns:
                    users_filter = ["Tous"] + list(logs_df['username'].unique())
                    selected_user = st.selectbox("Utilisateur", users_filter)
            
            with col2:
                if 'action' in logs_df.columns:
                    actions_filter = ["Toutes"] + list(logs_df['action'].unique())
                    selected_action = st.selectbox("Action", actions_filter)
            
            with col3:
                if 'level' in logs_df.columns:
                    levels_filter = ["Tous"] + list(logs_df['level'].unique())
                    selected_level = st.selectbox("Niveau", levels_filter)
            
            # Application des filtres
            filtered_logs = logs_df.copy()
            
            if selected_user != "Tous":
                filtered_logs = filtered_logs[filtered_logs['username'] == selected_user]
            
            if selected_action != "Toutes":
                filtered_logs = filtered_logs[filtered_logs['action'] == selected_action]
            
            if selected_level != "Tous":
                filtered_logs = filtered_logs[filtered_logs['level'] == selected_level]
            
            # Affichage
            st.dataframe(filtered_logs, use_container_width=True, height=500, hide_index=True)
            
            # Export
            col1, col2 = st.columns(2)
            
            with col1:
                if st.button("📥 Exporter les logs filtrés", use_container_width=True):
                    output = export_to_excel([filtered_logs], ["Logs"], "logs_filtres.xlsx")
                    if output:
                        create_download_button(
                            output,
                            "logs_filtres.xlsx",
                            "Télécharger Excel"
                        )
            
            with col2:
                if st.button("🗑️ Effacer les logs", use_container_width=True):
                    if st.checkbox("Confirmer la suppression"):
                        st.session_state.logs = []
                        st.success("Logs effacés")
                        st.rerun()
        else:
            st.info("Aucun log disponible")
    
    with tab3:
        st.markdown("### Historique des actions")
        
        if st.session_state.history:
            history_df = pd.DataFrame(st.session_state.history)
            
            # Filtre par utilisateur
            if 'username' in history_df.columns:
                users_filter = ["Tous"] + list(history_df['username'].unique())
                selected_user_hist = st.selectbox("Filtrer par utilisateur", users_filter, key="hist_user")
                
                if selected_user_hist != "Tous":
                    history_df = history_df[history_df['username'] == selected_user_hist]
            
            st.dataframe(history_df, use_container_width=True, height=500, hide_index=True)
            
            # Export
            if st.button("📥 Exporter l'historique", use_container_width=True):
                output = export_to_excel([history_df], ["Historique"], "historique_actions.xlsx")
                if output:
                    create_download_button(
                        output,
                        "historique_actions.xlsx",
                        "Télécharger Excel"
                    )
        else:
            st.info("Aucun historique disponible")
    
    with tab4:
        st.markdown("### Gestion des sauvegardes")
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("💾 Créer une sauvegarde", use_container_width=True):
                with st.spinner("Création de la sauvegarde..."):
                    # Créer un dossier backups si nécessaire
                    os.makedirs("backups", exist_ok=True)
                    
                    # Créer une sauvegarde simple
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    backup_file = f"backups/backup_{timestamp}.db"
                    
                    if os.path.exists(DB_FILE):
                        shutil.copy2(DB_FILE, backup_file)
                        st.success(f"Sauvegarde créée: {os.path.basename(backup_file)}")
                        st.balloons()
                    else:
                        st.error("Fichier de base de données introuvable")
        
        with col2:
            # Liste des sauvegardes
            backup_dir = "backups"
            if os.path.exists(backup_dir):
                backups = [f for f in os.listdir(backup_dir) if f.startswith("backup_") and f.endswith(".db")]
                
                if backups:
                    selected_backup = st.selectbox("Sauvegardes disponibles", backups)
                    
                    if st.button("🔄 Restaurer", use_container_width=True):
                        if st.checkbox("Confirmer la restauration"):
                            with st.spinner("Restauration en cours..."):
                                backup_path = os.path.join(backup_dir, selected_backup)
                                if os.path.exists(backup_path):
                                    shutil.copy2(backup_path, DB_FILE)
                                    st.success("Restauration réussie")
                                    st.rerun()
                                else:
                                    st.error("Fichier de sauvegarde introuvable")
        
        # Configuration des sauvegardes
        st.markdown("### Configuration des sauvegardes")
        
        col1, col2 = st.columns(2)
        
        with col1:
            backup_interval = st.number_input(
                "Intervalle (heures)",
                min_value=1,
                max_value=168,
                value=24,
                help="Intervalle entre les sauvegardes automatiques"
            )
        
        with col2:
            keep_backups = st.number_input(
                "Nombre de sauvegardes à conserver",
                min_value=1,
                max_value=50,
                value=10,
                help="Nombre maximum de sauvegardes à garder"
            )
        
        if st.button("💾 Sauvegarder la configuration", use_container_width=True):
            st.success("Configuration sauvegardée")
    
    with tab5:
        st.markdown("### Paramètres de sécurité")
        
        config = st.session_state.security_config
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### Politique de mot de passe")
            
            config['min_password_length'] = st.number_input(
                "Longueur minimale",
                min_value=6,
                max_value=20,
                value=config.get('min_password_length', 8)
            )
            
            config['require_uppercase'] = st.checkbox(
                "Requiert des majuscules",
                value=config.get('require_uppercase', True)
            )
            
            config['require_digit'] = st.checkbox(
                "Requiert des chiffres",
                value=config.get('require_digit', True)
            )
            
            config['require_special'] = st.checkbox(
                "Requiert des caractères spéciaux",
                value=config.get('require_special', True)
            )
        
        with col2:
            st.markdown("#### Verrouillage de compte")
            
            config['max_login_attempts'] = st.number_input(
                "Tentatives maximales",
                min_value=3,
                max_value=10,
                value=config.get('max_login_attempts', 5)
            )
            
            config['lockout_duration'] = st.number_input(
                "Durée de verrouillage (minutes)",
                min_value=5,
                max_value=1440,
                value=config.get('lockout_duration', 30)
            )
            
            config['session_timeout'] = st.number_input(
                "Timeout de session (minutes)",
                min_value=5,
                max_value=120,
                value=config.get('session_timeout', 30)
            )
            
            config['two_factor_enabled'] = st.checkbox(
                "Activer la double authentification",
                value=config.get('two_factor_enabled', False)
            )
        
        if st.button("💾 Sauvegarder les paramètres", type="primary", use_container_width=True):
            st.session_state.security_config = config
            st.success("Paramètres de sécurité mis à jour")
            log_action("Configuration", "Paramètres de sécurité modifiés")

def get_image_base64(image_path):
    """Convertit une image en base64"""
    try:
        with open(image_path, "rb") as f:
            return base64.b64encode(f.read()).decode()
    except:
        return None

def get_logo_src():
    """Récupère la source du logo en base64 ou SVG fallback"""
    # Essayer de charger l'image
    image_paths = ["logo_1.jpg", "assets/logo_1.jpg", "static/logo_1.jpg", "logo.png", "assets/logo.png"]
    
    for path in image_paths:
        if os.path.exists(path):
            logo_b64 = get_image_base64(path)
            if logo_b64:
                ext = path.split('.')[-1]
                return f"data:image/{ext};base64,{logo_b64}"
    
    # Fallback SVG si aucune image n'est trouvée
    return "data:image/svg+xml;base64,PHN2ZyB3aWR0aD0iMTAwIiBoZWlnaHQ9IjEwMCIgdmlld0JveD0iMCAwIDEwMCAxMDAiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyI+CiAgICA8ZGVmcz4KICAgICAgICA8bGluZWFyR3JhZGllbnQgaWQ9ImdyYWQiIHgxPSIwJSIgeTE9IjAlIiB4Mj0iMTAwJSIgeTI9IjEwMCUiPgogICAgICAgICAgICA8c3RvcCBvZmZzZXQ9IjAlIiBzdHlsZT0ic3RvcC1jb2xvcjojNjY3ZWVhO3N0b3Atb3BhY2l0eToxIiAvPgogICAgICAgICAgICA8c3RvcCBvZmZzZXQ9IjEwMCUiIHN0eWxlPSJzdG9wLWNvbG9yOiM3NjRiYTI7c3RvcC1vcGFjaXR5OjEiIC8+CiAgICAgICAgPC9saW5lYXJHcmFkaWVudD4KICAgIDwvZGVmcz4KICAgIDxyZWN0IHdpZHRoPSI5MCIgaGVpZ2h0PSI5MCIgeD0iNSIgeT0iNSIgcng9IjI1IiBmaWxsPSJ1cmwoI2dyYWQpIiAvPgogICAgPHRleHQgeD0iNTAiIHk9IjU1IiBmb250LXNpemU9IjQwIiB0ZXh0LWFuY2hvcj0ibWlkZGxlIiBmaWxsPSJ3aGl0ZSIgZm9udC13ZWlnaHQ9ImJvbGQiIGZvbnQtZmFtaWx5PSJBcmlhbCwgc2Fucy1zZXJpZiI+QVY8L3RleHQ+CiAgICA8Y2lyY2xlIGN4PSI1MCIgY3k9IjUwIiByPSI0NSIgZmlsbD0ibm9uZSIgc3Ryb2tlPSJ3aGl0ZSIgc3Ryb2tlLXdpZHRoPSIyIiBzdHJva2UtZGFzaGFycmF5PSI1LDYiIG9wYWNpdHk9IjAuNSIgLz4KPC9zdmc+"
# ======================== BARRE LATÉRALE ========================
def sidebar():
    """Affiche la barre latérale"""
    with st.sidebar:
        # ============ CSS PERSONNALISÉ ============
        st.markdown("""
        <style>
            /* Style de la barre latérale */
            section[data-testid="stSidebar"] {
                background: linear-gradient(180deg, #f8f9fa 0%, #e9ecef 100%);
            }
            
            /* Style du séparateur */
            hr {
                margin: 15px 0;
                border: none;
                height: 1px;
                background: linear-gradient(90deg, transparent, rgba(0,0,0,0.1), transparent);
            }
            
            /* ============ EN-TÊTE AVEC LOGO ============ */
            .sidebar-header {
                text-align: center;
                padding: 20px 15px 15px 15px;
                margin-bottom: 15px;
                background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
                border-radius: 15px;
                box-shadow: 0 4px 20px rgba(30, 60, 114, 0.3);
                position: relative;
                overflow: hidden;
            }
            
            .sidebar-header::before {
                content: '';
                position: absolute;
                top: -50%;
                right: -50%;
                width: 100%;
                height: 100%;
                background: radial-gradient(circle, rgba(255,255,255,0.1) 0%, transparent 70%);
                animation: pulse 3s ease-in-out infinite;
            }
            
            @keyframes pulse {
                0%, 100% { transform: scale(1); opacity: 0.5; }
                50% { transform: scale(1.2); opacity: 1; }
            }
            
            .sidebar-header .logo-container {
                position: relative;
                z-index: 1;
                display: flex;
                flex-direction: column;
                align-items: center;
            }
            
            .sidebar-header .logo-wrapper {
                position: relative;
                width: 100px;
                height: 100px;
                margin: 0 auto 10px auto;
            }
            
            .sidebar-header .logo-wrapper img {
                width: 100px;
                height: 100px;
                border-radius: 50%;
                border: 3px solid rgba(255,255,255,0.3);
                box-shadow: 0 4px 20px rgba(0,0,0,0.3);
                transition: all 0.5s cubic-bezier(0.4, 0, 0.2, 1);
                object-fit: cover;
            }
            
            .sidebar-header .logo-wrapper img:hover {
                transform: scale(1.1) rotate(-5deg);
                border-color: rgba(255,255,255,0.8);
                box-shadow: 0 8px 40px rgba(0,0,0,0.4);
            }
            
            .sidebar-header .logo-wrapper .logo-ring {
                position: absolute;
                top: -5px;
                left: -5px;
                right: -5px;
                bottom: -5px;
                border-radius: 50%;
                border: 3px solid transparent;
                border-top-color: #fff;
                border-right-color: #fff;
                animation: spin 3s linear infinite;
            }
            
            @keyframes spin {
                0% { transform: rotate(0deg); }
                100% { transform: rotate(360deg); }
            }
            
            .sidebar-header .app-title {
                position: relative;
                z-index: 1;
                color: white;
                font-size: 1.8em;
                font-weight: 700;
                margin: 5px 0 2px 0;
                text-shadow: 0 2px 10px rgba(0,0,0,0.3);
                letter-spacing: 2px;
            }
            
            .sidebar-header .app-subtitle {
                position: relative;
                z-index: 1;
                color: rgba(255,255,255,0.8);
                font-size: 0.85em;
                margin: 0;
                letter-spacing: 1px;
            }
            
            .sidebar-header .app-version {
                position: relative;
                z-index: 1;
                color: rgba(255,255,255,0.5);
                font-size: 0.7em;
                margin: 3px 0 0 0;
                background: rgba(255,255,255,0.1);
                padding: 2px 12px;
                border-radius: 12px;
                display: inline-block;
            }
            
            .sidebar-header .decoration-line {
                position: relative;
                z-index: 1;
                width: 50px;
                height: 2px;
                margin: 8px auto;
                background: linear-gradient(90deg, transparent, rgba(255,255,255,0.6), transparent);
            }
            
            /* ============ CARTE UTILISATEUR ============ */
            .user-card {
                background: rgba(255,255,255,0.95);
                padding: 15px;
                border-radius: 12px;
                margin-bottom: 15px;
                box-shadow: 0 2px 10px rgba(0,0,0,0.05);
                border: 1px solid rgba(255,255,255,0.2);
                transition: all 0.3s ease;
            }
            
            .user-card:hover {
                transform: translateY(-2px);
                box-shadow: 0 4px 20px rgba(0,0,0,0.1);
            }
            
            .user-card .user-name {
                color: #1e3c72;
                margin: 0;
                font-size: 1.1em;
                font-weight: 600;
            }
            
            .user-card .user-role {
                margin: 5px 0 0 0;
                font-size: 0.85em;
                font-weight: 500;
            }
            
            .user-card .user-role.admin {
                color: #dc3545;
            }
            
            .user-card .user-role.user {
                color: #28a745;
            }
            
            .user-card .user-time {
                color: #999;
                margin: 5px 0 0 0;
                font-size: 0.75em;
            }
            
            /* ============ CACHE LE LABEL RADIO ============ */
            .stRadio > label {
                display: none !important;
            }
            
            /* ============ STREAMLIT RADIO OVERRIDE ============ */
            div[data-testid="stRadio"] > div[role="radiogroup"] {
                gap: 2px;
            }
            
            div[data-testid="stRadio"] > div[role="radiogroup"] > label {
                padding: 0;
                margin: 0;
                background: transparent !important;
                border-radius: 10px;
            }
            
            div[data-testid="stRadio"] > div[role="radiogroup"] > label > div {
                display: none !important;
            }
            
            div[data-testid="stRadio"] > div[role="radiogroup"] > label > div:last-child {
                display: block !important;
                width: 100%;
            }
            
            /* Style des items du menu radio */
            div[data-testid="stRadio"] > div[role="radiogroup"] > label > div:last-child > div {
                padding: 0 !important;
            }
            
            div[data-testid="stRadio"] > div[role="radiogroup"] > label > div:last-child p {
                margin: 0 !important;
                padding: 10px 15px !important;
                border-radius: 10px !important;
                font-weight: 500 !important;
                color: #495057 !important;
                transition: all 0.3s ease !important;
                background: transparent !important;
            }
            
            div[data-testid="stRadio"] > div[role="radiogroup"] > label > div:last-child p:hover {
                background: rgba(30, 60, 114, 0.08) !important;
                transform: translateX(5px) !important;
                color: #1e3c72 !important;
            }
            
            div[data-testid="stRadio"] > div[role="radiogroup"] > label[data-selected="true"] > div:last-child p {
                background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%) !important;
                color: white !important;
                box-shadow: 0 4px 15px rgba(30, 60, 114, 0.3) !important;
                font-weight: 600 !important;
            }
            
            /* ============ PIED DE PAGE ============ */
            .sidebar-footer {
                text-align: center;
                color: #999;
                font-size: 0.75em;
                padding: 15px 10px;
                margin-top: 20px;
                border-top: 1px solid rgba(0,0,0,0.05);
            }
            
            .sidebar-footer .footer-logo {
                font-size: 1.5em;
                margin-bottom: 5px;
                opacity: 0.5;
            }
            
            .sidebar-footer .footer-text {
                margin: 2px 0;
            }
            
            .sidebar-footer .footer-divider {
                width: 30px;
                height: 1px;
                margin: 8px auto;
                background: linear-gradient(90deg, transparent, #ddd, transparent);
            }
        </style>
        """, unsafe_allow_html=True)
        
        # ============ EN-TÊTE AVEC LOGO ============
        # Utiliser st.image pour afficher le logo
        try:
            if os.path.exists("logo_1.jpg"):
                logo = Image.open("logo_1.jpg")
                st.image(logo, width=220)
            else:
                st.markdown("""
                <div style="text-align: center; font-size: 3em; margin-bottom: 10px;">
                    🏢
                </div>
                """, unsafe_allow_html=True)
        except:
            st.markdown("""
            <div style="text-align: center; font-size: 3em; margin-bottom: 10px;">
                🏢
            </div>
            """, unsafe_allow_html=True)
        
        st.markdown("""
        <div style="text-align: center; margin-bottom: 15px;">
            <!--h2 style="color: #1e3c72; margin: 5px 0 0 0; border-bottom: none;">AGC-VIE</h2-->
            <p style="color: #666; font-size: 0.85em; margin: 0;">Système de Gestion</p>
            <p style="color: #999; font-size: 0.7em; margin: 2px 0 0 0;">Version 2.0</p>
        </div>
        """, unsafe_allow_html=True)
        
        # ============ CARTE UTILISATEUR ============
        if st.session_state.authenticated:
            role_class = "admin" if st.session_state.role == "admin" else "user"
            role_icon = "👑" if st.session_state.role == "admin" else "👤"
            role_text = "Administrateur" if st.session_state.role == "admin" else "Utilisateur"
            
            st.markdown(f"""
            <div class="user-card">
                <div class="user-name">👤 {st.session_state.username}</div>
                <div class="user-role {role_class}">{role_icon} {role_text}</div>
                <div class="user-time">🕐 {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>
            </div>
            """, unsafe_allow_html=True)
        
        # ============ MENU DE NAVIGATION ============
        menu_options = {
            "Accueil": "🏠",
            "Gestion Technique": "📊",
            "Gestion Comptable": "💰",
            "Rapprochement Technique": "🔄",
            "Rapprochement Comptable": "📋",
            "Gestion 410 & 411": "📁",
            "Gestion Doublons": "🔍",
            "Gestion Production": "📄",
            "Statistiques": "📈",
        }
        
        # Options admin
        if st.session_state.role == "admin":
            menu_options["Administration"] = "⚙️"
        
        # Déconnexion en dernier
        menu_options["Déconnexion"] = "🚪"
        
        # Radio stylisé
        selected = st.radio(
            "Navigation",
            list(menu_options.keys()),
            format_func=lambda x: f"{menu_options[x]} {x}",
            key="navigation",
            label_visibility="collapsed"
        )
        
        # ============ PIED DE PAGE ============
        st.markdown("""
        <div class="sidebar-footer">
            <div class="footer-logo">🏢</div>
            <div class="footer-divider"></div>
            <div class="footer-text">© 2025 AGC-VIE</div>
            <div class="footer-text" style="color: #bbb;">Tous droits réservés</div>
            <div class="footer-divider"></div>
            <div class="footer-text" style="font-size: 0.7em; color: #ccc;">Made with ❤️</div>
        </div>
        """, unsafe_allow_html=True)
        
        return selected

# ======================== FONCTION PRINCIPALE ========================
def main():
    """Fonction principale de l'application"""
    
    # Appliquer les styles CSS
    apply_custom_css()
    
    # Vérification du timeout de session
    if st.session_state.authenticated:
        if check_session_timeout():
            return
    
    # Affichage de la page appropriée
    if not st.session_state.authenticated:
        page_login()
        return
    
    # Menu latéral
    selected = sidebar()
    
    # Mise à jour de la page courante
    st.session_state.page = selected
    
    # Navigation vers la page sélectionnée
    if selected == "Accueil":
        page_accueil()
    elif selected == "Gestion Technique":
        page_gestion_technique()
    elif selected == "Gestion Comptable":
        page_gestion_comptable()
    elif selected == "Rapprochement Technique":
        page_rapprochement_technique()
    elif selected == "Rapprochement Comptable":
        page_rapprochement_comptable()
    elif selected == "Gestion 410 & 411":
        page_gestion_410_411()
    elif selected == "Gestion Doublons":
        page_gestion_doublons()
    elif selected == "Gestion Production":
        page_gestion_production()
    elif selected == "Statistiques":
        page_statistiques()
    elif selected == "Administration":
        page_administration()
    elif selected == "Déconnexion":
        logout()

# ======================== POINT D'ENTRÉE ========================
if __name__ == "__main__":
    # Initialisation de la session
    init_session_state()
    
    # Exécution de l'application
    try:
        main()
    except Exception as e:
        st.error(f"Erreur critique: {str(e)}")
        log_action("Erreur critique", str(e), level="error")
        if os.getenv('ENVIRONMENT') == 'development':
            st.exception(e)
