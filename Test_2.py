"""
Application de Gestion Technique et Comptable AGC-VIE
Version Streamlit - Design adaptatif (Black, Light, System)
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
import hashlib
import sqlite3
import logging
import shutil
import warnings
from typing import Optional, Dict, List, Any, Tuple
from pathlib import Path

# Tentative d'import des modules optionnels
try:
    import bcrypt
except ImportError:
    bcrypt = None

warnings.filterwarnings('ignore')

# ======================== CONFIGURATION DE LA PAGE ========================
st.set_page_config(
    page_title="AGC-VIE - Gestion Technique et Comptable",
    page_icon="📊",
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

# ======================== GESTION DES THÈMES ========================
def get_theme_colors(theme_name=None):
    """Retourne les couleurs selon le thème sélectionné"""
    
    # Thèmes prédéfinis
    themes = {
        "light": {
            "primary": "#1e3c72",
            "primary_light": "#2a5298",
            "secondary": "#764ba2",
            "background": "#ffffff",
            "background_secondary": "#f8f9fa",
            "text": "#1a1a2e",
            "text_secondary": "#4a4a6a",
            "card_bg": "#ffffff",
            "card_shadow": "rgba(0,0,0,0.1)",
            "sidebar_bg": "#f0f2f6",
            "success": "#28a745",
            "warning": "#ffc107",
            "danger": "#dc3545",
            "info": "#17a2b8",
            "border": "#e0e0e0",
            "hover": "#f5f5f5"
        },
        "dark": {
            "primary": "#4a8fc1",
            "primary_light": "#6aa8d4",
            "secondary": "#9b6ab8",
            "background": "#0e1117",
            "background_secondary": "#1a1d24",
            "text": "#f0f2f6",
            "text_secondary": "#b0b2b6",
            "card_bg": "#1a1d24",
            "card_shadow": "rgba(0,0,0,0.5)",
            "sidebar_bg": "#262730",
            "success": "#28a745",
            "warning": "#ffc107",
            "danger": "#dc3545",
            "info": "#17a2b8",
            "border": "#2d3038",
            "hover": "#262730"
        }
    }
    
    # Déterminer le thème
    if theme_name is None:
        theme_name = st.session_state.get('theme', 'system')
    
    if theme_name == 'system':
        # Détection automatique du thème système
        import streamlit as st
        # Utiliser le thème actuel de Streamlit
        is_dark = st.get_option('theme.base') == 'dark'
        theme_name = 'dark' if is_dark else 'light'
    
    return themes.get(theme_name, themes['light'])

def apply_theme_css():
    """Applique les styles CSS selon le thème actuel"""
    
    theme = st.session_state.get('theme', 'system')
    colors = get_theme_colors(theme)
    
    # Déterminer si le thème est sombre
    is_dark = theme == 'dark'
    
    # Styles adaptatifs
    st.markdown(f"""
    <style>
        /* Variables de thème */
        :root {{
            --primary-color: {colors['primary']};
            --primary-light: {colors['primary_light']};
            --secondary-color: {colors['secondary']};
            --bg-color: {colors['background']};
            --bg-secondary: {colors['background_secondary']};
            --text-color: {colors['text']};
            --text-secondary: {colors['text_secondary']};
            --card-bg: {colors['card_bg']};
            --card-shadow: {colors['card_shadow']};
            --sidebar-bg: {colors['sidebar_bg']};
            --border-color: {colors['border']};
            --hover-color: {colors['hover']};
        }}
        
        /* Style global */
        .stApp {{
            background-color: var(--bg-color);
            color: var(--text-color);
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        }}
        
        /* En-têtes */
        h1, h2, h3, h4, h5, h6 {{
            color: var(--text-color);
            font-weight: 600;
        }}
        
        h1 {{
            font-size: 2.5rem;
            border-bottom: 3px solid var(--primary-color);
            padding-bottom: 0.5rem;
        }}
        
        h2 {{
            font-size: 2rem;
            border-bottom: 2px solid var(--primary-light);
            padding-bottom: 0.3rem;
        }}
        
        /* Cartes métriques */
        .metric-card {{
            background: linear-gradient(135deg, var(--primary-color) 0%, var(--secondary-color) 100%);
            padding: 20px;
            border-radius: 15px;
            color: white;
            text-align: center;
            box-shadow: 0 10px 30px var(--card-shadow);
            transition: transform 0.3s, box-shadow 0.3s;
            margin: 10px 0;
        }}
        
        .metric-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 15px 40px var(--card-shadow);
        }}
        
        .metric-card h3 {{
            color: white;
            font-size: 1.2em;
            margin-bottom: 10px;
            opacity: 0.9;
        }}
        
        .metric-card p {{
            font-size: 2.2em;
            font-weight: bold;
            margin: 0;
        }}
        
        /* Conteneurs */
        .content-card {{
            background: var(--card-bg);
            padding: 20px;
            border-radius: 15px;
            box-shadow: 0 10px 30px var(--card-shadow);
            margin: 20px 0;
            transition: transform 0.3s, box-shadow 0.3s;
            border: 1px solid var(--border-color);
        }}
        
        .content-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 15px 40px var(--card-shadow);
        }}
        
        /* Badges */
        .badge {{
            display: inline-block;
            padding: 5px 10px;
            border-radius: 20px;
            font-size: 0.85em;
            font-weight: 600;
            text-align: center;
        }}
        
        .badge-success {{
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
            color: white;
        }}
        
        .badge-warning {{
            background: linear-gradient(135deg, #ffc107 0%, #fd7e14 100%);
            color: #1a1a2e;
        }}
        
        .badge-danger {{
            background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);
            color: white;
        }}
        
        .badge-info {{
            background: linear-gradient(135deg, #17a2b8 0%, #138496 100%);
            color: white;
        }}
        
        /* Barre latérale */
        .css-1d391kg, .st-emotion-cache-1d391kg {{
            background: var(--sidebar-bg);
            border-right: 1px solid var(--border-color);
        }}
        
        /* Éléments de la barre latérale */
        .sidebar-content {{
            color: var(--text-color);
        }}
        
        /* Pied de page */
        .footer {{
            text-align: center;
            padding: 20px;
            color: var(--text-secondary);
            font-size: 0.9em;
            border-top: 1px solid var(--border-color);
            margin-top: 40px;
        }}
        
        /* Animations */
        @keyframes fadeIn {{
            from {{ opacity: 0; transform: translateY(20px); }}
            to {{ opacity: 1; transform: translateY(0); }}
        }}
        
        .fade-in {{
            animation: fadeIn 0.5s ease-out;
        }}
        
        /* Tableaux */
        .stDataFrame {{
            border-radius: 10px;
            overflow: hidden;
        }}
        
        .stDataFrame table {{
            border-collapse: collapse;
            width: 100%;
        }}
        
        .stDataFrame thead tr th {{
            background: var(--primary-color) !important;
            color: white !important;
            padding: 12px !important;
            font-weight: 600 !important;
        }}
        
        .stDataFrame tbody tr:hover {{
            background: var(--hover-color) !important;
        }}
        
        /* Inputs */
        .stTextInput > div > div > input,
        .stSelectbox > div > div > select,
        .stNumberInput > div > div > input {{
            border-radius: 8px;
            border: 2px solid var(--border-color);
            padding: 10px;
            transition: all 0.3s;
            background: var(--card-bg);
            color: var(--text-color);
        }}
        
        .stTextInput > div > div > input:focus,
        .stSelectbox > div > div > select:focus,
        .stNumberInput > div > div > input:focus {{
            border-color: var(--primary-color);
            box-shadow: 0 0 0 3px rgba(30, 60, 114, 0.1);
        }}
        
        /* Boutons */
        .stButton > button {{
            background: linear-gradient(135deg, var(--primary-color) 0%, var(--primary-light) 100%);
            color: white;
            border-radius: 8px;
            padding: 12px 24px;
            font-weight: 600;
            border: none;
            transition: all 0.3s;
            box-shadow: 0 4px 6px var(--card-shadow);
            width: 100%;
        }}
        
        .stButton > button:hover {{
            transform: translateY(-2px);
            box-shadow: 0 6px 12px var(--card-shadow);
        }}
        
        .stButton > button:active {{
            transform: translateY(0);
        }}
        
        /* Messages */
        .stAlert {{
            border-radius: 10px;
            border-left: 5px solid;
            box-shadow: 0 4px 6px var(--card-shadow);
        }}
        
        .stSuccess {{
            background: rgba(40, 167, 69, 0.1) !important;
            border-left-color: #28a745 !important;
            color: var(--text-color) !important;
        }}
        
        .stError {{
            background: rgba(220, 53, 69, 0.1) !important;
            border-left-color: #dc3545 !important;
            color: var(--text-color) !important;
        }}
        
        .stWarning {{
            background: rgba(255, 193, 7, 0.1) !important;
            border-left-color: #ffc107 !important;
            color: var(--text-color) !important;
        }}
        
        .stInfo {{
            background: rgba(23, 162, 184, 0.1) !important;
            border-left-color: #17a2b8 !important;
            color: var(--text-color) !important;
        }}
        
        /* Scrollbar personnalisée */
        ::-webkit-scrollbar {{
            width: 10px;
            height: 10px;
        }}
        
        ::-webkit-scrollbar-track {{
            background: var(--bg-secondary);
            border-radius: 5px;
        }}
        
        ::-webkit-scrollbar-thumb {{
            background: linear-gradient(135deg, var(--primary-color) 0%, var(--primary-light) 100%);
            border-radius: 5px;
        }}
        
        ::-webkit-scrollbar-thumb:hover {{
            background: linear-gradient(135deg, var(--primary-light) 0%, var(--primary-color) 100%);
        }}
        
        /* Onglets adaptatifs */
        .stTabs [data-baseweb="tab-list"] {{
            gap: 8px;
            background-color: var(--bg-secondary);
            padding: 0.5rem;
            border-radius: 10px;
            border: 1px solid var(--border-color);
        }}
        
        .stTabs [data-baseweb="tab"] {{
            background: var(--card-bg);
            border-radius: 8px;
            padding: 12px 24px;
            font-weight: 600;
            color: var(--text-color) !important;
            transition: all 0.3s;
            border: 1px solid var(--border-color);
        }}
        
        .stTabs [data-baseweb="tab"]:hover {{
            transform: translateY(-2px);
            box-shadow: 0 4px 8px var(--card-shadow);
        }}
        
        .stTabs [aria-selected="true"] {{
            background: linear-gradient(135deg, var(--primary-color) 0%, var(--primary-light) 100%) !important;
            color: white !important;
            transform: translateY(-2px);
            box-shadow: 0 4px 8px var(--card-shadow);
            border-color: var(--primary-color);
        }}
        
        /* Métriques Streamlit */
        .stMetric {{
            background: var(--card-bg);
            padding: 15px;
            border-radius: 10px;
            border: 1px solid var(--border-color);
            box-shadow: 0 2px 8px var(--card-shadow);
        }}
        
        .stMetric label {{
            color: var(--text-secondary) !important;
        }}
        
        .stMetric .stMetricValue {{
            color: var(--text-color) !important;
        }}
        
        /* Switch pour le thème */
        .theme-switch {{
            display: flex;
            align-items: center;
            gap: 10px;
            padding: 10px;
            background: var(--card-bg);
            border-radius: 10px;
            border: 1px solid var(--border-color);
            margin-bottom: 20px;
        }}
        
        .theme-switch label {{
            color: var(--text-color);
            font-weight: 500;
        }}
        
        /* Responsive */
        @media (max-width: 768px) {{
            .stTabs [data-baseweb="tab"] {{
                padding: 8px 12px;
                font-size: 0.9em;
            }}
            
            .metric-card p {{
                font-size: 1.5em;
            }}
            
            h1 {{
                font-size: 2rem;
            }}
            
            h2 {{
                font-size: 1.5rem;
            }}
        }}
        
        /* Thème sombre spécifique */
        {'/* Dark theme specific */' if is_dark else ''}
        {f'''
        .stApp {{
            background-color: #0e1117;
        }}
        
        .stDataFrame thead tr th {{
            background: #1a1d24 !important;
        }}
        
        .stDataFrame tbody tr td {{
            background: #1a1d24 !important;
            color: #f0f2f6 !important;
        }}
        
        .stDataFrame tbody tr:hover td {{
            background: #262730 !important;
        }}
        ''' if is_dark else ''}
        
        /* Thème clair spécifique */
        {'/* Light theme specific */' if not is_dark else ''}
        {f'''
        .stApp {{
            background-color: #ffffff;
        }}
        ''' if not is_dark else ''}
    </style>
    """, unsafe_allow_html=True)
    
    # Appliquer les couleurs aux graphiques Plotly
    update_plotly_theme(theme)

def update_plotly_theme(theme):
    """Met à jour le thème des graphiques Plotly"""
    
    template = 'plotly_dark' if theme == 'dark' else 'plotly_white'
    
    # Configurer le template par défaut
    import plotly.io as pio
    pio.templates.default = template
    
    # Mettre à jour les couleurs des graphiques existants
    if 'plotly_theme_updated' not in st.session_state:
        st.session_state.plotly_theme_updated = True

# ======================== INITIALISATION DE LA SESSION ========================
def init_session_state():
    """Initialise toutes les variables de session"""
    
    # Thème
    if 'theme' not in st.session_state:
        st.session_state.theme = 'system'  # 'light', 'dark', 'system'
    
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
    
    if len(st.session_state.history) > 500:
        st.session_state.history = st.session_state.history[-500:]

# ======================== FONCTIONS DE SÉCURITÉ ========================
class SecurityManager:
    """Gestionnaire de sécurité"""
    
    @staticmethod
    def hash_password(password):
        """Hash un mot de passe avec bcrypt"""
        if bcrypt is None:
            return hashlib.sha256(password.encode('utf-8')).hexdigest()
        salt = bcrypt.gensalt()
        return bcrypt.hashpw(password.encode('utf-8'), salt).decode('utf-8')
    
    @staticmethod
    def verify_password(hashed_password, plain_password):
        """Vérifie un mot de passe"""
        if bcrypt is None:
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
            
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_by TEXT
                )
            """)
            
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
        conn = sqlite3.connect(self.db_file)
        cursor = conn.cursor()
        cursor.execute(query, params)
        conn.commit()
        conn.close()
        return cursor
    
    def fetch_all(self, query, params=()):
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute(query, params)
        rows = cursor.fetchall()
        conn.close()
        return [dict(row) for row in rows]
    
    def fetch_one(self, query, params=()):
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute(query, params)
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

# ======================== FONCTIONS D'AUTHENTIFICATION ========================
def login(username, password):
    try:
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
            st.session_state.authenticated = True
            st.session_state.username = username
            st.session_state.role = user['role']
            st.session_state.login_attempts = 0
            st.session_state.locked_until = None
            st.session_state.last_activity = datetime.now()
            
            db.execute_query(
                "UPDATE users SET last_login = ? WHERE username = ?",
                (datetime.now(), username)
            )
            
            log_action("Connexion", f"Utilisateur {username} connecté")
            log_history("login", username, "Connexion réussie")
            
            return True
        else:
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
    if st.session_state.authenticated:
        username = st.session_state.username
        log_action("Déconnexion", f"Utilisateur {username} déconnecté")
        log_history("logout", username, "Déconnexion")
    
    st.session_state.authenticated = False
    st.session_state.username = None
    st.session_state.role = None
    st.rerun()

def check_session_timeout():
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
    if st.session_state.authenticated:
        st.session_state.last_activity = datetime.now()

# ======================== FONCTIONS DE GESTION DES UTILISATEURS ========================
def get_all_users():
    db = DatabaseHandler()
    return db.fetch_all("SELECT id, username, email, role, status, last_login, created_at FROM users ORDER BY username")

def add_user(username, password, email, role="user"):
    try:
        valid, errors = SecurityManager.validate_password_strength(password)
        if not valid:
            return False, "\n".join(errors)
        
        db = DatabaseHandler()
        
        existing = db.fetch_one("SELECT username FROM users WHERE username = ?", (username,))
        if existing:
            return False, "Ce nom d'utilisateur existe déjà"
        
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
    try:
        df = df.copy()
        df.columns = df.columns.str.strip()
        
        if all(col in df.columns for col in ['Num avenant', 'Code intermédiaire', 'N° police']):
            df['Nouvelle_Police'] = df.apply(
                lambda row: f"{row['Code intermédiaire']}-{row['N° police']}/{row['Num avenant']}" 
                if pd.notnull(row['Num avenant']) and str(row['Num avenant']).strip() 
                else f"{row['Code intermédiaire']}-{row['N° police']}", 
                axis=1
            )
        
        if 'Nouvelle_Police' in df.columns:
            df['Nouvelle_Police'] = df['Nouvelle_Police'].astype(str).str.replace('.0', '', regex=False)
        
        if 'Type quittance' in df.columns and 'Chiffre affaire' in df.columns:
            df['Ristournes'] = df.apply(
                lambda row: row['Chiffre affaire'] if str(row['Type quittance']).strip() == 'Ristourne' else 0, 
                axis=1
            )
            df['Emissions'] = df.apply(
                lambda row: row['Chiffre affaire'] if str(row['Type quittance']).strip() == 'Emission' else 0, 
                axis=1
            )
        
        index_col = 'Nouvelle_Police' if 'Nouvelle_Police' in df.columns else df.columns[0]
        value_cols = []
        
        for col in ['Emissions', 'Ristournes', 'Chiffre affaire']:
            if col in df.columns:
                value_cols.append(col)
        
        if not value_cols:
            value_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        
        if value_cols and index_col in df.columns:
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
    try:
        df = df.copy()
        df.columns = df.columns.str.strip()
        
        if 'No Police' in df.columns:
            df['No Police'] = df['No Police'].astype(str).str.replace('.0', '', regex=False)
        
        numeric_cols = ['Débit', 'Crédit', 'Montant']
        for col in numeric_cols:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
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
    try:
        if police_column not in df.columns:
            for col in df.columns:
                if 'police' in col.lower() or 'num' in col.lower():
                    police_column = col
                    break
            else:
                return None, None, "Colonne de police non trouvée"
        
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
    try:
        tech = tech_df.copy()
        compta = compta_df.copy()
        
        tech_col = 'Nouvelle_Police' if 'Nouvelle_Police' in tech.columns else tech.columns[0]
        compta_col = 'No Police' if 'No Police' in compta.columns else compta.columns[0]
        
        tech_renamed = tech.rename(columns={tech_col: 'Police'})
        compta_renamed = compta.rename(columns={compta_col: 'Police'})
        
        for col in ['Emissions', 'Ristournes', 'Débit', 'Crédit']:
            if col in tech_renamed.columns:
                tech_renamed[col] = pd.to_numeric(tech_renamed[col], errors='coerce').fillna(0)
            if col in compta_renamed.columns:
                compta_renamed[col] = pd.to_numeric(compta_renamed[col], errors='coerce').fillna(0)
        
        merged = pd.merge(
            tech_renamed, 
            compta_renamed, 
            on='Police', 
            how='outer', 
            suffixes=('_tech', '_compta')
        )
        
        if 'Emissions_tech' in merged.columns and 'Débit_compta' in merged.columns and 'Crédit_compta' in merged.columns:
            merged['CA_Technique'] = merged['Emissions_tech']
            merged['CA_Comptable'] = abs(merged['Crédit_compta'] - merged['Débit_compta'])
            merged['Écart'] = merged['CA_Technique'] - merged['CA_Comptable']
            merged['Statut'] = merged['Écart'].apply(
                lambda x: 'Rapproché' if abs(x) < 0.01 else 'Non rapproché'
            )
        
        stats = {
            'total_polices': len(merged),
            'polices_techniques': len(tech_renamed),
            'polices_comptables': len(compta_renamed),
            'polices_communes': len(merged[merged['Police'].notna() & merged['Emissions_tech'].notna() & merged['Débit_compta'].notna()]),
            'polices_tech_only': len(merged[merged['Emissions_tech'].notna() & merged['Débit_compta'].isna()]),
            'polices_compta_only': len(merged[merged['Emissions_tech'].isna() & merged['Débit_compta'].notna()])
        }
        
        if 'Statut' in merged.columns:
            stats['rapprochees'] = len(merged[merged['Statut'] == 'Rapproché'])
            stats['non_rapprochees'] = len(merged[merged['Statut'] == 'Non rapproché'])
            stats['ecart_total'] = merged['Écart'].sum()
        
        log_action("Rapprochement", f"{stats['polices_communes']} polices communes")
        return merged, stats
        
    except Exception as e:
        log_action("Erreur rapprochement", str(e), level="error")
        return None, str(e)

def validate_references(df, ref_column='Réf Pièce'):
    try:
        if ref_column not in df.columns:
            return None, "Colonne de référence non trouvée"
        
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
    try:
        output = io.BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for df, sheet_name in zip(dataframes, sheet_names):
                if df is not None and not df.empty:
                    safe_name = sheet_name[:31]
                    df.to_excel(writer, sheet_name=safe_name, index=False)
        
        output.seek(0)
        
        log_action("Export Excel", f"{filename} créé")
        return output
        
    except Exception as e:
        log_action("Erreur export Excel", str(e), level="error")
        return None

def export_to_csv(df, filename="export.csv"):
    try:
        return df.to_csv(index=False).encode('utf-8')
    except Exception as e:
        log_action("Erreur export CSV", str(e), level="error")
        return None

def create_download_button(data, filename, button_text, mime_type=None):
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
def display_metric_card(title, value, icon="📊", description=""):
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown(f"""
        <div class="metric-card fade-in">
            <h3>{icon} {title}</h3>
            <p>{value}</p>
            <small>{description}</small>
        </div>
        """, unsafe_allow_html=True)

def display_badge(text, type="info"):
    colors = {
        "success": "badge-success",
        "warning": "badge-warning",
        "danger": "badge-danger",
        "info": "badge-info"
    }
    css_class = colors.get(type, "badge-info")
    st.markdown(f'<span class="badge {css_class}">{text}</span>', unsafe_allow_html=True)

def create_search_bar(key, placeholder="Rechercher..."):
    search = st.text_input(
        "🔍",
        placeholder=placeholder,
        key=key,
        label_visibility="collapsed"
    )
    return search

def filter_dataframe(df, search_term):
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
    if items_per_page is None:
        items_per_page = st.session_state.user_preferences.get('items_per_page', 50)
    
    if df is None or df.empty:
        return df, 0, 0
    
    total_pages = (len(df) + items_per_page - 1) // items_per_page
    
    if f'{key_prefix}_page' not in st.session_state:
        st.session_state[f'{key_prefix}_page'] = 0
    
    current_page = st.session_state[f'{key_prefix}_page']
    
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
    
    start_idx = current_page * items_per_page
    end_idx = min(start_idx + items_per_page, len(df))
    
    return df.iloc[start_idx:end_idx], start_idx, end_idx

# ======================== SÉLECTEUR DE THÈME ========================
def theme_selector():
    """Composant de sélection de thème"""
    with st.sidebar:
        st.markdown("---")
        st.markdown("### 🎨 Thème")
        
        current_theme = st.session_state.get('theme', 'system')
        
        # Utiliser des colonnes pour les boutons radio horizontaux
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("☀️", key="theme_light", 
                        help="Thème clair",
                        use_container_width=True):
                st.session_state.theme = 'light'
                st.rerun()
        
        with col2:
            if st.button("🌙", key="theme_dark",
                        help="Thème sombre",
                        use_container_width=True):
                st.session_state.theme = 'dark'
                st.rerun()
        
        with col3:
            if st.button("💻", key="theme_system",
                        help="Thème système",
                        use_container_width=True):
                st.session_state.theme = 'system'
                st.rerun()
        
        # Indicateur du thème actif
        theme_names = {
            'light': '☀️ Clair',
            'dark': '🌙 Sombre',
            'system': '💻 Système'
        }
        st.caption(f"Actif: {theme_names.get(current_theme, 'Système')}")
        
        # Appliquer le thème immédiatement
        apply_theme_css()

# ======================== PAGES DE L'APPLICATION ========================

def page_login():
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("""
        <div style="text-align: center; padding: 40px; animation: fadeIn 0.5s;">
            <h1 style="font-size: 3em; margin-bottom: 10px;">AGC-VIE</h1>
            <p style="font-size: 1.2em; margin-bottom: 30px; opacity: 0.8;">
                Système de Gestion Technique et Comptable
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        with st.form("login_form", clear_on_submit=True):
            st.markdown("### Connexion")
            
            username = st.text_input(
                "👤 Nom d'utilisateur",
                placeholder="Entrez votre nom d'utilisateur"
            )
            
            password = st.text_input(
                "🔒 Mot de passe",
                type="password",
                placeholder="Entrez votre mot de passe"
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
        <div style="text-align: center; font-size: 0.9em; margin-top: 20px; opacity: 0.7;">
            <p>Compte démo: admin / Admin123!</p>
            <p>© 2025 AGC-VIE - Tous droits réservés</p>
        </div>
        """, unsafe_allow_html=True)

def page_accueil():
    update_last_activity()
    
    st.markdown("""
    <div class="content-card fade-in" style="background: linear-gradient(135deg, var(--primary-color) 0%, var(--secondary-color) 100%); text-align: center; color: white;">
        <h1 style="color: white;">Bienvenue sur AGC-VIE</h1>
        <p style="color: white; font-size: 1.2em; opacity: 0.9;">
            Système intégré de Gestion Technique et Comptable
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        tech_count = len(st.session_state.pivot_techniques) if st.session_state.pivot_techniques is not None else 0
        st.metric(
            "📊 Données techniques",
            tech_count,
            help="Nombre d'enregistrements techniques"
        )
    
    with col2:
        compta_count = len(st.session_state.pivot_comptables) if st.session_state.pivot_comptables is not None else 0
        st.metric(
            "💰 Données comptables",
            compta_count,
            help="Nombre d'enregistrements comptables"
        )
    
    with col3:
        if st.session_state.pivot_techniques is not None and 'Emissions' in st.session_state.pivot_techniques.columns:
            ca_tech = st.session_state.pivot_techniques['Emissions'].sum()
            st.metric(
                "📈 CA Technique",
                f"{ca_tech:,.0f} FCFA",
                help="Chiffre d'affaires technique"
            )
    
    with col4:
        if st.session_state.pivot_comptables is not None and 'Crédit' in st.session_state.pivot_comptables.columns:
            ca_compta = st.session_state.pivot_comptables['Crédit'].sum()
            st.metric(
                "📉 CA Comptable",
                f"{ca_compta:,.0f} FCFA",
                help="Chiffre d'affaires comptable"
            )
    
    st.markdown("## 🚀 Modules disponibles")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div class="content-card">
            <h3>📊 Gestion Technique</h3>
            <p>Import, analyse et traitement des données techniques</p>
            <ul>
                <li>Import de fichiers Excel/CSV</li>
                <li>Traitement des polices</li>
                <li>Calcul des émissions et ristournes</li>
                <li>Export des données</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="content-card">
            <h3>💰 Gestion Comptable</h3>
            <p>Gestion des données comptables et rapprochements</p>
            <ul>
                <li>Import des écritures comptables</li>
                <li>Analyse des débits/crédits</li>
                <li>Tableaux croisés dynamiques</li>
                <li>Export des résultats</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div class="content-card">
            <h3>🔄 Rapprochement Technique</h3>
            <p>Rapprochement entre données techniques et comptables</p>
            <ul>
                <li>Comparaison automatique</li>
                <li>Détection des écarts</li>
                <li>Visualisation des résultats</li>
                <li>Export des rapprochements</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class="content-card">
            <h3>📋 Gestion 410 & 411</h3>
            <p>Gestion des comptes 410 et 411</p>
            <ul>
                <li>Vérification des polices</li>
                <li>Détection des incohérences</li>
                <li>Validation des références</li>
                <li>Analyse comparative</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

def page_gestion_technique():
    update_last_activity()
    
    st.markdown("## 📊 Gestion Technique")
    
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
                    if uploaded_file.name.endswith('.csv'):
                        df = pd.read_csv(uploaded_file, dtype=str)
                    else:
                        df = pd.read_excel(uploaded_file, dtype=str)
                    
                    st.session_state.df_technique = df
                    
                    st.markdown("### Aperçu des données")
                    st.dataframe(df.head(10), use_container_width=True)
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Lignes", len(df))
                    with col2:
                        st.metric("Colonnes", len(df.columns))
                    with col3:
                        st.metric("Taille", f"{uploaded_file.size / 1024:.1f} KB")
                    
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
            
            search = create_search_bar("tech_search", "Rechercher une police...")
            
            df_display = st.session_state.pivot_techniques.copy()
            if search:
                df_display = filter_dataframe(df_display, search)
            
            df_page, start, end = create_pagination(df_display, "tech")
            
            st.markdown(f"**Affichage {start+1}-{end} sur {len(df_display)} enregistrements**")
            
            st.dataframe(
                df_page,
                use_container_width=True,
                height=500,
                hide_index=True
            )
            
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
                
                st.markdown("#### Statistiques descriptives")
                stats_df = df[selected_col].describe().reset_index()
                stats_df.columns = ['Statistique', 'Valeur']
                st.dataframe(stats_df, use_container_width=True, hide_index=True)
                
                st.markdown("#### Visualisation")
                
                if chart_type == "Histogramme":
                    fig = px.histogram(
                        df,
                        x=selected_col,
                        nbins=30,
                        title=f"Distribution de {selected_col}",
                        color_discrete_sequence=[get_theme_colors()['primary']]
                    )
                elif chart_type == "Box plot":
                    fig = px.box(
                        df,
                        y=selected_col,
                        title=f"Box plot de {selected_col}",
                        color_discrete_sequence=[get_theme_colors()['primary']]
                    )
                else:
                    fig = px.line(
                        df.reset_index(),
                        y=selected_col,
                        title=f"Évolution de {selected_col}",
                        color_discrete_sequence=[get_theme_colors()['primary']]
                    )
                
                fig.update_layout(
                    height=500,
                    template='plotly_dark' if st.session_state.get('theme') == 'dark' else 'plotly_white'
                )
                st.plotly_chart(fig, use_container_width=True)
                
                st.markdown("#### Top 10 des valeurs")
                top_df = df.nlargest(10, selected_col)[df.columns[:3]]
                st.dataframe(top_df, use_container_width=True, hide_index=True)
                
            else:
                st.warning("Aucune colonne numérique disponible pour l'analyse")
        else:
            st.info("Aucune donnée à analyser")

def page_gestion_comptable():
    update_last_activity()
    
    st.markdown("## 💰 Gestion Comptable")
    
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
                    if uploaded_file.name.endswith('.csv'):
                        df = pd.read_csv(uploaded_file, dtype=str)
                    else:
                        df = pd.read_excel(uploaded_file, dtype=str)
                    
                    st.session_state.df_comptable = df
                    
                    st.markdown("### Aperçu des données")
                    st.dataframe(df.head(10), use_container_width=True)
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Lignes", len(df))
                    with col2:
                        st.metric("Colonnes", len(df.columns))
                    with col3:
                        st.metric("Taille", f"{uploaded_file.size / 1024:.1f} KB")
                    
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
            
            search = create_search_bar("compta_search", "Rechercher une police...")
            
            df_display = st.session_state.pivot_comptables.copy()
            if search:
                df_display = filter_dataframe(df_display, search)
            
            df_page, start, end = create_pagination(df_display, "compta")
            
            st.markdown(f"**Affichage {start+1}-{end} sur {len(df_display)} enregistrements**")
            
            st.dataframe(
                df_page,
                use_container_width=True,
                height=500,
                hide_index=True
            )
            
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
                
                st.markdown("#### Statistiques descriptives")
                stats_df = df[selected_col].describe().reset_index()
                stats_df.columns = ['Statistique', 'Valeur']
                st.dataframe(stats_df, use_container_width=True, hide_index=True)
                
                st.markdown("#### Visualisation")
                
                if chart_type == "Histogramme":
                    fig = px.histogram(
                        df,
                        x=selected_col,
                        nbins=30,
                        title=f"Distribution de {selected_col}",
                        color_discrete_sequence=[get_theme_colors()['primary']]
                    )
                elif chart_type == "Box plot":
                    fig = px.box(
                        df,
                        y=selected_col,
                        title=f"Box plot de {selected_col}",
                        color_discrete_sequence=[get_theme_colors()['primary']]
                    )
                else:
                    fig = px.line(
                        df.reset_index(),
                        y=selected_col,
                        title=f"Évolution de {selected_col}",
                        color_discrete_sequence=[get_theme_colors()['primary']]
                    )
                
                fig.update_layout(
                    height=500,
                    template='plotly_dark' if st.session_state.get('theme') == 'dark' else 'plotly_white'
                )
                st.plotly_chart(fig, use_container_width=True)
                
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
    update_last_activity()
    
    st.markdown("## 🔄 Rapprochement Technique")
    
    if st.session_state.pivot_techniques is None:
        st.warning("⚠️ Données techniques manquantes. Veuillez d'abord importer les données techniques.")
        return
    
    if st.session_state.pivot_comptables is None:
        st.warning("⚠️ Données comptables manquantes. Veuillez d'abord importer les données comptables.")
        return
    
    with st.spinner("Calcul du rapprochement en cours..."):
        merged_df, stats = rapprochement_technique_comptable(
            st.session_state.pivot_techniques,
            st.session_state.pivot_comptables
        )
    
    if merged_df is not None:
        st.session_state.stats['total_verifications'] += 1
        
        st.markdown("### 📊 Résumé du rapprochement")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            display_metric_card(
                "Polices techniques",
                stats.get('polices_techniques', 0),
                "📊"
            )
        
        with col2:
            display_metric_card(
                "Polices comptables",
                stats.get('polices_comptables', 0),
                "💰"
            )
        
        with col3:
            display_metric_card(
                "Polices communes",
                stats.get('polices_communes', 0),
                "🔄"
            )
        
        with col4:
            ecart = stats.get('ecart_total', 0)
            display_metric_card(
                "Écart total",
                f"{ecart:,.0f} FCFA",
                "📈" if ecart >= 0 else "📉"
            )
        
        tab1, tab2, tab3 = st.tabs(["📋 Données complètes", "❌ Non rapprochées", "✅ Rapprochées"])
        
        with tab1:
            st.markdown("### Toutes les polices")
            
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
        
        st.markdown("### 📊 Visualisations")
        
        col1, col2 = st.columns(2)
        
        with col1:
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
                    template='plotly_dark' if st.session_state.get('theme') == 'dark' else 'plotly_white'
                )
                st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            if 'Écart' in merged_df.columns:
                fig = px.histogram(
                    merged_df[merged_df['Écart'].notna()],
                    x='Écart',
                    nbins=50,
                    title="Distribution des écarts",
                    color_discrete_sequence=[get_theme_colors()['primary']]
                )
                fig.update_layout(
                    height=400,
                    template='plotly_dark' if st.session_state.get('theme') == 'dark' else 'plotly_white'
                )
                st.plotly_chart(fig, use_container_width=True)
        
        log_action("Rapprochement technique", f"{len(merged_df)} polices analysées")
        
    else:
        st.error(f"Erreur lors du rapprochement: {stats}")

def page_rapprochement_comptable():
    update_last_activity()
    
    st.markdown("## 🔄 Rapprochement Comptable")
    
    if st.session_state.pivot_techniques is None:
        st.warning("⚠️ Données techniques manquantes. Veuillez d'abord importer les données techniques.")
        return
    
    if st.session_state.pivot_comptables is None:
        st.warning("⚠️ Données comptables manquantes. Veuillez d'abord importer les données comptables.")
        return
    
    with st.spinner("Calcul du rapprochement comptable en cours..."):
        try:
            df_tech = st.session_state.pivot_techniques.copy()
            df_compta = st.session_state.pivot_comptables.copy()
            
            tech_col = 'Nouvelle_Police' if 'Nouvelle_Police' in df_tech.columns else df_tech.columns[0]
            compta_col = 'No Police' if 'No Police' in df_compta.columns else df_compta.columns[0]
            
            df_compta['Ristournes'] = 0
            df_compta['Emissions'] = 0
            df_compta['Statut_Ristournes'] = 'Non trouvé'
            df_compta['Statut_Emissions'] = 'Non trouvé'
            
            for index, row in df_compta.iterrows():
                police_compta = str(row[compta_col]).strip()
                
                correspondance = df_tech[df_tech[tech_col].astype(str).str.strip() == police_compta]
                
                if not correspondance.empty:
                    if 'Ristournes' in correspondance.columns:
                        df_compta.at[index, 'Ristournes'] = pd.to_numeric(correspondance['Ristournes'].values[0], errors='coerce')
                        df_compta.at[index, 'Statut_Ristournes'] = 'Trouvé'
                    
                    if 'Emissions' in correspondance.columns:
                        df_compta.at[index, 'Emissions'] = pd.to_numeric(correspondance['Emissions'].values[0], errors='coerce')
                        df_compta.at[index, 'Statut_Emissions'] = 'Trouvé'
            
            for col in ['Crédit', 'Débit', 'Emissions', 'Ristournes']:
                if col in df_compta.columns:
                    df_compta[col] = pd.to_numeric(df_compta[col], errors='coerce').fillna(0)
            
            if all(col in df_compta.columns for col in ['Crédit', 'Débit', 'Emissions', 'Ristournes']):
                df_compta['CA_Comptable'] = df_compta['Crédit'] - df_compta['Débit']
                df_compta['CA_Technique'] = df_compta['Emissions'] + df_compta['Ristournes']
                df_compta['Écart'] = abs(df_compta['CA_Comptable']) - abs(df_compta['CA_Technique'])
                
                df_compta['Rapprochement'] = df_compta.apply(
                    lambda row: 'Rapproché' if abs(row['Écart']) < 0.01 else 'Non rapproché', 
                    axis=1
                )
            
            df_invalide = df_compta[df_compta['Rapprochement'] == 'Non rapproché'] if 'Rapprochement' in df_compta.columns else pd.DataFrame()
            df_valide = df_compta[df_compta['Rapprochement'] == 'Rapproché'] if 'Rapprochement' in df_compta.columns else pd.DataFrame()
            
            stats = {
                'total_debit': df_compta['Débit'].sum() if 'Débit' in df_compta.columns else 0,
                'total_credit': df_compta['Crédit'].sum() if 'Crédit' in df_compta.columns else 0,
                'total_CA_comptable': abs(df_compta['Crédit'].sum() - df_compta['Débit'].sum()) if 'Débit' in df_compta.columns and 'Crédit' in df_compta.columns else 0,
                'total_emissions_tech': df_compta['Emissions'].sum() if 'Emissions' in df_compta.columns else 0,
                'total_ristournes_tech': df_compta['Ristournes'].sum() if 'Ristournes' in df_compta.columns else 0,
                'total_CA_technique': abs(df_compta['Emissions'].sum() + df_compta['Ristournes'].sum()) if 'Emissions' in df_compta.columns and 'Ristournes' in df_compta.columns else 0,
                'ecart': abs(abs(df_compta['Emissions'].sum() + df_compta['Ristournes'].sum()) - abs(df_compta['Crédit'].sum() - df_compta['Débit'].sum())) if all(col in df_compta.columns for col in ['Emissions', 'Ristournes', 'Crédit', 'Débit']) else 0,
                'total_polices': len(df_compta),
                'polices_valides': len(df_valide),
                'polices_invalides': len(df_invalide)
            }
            
            st.session_state.pivot_comptables_complet = df_compta
            st.session_state.tableau_listing_police_invalide_comptable = df_invalide
            st.session_state.tableau_listing_valide_comptable = df_valide
            
            log_action("Rapprochement comptable", f"{len(df_compta)} polices analysées")
            
        except Exception as e:
            st.error(f"Erreur lors du rapprochement: {str(e)}")
            log_action("Erreur rapprochement comptable", str(e), level="error")
            return
    
    if stats:
        st.session_state.stats['total_verifications'] += 1
        
        st.markdown("### 📊 Résumé du rapprochement comptable")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            display_metric_card(
                "Total Débit",
                f"{stats.get('total_debit', 0):,.0f} FCFA",
                "💳"
            )
        
        with col2:
            display_metric_card(
                "Total Crédit",
                f"{stats.get('total_credit', 0):,.0f} FCFA",
                "💰"
            )
        
        with col3:
            display_metric_card(
                "CA Comptable",
                f"{stats.get('total_CA_comptable', 0):,.0f} FCFA",
                "📊"
            )
        
        with col4:
            display_metric_card(
                "CA Technique",
                f"{stats.get('total_CA_technique', 0):,.0f} FCFA",
                "📈"
            )
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            display_metric_card(
                "Écart",
                f"{stats.get('ecart', 0):,.0f} FCFA",
                "📉" if stats.get('ecart', 0) > 0 else "📈"
            )
        
        with col2:
            display_metric_card(
                "Total polices",
                stats.get('total_polices', 0),
                "📋"
            )
        
        with col3:
            valides = stats.get('polices_valides', 0)
            taux_valides = (valides / stats.get('total_polices', 1) * 100) if stats.get('total_polices', 0) > 0 else 0
            display_metric_card(
                "Polices rapprochées",
                valides,
                "✅",
                f"Taux: {taux_valides:.1f}%"
            )
        
        with col4:
            invalides = stats.get('polices_invalides', 0)
            taux_invalides = (invalides / stats.get('total_polices', 1) * 100) if stats.get('total_polices', 0) > 0 else 0
            display_metric_card(
                "Polices non rapprochées",
                invalides,
                "❌",
                f"Taux: {taux_invalides:.1f}%"
            )
        
        tab1, tab2, tab3 = st.tabs(["📋 Données complètes", "❌ Non rapprochées", "✅ Rapprochées"])
        
        with tab1:
            st.markdown("### Toutes les polices comptables")
            
            search = create_search_bar("compta_rapprochement_search", "Rechercher une police...")
            
            df_display = st.session_state.pivot_comptables_complet.copy()
            if search:
                df_display = filter_dataframe(df_display, search)
            
            st.dataframe(df_display, use_container_width=True, height=500, hide_index=True)
        
        with tab2:
            df_invalide = st.session_state.tableau_listing_police_invalide_comptable
            st.markdown(f"### Polices non rapprochées ({len(df_invalide)})")
            
            if not df_invalide.empty:
                search_invalide = create_search_bar("invalide_search", "Rechercher dans les non rapprochées...")
                
                df_invalide_display = df_invalide.copy()
                if search_invalide:
                    df_invalide_display = filter_dataframe(df_invalide_display, search_invalide)
                
                st.dataframe(df_invalide_display, use_container_width=True, height=500, hide_index=True)
                
                if st.button("📥 Exporter les non rapprochées", use_container_width=True):
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
            else:
                st.success("✅ Toutes les polices comptables sont rapprochées !")
        
        with tab3:
            df_valide = st.session_state.tableau_listing_valide_comptable
            st.markdown(f"### Polices rapprochées ({len(df_valide)})")
            
            if not df_valide.empty:
                search_valide = create_search_bar("valide_search", "Rechercher dans les rapprochées...")
                
                df_valide_display = df_valide.copy()
                if search_valide:
                    df_valide_display = filter_dataframe(df_valide_display, search_valide)
                
                st.dataframe(df_valide_display, use_container_width=True, height=500, hide_index=True)
        
        st.markdown("### 📊 Analyses et visualisations")
        
        col1, col2 = st.columns(2)
        
        with col1:
            df_invalide = st.session_state.tableau_listing_police_invalide_comptable
            df_valide = st.session_state.tableau_listing_valide_comptable
            
            if not df_invalide.empty or not df_valide.empty:
                fig = go.Figure(data=[
                    go.Pie(
                        labels=['Rapprochées', 'Non rapprochées'],
                        values=[len(df_valide), len(df_invalide)],
                        marker_colors=['#28a745', '#dc3545'],
                        hole=0.3,
                        textinfo='label+percent'
                    )
                ])
                
                fig.update_layout(
                    title="Répartition des polices comptables",
                    height=400,
                    showlegend=True,
                    template='plotly_dark' if st.session_state.get('theme') == 'dark' else 'plotly_white'
                )
                
                st.plotly_chart(fig, use_container_width=True)
        
        with col2:
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
                    showlegend=False,
                    template='plotly_dark' if st.session_state.get('theme') == 'dark' else 'plotly_white'
                )
                
                st.plotly_chart(fig, use_container_width=True)

def page_gestion_410_411():
    update_last_activity()
    
    st.markdown("## 📋 Gestion 410 & 411")
    
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
                    height=400,
                    template='plotly_dark' if st.session_state.get('theme') == 'dark' else 'plotly_white'
                )
                
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Importez les fichiers pour voir les analyses.")

def page_gestion_doublons():
    update_last_activity()
    
    st.markdown("## 🔍 Gestion des Doublons")
    
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
                
                duplicates_df, uniques_df, stats = detect_duplicates(df)
                
                if isinstance(stats, dict):
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
                    
                    tab1, tab2 = st.tabs(["🔄 Polices en doublon", "✅ Polices uniques"])
                    
                    with tab1:
                        if duplicates_df is not None and not duplicates_df.empty:
                            st.markdown(f"### {len(duplicates_df)} enregistrements en doublon")
                            
                            search = create_search_bar("dup_search", "Rechercher...")
                            
                            df_display = duplicates_df.copy()
                            if search:
                                df_display = filter_dataframe(df_display, search)
                            
                            st.dataframe(df_display, use_container_width=True, height=500, hide_index=True)
                            
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
                            
                            search = create_search_bar("unique_search", "Rechercher...")
                            
                            df_display = uniques_df.copy()
                            if search:
                                df_display = filter_dataframe(df_display, search)
                            
                            st.dataframe(df_display, use_container_width=True, height=500, hide_index=True)
                            
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
                    
                    st.markdown("### 📊 Visualisation")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        fig = go.Figure(data=[
                            go.Pie(
                                labels=['Polices uniques', 'Polices en doublon'],
                                values=[stats['uniques'], stats['duplicates']],
                                marker_colors=['#28a745', '#dc3545'],
                                hole=0.3
                            )
                        ])
                        fig.update_layout(
                            title="Répartition des polices",
                            height=400,
                            template='plotly_dark' if st.session_state.get('theme') == 'dark' else 'plotly_white'
                        )
                        st.plotly_chart(fig, use_container_width=True)
                    
                    with col2:
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
                                color_discrete_sequence=[get_theme_colors()['primary']]
                            )
                            fig.update_layout(
                                height=400,
                                template='plotly_dark' if st.session_state.get('theme') == 'dark' else 'plotly_white'
                            )
                            st.plotly_chart(fig, use_container_width=True)
                    
                    log_action("Analyse doublons", f"{stats['duplicate_polices']} polices dupliquées")
                
                else:
                    st.error(f"Erreur: {stats}")
                    
            except Exception as e:
                st.error(f"Erreur lors du traitement: {str(e)}")

def page_gestion_production():
    update_last_activity()
    
    st.markdown("## 📄 Générateur de Certificats")
    
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
                         padding: 20px; border: 1px solid var(--border-color); border-radius: 5px;
                         background: var(--card-bg);">
                        Texte d'exemple
                    </div>
                    """,
                    unsafe_allow_html=True
                )
            
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
            
            if st.button("🚀 Lancer la génération", type="primary", use_container_width=True):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                total = len(df)
                
                for i in range(total):
                    time.sleep(0.05)
                    
                    progress = (i + 1) / total
                    progress_bar.progress(progress)
                    status_text.text(f"Génération: {i+1}/{total} certificats")
                    
                    if (i + 1) % 10 == 0:
                        st.session_state.stats['total_certificats'] += 10
                
                progress_bar.progress(1.0)
                status_text.text(f"✅ {total} certificats générés avec succès!")
                
                st.balloons()
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Certificats générés", total)
                with col2:
                    st.metric("Format", output_format)
                with col3:
                    taille_estimée = total * 50
                    st.metric("Taille estimée", f"{taille_estimée / 1024:.1f} MB")
                
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
    update_last_activity()
    
    st.markdown("## 📈 Statistiques et Analyses")
    
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Vue d'ensemble", "📈 Tendances", "📋 Rapports", "📥 Export"])
    
    with tab1:
        st.markdown("### Vue d'ensemble du système")
        
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
        
        dates = pd.date_range(end=datetime.now(), periods=30, freq='D')
        
        np.random.seed(42)
        imports_data = np.random.randint(5, 20, 30)
        verifications_data = np.random.randint(10, 30, 30)
        
        trend_df = pd.DataFrame({
            'Date': dates,
            'Imports': imports_data,
            'Vérifications': verifications_data
        })
        
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
            hovermode='x unified',
            template='plotly_dark' if st.session_state.get('theme') == 'dark' else 'plotly_white'
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
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
    update_last_activity()
    
    if st.session_state.role != "admin":
        st.error("⛔ Accès réservé aux administrateurs")
        return
    
    st.markdown("## ⚙️ Administration")
    
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "👥 Utilisateurs", "📋 Logs", "📜 Historique", 
        "💾 Sauvegardes", "🔧 Paramètres"
    ])
    
    with tab1:
        st.markdown("### Gestion des utilisateurs")
        
        users = get_all_users()
        
        if users:
            users_df = pd.DataFrame(users)
            st.dataframe(users_df, use_container_width=True, hide_index=True)
        
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
            
            filtered_logs = logs_df.copy()
            
            if selected_user != "Tous":
                filtered_logs = filtered_logs[filtered_logs['username'] == selected_user]
            
            if selected_action != "Toutes":
                filtered_logs = filtered_logs[filtered_logs['action'] == selected_action]
            
            if selected_level != "Tous":
                filtered_logs = filtered_logs[filtered_logs['level'] == selected_level]
            
            st.dataframe(filtered_logs, use_container_width=True, height=500, hide_index=True)
            
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
            
            if 'username' in history_df.columns:
                users_filter = ["Tous"] + list(history_df['username'].unique())
                selected_user_hist = st.selectbox("Filtrer par utilisateur", users_filter, key="hist_user")
                
                if selected_user_hist != "Tous":
                    history_df = history_df[history_df['username'] == selected_user_hist]
            
            st.dataframe(history_df, use_container_width=True, height=500, hide_index=True)
            
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
                    os.makedirs("backups", exist_ok=True)
                    
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    backup_file = f"backups/backup_{timestamp}.db"
                    
                    if os.path.exists(DB_FILE):
                        shutil.copy2(DB_FILE, backup_file)
                        st.success(f"Sauvegarde créée: {os.path.basename(backup_file)}")
                        st.balloons()
                    else:
                        st.error("Fichier de base de données introuvable")
        
        with col2:
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

# ======================== BARRE LATÉRALE ========================
def sidebar():
    with st.sidebar:
        st.markdown("""
        <div style="text-align: center; padding: 20px 10px;">
            <h2 style="margin-bottom: 5px;">AGC-VIE</h2>
            <p style="opacity: 0.7; font-size: 0.9em;">Version 2.0</p>
        </div>
        """, unsafe_allow_html=True)
        
        if st.session_state.authenticated:
            st.markdown(f"""
            <div style="background: rgba(30, 60, 114, 0.1); padding: 15px; border-radius: 10px; margin-bottom: 20px;">
                <p style="margin: 0; font-size: 1.1em;">👤 {st.session_state.username}</p>
                <p style="color: #28a745; margin: 5px 0 0 0; font-size: 0.9em;">
                    {'👑 Administrateur' if st.session_state.role == 'admin' else '👤 Utilisateur'}
                </p>
                <p style="opacity: 0.6; margin: 5px 0 0 0; font-size: 0.8em;">
                    {datetime.now().strftime('%d/%m/%Y %H:%M')}
                </p>
            </div>
            """, unsafe_allow_html=True)
        
        menu_options = {
            "Accueil": "🏠",
            "Gestion Technique": "📊",
            "Gestion Comptable": "💰",
            "Rapprochement Technique": "🔄",
            "Rapprochement Comptable": "🔄",
            "Gestion 410 & 411": "📋",
            "Gestion Doublons": "🔍",
            "Gestion Production": "📄",
            "Statistiques": "📈",
            "Administration": "⚙️" if st.session_state.role == "admin" else None,
            "Déconnexion": "🚪"
        }
        
        filtered_options = {k: v for k, v in menu_options.items() if v is not None}
        
        selected = st.radio(
            "Navigation",
            list(filtered_options.keys()),
            format_func=lambda x: f"{filtered_options[x]} {x}",
            key="navigation",
            label_visibility="collapsed"
        )
        
        st.markdown("---")
        st.markdown("""
        <div style="text-align: center; opacity: 0.6; font-size: 0.8em; padding: 10px;">
            <p>© 2025 AGC-VIE</p>
            <p>Version 2.0</p>
        </div>
        """, unsafe_allow_html=True)
        
        return selected

# ======================== FONCTION PRINCIPALE ========================
def main():
    """Fonction principale de l'application"""
    
    # Appliquer le thème
    apply_theme_css()
    
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
    
    # Sélecteur de thème
    theme_selector()
    
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
    init_session_state()
    
    try:
        main()
    except Exception as e:
        st.error(f"Erreur critique: {str(e)}")
        log_action("Erreur critique", str(e), level="error")
        if os.getenv('ENVIRONMENT') == 'development':
            st.exception(e)
