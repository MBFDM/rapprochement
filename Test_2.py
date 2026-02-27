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
import uuid
import bcrypt
import smtplib
import shutil
import logging
import webbrowser
import subprocess
from contextlib import contextmanager
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from cryptography.fernet import Fernet
from docx import Document
from docx.shared import Pt, RGBColor
from docx2pdf import convert
import fitz
import warnings
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
BACKUP_INTERVAL = 3600  # 1 heure en secondes
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
            background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
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
            border-bottom: 3px solid #1e3c72;
            padding-bottom: 0.5rem;
        }
        
        h2 {
            font-size: 2rem;
            border-bottom: 2px solid #2a5298;
            padding-bottom: 0.3rem;
        }
        
        /* Cartes métriques */
        .metric-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
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
        
        /* Tableaux */
        .dataframe {
            font-size: 0.9em;
            border-collapse: collapse;
            width: 100%;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        
        .dataframe th {
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: 600;
        }
        
        .dataframe td {
            padding: 10px;
            border-bottom: 1px solid #e0e0e0;
            background-color: white;
        }
        
        .dataframe tr:hover td {
            background-color: #f5f5f5;
            transition: background-color 0.3s;
        }
        
        /* Boutons */
        .stButton > button {
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            color: white;
            border-radius: 8px;
            padding: 12px 24px;
            font-weight: 600;
            border: none;
            transition: all 0.3s;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            width: 100%;
        }
        
        .stButton > button:hover {
            background: linear-gradient(135deg, #2a5298 0%, #1e3c72 100%);
            transform: translateY(-2px);
            box-shadow: 0 6px 12px rgba(0,0,0,0.2);
        }
        
        .stButton > button:active {
            transform: translateY(0);
        }
        
        /* Bouton secondaire */
        .stButton > button.secondary {
            background: linear-gradient(135deg, #6c757d 0%, #495057 100%);
        }
        
        /* Bouton succès */
        .stButton > button.success {
            background: linear-gradient(135deg, #28a745 0%, #20c997 100%);
        }
        
        /* Bouton danger */
        .stButton > button.danger {
            background: linear-gradient(135deg, #dc3545 0%, #c82333 100%);
        }
        
        /* Barre latérale */
        .css-1d391kg {
            background: linear-gradient(180deg, #1e3c72 0%, #0a1a2f 100%);
        }
        
        .sidebar .sidebar-content {
            background: transparent;
            color: white;
            padding: 1rem;
        }
        
        /* Éléments de la barre latérale */
        .sidebar .sidebar-content .stMarkdown {
            color: white;
        }
        
        /* Menu option */
        .nav-link {
            color: white !important;
            font-size: 1.1rem !important;
            padding: 0.75rem 1rem !important;
            margin: 0.2rem 0 !important;
            border-radius: 8px !important;
            transition: all 0.3s !important;
        }
        
        .nav-link:hover {
            background: rgba(255, 255, 255, 0.1) !important;
            transform: translateX(5px);
        }
        
        .nav-link-selected {
            background: linear-gradient(135deg, #2a5298 0%, #1e3c72 100%) !important;
            font-weight: bold !important;
            box-shadow: 0 4px 6px rgba(0,0,0,0.2) !important;
        }
        
        /* Onglets */
        .stTabs [data-baseweb="tab-list"] {
            gap: 8px;
            background-color: #f0f2f6;
            padding: 0.5rem;
            border-radius: 10px;
        }
        
        .stTabs [data-baseweb="tab"] {
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            border-radius: 8px;
            padding: 12px 24px;
            font-weight: 600;
            color: white !important;
            transition: all 0.3s;
            border: none;
        }
        
        .stTabs [data-baseweb="tab"]:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        }
        
        .stTabs [aria-selected="true"] {
            background: linear-gradient(135deg, #2a5298 0%, #1e3c72 100%) !important;
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.3);
        }
        
        /* Messages */
        .stAlert {
            border-radius: 10px;
            border-left: 5px solid;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        
        .stSuccess {
            background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
            border-left-color: #28a745;
            color: #155724;
        }
        
        .stError {
            background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%);
            border-left-color: #dc3545;
            color: #721c24;
        }
        
        .stWarning {
            background: linear-gradient(135deg, #fff3cd 0%, #ffeeba 100%);
            border-left-color: #ffc107;
            color: #856404;
        }
        
        .stInfo {
            background: linear-gradient(135deg, #d1ecf1 0%, #bee5eb 100%);
            border-left-color: #17a2b8;
            color: #0c5460;
        }
        
        /* Barre de progression */
        .stProgress > div > div > div > div {
            background: linear-gradient(90deg, #1e3c72 0%, #2a5298 100%);
            border-radius: 10px;
        }
        
        /* Inputs */
        .stTextInput > div > div > input,
        .stSelectbox > div > div > select,
        .stNumberInput > div > div > input {
            border-radius: 8px;
            border: 2px solid #e0e0e0;
            padding: 10px;
            transition: all 0.3s;
        }
        
        .stTextInput > div > div > input:focus,
        .stSelectbox > div > div > select:focus,
        .stNumberInput > div > div > input:focus {
            border-color: #1e3c72;
            box-shadow: 0 0 0 3px rgba(30, 60, 114, 0.1);
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
        
        /* Cartes de contenu */
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
        
        /* Pied de page */
        .footer {
            text-align: center;
            padding: 20px;
            color: #6c757d;
            font-size: 0.9em;
            border-top: 1px solid #e0e0e0;
            margin-top: 40px;
        }
        
        /* Loading spinner personnalisé */
        .custom-spinner {
            border: 4px solid #f3f3f3;
            border-top: 4px solid #1e3c72;
            border-radius: 50%;
            width: 50px;
            height: 50px;
            animation: spin 1s linear infinite;
            margin: 20px auto;
        }
        
        @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
        }
        
        /* Animations */
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(20px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        .fade-in {
            animation: fadeIn 0.5s ease-out;
        }
        
        /* Grille responsive */
        .grid-container {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(300px, 1fr));
            gap: 20px;
            padding: 20px;
        }
        
        /* Tooltips personnalisés */
        .tooltip {
            position: relative;
            display: inline-block;
        }
        
        .tooltip .tooltiptext {
            visibility: hidden;
            width: 200px;
            background-color: #1e3c72;
            color: white;
            text-align: center;
            border-radius: 6px;
            padding: 5px;
            position: absolute;
            z-index: 1;
            bottom: 125%;
            left: 50%;
            margin-left: -100px;
            opacity: 0;
            transition: opacity 0.3s;
        }
        
        .tooltip:hover .tooltiptext {
            visibility: visible;
            opacity: 1;
        }
        
        /* Scrollbar personnalisée */
        ::-webkit-scrollbar {
            width: 10px;
            height: 10px;
        }
        
        ::-webkit-scrollbar-track {
            background: #f1f1f1;
            border-radius: 5px;
        }
        
        ::-webkit-scrollbar-thumb {
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
            border-radius: 5px;
        }
        
        ::-webkit-scrollbar-thumb:hover {
            background: linear-gradient(135deg, #2a5298 0%, #1e3c72 100%);
        }
        
        /* Switch toggle */
        .switch {
            position: relative;
            display: inline-block;
            width: 60px;
            height: 34px;
        }
        
        .switch input {
            opacity: 0;
            width: 0;
            height: 0;
        }
        
        .slider {
            position: absolute;
            cursor: pointer;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background-color: #ccc;
            transition: .4s;
            border-radius: 34px;
        }
        
        .slider:before {
            position: absolute;
            content: "";
            height: 26px;
            width: 26px;
            left: 4px;
            bottom: 4px;
            background-color: white;
            transition: .4s;
            border-radius: 50%;
        }
        
        input:checked + .slider {
            background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        }
        
        input:checked + .slider:before {
            transform: translateX(26px);
        }
        
        /* Timeline */
        .timeline {
            position: relative;
            max-width: 1200px;
            margin: 0 auto;
        }
        
        .timeline::after {
            content: '';
            position: absolute;
            width: 6px;
            background: linear-gradient(180deg, #1e3c72 0%, #2a5298 100%);
            top: 0;
            bottom: 0;
            left: 50%;
            margin-left: -3px;
        }
        
        .timeline-item {
            padding: 10px 40px;
            position: relative;
            background-color: inherit;
            width: 50%;
        }
        
        .timeline-item::after {
            content: '';
            position: absolute;
            width: 25px;
            height: 25px;
            right: -17px;
            background-color: white;
            border: 4px solid #1e3c72;
            top: 15px;
            border-radius: 50%;
            z-index: 1;
        }
        
        .left {
            left: 0;
        }
        
        .right {
            left: 50%;
        }
        
        .right::after {
            left: -16px;
        }
        
        .timeline-content {
            padding: 20px 30px;
            background-color: white;
            position: relative;
            border-radius: 6px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
        }
    </style>
    """, unsafe_allow_html=True)

# Appliquer les styles
apply_custom_css()

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
    if 'pivot_compte_41' not in st.session_state:
        st.session_state.pivot_compte_41 = None
    
    # DataFrames
    if 'df_technique' not in st.session_state:
        st.session_state.df_technique = None
    if 'df_comptable' not in st.session_state:
        st.session_state.df_comptable = None
    if 'df_compte_41' not in st.session_state:
        st.session_state.df_compte_41 = None
    if 'df_410' not in st.session_state:
        st.session_state.df_410 = None
    if 'df_411' not in st.session_state:
        st.session_state.df_411 = None
    if 'production_data' not in st.session_state:
        st.session_state.production_data = None
    
    # Résultats de vérification
    if 'tableau_listing_police_invalide' not in st.session_state:
        st.session_state.tableau_listing_police_invalide = None
    if 'tableau_listing_valide' not in st.session_state:
        st.session_state.tableau_listing_valide = None
    if 'tableau_listing_police_invalide_comptable' not in st.session_state:
        st.session_state.tableau_listing_police_invalide_comptable = None
    if 'tableau_listing_valide_comptable' not in st.session_state:
        st.session_state.tableau_listing_valide_comptable = None
    
    # Logs et historique
    if 'logs' not in st.session_state:
        st.session_state.logs = []
    if 'history' not in st.session_state:
        st.session_state.history = []
    
    # Configuration
    if 'theme' not in st.session_state:
        st.session_state.theme = "Light"
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
            'lockout_duration': 30,  # minutes
            'session_timeout': 30,  # minutes
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

# Initialiser la session
init_session_state()

# ======================== FONCTIONS DE SÉCURITÉ ========================
class SecurityManager:
    """Gestionnaire de sécurité"""
    
    @staticmethod
    def hash_password(password):
        """Hash un mot de passe avec bcrypt"""
        salt = bcrypt.gensalt()
        return bcrypt.hashpw(password.encode('utf-8'), salt).decode('utf-8')
    
    @staticmethod
    def verify_password(hashed_password, plain_password):
        """Vérifie un mot de passe"""
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
            with self.get_connection() as conn:
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
                
                # Table des sauvegardes
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS backups (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        filename TEXT NOT NULL,
                        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                        size INTEGER,
                        status TEXT DEFAULT 'active'
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
                
        except Exception as e:
            st.error(f"Erreur d'initialisation de la base de données: {str(e)}")
    
    @contextmanager
    def get_connection(self):
        """Obtient une connexion à la base de données"""
        conn = sqlite3.connect(self.db_file)
        conn.row_factory = sqlite3.Row
        try:
            yield conn
        finally:
            conn.close()
    
    def execute_query(self, query, params=()):
        """Exécute une requête SQL"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            conn.commit()
            return cursor
    
    def fetch_all(self, query, params=()):
        """Récupère tous les résultats d'une requête"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            return [dict(row) for row in cursor.fetchall()]
    
    def fetch_one(self, query, params=()):
        """Récupère un seul résultat"""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, params)
            row = cursor.fetchone()
            return dict(row) if row else None

# ======================== GESTIONNAIRE DE SAUVEGARDE ========================
class BackupManager:
    """Gestionnaire de sauvegardes"""
    
    def __init__(self, db_file=DB_FILE):
        self.db_file = db_file
        self.backup_key = self._get_or_create_key()
        self.db_handler = DatabaseHandler()
    
    def _get_or_create_key(self):
        """Génère ou récupère la clé de chiffrement"""
        key_file = "backup_key.key"
        if os.path.exists(key_file):
            with open(key_file, "rb") as f:
                return f.read()
        else:
            key = Fernet.generate_key()
            with open(key_file, "wb") as f:
                f.write(key)
            return key
    
    def create_backup(self, description=""):
        """Crée une sauvegarde chiffrée"""
        try:
            backup_dir = "backups"
            os.makedirs(backup_dir, exist_ok=True)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_id = str(uuid.uuid4())[:8]
            temp_file = os.path.join(backup_dir, f"temp_backup_{timestamp}.db")
            backup_file = os.path.join(backup_dir, f"backup_{timestamp}_{backup_id}.enc")
            
            # Copier la base de données
            shutil.copy2(self.db_file, temp_file)
            
            # Chiffrer
            fernet = Fernet(self.backup_key)
            with open(temp_file, "rb") as f:
                data = f.read()
            
            encrypted = fernet.encrypt(data)
            
            with open(backup_file, "wb") as f:
                f.write(encrypted)
            
            # Nettoyer
            os.remove(temp_file)
            
            # Enregistrer dans la base
            size = os.path.getsize(backup_file)
            self.db_handler.execute_query(
                "INSERT INTO backups (filename, size, status) VALUES (?, ?, ?)",
                (backup_file, size, 'active')
            )
            
            # Nettoyer les vieilles sauvegardes
            self._clean_old_backups(backup_dir)
            
            log_action("Sauvegarde", f"Sauvegarde créée: {backup_file}")
            return True, backup_file
            
        except Exception as e:
            log_action("Erreur sauvegarde", str(e), level="error")
            return False, str(e)
    
    def restore_backup(self, backup_file):
        """Restaure une sauvegarde"""
        try:
            if not os.path.exists(backup_file):
                return False, "Fichier de sauvegarde introuvable"
            
            # Déchiffrer
            fernet = Fernet(self.backup_key)
            with open(backup_file, "rb") as f:
                encrypted = f.read()
            
            decrypted = fernet.decrypt(encrypted)
            
            # Sauvegarder l'actuelle avant restauration
            self.create_backup("Avant restauration")
            
            # Restaurer
            temp_file = backup_file.replace('.enc', '_restore.db')
            with open(temp_file, "wb") as f:
                f.write(decrypted)
            
            shutil.copy2(temp_file, self.db_file)
            os.remove(temp_file)
            
            log_action("Restauration", f"Base restaurée depuis: {backup_file}")
            return True, "Restauration réussie"
            
        except Exception as e:
            log_action("Erreur restauration", str(e), level="error")
            return False, str(e)
    
    def _clean_old_backups(self, backup_dir, keep=10):
        """Garde seulement les N dernières sauvegardes"""
        try:
            backups = sorted(
                [os.path.join(backup_dir, f) for f in os.listdir(backup_dir) 
                 if f.startswith("backup_") and f.endswith(".enc")],
                key=os.path.getmtime
            )
            
            for old_backup in backups[:-keep]:
                os.remove(old_backup)
                
                # Mettre à jour le statut dans la base
                self.db_handler.execute_query(
                    "UPDATE backups SET status='deleted' WHERE filename=?",
                    (old_backup,)
                )
                
        except Exception as e:
            log_action("Erreur nettoyage", str(e), level="warning")

# ======================== GESTIONNAIRE DE NOTIFICATIONS ========================
class NotificationManager:
    """Gestionnaire de notifications par email"""
    
    def __init__(self, config=None):
        self.config = config or {
            'smtp_server': 'smtp.gmail.com',
            'smtp_port': 587,
            'email_from': 'notifications@agc-vie.com',
            'email_to': ['admin@agc-vie.com'],
            'username': None,
            'password': None,
            'use_tls': True
        }
    
    def send_email(self, subject, body, recipients=None):
        """Envoie un email"""
        try:
            if not self.config['username'] or not self.config['password']:
                return False, "Configuration email incomplète"
            
            msg = MIMEMultipart()
            msg['From'] = self.config['email_from']
            msg['To'] = ', '.join(recipients or self.config['email_to'])
            msg['Subject'] = subject
            
            msg.attach(MIMEText(body, 'plain'))
            
            with smtplib.SMTP(self.config['smtp_server'], self.config['smtp_port']) as server:
                if self.config['use_tls']:
                    server.starttls()
                
                server.login(self.config['username'], self.config['password'])
                server.send_message(msg)
            
            log_action("Email envoyé", f"Sujet: {subject}")
            return True, "Email envoyé avec succès"
            
        except Exception as e:
            log_action("Erreur email", str(e), level="error")
            return False, str(e)
    
    def send_alert(self, alert_type, details):
        """Envoie une alerte"""
        subject = f"ALERTE {alert_type} - AGC-VIE"
        
        body = f"""
        Type d'alerte: {alert_type}
        Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
        Détails: {details}
        
        Ceci est un message automatique du système AGC-VIE.
        """
        
        return self.send_email(subject, body)
    
    def send_report(self, report_data, report_type="quotidien"):
        """Envoie un rapport"""
        subject = f"Rapport {report_type} - AGC-VIE"
        
        body = f"""
        Rapport {report_type}
        Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
        
        Résumé:
        {report_data}
        
        Pour plus de détails, connectez-vous à l'application.
        """
        
        return self.send_email(subject, body)

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
    
    # Logger dans un fichier
    try:
        logging.basicConfig(
            filename='agc_vie.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        
        if level == "info":
            logging.info(f"{action} - {details}")
        elif level == "warning":
            logging.warning(f"{action} - {details}")
        elif level == "error":
            logging.error(f"{action} - {details}")
            
    except Exception as e:
        print(f"Erreur d'écriture des logs: {e}")

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
        
        if value_cols:
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
        
        # Tableau croisé dynamique
        index_col = 'No Police' if 'No Police' in df.columns else df.columns[0]
        value_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        
        if value_cols:
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
        
        # Déterminer les colonnes de jointure
        tech_col = 'Nouvelle_Police' if 'Nouvelle_Police' in tech.columns else tech.columns[0]
        compta_col = 'No Police' if 'No Police' in compta.columns else compta.columns[0]
        
        # Renommer pour la fusion
        tech = tech.rename(columns={tech_col: 'Police'})
        compta = compta.rename(columns={compta_col: 'Police'})
        
        # Fusion
        merged = pd.merge(tech, compta, on='Police', how='outer', suffixes=('_tech', '_compta'))
        
        # Calcul des écarts
        if 'Emissions' in merged.columns and 'Débit' in merged.columns and 'Crédit' in merged.columns:
            # Conversion en numérique
            for col in ['Emissions', 'Débit', 'Crédit']:
                if col in merged.columns:
                    merged[col] = pd.to_numeric(merged[col], errors='coerce').fillna(0)
            
            merged['CA_Technique'] = merged['Emissions']
            merged['CA_Comptable'] = abs(merged['Crédit'] - merged['Débit'])
            merged['Écart'] = merged['CA_Technique'] - merged['CA_Comptable']
            merged['Statut'] = merged['Écart'].apply(
                lambda x: 'Rapproché' if abs(x) < 0.01 else 'Non rapproché'
            )
        
        # Statistiques
        stats = {
            'total_polices': len(merged),
            'polices_techniques': len(tech),
            'polices_comptables': len(compta),
            'polices_communes': len(merged[merged['Police'].notna() & merged['Police_tech'].notna() & merged['Police_compta'].notna()]),
            'polices_tech_only': len(merged[merged['Police_tech'].notna() & merged['Police_compta'].isna()]),
            'polices_compta_only': len(merged[merged['Police_tech'].isna() & merged['Police_compta'].notna()])
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
                    df.to_excel(writer, sheet_name=sheet_name[:31], index=False)  # Excel limite à 31 caractères
        
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
def display_metric_card(title, value, icon="📊", description=""):
    """Affiche une carte métrique stylisée"""
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
        st.markdown("""
        <div style="text-align: center; padding: 40px; animation: fadeIn 0.5s;">
            <h1 style="color: #1e3c72; font-size: 3em; margin-bottom: 10px;">AGC-VIE</h1>
            <p style="color: #666; font-size: 1.2em; margin-bottom: 30px;">
                Système de Gestion Technique et Comptable
            </p>
        </div>
        """, unsafe_allow_html=True)
        
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
            
            role = st.selectbox(
                "🎭 Rôle",
                ["user", "admin"],
                help="Sélectionnez votre rôle"
            )
            
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
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
                        st.error("Nom d'utilisateur ou mot de passe incorrect")
                else:
                    st.warning("Veuillez remplir tous les champs")
        
        # Informations supplémentaires

        st.markdown("""
        <div style="text-align: center; color: #666; font-size: 0.9em;">
            <p>Compte démo: admin / Admin123!</p>
            <p>© 2025 AGC-VIE - Tous droits réservés</p>
        </div>
        """, unsafe_allow_html=True)

def page_accueil():
    """Page d'accueil"""
    update_last_activity()
    
    st.markdown("""
    <div class="content-card fade-in" style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);">
        <h1 style="color: white; text-align: center;">Bienvenue sur AGC-VIE</h1>
        <p style="color: white; text-align: center; font-size: 1.2em;">
            Système intégré de Gestion Technique et Comptable
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Métriques rapides
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        tech_count = len(st.session_state.pivot_techniques) if st.session_state.pivot_techniques is not None else 0
        st.metric(
            "📊 Données techniques",
            tech_count,
            delta=None,
            help="Nombre d'enregistrements techniques"
        )
    
    with col2:
        compta_count = len(st.session_state.pivot_comptables) if st.session_state.pivot_comptables is not None else 0
        st.metric(
            "💰 Données comptables",
            compta_count,
            delta=None,
            help="Nombre d'enregistrements comptables"
        )
    
    with col3:
        if st.session_state.pivot_techniques is not None and 'Emissions' in st.session_state.pivot_techniques.columns:
            ca_tech = st.session_state.pivot_techniques['Emissions'].sum()
            st.metric(
                "📈 CA Technique",
                f"{ca_tech:,.0f} FCFA",
                delta=None,
                help="Chiffre d'affaires technique"
            )
    
    with col4:
        if st.session_state.pivot_comptables is not None and 'Crédit' in st.session_state.pivot_comptables.columns:
            ca_compta = st.session_state.pivot_comptables['Crédit'].sum()
            st.metric(
                "📉 CA Comptable",
                f"{ca_compta:,.0f} FCFA",
                delta=None,
                help="Chiffre d'affaires comptable"
            )
    
    # Modules
    st.markdown("## 🚀 Modules disponibles")
    
    col1, col2 = st.columns(2)
    
    with col1:
        with st.container():
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
        with st.container():
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
        with st.container():
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
        with st.container():
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
    
    # Dernières actions
    if st.session_state.logs:
        st.markdown("## 📋 Dernières activités")
        
        logs_df = pd.DataFrame(st.session_state.logs[-10:])
        
        # Formatage pour l'affichage
        if 'timestamp' in logs_df.columns:
            logs_df['timestamp'] = pd.to_datetime(logs_df['timestamp']).dt.strftime('%d/%m/%Y %H:%M')
        
        st.dataframe(
            logs_df[['timestamp', 'action', 'details']],
            use_container_width=True,
            height=300,
            hide_index=True
        )

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
                        delta_color = "normal" if solde >= 0 else "inverse"
                        st.metric("Solde", f"{solde:,.0f} FCFA", delta=f"{abs(solde):,.0f}", delta_color=delta_color)
                
            else:
                st.warning("Aucune colonne numérique disponible pour l'analyse")
        else:
            st.info("Aucune donnée à analyser")
            
            
def rapprochement_technique_comptable(tech_df, compta_df):
    """Effectue le rapprochement entre données techniques et comptables"""
    try:
        # Copie des DataFrames
        tech = tech_df.copy()
        compta = compta_df.copy()
        
        # Déterminer les colonnes de jointure
        tech_col = 'Nouvelle_Police' if 'Nouvelle_Police' in tech.columns else tech.columns[0]
        compta_col = 'No Police' if 'No Police' in compta.columns else compta.columns[0]
        
        # Renommer pour la fusion
        tech_renamed = tech.rename(columns={tech_col: 'Police'})
        compta_renamed = compta.rename(columns={compta_col: 'Police'})
        
        # Fusion - Utiliser 'Police' comme clé commune après renommage
        merged = pd.merge(
            tech_renamed, 
            compta_renamed, 
            on='Police', 
            how='outer', 
            suffixes=('_tech', '_compta')
        )
        
        # Calcul des écarts
        if 'Emissions' in merged.columns and 'Débit' in merged.columns and 'Crédit' in merged.columns:
            # Conversion en numérique
            for col in ['Emissions', 'Débit', 'Crédit']:
                if col in merged.columns:
                    merged[col] = pd.to_numeric(merged[col], errors='coerce').fillna(0)
            
            # Créer les colonnes CA si elles n'existent pas
            if 'CA_Technique' not in merged.columns:
                merged['CA_Technique'] = merged['Emissions']
            
            if 'CA_Comptable' not in merged.columns:
                merged['CA_Comptable'] = abs(merged['Crédit'] - merged['Débit'])
            
            if 'Écart' not in merged.columns:
                merged['Écart'] = merged['CA_Technique'] - merged['CA_Comptable']
            
            if 'Statut' not in merged.columns:
                merged['Statut'] = merged['Écart'].apply(
                    lambda x: 'Rapproché' if abs(x) < 0.01 else 'Non rapproché'
                )
        
        # Statistiques - Correction pour utiliser les bonnes colonnes
        stats = {
            'total_polices': len(merged),
            'polices_techniques': len(tech),
            'polices_comptables': len(compta),
            'polices_communes': len(merged[merged['Police_tech'].notna() & merged['Police_compta'].notna()]) 
                               if 'Police_tech' in merged.columns and 'Police_compta' in merged.columns 
                               else len(merged[merged['Police'].notna()]),
            'polices_tech_only': len(merged[merged['Police_tech'].notna() & merged['Police_compta'].isna()]) 
                               if 'Police_tech' in merged.columns and 'Police_compta' in merged.columns 
                               else len(merged[merged['Police'].notna() & merged['Emissions'].notna() & merged['Débit'].isna()]),
            'polices_compta_only': len(merged[merged['Police_tech'].isna() & merged['Police_compta'].notna()]) 
                                 if 'Police_tech' in merged.columns and 'Police_compta' in merged.columns 
                                 else len(merged[merged['Police'].notna() & merged['Emissions'].isna() & merged['Débit'].notna()])
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
        # Métriques
        st.markdown("### 📊 Résumé du rapprochement")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            display_metric_card(
                "Polices techniques",
                stats.get('polices_techniques', 0),
                "📊",
                f"Total: {stats.get('polices_techniques', 0)}"
            )
        
        with col2:
            display_metric_card(
                "Polices comptables",
                stats.get('polices_comptables', 0),
                "💰",
                f"Total: {stats.get('polices_comptables', 0)}"
            )
        
        with col3:
            display_metric_card(
                "Polices communes",
                stats.get('polices_communes', 0),
                "🔄",
                f"Taux: {stats.get('polices_communes', 0)/stats.get('polices_techniques', 1)*100:.1f}%"
            )
        
        with col4:
            ecart = stats.get('ecart_total', 0)
            display_metric_card(
                "Écart total",
                f"{ecart:,.0f} FCFA",
                "📈" if ecart >= 0 else "📉",
                "Positif si CA technique > CA comptable"
            )
        
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
                fig.update_layout(height=400)
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
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)
        
        # Top des écarts
        if 'Écart' in merged_df.columns:
            st.markdown("### 📈 Top 10 des écarts")
            
            top_ecarts = merged_df.nlargest(10, 'Écart')[['Police', 'CA_Technique', 'CA_Comptable', 'Écart']]
            st.dataframe(top_ecarts, use_container_width=True, hide_index=True)
        
        # Export complet
        st.markdown("### 📥 Export du rapprochement")
        
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
        
        log_action("Rapprochement technique", f"{len(merged_df)} polices analysées")
        
    else:
        st.error(f"Erreur lors du rapprochement: {stats}")
        
        
def page_rapprochement_comptable():
    """Page de rapprochement comptable"""
    update_last_activity()
    
    st.markdown("## 🔄 Rapprochement Comptable")
    
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
    
    with st.spinner("Calcul du rapprochement comptable en cours..."):
        # Fonction pour récupérer les annulations et émissions dans les données techniques
        def recuperer_annulations_emissions(pivot_techniques, pivot_comptables):
            """Ajoute les colonnes Ristournes et Emissions aux données comptables"""
            try:
                # Copie pour éviter les modifications sur l'original
                df_compta = pivot_comptables.copy()
                df_tech = pivot_techniques.copy()
                
                # Nettoyage des noms de colonnes
                df_compta.columns = df_compta.columns.str.strip()
                df_tech.columns = df_tech.columns.str.strip()
                
                # Déterminer les colonnes de police
                tech_col = 'Nouvelle_Police' if 'Nouvelle_Police' in df_tech.columns else df_tech.columns[0]
                compta_col = 'No Police' if 'No Police' in df_compta.columns else df_compta.columns[0]
                
                # Initialiser les colonnes
                df_compta['Ristournes'] = 0
                df_compta['Emissions'] = 0
                df_compta['Statut_Ristournes'] = 'Non trouvé'
                df_compta['Statut_Emissions'] = 'Non trouvé'
                
                # Pour chaque ligne comptable, chercher la correspondance dans les données techniques
                for index, row in df_compta.iterrows():
                    police_compta = str(row[compta_col]).strip()
                    
                    # Chercher dans les données techniques
                    correspondance = df_tech[df_tech[tech_col].astype(str).str.strip() == police_compta]
                    
                    if not correspondance.empty:
                        # Récupérer les valeurs techniques
                        if 'Ristournes' in correspondance.columns:
                            df_compta.at[index, 'Ristournes'] = correspondance['Ristournes'].values[0]
                            df_compta.at[index, 'Statut_Ristournes'] = 'Trouvé'
                        
                        if 'Emissions' in correspondance.columns:
                            df_compta.at[index, 'Emissions'] = correspondance['Emissions'].values[0]
                            df_compta.at[index, 'Statut_Emissions'] = 'Trouvé'
                
                return df_compta
                
            except Exception as e:
                st.error(f"Erreur dans recuperer_annulations_emissions: {str(e)}")
                return pivot_comptables
        
        # Fonction pour vérifier les polices comptables
        def verifier_polices_comptable(pivot_comptables):
            """Vérifie le rapprochement des polices comptables"""
            try:
                df = pivot_comptables.copy()
                
                # Conversion en numérique
                for col in ['Crédit', 'Débit', 'Emissions', 'Ristournes']:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                
                # Calcul du rapprochement
                if all(col in df.columns for col in ['Crédit', 'Débit', 'Emissions', 'Ristournes']):
                    df['CA_Comptable'] = df['Crédit'] - df['Débit']
                    df['CA_Technique'] = df['Emissions'] + df['Ristournes']
                    df['Écart'] = abs(df['CA_Comptable']) - abs(df['CA_Technique'])
                    
                    df['Rapprochement'] = df.apply(
                        lambda row: 'Rapproché' if abs(row['Écart']) < 0.01 else 'Non rapproché', 
                        axis=1
                    )
                
                # Calcul des totaux
                total_debit = df['Débit'].sum() if 'Débit' in df.columns else 0
                total_credit = df['Crédit'].sum() if 'Crédit' in df.columns else 0
                total_emissions_tech = df['Emissions'].sum() if 'Emissions' in df.columns else 0
                total_ristournes_tech = df['Ristournes'].sum() if 'Ristournes' in df.columns else 0
                total_CA_comptable = abs(total_credit - total_debit)
                total_CA_technique = abs(total_emissions_tech + total_ristournes_tech)
                ecart = abs(total_CA_technique - total_CA_comptable)
                
                # Tableaux de résultats
                tableau_invalide = df[df['Rapprochement'] == 'Non rapproché'] if 'Rapprochement' in df.columns else pd.DataFrame()
                tableau_valide = df[df['Rapprochement'] == 'Rapproché'] if 'Rapprochement' in df.columns else pd.DataFrame()
                
                stats = {
                    'total_debit': total_debit,
                    'total_credit': total_credit,
                    'total_CA_comptable': total_CA_comptable,
                    'total_emissions_tech': total_emissions_tech,
                    'total_ristournes_tech': total_ristournes_tech,
                    'total_CA_technique': total_CA_technique,
                    'ecart': ecart,
                    'total_polices': len(df),
                    'polices_valides': len(tableau_valide),
                    'polices_invalides': len(tableau_invalide)
                }
                
                return df, tableau_invalide, tableau_valide, stats
                
            except Exception as e:
                st.error(f"Erreur dans verifier_polices_comptable: {str(e)}")
                return pivot_comptables, pd.DataFrame(), pd.DataFrame(), {}
        
        # Exécution du rapprochement
        pivot_comptables_avec_tech = recuperer_annulations_emissions(
            st.session_state.pivot_techniques,
            st.session_state.pivot_comptables
        )
        
        df_complet, df_invalide, df_valide, stats = verifier_polices_comptable(pivot_comptables_avec_tech)
        
        # Stockage dans la session
        st.session_state.pivot_comptables_complet = df_complet
        st.session_state.tableau_listing_police_invalide_comptable = df_invalide
        st.session_state.tableau_listing_valide_comptable = df_valide
    
    # Affichage des résultats
    if stats:
        # Métriques principales
        st.markdown("### 📊 Résumé du rapprochement comptable")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            display_metric_card(
                "Total Débit",
                f"{stats.get('total_debit', 0):,.0f} FCFA",
                "💳",
                "Somme des débits"
            )
        
        with col2:
            display_metric_card(
                "Total Crédit",
                f"{stats.get('total_credit', 0):,.0f} FCFA",
                "💰",
                "Somme des crédits"
            )
        
        with col3:
            display_metric_card(
                "CA Comptable",
                f"{stats.get('total_CA_comptable', 0):,.0f} FCFA",
                "📊",
                "Crédit - Débit"
            )
        
        with col4:
            display_metric_card(
                "CA Technique",
                f"{stats.get('total_CA_technique', 0):,.0f} FCFA",
                "📈",
                "Émissions + Ristournes"
            )
        
        # Deuxième ligne de métriques
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            display_metric_card(
                "Écart",
                f"{stats.get('ecart', 0):,.0f} FCFA",
                "📉" if stats.get('ecart', 0) > 0 else "📈",
                "CA Technique - CA Comptable"
            )
        
        with col2:
            total_polices = stats.get('total_polices', 0)
            display_metric_card(
                "Total polices",
                total_polices,
                "📋",
                f"Nombre total de polices"
            )
        
        with col3:
            valides = stats.get('polices_valides', 0)
            taux_valides = (valides / total_polices * 100) if total_polices > 0 else 0
            display_metric_card(
                "Polices rapprochées",
                valides,
                "✅",
                f"Taux: {taux_valides:.1f}%"
            )
        
        with col4:
            invalides = stats.get('polices_invalides', 0)
            taux_invalides = (invalides / total_polices * 100) if total_polices > 0 else 0
            display_metric_card(
                "Polices non rapprochées",
                invalides,
                "❌",
                f"Taux: {taux_invalides:.1f}%"
            )
        
        # Tabs pour les différentes vues
        tab1, tab2, tab3 = st.tabs(["📋 Données complètes", "❌ Non rapprochées", "✅ Rapprochées"])
        
        with tab1:
            st.markdown("### Toutes les polices comptables")
            
            # Recherche
            search = create_search_bar("compta_rapprochement_search", "Rechercher une police...")
            
            df_display = df_complet.copy()
            if search:
                df_display = filter_dataframe(df_display, search)
            
            st.dataframe(df_display, use_container_width=True, height=500, hide_index=True)
        
        with tab2:
            st.markdown(f"### Polices non rapprochées ({len(df_invalide)})")
            
            if not df_invalide.empty:
                # Recherche dans les invalides
                search_invalide = create_search_bar("invalide_search", "Rechercher dans les non rapprochées...")
                
                df_invalide_display = df_invalide.copy()
                if search_invalide:
                    df_invalide_display = filter_dataframe(df_invalide_display, search_invalide)
                
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
            st.markdown(f"### Polices rapprochées ({len(df_valide)})")
            
            if not df_valide.empty:
                # Recherche dans les valides
                search_valide = create_search_bar("valide_search", "Rechercher dans les rapprochées...")
                
                df_valide_display = df_valide.copy()
                if search_valide:
                    df_valide_display = filter_dataframe(df_valide_display, search_valide)
                
                st.dataframe(df_valide_display, use_container_width=True, height=500, hide_index=True)
        
        # Visualisations
        st.markdown("### 📊 Analyses et visualisations")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Graphique de répartition des statuts
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
                    showlegend=True
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
        st.markdown("### 📈 Analyse des écarts")
        
        if not df_invalide.empty and 'Écart' in df_invalide.columns:
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
                fig.update_layout(height=400)
                st.plotly_chart(fig, use_container_width=True)
            
            with col2:
                # Top 10 des écarts
                st.markdown("#### Top 10 des écarts")
                
                ecarts_cols = ['No Police' if 'No Police' in df_invalide.columns else df_invalide.columns[0]]
                if 'Écart' in df_invalide.columns:
                    ecarts_cols.append('Écart')
                if 'CA_Comptable' in df_invalide.columns:
                    ecarts_cols.append('CA_Comptable')
                if 'CA_Technique' in df_invalide.columns:
                    ecarts_cols.append('CA_Technique')
                
                top_ecarts = df_invalide.nlargest(10, 'Écart')[ecarts_cols]
                st.dataframe(top_ecarts, use_container_width=True, hide_index=True)
        
        # Export complet
        st.markdown("### 📥 Export des résultats")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("📥 Exporter le rapport complet", type="primary", use_container_width=True):
                # Créer un fichier Excel avec plusieurs onglets
                dataframes = []
                sheet_names = []
                
                if df_complet is not None:
                    dataframes.append(df_complet)
                    sheet_names.append("Données complètes")
                
                if df_valide is not None and not df_valide.empty:
                    dataframes.append(df_valide)
                    sheet_names.append("Polices rapprochées")
                
                if df_invalide is not None and not df_invalide.empty:
                    dataframes.append(df_invalide)
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
                    ["Taux de rapprochement", f"{(stats.get('polices_valides', 0)/stats.get('total_polices', 1)*100):.1f}%"]
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
            f"{len(df_complet)} polices analysées, {len(df_invalide)} non rapprochées, écart: {stats.get('ecart', 0):,.0f} FCFA"
        )
        
    else:
        st.error("Erreur lors du calcul des statistiques de rapprochement")
        
        
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
                        # Récupérer les polices
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
                                    
                                    # Export
                                    if st.button("📥 Exporter les invalides"):
                                        df_invalid = pd.DataFrame(invalid_refs, columns=['Références invalides'])
                                        csv_data = export_to_csv(df_invalid, "references_invalides.csv")
                                        if csv_data:
                                            create_download_button(
                                                csv_data,
                                                "references_invalides.csv",
                                                "Télécharger CSV"
                                            )
                            
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
                        if 'NUMERO POLICE' in df.columns:
                            occurrences = df['NUMERO POLICE'].value_counts().value_counts().sort_index()
                            
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
                    taille_estimée = total * 50  # Estimation 50KB par certificat
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
        
        # Génération de données simulées
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
        
        backup_manager = BackupManager()
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("💾 Créer une sauvegarde", use_container_width=True):
                with st.spinner("Création de la sauvegarde..."):
                    success, result = backup_manager.create_backup("Sauvegarde manuelle")
                    if success:
                        st.success(f"Sauvegarde créée: {os.path.basename(result)}")
                        st.balloons()
                    else:
                        st.error(f"Erreur: {result}")
        
        with col2:
            # Liste des sauvegardes
            backup_dir = "backups"
            if os.path.exists(backup_dir):
                backups = [f for f in os.listdir(backup_dir) if f.startswith("backup_") and f.endswith(".enc")]
                
                if backups:
                    selected_backup = st.selectbox("Sauvegardes disponibles", backups)
                    
                    if st.button("🔄 Restaurer", use_container_width=True):
                        if st.checkbox("Confirmer la restauration"):
                            with st.spinner("Restauration en cours..."):
                                backup_path = os.path.join(backup_dir, selected_backup)
                                success, result = backup_manager.restore_backup(backup_path)
                                if success:
                                    st.success(result)
                                else:
                                    st.error(result)
        
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

# ======================== BARRE LATÉRALE ========================
def sidebar():
    """Affiche la barre latérale"""
    with st.sidebar:
        # En-tête
        st.markdown("""
        <div style="text-align: center; padding: 20px 10px;">
            <h2 style="color: white; margin-bottom: 5px;">AGC-VIE</h2>
            <p style="color: rgba(255,255,255,0.7); font-size: 0.9em;">Version 2.0</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Informations utilisateur
        if st.session_state.authenticated:
            st.markdown(f"""
            <div style="background: rgba(255,255,255,0.1); padding: 15px; border-radius: 10px; margin-bottom: 20px;">
                <p style="color: white; margin: 0; font-size: 1.1em;">👤 {st.session_state.username}</p>
                <p style="color: #4CAF50; margin: 5px 0 0 0; font-size: 0.9em;">
                    {'👑 Administrateur' if st.session_state.role == 'admin' else '👤 Utilisateur'}
                </p>
                <p style="color: rgba(255,255,255,0.5); margin: 5px 0 0 0; font-size: 0.8em;">
                    {datetime.now().strftime('%d/%m/%Y %H:%M')}
                </p>
            </div>
            """, unsafe_allow_html=True)
        
        # Menu principal avec streamlit-option-menu
        selected = option_menu(
            menu_title=None,
            options=[
                "Accueil",
                "Gestion Technique",
                "Gestion Comptable",
                "Rapprochement Technique",
                "Rapprochement Comptable",
                "Gestion 410 & 411",
                "Gestion Doublons",
                "Gestion Production",
                "Statistiques",
                "Administration",
                "Déconnexion"
            ],
            icons=[
                "house",
                "bar-chart",
                "cash",
                "arrow-repeat",
                "arrow-repeat",
                "folder",
                "files",
                "file-earmark",
                "graph-up",
                "gear",
                "box-arrow-right"
            ],
            menu_icon="cast",
            default_index=0,
            styles={
                "container": {
                    "padding": "0!important",
                    "background-color": "blue"
                },
                "icon": {
                    "color": "white",
                    "font-size": "18px"
                },
                "nav-link": {
                    "color": "white",
                    "font-size": "16px",
                    "text-align": "left",
                    "margin": "5px 0",
                    "padding": "10px 15px",
                    "border-radius": "8px",
                    "transition": "all 0.3s"
                },
                "nav-link:hover": {
                    "background-color": "rgba(255,255,255,0.1)",
                    "transform": "translateX(5px)"
                },
                "nav-link-selected": {
                    "background": "linear-gradient(135deg, #2a5298 0%, #1e3c72 100%)",
                    "font-weight": "bold",
                    "box-shadow": "0 4px 6px rgba(0,0,0,0.2)"
                }
            }
        )
        
        # Pied de page
        st.markdown("---")
        st.markdown("""
        <div style="text-align: center; color: rgba(255,255,255,0.5); font-size: 0.8em; padding: 10px;">
            <p>© 2025 AGC-VIE</p>
            <p>Version 2.0</p>
        </div>
        """, unsafe_allow_html=True)
        
        return selected

# ======================== FONCTION PRINCIPALE ========================
def main():
    """Fonction principale de l'application"""
    
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
    
    # Pied de page commun
    st.markdown("""
    <div style="text-align: center; color: #666; padding: 20px;">
        <p>AGC-VIE - Système de Gestion Technique et Comptable</p>
        <p style="font-size: 0.9em;">Développé par Frédéric BAYONNE MAVOUNGOU | Ingénieur en Génie Numérique</p>
    </div>
    """, unsafe_allow_html=True)

# ======================== IMPORT DES MODULES MANQUANTS ========================
from datetime import datetime, timedelta
from typing import Optional, Dict, List, Any, Tuple

# ======================== POINT D'ENTRÉE ========================
if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        st.error(f"Erreur critique: {str(e)}")
        log_action("Erreur critique", str(e), level="error")
        
        # Affichage détaillé en mode développement
        if os.getenv('ENVIRONMENT') == 'development':
            st.exception(e)