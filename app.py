import base64
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
import html
import io
import json
import math
import os
import re
import urllib.parse
import urllib3
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import requests
import streamlit as st
import tabelas_engine as te
import importlib
try:
    importlib.reload(te)
except Exception:
    pass

CONTATOS_TRANSPORTADORAS = {
    "braspress": "551122233500",
    "brasul": "554399652615",
    "sudoeste": "554398660050",
    "coopex": "554388705800",
    "rico": "554388705800",
    "aragão": "554388705800",
    "aragao": "554388705800",
    "alfa": "554230354979",
    "tecmar": "554399680511",
    "vip": "554333560099",
    "bertolini": "555421023000",
    "plav": "55438416965"
}

def obter_numero_wa(nome_transportadora):
    nome_limpo = str(nome_transportadora).lower().strip()
    for chave, numero in CONTATOS_TRANSPORTADORAS.items():
        if chave in nome_limpo:
            return numero
    return ""

def gerar_badge_logo_html(nome_transp: str) -> str:
    """Gera bloco HTML com logotipo da transportadora para os cards e tabelas."""
    try:
        if hasattr(te, "gerar_badge_logo_html"):
            return te.gerar_badge_logo_html(nome_transp)
    except Exception:
        pass
    return f'<div style="background-color: #f0f2f6; padding: 5px 10px; border-radius: 5px; display: inline-block; font-weight: bold; color: #31333F;">{nome_transp}</div>'

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Configuração da página
st.set_page_config(
    page_title="Sistema de Cotações Logísticas — Next Cable",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

APP_VERSION = "Sistema de Cotações Logísticas"
HISTORICO_FILE = "historico_cotacoes.json"
ENTREGAS_FILE = "entregas.xlsx"
BROWSER_USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
)

# Estilo Visual Enterprise / SaaS Moderno
st.markdown("""
<style>
#MainMenu, footer, header {visibility: hidden;}
.stApp {
    background: linear-gradient(180deg, #f8fafc 0%, #f1f5f9 100%);
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
}

/* Sidebar Estilo SaaS Corporativo */
section[data-testid="stSidebar"] {
    background-color: #0b1329 !important;
    color: #f8fafc !important;
    border-right: 1px solid #1e293b !important;
}
section[data-testid="stSidebar"] .stButton > button {
    border-radius: 8px !important;
    font-size: 0.95rem !important;
    font-weight: 700 !important;
    text-align: left !important;
    justify-content: flex-start !important;
    padding: 10px 14px !important;
    margin-bottom: 5px !important;
    transition: all 0.15s ease !important;
    min-height: 42px !important;
}
section[data-testid="stSidebar"] .stButton > button[kind="secondary"] {
    background: #131c31 !important;
    color: #94a3b8 !important;
    border: 1px solid transparent !important;
    font-weight: 600 !important;
}
section[data-testid="stSidebar"] .stButton > button[kind="secondary"]:hover {
    background: #1e2942 !important;
    color: #e2e8f0 !important;
    border-color: #334155 !important;
}
section[data-testid="stSidebar"] .stButton > button[kind="primary"] {
    background: linear-gradient(135deg, #0284c7 0%, #0369a1 100%) !important;
    color: #ffffff !important;
    border: 1px solid #38bdf8 !important;
    font-weight: 800 !important;
    font-size: 1.05rem !important;
    box-shadow: 0 4px 14px rgba(2, 132, 199, 0.45) !important;
}

/* Formulários e Inputs */
/* Fix Expander Contrast in Dark Theme */
section[data-testid="stSidebar"] [data-testid="stExpander"] {
    background-color: #131c31 !important;
    border: 1px solid #1e293b !important;
    border-radius: 8px !important;
}
section[data-testid="stSidebar"] [data-testid="stExpander"] summary,
section[data-testid="stSidebar"] [data-testid="stExpander"] summary p,
section[data-testid="stSidebar"] [data-testid="stExpander"] summary svg {
    color: #e2e8f0 !important;
}
section[data-testid="stSidebar"] [data-testid="stExpander"] .streamlit-expanderContent {
    background-color: #131c31 !important;
    color: #f8fafc !important;
}
.stTextInput label, .stNumberInput label, .stSelectbox label {
    color: #1e293b !important;
    font-weight: 700 !important;
    font-size: 0.77rem !important;
    text-transform: uppercase !important;
    letter-spacing: 0.04em !important;
}
.stTextInput > div > div > input,
.stNumberInput > div > div > input {
    background: #ffffff !important;
    border: 1.5px solid #cbd5e1 !important;
    border-radius: 10px !important;
    min-height: 42px !important;
    font-size: 0.92rem !important;
}
.stTextInput > div > div > input:focus,
.stNumberInput > div > div > input:focus {
    border-color: #0284c7 !important;
    box-shadow: 0 0 0 3px rgba(2, 132, 199, 0.15) !important;
}

/* Botão de Ação Primária (Super CTA) */
.btn-cta > div > button {
    background-color: #FF0000 !important;
    background-image: none !important;
    color: #FFFFFF !important;
    font-weight: bold !important;
    font-size: 1.05rem !important;
    letter-spacing: 0.02em !important;
    border-radius: 14px !important;
    height: 54px !important;
    border: none !important;
    box-shadow: 0 8px 24px rgba(255, 0, 0, 0.32) !important;
    transition: all 0.25s ease !important;
}
.btn-cta > div > button:hover {
    filter: brightness(1.08) !important;
    transform: translateY(-2px) !important;
    box-shadow: 0 12px 28px rgba(37, 99, 235, 0.42) !important;
}

/* Containers com Cards Suaves */
div[data-testid="stVerticalBlockBorderWrapper"] {
    background: #ffffff !important;
    border: 1px solid #e2e8f0 !important;
    border-radius: 14px !important;
    box-shadow: 0 4px 16px -2px rgba(0, 0, 0, 0.04) !important;
}

.block-container { padding-top: 1.2rem !important; max-width: 1380px !important; }

/* Badges de Ranking e Smart Tags */
.badge-ranking {
    display: inline-flex;
    align-items: center;
    gap: 5px;
    font-size: 0.72rem;
    font-weight: 800;
    padding: 4px 10px;
    border-radius: 9999px;
    letter-spacing: 0.03em;
    text-transform: uppercase;
    white-space: nowrap;
}
.badge-pos1 {
    background: #059669;
    color: #ffffff;
    box-shadow: 0 2px 6px rgba(5, 150, 105, 0.3);
}
.badge-pos2 {
    background: #0284c7;
    color: #ffffff;
    box-shadow: 0 2px 6px rgba(2, 132, 199, 0.3);
}
.badge-pos3 {
    background: #475569;
    color: #ffffff;
    box-shadow: 0 2px 6px rgba(71, 85, 105, 0.3);
}
.badge-fast {
    background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);
    color: #ffffff;
    box-shadow: 0 2px 6px rgba(217, 119, 6, 0.35);
}

.card-ranking-pos1 {
    border-top: 3px solid #10b981;
}
.card-ranking-pos2 {
    border-top: 3px solid #0284c7;
}
.card-ranking-pos3 {
    border-top: 3px solid #64748b;
}
</style>
""", unsafe_allow_html=True)


def _secret(chave, padrao=""):
    try:
        if hasattr(st, "secrets") and chave in st.secrets:
            return st.secrets[chave]
    except Exception:
        pass
    return padrao


USUARIO_BRASPRESS = _secret("USUARIO_BRASPRESS", "")
SENHA_BRASPRESS = _secret("SENHA_BRASPRESS", "")
DOMINIO_AGEX = _secret("DOMINIO_AGEX", "")
USUARIO_AGEX = _secret("USUARIO_AGEX", "")
SENHA_AGEX = _secret("SENHA_AGEX", "")
DOMINIO_COPEX = _secret("DOMINIO_COPEX", "")
USUARIO_COPEX = _secret("USUARIO_COPEX", "")
SENHA_COPEX = _secret("SENHA_COPEX", "")
DOMINIO_LOGBG = _secret("DOMINIO_LOGBG", "")
USUARIO_LOGBG = _secret("USUARIO_LOGBG", "")
SENHA_LOGBG = _secret("SENHA_LOGBG", "")
MERCADORIA_LOGBG = int(_secret("MERCADORIA_LOGBG", 110))

COBERTURA = {
    "Braspress": "API nacional oficial (quase todo o Brasil)",
    "Coopex": "PR, SC, RS, SP (API SSW oficial + Tabela fallback)",
    "Agex": "PR, SC, RS, SP, MS (Tabela negociada)",
    "Princesa dos Campos": "PR, SC, RS, SP (Tabela negociada)",
    "TW Transportes": "PR, SC, RS, SP (Tabela negociada)",
    "Envia Rápido": "RJ, SP, MG, BA, PR, GO, MT, MS, DF, Nordeste e mais (Tabela)",
    "LOGBG": "API SSW (Credencial em validação)",
}

TABELAS_INFO = {
    "Braspress": "API Oficial REST v1 (Homologada em Tempo Real)",
    "Coopex": "API Oficial SSW SOAP (Homologada em Tempo Real)",
    "Agex": "Tabela Contratual V. 01/2024 (Tarifa Negociada)",
    "Princesa dos Campos": "Tabela Contratual V. 11/2023 (Tarifa Negociada)",
    "TW Transportes": "Tabela Contratual V. 10/2023 (Tarifa Negociada)",
    "Envia Rápido": "Tabela Contratual V. 02/2024 (Tarifa Negociada)",
    "LOGBG": "API SSW SOAP (Em validação de credencial)",
}

TABELAS = {
    "AGEX": {
        "cubagem": 300,
        "faixas_peso": [10, 30, 50, 70, 100],
        "rotas": {
            ("PR", "CAPITAL"): {"vals": [36.44, 39.88, 44.10, 48.50, 53.34], "exc": 0.292, "adv": 0.0030, "gris": 0.0020},
            ("PR", "INT1"):    {"vals": [41.91, 46.09, 50.71, 55.77, 61.36], "exc": 0.486, "adv": 0.0030, "gris": 0.0020},
            ("PR", "INT2"):    {"vals": [48.18, 53.00, 58.30, 64.88, 70.53], "exc": 0.528, "adv": 0.0030, "gris": 0.0020},
            ("SC", "CAPITAL"): {"vals": [48.18, 53.00, 58.30, 64.88, 70.53], "exc": 0.528, "adv": 0.0030, "gris": 0.0020},
            ("SC", "INT1"):    {"vals": [48.18, 53.00, 58.30, 64.88, 70.53], "exc": 0.528, "adv": 0.0030, "gris": 0.0020},
            ("SC", "INT2"):    {"vals": [55.66, 61.23, 67.34, 74.09, 81.49], "exc": 0.694, "adv": 0.0030, "gris": 0.0020},
            ("RS", "CAPITAL"): {"vals": [63.67, 70.03, 77.02, 84.74, 93.21], "exc": 0.694, "adv": 0.0030, "gris": 0.0020},
            ("RS", "INT1"):    {"vals": [76.40, 84.02, 92.43, 101.68, 111.84], "exc": 0.972, "adv": 0.0030, "gris": 0.0020},
            ("RS", "INT2"):    {"vals": [89.11, 98.04, 107.83, 117.23, 129.95], "exc": 1.111, "adv": 0.0030, "gris": 0.0020},
            ("SP", "CAPITAL"): {"vals": [60.72, 69.53, 76.49, 84.14, 92.55], "exc": 0.972, "adv": 0.0030, "gris": 0.0020},
            ("SP", "INT1"):    {"vals": [54.64, 60.12, 66.11, 72.73, 80.01], "exc": 0.833, "adv": 0.0030, "gris": 0.0020},
            ("SP", "INT2"):    {"vals": [54.64, 60.12, 66.11, 72.73, 80.01], "exc": 0.833, "adv": 0.0030, "gris": 0.0020},
            ("MS", "CAPITAL"): {"vals": [58.12, 63.93, 70.33, 77.35, 85.10], "exc": 1.250, "adv": 0.0030, "gris": 0.0020},
            ("MS", "INT1"):    {"vals": [63.31, 69.62, 76.60, 84.25, 92.68], "exc": 1.319, "adv": 0.0035, "gris": 0.0020},
            ("MS", "INT2"):    {"vals": [76.40, 84.02, 92.43, 101.68, 111.84], "exc": 1.389, "adv": 0.0040, "gris": 0.0020},
        },
        "pedagio": {"PR": 6.38, "SC": 6.52, "SP": 6.52, "RS": 7.23, "MS": 12.13},
        "gris_min": 4.70, "adv_min": 0.0, "tas": 0.0, "despacho": 0.0, "default_regiao": "INT1",
    },
    "COOPEX": {
        "cubagem": 300,
        "faixas_peso": [10, 20, 40, 60, 100],
        "rotas": {
            ("PR", "EXPRESS"): {"vals": [17.36, 21.49, 25.76, 29.61, 46.48], "exc": 0.381, "adv": 0.0036, "gris": 0.0018},
            ("PR", "INT1"):    {"vals": [22.02, 26.45, 29.22, 33.53, 50.91], "exc": 0.423, "adv": 0.0036, "gris": 0.0018},
            ("PR", "INT2"):    {"vals": [34.07, 37.92, 44.27, 49.15, 64.44], "exc": 0.510, "adv": 0.0036, "gris": 0.0018},
            ("SC", "LITORAL_N"): {"vals": [26.98, 33.00, 38.65, 44.16, 59.09], "exc": 0.554, "adv": 0.0036, "gris": 0.0018},
            ("SC", "LITORAL_S"): {"vals": [31.56, 35.72, 42.39, 49.96, 64.96], "exc": 0.576, "adv": 0.0036, "gris": 0.0018},
            ("SC", "OESTE"):     {"vals": [36.22, 44.29, 51.98, 60.13, 72.73], "exc": 0.614, "adv": 0.0036, "gris": 0.0018},
            ("RS", "CAPITAL"): {"vals": [37.70, 45.02, 54.22, 61.39, 73.15], "exc": 0.633, "adv": 0.0036, "gris": 0.0018},
            ("RS", "INT1"):    {"vals": [39.40, 47.40, 56.49, 65.35, 77.60], "exc": 0.649, "adv": 0.0036, "gris": 0.0018},
            ("RS", "FRONTEIRA"): {"vals": [57.79, 62.97, 74.40, 86.42, 102.25], "exc": 0.826, "adv": 0.0036, "gris": 0.0018},
            ("SP", "INT1"):    {"vals": [32.64, 36.96, 43.85, 51.68, 67.20], "exc": 0.599, "adv": 0.0036, "gris": 0.0018},
            ("SP", "INT2"):    {"vals": [35.91, 40.65, 48.24, 56.85, 73.92], "exc": 0.628, "adv": 0.0036, "gris": 0.0018},
        },
        "pedagio": {"PR": 4.5, "SC": 4.5, "RS": 4.5, "SP": 4.5},
        "gris_min": 3.15, "adv_min": 4.9, "tas": 3.15, "despacho": 0.0, "default_regiao": "INT1",
    },
    "PRINCESA": {
        "cubagem": 300,
        "faixas_peso": [10, 20, 30, 50, 70, 100],
        "rotas": {
            ("SP", "CAPITAL"): {"vals": [42.15, 45.96, 49.80, 57.47, 61.31, 76.63], "exc": 0.77, "adv": 0.0033, "gris": 0.0011, "ped": 5.9},
            ("PR", "CAPITAL"): {"vals": [29.65, 32.35, 35.05, 40.44, 43.16, 53.91], "exc": 0.55, "adv": 0.0033, "gris": 0.0011, "ped": 4.97},
            ("SC", "CAPITAL"): {"vals": [42.58, 49.38, 51.07, 59.58, 68.11, 85.11], "exc": 0.87, "adv": 0.0033, "gris": 0.0011, "ped": 5.9},
            ("RS", "CAPITAL"): {"vals": [54.89, 61.01, 73.21, 85.42, 97.62, 122.02], "exc": 1.23, "adv": 0.0033, "gris": 0.0011, "ped": 5.9},
            ("RS", "INT1"):    {"vals": [72.37, 80.39, 96.50, 112.57, 128.65, 160.79], "exc": 1.62, "adv": 0.0033, "gris": 0.0011, "ped": 5.9},
            ("SC", "INT1"):    {"vals": [57.24, 66.40, 68.68, 80.13, 91.56, 114.46], "exc": 1.15, "adv": 0.0033, "gris": 0.0011, "ped": 5.9},
            ("SP", "INT1"):    {"vals": [45.20, 49.31, 53.41, 61.63, 65.73, 82.18], "exc": 0.80, "adv": 0.0033, "gris": 0.0011, "ped": 5.9},
        },
        "pedagio": {}, "gris_min": 0.0, "adv_min": 0.0, "tas": 0.0, "despacho": 0.0, "default_regiao": "INT1",
    },
    "TW": {
        "cubagem": 300,
        "faixas_peso": [10, 20, 30, 50, 70, 100],
        "rotas": {
            ("RS", "CAPITAL"): {"vals": [31.29, 37.86, 44.62, 61.23, 76.12, 112.50], "exc": 1.351, "adv": 0.0040, "gris": 0.0020, "ped": 4.00},
            ("RS", "INT1"):    {"vals": [32.26, 39.12, 46.18, 63.71, 85.71, 107.53], "exc": 1.250, "adv": 0.0035, "gris": 0.0020, "ped": 4.00},
            ("SC", "CAPITAL"): {"vals": [26.80, 31.99, 37.17, 51.24, 64.74, 80.41], "exc": 0.980, "adv": 0.0025, "gris": 0.0020, "ped": 4.00},
            ("SC", "INT1"):    {"vals": [28.25, 35.25, 41.44, 60.56, 74.40, 95.38], "exc": 1.105, "adv": 0.0030, "gris": 0.0020, "ped": 4.00},
            ("PR", "CAPITAL"): {"vals": [22.81, 28.05, 32.52, 45.41, 55.79, 69.53], "exc": 0.770, "adv": 0.0025, "gris": 0.0020, "ped": 4.00},
            ("PR", "INT1"):    {"vals": [23.49, 27.54, 31.59, 44.06, 53.78, 67.40], "exc": 0.804, "adv": 0.0025, "gris": 0.0020, "ped": 4.00},
            ("PR", "INT2"):    {"vals": [16.81, 20.54, 23.26, 29.53, 32.81, 38.63], "exc": 0.355, "adv": 0.0020, "gris": 0.0020, "ped": 4.00},
            ("SP", "CAPITAL"): {"vals": [36.24, 42.56, 48.82, 66.92, 76.63, 98.29], "exc": 0.705, "adv": 0.0025, "gris": 0.0020, "ped": 4.00},
            ("SP", "INT1"):    {"vals": [39.14, 45.89, 52.66, 74.43, 89.62, 112.44], "exc": 0.844, "adv": 0.0025, "gris": 0.0020, "ped": 4.00},
            ("SP", "INT2"):    {"vals": [40.91, 47.75, 54.87, 78.17, 94.62, 118.75], "exc": 0.913, "adv": 0.0025, "gris": 0.0020, "ped": 4.00},
        },
        "pedagio": {}, "gris_min": 5.00, "adv_min": 0.0, "tas": 4.00, "despacho": 0.0, "default_regiao": "INT1",
    },
    "ENVIA_RAPIDO": {
        "cubagem": 300,
        "faixas_peso": [20, 35, 50, 70, 100],
        "rotas": {
            ("RJ", "GERAL"): {"vals": [220, 230, 235, 255, 285], "exc_ton": 2400, "adv": 0.0030, "gris": 0.0020, "ped": 10, "desp": 10},
            ("SP", "GERAL"): {"vals": [85, 90, 95, 105, 115], "exc_ton": 950, "adv": 0.0030, "gris": 0.0020, "ped": 10, "desp": 10},
            ("SP", "CAPITAL"): {"vals": [105, 115, 125, 135, 145], "exc_ton": 1000, "adv": 0.0030, "gris": 0.0020, "ped": 10, "desp": 10},
            ("MG", "GERAL"): {"vals": [220, 235, 245, 260, 290], "exc_ton": 2400, "adv": 0.0030, "gris": 0.0020, "ped": 10, "desp": 20},
            ("BA", "GERAL"): {"vals": [255, 265, 275, 285, 320], "exc_ton": 2500, "adv": 0.0050, "gris": 0.0030, "ped": 10, "desp": 10},
            ("ES", "GERAL"): {"vals": [255, 265, 275, 285, 320], "exc_ton": 2500, "adv": 0.0050, "gris": 0.0020, "ped": 10, "desp": 10},
            ("PR", "GERAL"): {"vals": [40, 45, 50, 55, 60], "exc_ton": 430, "adv": 0.0020, "gris": 0.0010, "ped": 5, "desp": 5},
            ("PR", "CAPITAL"): {"vals": [40, 45, 50, 55, 60], "exc_ton": 430, "adv": 0.0020, "gris": 0.0010, "ped": 0, "desp": 5},
            ("GO", "GERAL"): {"vals": [140, 155, 165, 175, 190], "exc_ton": 1150, "adv": 0.0030, "gris": 0.0020, "ped": 10, "desp": 10},
            ("MT", "GERAL"): {"vals": [190, 215, 235, 258, 270], "exc_ton": 2400, "adv": 0.0050, "gris": 0.0030, "ped": 12, "desp": 20, "tas": 15},
            ("MS", "GERAL"): {"vals": [85, 90, 100, 115, 120], "exc_ton": 850, "adv": 0.0030, "gris": 0.0020, "ped": 5, "desp": 5, "tas": 5},
            ("DF", "GERAL"): {"vals": [110, 125, 145, 165, 180], "exc_ton": 1150, "adv": 0.0030, "gris": 0.0020, "ped": 12, "desp": 10, "tas": 10},
            ("RO", "GERAL"): {"vals": [110, 125, 135, 150, 180], "exc_ton": 1150, "adv": 0.0030, "gris": 0.0020, "ped": 10, "desp": 10, "tas": 10},
            ("PB", "GERAL"): {"vals": [255, 265, 275, 285, 320], "exc_ton": 2500, "adv": 0.0030, "gris": 0.0020, "ped": 12, "desp": 20},
            ("PE", "GERAL"): {"vals": [255, 265, 275, 285, 320], "exc_ton": 2500, "adv": 0.0030, "gris": 0.0020, "ped": 12, "desp": 20},
            ("CE", "GERAL"): {"vals": [255, 265, 275, 285, 320], "exc_ton": 2500, "adv": 0.0030, "gris": 0.0020, "ped": 12, "desp": 20},
            ("RN", "GERAL"): {"vals": [255, 265, 275, 285, 320], "exc_ton": 2500, "adv": 0.0030, "gris": 0.0020, "ped": 12, "desp": 20},
            ("SE", "GERAL"): {"vals": [255, 265, 275, 285, 320], "exc_ton": 2500, "adv": 0.0030, "gris": 0.0020, "ped": 12, "desp": 20},
            ("AL", "GERAL"): {"vals": [255, 265, 275, 285, 320], "exc_ton": 2500, "adv": 0.0030, "gris": 0.0020, "ped": 12, "desp": 20},
            ("MA", "GERAL"): {"vals": [255, 265, 275, 285, 320], "exc_ton": 2500, "adv": 0.0030, "gris": 0.0020, "ped": 12, "desp": 20},
            ("PA", "GERAL"): {"vals": [255, 265, 275, 285, 320], "exc_ton": 2500, "adv": 0.0030, "gris": 0.0002, "ped": 12, "desp": 20},
            ("TO", "GERAL"): {"vals": [85, 90, 95, 100, 110], "exc_ton": 950, "adv": 0.0030, "gris": 0.0020, "ped": 10, "desp": 10},
        },
        "pedagio": {}, "gris_min": 0.0, "adv_min": 0.0, "tas": 0.0, "despacho": 0.0, "default_regiao": "GERAL",
    },
}

CAPITAIS = {
    "curitiba": ("PR", "CAPITAL"), "londrina": ("PR", "CAPITAL"), "maringa": ("PR", "CAPITAL"),
    "maringá": ("PR", "CAPITAL"), "florianopolis": ("SC", "CAPITAL"), "florianópolis": ("SC", "CAPITAL"),
    "joinville": ("SC", "CAPITAL"), "porto alegre": ("RS", "CAPITAL"),
    "sao paulo": ("SP", "CAPITAL"), "são paulo": ("SP", "CAPITAL"),
    "rio de janeiro": ("RJ", "GERAL"), "campinas": ("SP", "INT1"),
    "campo grande": ("MS", "CAPITAL"), "brasilia": ("DF", "GERAL"), "brasília": ("DF", "GERAL"),
}


# ===================== PERSISTÊNCIA DE HISTÓRICO =====================
def carregar_historico():
    if "historico" not in st.session_state:
        st.session_state.historico = []
        if os.path.exists(HISTORICO_FILE):
            try:
                with open(HISTORICO_FILE, "r", encoding="utf-8") as f:
                    st.session_state.historico = json.load(f)
            except Exception:
                st.session_state.historico = []
    return st.session_state.historico


def salvar_historico_item(item):
    historico = carregar_historico()
    historico.insert(0, item)
    if len(historico) > 100:
        historico = historico[:100]
    st.session_state.historico = historico
    try:
        with open(HISTORICO_FILE, "w", encoding="utf-8") as f:
            json.dump(historico, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def limpar_historico_arquivo():
    st.session_state.historico = []
    try:
        if os.path.exists(HISTORICO_FILE):
            os.remove(HISTORICO_FILE)
    except Exception:
        pass


# ===================== UTILITÁRIOS =====================
def formatar_moeda(valor):
    if valor is None or pd.isna(valor):
        return "-"
    return f"R$ {float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def limpar_mensagem_ssw(texto):
    if not texto:
        return ""
    texto = html.unescape(str(texto))
    texto = re.sub(r"<[^>]+>", " ", texto)
    return re.sub(r"\s+", " ", texto).strip()


def limpar_documento(valor):
    return re.sub(r"\D", "", str(valor or ""))


def limpar_cep(valor):
    return int(re.sub(r"\D", "", str(valor or "0")))


def xml_escape(valor):
    return html.escape(str(valor or ""), quote=True)


@st.cache_data(ttl=3600)
def buscar_endereco_cep(cep):
    """ViaCEP — retorna endereço completo."""
    cep_limpo = re.sub(r"\D", "", str(cep or ""))
    if len(cep_limpo) != 8:
        return {"ok": False, "texto": "", "cidade": "", "uf": "", "logradouro": "", "bairro": ""}
    try:
        headers = {"User-Agent": BROWSER_USER_AGENT}
        res = requests.get(f"https://viacep.com.br/ws/{cep_limpo}/json/", headers=headers, timeout=6)
        if res.status_code == 200:
            d = res.json()
            if "erro" not in d:
                logradouro = d.get("logradouro", "")
                bairro = d.get("bairro", "")
                cidade = d.get("localidade", "")
                uf = d.get("uf", "")
                partes = [p for p in [logradouro, bairro, f"{cidade}/{uf}" if cidade else ""] if p]
                return {
                    "ok": True,
                    "texto": " · ".join(partes),
                    "cidade": cidade,
                    "uf": uf,
                    "logradouro": logradouro,
                    "bairro": bairro,
                    "cep": d.get("cep", cep_limpo),
                }
    except Exception:
        pass
    return {"ok": False, "texto": "", "cidade": "", "uf": "", "logradouro": "", "bairro": ""}


@st.cache_data(ttl=3600)
def buscar_empresa_cnpj(cnpj):
    """BrasilAPI — razão social, CEP e endereço completo do CNPJ."""
    cnpj_limpo = limpar_documento(cnpj)
    if len(cnpj_limpo) != 14:
        return {"ok": False, "texto": "", "cep": "", "razao": ""}
    try:
        headers = {"User-Agent": BROWSER_USER_AGENT}
        res = requests.get(f"https://brasilapi.com.br/api/cnpj/v1/{cnpj_limpo}", headers=headers, timeout=12)
        if res.status_code == 200:
            d = res.json()
            nome = d.get("razao_social") or d.get("nome_fantasia") or ""
            logradouro = d.get("logradouro", "")
            numero = d.get("numero", "")
            bairro = d.get("bairro", "")
            cidade = d.get("municipio", "")
            uf = d.get("uf", "")
            cep = re.sub(r"\D", "", str(d.get("cep", "")))
            end = " ".join(x for x in [logradouro, numero] if x).strip()
            partes = [p for p in [nome, end, bairro, f"{cidade}/{uf}" if cidade else "", f"CEP {cep}" if cep else ""] if p]
            return {
                "ok": True,
                "texto": " · ".join(partes),
                "razao": nome,
                "cidade": cidade,
                "uf": uf,
                "cep": cep,
                "logradouro": logradouro,
                "bairro": bairro,
            }
    except Exception:
        pass
    return {"ok": False, "texto": "", "cep": "", "razao": ""}


def mapear_regiao(uf, cidade, transportadora):
    uf = (uf or "").upper().strip()
    cidade_n = (cidade or "").lower().strip()
    tab = TABELAS.get(transportadora, {})
    rotas = tab.get("rotas", {})

    if cidade_n in CAPITAIS:
        u, r = CAPITAIS[cidade_n]
        if u == uf and (uf, r) in rotas:
            return uf, r

    if transportadora == "ENVIA_RAPIDO":
        if (uf, "CAPITAL") in rotas and cidade_n in ["sao paulo", "são paulo", "curitiba", "londrina"]:
            return uf, "CAPITAL"
        if (uf, "GERAL") in rotas:
            return uf, "GERAL"

    if transportadora == "COOPEX" and uf == "PR":
        express = ["apucarana", "arapongas", "cornelio procopio", "cornélio procópio", "londrina"]
        if any(x in cidade_n for x in express) and (uf, "EXPRESS") in rotas:
            return uf, "EXPRESS"
    if transportadora == "COOPEX" and uf == "SC":
        if any(x in cidade_n for x in ["florianopolis", "florianópolis", "tubarao", "criciuma", "criciúma"]):
            return uf, "LITORAL_S"
        if any(x in cidade_n for x in ["joinville", "blumenau", "itajai"]):
            return uf, "LITORAL_N"
        if (uf, "OESTE") in rotas:
            return uf, "OESTE"

    default = tab.get("default_regiao", "INT1")
    if (uf, default) in rotas:
        return uf, default
    for (u, r) in rotas:
        if u == uf:
            return u, r
    return None, None


def calcular_frete_tabela(nome, peso, volume_m3, valor_nf, uf_destino, cidade_destino):
    tab = TABELAS.get(nome)
    if not tab:
        return None

    uf, regiao = mapear_regiao(uf_destino, cidade_destino, nome)
    if not uf or not regiao:
        ufs_ok = sorted({u for (u, _) in tab.get("rotas", {})})
        return {
            "Transportadora": nome,
            "Nº Cotação": "Tabela",
            "Valor Frete (R$)": None,
            "Prazo (Dias Úteis)": 0,
            "Status": f"Não atende {uf_destino or '?'} (cobre: {', '.join(ufs_ok)})",
            "Fonte": "tabela",
            "Composição": None,
        }

    rota = tab["rotas"][(uf, regiao)]
    fator = tab.get("cubagem", 300)
    peso_cubado = volume_m3 * fator if volume_m3 else 0
    peso_calc = max(float(peso), peso_cubado)

    faixas = tab["faixas_peso"]
    vals = rota["vals"]
    frete_peso = None
    faixa_atingida = ""
    for i, limite in enumerate(faixas):
        if peso_calc <= limite:
            frete_peso = vals[i]
            faixa_atingida = f"Até {limite} kg"
            break
    if frete_peso is None:
        if "exc_ton" in rota:
            exc_calc = ((peso_calc - faixas[-1]) / 1000.0) * rota["exc_ton"]
            frete_peso = vals[-1] + exc_calc
            faixa_atingida = f"Acima de {faixas[-1]} kg (+ R$ {rota['exc_ton']:.2f}/ton)"
        else:
            exc_calc = (peso_calc - faixas[-1]) * rota.get("exc", 0)
            frete_peso = vals[-1] + exc_calc
            faixa_atingida = f"Acima de {faixas[-1]} kg (+ R$ {rota.get('exc', 0):.2f}/kg)"

    adv_pct = rota.get("adv", 0)
    gris_pct = rota.get("gris", 0)
    frete_valor = float(valor_nf) * adv_pct
    if tab.get("adv_min"):
        frete_valor = max(frete_valor, tab["adv_min"])
    gris = float(valor_nf) * gris_pct
    if tab.get("gris_min"):
        gris = max(gris, tab["gris_min"])

    ped_unit = rota.get("ped", tab.get("pedagio", {}).get(uf, 0))
    fracoes = max(1, math.ceil(peso_calc / 100)) if peso_calc > 0 else 1
    pedagio = fracoes * ped_unit
    despacho = rota.get("desp", tab.get("despacho", 0))
    tas = rota.get("tas", tab.get("tas", 0))
    total = frete_peso + frete_valor + gris + pedagio + despacho + tas

    composicao = {
        "Região / Rota": f"{uf}/{regiao}",
        "Faixa Aplicada": faixa_atingida,
        "Peso Tarifado": f"{peso_calc:.2f} kg ({'Cubado' if peso_cubado > float(peso) else 'Real'})",
        "Frete Peso": round(frete_peso, 2),
        "Ad Valorem": round(frete_valor, 2),
        "Taxa Ad Valorem": f"{adv_pct * 100:.2f}%" if adv_pct else "Isento",
        "GRIS": round(gris, 2),
        "Taxa GRIS": f"{gris_pct * 100:.2f}%" if gris_pct else "Isento",
        "Pedágio": round(pedagio, 2),
        "Frações Pedágio": f"{fracoes}x (R$ {ped_unit:.2f}/fração)",
        "Despacho": round(despacho, 2),
        "TAS": round(tas, 2),
        "Total": round(total, 2),
    }

    return {
        "Transportadora": nome,
        "Nº Cotação": f"Tabela {uf}/{regiao}",
        "Valor Frete (R$)": round(total, 2),
        "Prazo (Dias Úteis)": 0,
        "Status": "Tarifa negociada estimada",
        "Fonte": "tabela",
        "Composição": composicao,
    }


# ===================== INTEGRAÇÃO BRASPRESS (OFICIAL API) =====================
def cotar_braspress(cep_origem, cep_destino, peso, valor_nf, cubagem_lista, cnpj_remetente, doc_destinatario):
    token = base64.b64encode(f"{USUARIO_BRASPRESS}:{SENHA_BRASPRESS}".encode()).decode()
    headers = {
        "Authorization": f"Basic {token}",
        "Content-Type": "application/json",
        "Accept": "application/json",
        "User-Agent": BROWSER_USER_AGENT,
    }
    url = "https://api.braspress.com/v1/cotacao/calcular/json"
    remetente_limpo = limpar_documento(cnpj_remetente)
    doc_dest_limpo = limpar_documento(doc_destinatario)
    if not doc_dest_limpo or len(doc_dest_limpo) < 11:
        doc_dest_limpo = remetente_limpo
    total_volumes = sum(int(item.get("Volumes", 0)) for item in cubagem_lista)
    payload = {
        "CnpjRemetente": remetente_limpo,
        "CnpjDestinatario": doc_dest_limpo,
        "Modal": "R",
        "TipoFrete": 1,
        "CepOrigem": limpar_cep(cep_origem),
        "CepDestino": limpar_cep(cep_destino),
        "Peso": float(peso),
        "ValorDeclarado": float(valor_nf),
        "Volumes": int(total_volumes) if total_volumes > 0 else 1,
        "Cubagem": cubagem_lista,
    }
    try:
        response = requests.post(url, json=payload, headers=headers, timeout=18)
        if response.status_code in [200, 201]:
            dados = response.json()
            prazo_dias = (
                dados.get("prazoEntrega") or dados.get("prazo")
                or dados.get("prazoDias") or dados.get("prazoDiasUteis") or 0
            )
            val_total = float(dados.get("totalFrete", dados.get("valorTotal", 0.0)))
            return {
                "Transportadora": "Braspress",
                "Nº Cotação": str(dados.get("id", dados.get("numeroCotacao", "N/A"))),
                "Valor Frete (R$)": val_total,
                "Prazo (Dias Úteis)": int(prazo_dias),
                "Status": "API oficial confirmada",
                "Fonte": "api",
                "Composição": None,
            }
        try:
            err_data = response.json()
            err_msg = err_data.get("message") or err_data.get("error") or str(err_data)
        except Exception:
            err_msg = response.text[:60]
        return {
            "Transportadora": "Braspress", "Nº Cotação": "-", "Valor Frete (R$)": None,
            "Prazo (Dias Úteis)": 0, "Status": f"API {response.status_code}: {err_msg[:45]}", "Fonte": "api",
            "Composição": None,
        }
    except Exception as e:
        return {
            "Transportadora": "Braspress", "Nº Cotação": "-", "Valor Frete (R$)": None,
            "Prazo (Dias Úteis)": 0, "Status": f"API falha de conexão ({str(e)[:40]})", "Fonte": "api",
            "Composição": None,
        }


def rastrear_braspress(nf, cnpj_remetente):
    """Rastreia NF na Braspress com cabeçalho de navegador para evitar HTTP 403."""
    token = base64.b64encode(f"{USUARIO_BRASPRESS}:{SENHA_BRASPRESS}".encode()).decode()
    headers = {
        "Authorization": f"Basic {token}",
        "Content-Type": "application/json",
        "Accept": "application/json",
        "User-Agent": BROWSER_USER_AGENT,
    }
    nf_limpa = re.sub(r"\D", "", str(nf or ""))
    cnpj_limpo = limpar_documento(cnpj_remetente)
    url = f"https://api.braspress.com/v1/tracking/remetente/{cnpj_limpo}/nf/{nf_limpa}"
    try:
        r = requests.get(url, headers=headers, timeout=15)
        if r.status_code in [200, 201]:
            data = r.json()
            eventos = data.get("tracking", data.get("eventos", []))
            if not eventos and isinstance(data, list):
                eventos = data

            status_geral = "Em processamento"
            ultima_ocorrencia = ""
            data_ocorrencia = ""
            data_entrega = data.get("dataEntrega", "")
            previsao_entrega = data.get("previsaoEntrega", data.get("dataPrevisao", ""))

            if eventos and len(eventos) > 0:
                ultimo_evt = eventos[-1]
                descricao_evt = str(ultimo_evt.get("descricao", ultimo_evt.get("status", "")))
                data_ocorrencia = str(ultimo_evt.get("data", ultimo_evt.get("dataHora", "")))
                ultima_ocorrencia = descricao_evt

                desc_lower = descricao_evt.lower()
                if "entregue" in desc_lower or "entrega realizada" in desc_lower:
                    status_geral = "Entregue"
                    if not data_entrega:
                        data_entrega = data_ocorrencia
                elif "saiu para entrega" in desc_lower or "em rota" in desc_lower:
                    status_geral = "Saiu para Entrega"
                elif "retido" in desc_lower or "avaria" in desc_lower or "recusad" in desc_lower or "devolu" in desc_lower:
                    status_geral = "Ocorrência / Retido"
                elif "transito" in desc_lower or "transferencia" in desc_lower or "coletado" in desc_lower:
                    status_geral = "Em Trânsito"
            elif data.get("status"):
                status_geral = str(data.get("status"))

            return {
                "sucesso": True,
                "codigo": r.status_code,
                "nf": nf_limpa,
                "status": status_geral,
                "previsao": previsao_entrega,
                "data_entrega": data_entrega,
                "ultima_ocorrencia": ultima_ocorrencia,
                "data_ocorrencia": data_ocorrencia,
                "destino": data.get("cidadeDestino", data.get("destino", "")),
                "valor_frete": data.get("totalFrete", data.get("valorFrete", "")),
                "eventos": eventos,
                "raw": data,
            }
        return {
            "sucesso": False,
            "codigo": r.status_code,
            "nf": nf_limpa,
            "status": f"NF não localizada (HTTP {r.status_code})" if r.status_code == 404 else f"Erro {r.status_code}",
            "previsao": "",
            "data_entrega": "",
            "ultima_ocorrencia": "Sem registro recente na transportadora",
            "data_ocorrencia": "",
            "destino": "",
            "valor_frete": "",
            "eventos": [],
            "raw": r.text,
        }
    except Exception as e:
        return {
            "sucesso": False,
            "codigo": 500,
            "nf": nf_limpa,
            "status": f"Falha de conexão ({str(e)[:35]})",
            "previsao": "",
            "data_entrega": "",
            "ultima_ocorrencia": str(e)[:50],
            "data_ocorrencia": "",
            "destino": "",
            "valor_frete": "",
            "eventos": [],
            "raw": str(e),
        }


# ===================== INTEGRAÇÃO SSW WSDL (OFICIAL API) =====================
def cotar_ssw_wsdl(
    nome_transportadora, dominio, usuario, senha,
    cep_origem, cep_destino, peso, valor_nf, cubagem_lista,
    cnpj_remetente, doc_destinatario, mercadoria=1,
):
    url = "https://ssw.inf.br/ws/sswCotacao/index.php"
    headers = {
        "Content-Type": "text/xml; charset=utf-8",
        "SOAPAction": "urn:sswinfbr.sswCotacao#cotacao",
        "User-Agent": BROWSER_USER_AGENT,
    }
    remetente = limpar_documento(cnpj_remetente)
    destinatario = limpar_documento(doc_destinatario)
    if not destinatario:
        destinatario = remetente
    cep_o = limpar_cep(cep_origem)
    cep_d = limpar_cep(cep_destino)
    total_volumes = sum(int(item.get("Volumes", 0)) for item in cubagem_lista)
    if total_volumes <= 0:
        total_volumes = 1

    volume_total = 0.0
    for item in cubagem_lista:
        try:
            volume_total += (
                float(item.get("Altura", 0)) * float(item.get("Largura", 0))
                * float(item.get("Comprimento", 0)) * int(item.get("Volumes", 1))
            )
        except Exception:
            pass

    primeiro = cubagem_lista[0] if cubagem_lista else {"Altura": 0.3, "Largura": 0.3, "Comprimento": 0.4}
    altura = float(primeiro.get("Altura", 0.3))
    largura = float(primeiro.get("Largura", 0.3))
    comprimento = float(primeiro.get("Comprimento", 0.4))

    def montar_envelope(pagador):
        return f"""<?xml version="1.0" encoding="UTF-8"?>
<soapenv:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
 xmlns:xsd="http://www.w3.org/2001/XMLSchema"
 xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
 xmlns:urn="urn:sswinfbr.sswCotacao">
  <soapenv:Header/>
  <soapenv:Body>
    <urn:cotar soapenv:encodingStyle="http://schemas.xmlsoap.org/soap/encoding/">
      <dominio xsi:type="xsd:string">{xml_escape(dominio)}</dominio>
      <login xsi:type="xsd:string">{xml_escape(usuario)}</login>
      <senha xsi:type="xsd:string">{xml_escape(senha)}</senha>
      <cnpjPagador xsi:type="xsd:string">{xml_escape(pagador)}</cnpjPagador>
      <cepOrigem xsi:type="xsd:integer">{cep_o}</cepOrigem>
      <cepDestino xsi:type="xsd:integer">{cep_d}</cepDestino>
      <valorNF xsi:type="xsd:decimal">{float(valor_nf):.2f}</valorNF>
      <quantidade xsi:type="xsd:integer">{total_volumes}</quantidade>
      <peso xsi:type="xsd:decimal">{float(peso):.3f}</peso>
      <volume xsi:type="xsd:decimal">{volume_total:.4f}</volume>
      <mercadoria xsi:type="xsd:integer">{int(mercadoria)}</mercadoria>
      <cnpjDestinatario xsi:type="xsd:string">{xml_escape(destinatario)}</cnpjDestinatario>
      <coletar xsi:type="xsd:string">S</coletar>
      <entDificil xsi:type="xsd:string">N</entDificil>
      <destContribuinte xsi:type="xsd:string">S</destContribuinte>
      <qtdePares xsi:type="xsd:integer">0</qtdePares>
      <altura xsi:type="xsd:decimal">{altura:.3f}</altura>
      <largura xsi:type="xsd:decimal">{largura:.3f}</largura>
      <comprimento xsi:type="xsd:decimal">{comprimento:.3f}</comprimento>
      <fatorMultiplicador xsi:type="xsd:integer">1</fatorMultiplicador>
      <cnpjRemetente xsi:type="xsd:string">{xml_escape(remetente)}</cnpjRemetente>
    </urn:cotar>
  </soapenv:Body>
</soapenv:Envelope>"""

    tentativas = [remetente, destinatario]
    ultimo_erro = ""

    for pagador_atual in tentativas:
        envelope = montar_envelope(pagador_atual)
        try:
            resp = requests.post(url, data=envelope.encode("utf-8"), headers=headers, timeout=14)
            if resp.status_code >= 500:
                ultimo_erro = f"HTTP {resp.status_code}"
                continue
            root_soap = ET.fromstring(resp.content)
            return_elem = next((e for e in root_soap.iter() if e.tag.endswith("return")), None)
            if return_elem is None or not (return_elem.text or "").strip():
                ultimo_erro = "Retorno vazio da API"
                continue

            cotacao = ET.fromstring(return_elem.text.strip())

            def find_text(tag):
                el = cotacao.find(tag)
                if el is not None and el.text:
                    return el.text.strip()
                for item in cotacao.iter():
                    if item.tag.split("}")[-1] == tag and item.text:
                        return item.text.strip()
                return ""

            erro = int(find_text("erro") or -1)
            mensagem = limpar_mensagem_ssw(find_text("mensagem"))

            if erro in (0, 1):
                frete_texto = (find_text("totalFrete") or "0").strip()
                if "," in frete_texto and "." in frete_texto:
                    frete_texto = frete_texto.replace(".", "").replace(",", ".")
                elif "," in frete_texto:
                    frete_texto = frete_texto.replace(",", ".")
                frete_val = float(frete_texto)
                prazo_val = int(float(find_text("prazo") or 0))

                status = "API oficial confirmada" if erro == 0 else (
                    f"API com ressalva: {mensagem[:65]}" if mensagem else "API oficial com ressalva"
                )

                def parse_moeda(tag):
                    v = find_text(tag) or "0"
                    v = v.replace(".", "").replace(",", ".") if ("," in v and "." in v) else v.replace(",", ".")
                    try:
                        return float(v)
                    except Exception:
                        return 0.0

                comp_ssw = {
                    "Região / Rota": find_text("tabCalculo") or "Tabela SSW",
                    "Faixa Aplicada": f"Peso Cálculo: {find_text('pesoCalculo') or peso} kg",
                    "Peso Tarifado": f"{find_text('pesoCalculo') or peso} kg",
                    "Frete Peso": parse_moeda("fretePeso"),
                    "Ad Valorem": parse_moeda("freteValor"),
                    "Taxa Ad Valorem": "-",
                    "GRIS": parse_moeda("gris"),
                    "Taxa GRIS": "-",
                    "Pedágio": parse_moeda("pedagio"),
                    "Frações Pedágio": "-",
                    "Despacho": parse_moeda("despacho"),
                    "TAS": parse_moeda("tas"),
                    "Impostos": parse_moeda("impostos"),
                    "Total": frete_val,
                }

                return {
                    "Transportadora": nome_transportadora,
                    "Nº Cotação": "SSW Oficial",
                    "Valor Frete (R$)": frete_val,
                    "Prazo (Dias Úteis)": prazo_val,
                    "Status": status,
                    "Fonte": "api",
                    "Composição": comp_ssw,
                }

            ultimo_erro = mensagem if mensagem else f"Erro SSW {erro}"
            msg_lower = ultimo_erro.lower()
            if any(termo in msg_lower for termo in ["invalido", "inválido", "senha", "dominio", "domínio", "bloqueado", "nao cadastrado"]):
                break

        except Exception as e:
            ultimo_erro = str(e)[:70]
            break

    return {
        "Transportadora": nome_transportadora,
        "Nº Cotação": "-",
        "Valor Frete (R$)": None,
        "Prazo (Dias Úteis)": 0,
        "Status": ultimo_erro[:75] if ultimo_erro else "Falha SSW",
        "Fonte": "api",
        "Composição": None,
    }


# ===================== GERADORES DE EXCEL PROFISSIONAIS =====================
def gerar_excel_cotacao(dados_gerais, df_resultados):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Comparativo de Frete"
    ws.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    sub_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    best_fill = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    info_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

    font_title = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    font_header = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Segoe UI", size=10, bold=True, color="0F172A")
    font_normal = Font(name="Segoe UI", size=10, color="1E293B")
    font_best = Font(name="Segoe UI", size=10, bold=True, color="166534")

    border_thin = Border(
        left=Side(style="thin", color="CBD5E1"),
        right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"),
        bottom=Side(style="thin", color="CBD5E1"),
    )

    ws.merge_cells("A1:G1")
    title_cell = ws["A1"]
    title_cell.value = "SISTEMA DE COTAÇÕES LOGÍSTICAS - NEXT CABLE"
    title_cell.font = font_title
    title_cell.fill = header_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    ws.merge_cells("A2:G2")
    ws["A2"].value = "PARÂMETROS DA COTAÇÃO"
    ws["A2"].font = Font(name="Segoe UI", size=10, bold=True, color="0F172A")
    ws["A2"].fill = info_fill
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 20

    ws["A3"] = "Data/Hora:"
    ws["B3"] = dados_gerais.get("data_hora", "")
    ws["A4"] = "Origem:"
    ws["B4"] = f"{dados_gerais.get('origem_cep', '')} ({dados_gerais.get('origem_texto', '')})"
    ws["A5"] = "Destino:"
    ws["B5"] = f"{dados_gerais.get('destino_cep', '')} ({dados_gerais.get('destino_texto', '')})"

    ws["D3"] = "Valor da NF:"
    ws["E3"] = f"R$ {dados_gerais.get('valor_nf', 0):,.2f}"
    ws["D4"] = "Peso Real:"
    ws["E4"] = f"{dados_gerais.get('peso', 0):.2f} kg"
    ws["D5"] = "Peso Tarifado:"
    ws["E5"] = f"{dados_gerais.get('peso_tarifado', 0):.2f} kg (Cubado: {dados_gerais.get('peso_cubado', 0):.2f} kg)"

    for r in range(3, 6):
        ws[f"A{r}"].font = font_bold
        ws[f"B{r}"].font = font_normal
        ws[f"D{r}"].font = font_bold
        ws[f"E{r}"].font = font_normal

    headers = ["Classificação", "Transportadora", "Nº Cotação", "Tipo", "Valor Frete", "Prazo", "Status / Detalhe"]
    start_row = 7
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=start_row, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = sub_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border_thin
    ws.row_dimensions[start_row].height = 24

    for idx, row in df_resultados.iterrows():
        curr_row = start_row + 1 + idx
        is_best = (idx == 0 and pd.notna(row.get("Valor Frete (R$)")))

        pos_txt = "★ MELHOR" if is_best else (f"{idx+1}º Lugar" if pd.notna(row.get("Valor Frete (R$)")) else "—")
        val_txt = f"R$ {row['Valor Frete (R$)']:,.2f}" if pd.notna(row.get("Valor Frete (R$)")) else "—"
        prazo_num = int(row.get("Prazo (Dias Úteis)") or 0) if pd.notna(row.get("Prazo (Dias Úteis)")) else 0
        prazo_txt = f"{prazo_num} dias úteis" if prazo_num > 0 else "A confirmar"
        tipo_txt = "API Oficial" if row.get("Fonte") == "api" and pd.notna(row.get("Valor Frete (R$)")) else (
            "Tabela Estimada" if row.get("Fonte") == "tabela" and pd.notna(row.get("Valor Frete (R$)")) else "—"
        )

        cells = [
            ws.cell(row=curr_row, column=1, value=pos_txt),
            ws.cell(row=curr_row, column=2, value=str(row.get("Transportadora", ""))),
            ws.cell(row=curr_row, column=3, value=str(row.get("Nº Cotação", ""))),
            ws.cell(row=curr_row, column=4, value=tipo_txt),
            ws.cell(row=curr_row, column=5, value=val_txt),
            ws.cell(row=curr_row, column=6, value=prazo_txt),
            ws.cell(row=curr_row, column=7, value=str(row.get("Status", ""))),
        ]

        for cell in cells:
            cell.font = font_best if is_best else font_normal
            if is_best:
                cell.fill = best_fill
            cell.border = border_thin
            cell.alignment = Alignment(vertical="center")

        cells[0].alignment = Alignment(horizontal="center", vertical="center")
        cells[4].alignment = Alignment(horizontal="right", vertical="center")
        cells[5].alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[curr_row].height = 20

    for col_idx in range(1, len(headers) + 1):
        col_letter = get_column_letter(col_idx)
        max_len = max(len(str(ws.cell(row=r, column=col_idx).value or "")) for r in range(start_row, start_row + 1 + len(df_resultados)))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def gerar_excel_entregas(df_entregas):
    """Gera planilha profissional de monitoramento de entregas."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Monitoramento de Entregas"
    ws.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    font_title = Font(name="Segoe UI", size=12, bold=True, color="FFFFFF")
    font_header = Font(name="Segoe UI", size=9, bold=True, color="FFFFFF")
    font_normal = Font(name="Segoe UI", size=9, color="1E293B")

    border_thin = Border(
        left=Side(style="thin", color="CBD5E1"), right=Side(style="thin", color="CBD5E1"),
        top=Side(style="thin", color="CBD5E1"), bottom=Side(style="thin", color="CBD5E1"),
    )

    ws.merge_cells("A1:I1")
    t = ws["A1"]
    t.value = f"RELATÓRIO DE MONITORAMENTO DE ENTREGAS - NEXT CABLE ({datetime.now():%d/%m/%Y %H:%M})"
    t.font = font_title
    t.fill = header_fill
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    colunas = ["CNPJ", "Nota Fiscal", "Status", "Previsão de Entrega", "Data de Entrega", "Última Ocorrência", "Data Ocorrência", "Destino", "Valor Frete"]
    for col_idx, col_name in enumerate(colunas, 1):
        c = ws.cell(row=2, column=col_idx, value=col_name)
        c.font = font_header
        c.fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border_thin
    ws.row_dimensions[2].height = 22

    fill_entregue = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    fill_transito = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    fill_erro = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

    for row_idx, r in df_entregas.iterrows():
        curr_row = 3 + row_idx
        st_text = str(r.get("Status", ""))
        row_fill = None
        if "entregue" in st_text.lower():
            row_fill = fill_entregue
        elif "trânsito" in st_text.lower() or "transito" in st_text.lower() or "saiu" in st_text.lower():
            row_fill = fill_transito
        elif "ocorrência" in st_text.lower() or "erro" in st_text.lower() or "retido" in st_text.lower():
            row_fill = fill_erro

        for col_idx, col_name in enumerate(colunas, 1):
            val = str(r.get(col_name, "") if pd.notna(r.get(col_name)) else "")
            c = ws.cell(row=curr_row, column=col_idx, value=val)
            c.font = font_normal
            c.border = border_thin
            c.alignment = Alignment(vertical="center")
            if row_fill and col_name == "Status":
                c.fill = row_fill
        ws.row_dimensions[curr_row].height = 19

    for col_idx in range(1, len(colunas) + 1):
        col_letter = get_column_letter(col_idx)
        max_len = max(len(str(ws.cell(row=r, column=col_idx).value or "")) for r in range(2, 3 + len(df_entregas)))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 13)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ====================== SIDEBAR ENTERPRISE ======================
with st.sidebar:
    st.markdown(
        """
        <div style="padding:10px 0 16px 0;border-bottom:1px solid #334155;margin-bottom:16px;">
          <div style="font-size:1.15rem;font-weight:800;color:#ffffff;letter-spacing:-0.02em;">
            NEXT CABLE
          </div>
          <div style="font-size:0.75rem;color:#94a3b8;font-weight:500;margin-top:2px;">
            Sistema de Cotações & Logística
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("<div style='font-size:0.75rem;font-weight:800;color:#FFFFFF;margin-bottom:8px;letter-spacing:0.04em;'>MÓDULOS DO SISTEMA</div>", unsafe_allow_html=True)

    modulos = [
        "Nova Cotação",
        "Rastreamento de Pedidos",
        "Histórico de Cotações (Em Breve)",
    ]

    if "active_module" not in st.session_state or st.session_state.active_module not in modulos:
        st.session_state.active_module = "Nova Cotação"

    for m in modulos:
        is_active = (st.session_state.active_module == m)
        btn_type = "primary" if is_active else "secondary"
        if st.button(m, key=f"nav_{m}", use_container_width=True, type=btn_type):
            st.session_state.active_module = m
            st.rerun()

    aba = st.session_state.active_module

    st.markdown("<div style='margin-top:28px;padding-top:14px;border-top:1px solid #1e293b;'></div>", unsafe_allow_html=True)
    st.markdown("<div style='font-size:0.75rem;font-weight:800;color:#FFFFFF;margin-bottom:8px;letter-spacing:0.04em;'>CONEXÕES & TABELAS</div>", unsafe_allow_html=True)

    with st.expander("Abrangência das 16 Transportadoras"):
        for k, cfg in te.TRANSPORTADORAS_CONFIG.items():
            tags_html = " ".join([f"<span style='background:#1e293b;color:#38bdf8;padding:2px 5px;border-radius:4px;font-size:0.68rem;font-weight:700;margin-right:2px;'>{t}</span>" for t in cfg["abrangencia_tags"]])
            st.markdown(
                f"<div style='margin-bottom:10px;border-bottom:1px solid #1e293b;padding-bottom:6px;'>"
                f"<div style='font-weight:700;color:#f8fafc;font-size:0.8rem;'>{cfg['nome']}</div>"
                f"<div style='margin:3px 0;'>{tags_html}</div>"
                f"<div style='font-size:0.7rem;color:#94a3b8;'>{cfg['descricao']}</div>"
                f"</div>",
                unsafe_allow_html=True
            )

    st.caption(f"{APP_VERSION.split('—')[0].strip()} · Next Cable")


# ==============================================================================
# MÓDULO 1: NOVA COTAÇÃO (COM SMART TRIAGE, GRID DINÂMICO E AUTO-COMPLETE)
# ==============================================================================
if aba == "Nova Cotação":
    # Barra elegante e compacta de Origem
    col_orig_banner, col_orig_btn = st.columns([3.5, 1])
    with col_orig_banner:
        cnpj_rem_default = st.session_state.get("cnpj_rem_custom", "")
        cep_rem_default = st.session_state.get("cep_rem_custom", "")
        st.markdown(
            f"""
            <div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:10px;padding:10px 16px;
                        display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 6px rgba(0,0,0,0.02);margin-bottom:14px;">
              <div>
                <div style="font-weight:700;color:#0f172a;font-size:0.88rem;letter-spacing:-0.01em;">
                  ORIGEM: Matriz Londrina/PR — Next Cable
                </div>
                <div style="color:#64748b;font-size:0.77rem;margin-top:2px;">
                  CNPJ: {cnpj_rem_default} · CEP: {cep_rem_default[:5]}-{cep_rem_default[5:]} · Rodocentro
                </div>
              </div>
              <span style="background:#f1f5f9;color:#475569;padding:4px 8px;border-radius:6px;font-size:0.72rem;font-weight:700;">EXPEDIÇÃO PADRÃO</span>
            </div>
            """,
            unsafe_allow_html=True,
        )

    with col_orig_btn:
        with st.expander("Alterar Origem"):
            cnpj_remetente = st.text_input("CNPJ Origem", value=cnpj_rem_default, key="cnpj_rem_inp")
            cep_origem = st.text_input("CEP Origem", value=cep_rem_default, key="cep_o_inp")
            st.session_state.cnpj_rem_custom = cnpj_remetente
            st.session_state.cep_rem_custom = cep_origem

    cnpj_remetente = st.session_state.get("cnpj_rem_custom", "")
    cep_origem = st.session_state.get("cep_rem_custom", "")
    info_o = buscar_endereco_cep(cep_origem)

    # 1. Dados de Destinatário e Dados Fiscais em 2 colunas equilibradas
    col_dest, col_nf = st.columns([1.6, 1], gap="medium")

    with col_dest:
        with st.container(border=True):
            st.markdown(
                '<div style="font-size:0.8rem;font-weight:800;color:#0f172a;margin-bottom:10px;letter-spacing:0.04em;">'
                'DESTINATÁRIO</div>',
                unsafe_allow_html=True,
            )
            doc_destinatario = st.text_input("CNPJ / CPF do Destinatário", value=st.session_state.get("doc_dest_val", ""), key="doc_dest")
            doc_limpo = limpar_documento(doc_destinatario)

            # Auto-complete inteligente por CNPJ
            info_emp_dest = None
            if len(doc_limpo) == 14:
                info_emp_dest = buscar_empresa_cnpj(doc_limpo)
                if info_emp_dest.get("ok"):
                    if info_emp_dest.get("cep") and st.session_state.get("ultimo_cnpj_buscado") != doc_limpo:
                        st.session_state.cep_d_val = info_emp_dest["cep"]
                        st.session_state.ultimo_cnpj_buscado = doc_limpo

                    st.markdown(
                        f"""
                        <div style="background:#f0fdf4;border-left:3px solid #10b981;border-radius:6px;padding:8px 10px;margin:8px 0;font-size:0.78rem;color:#166534;">
                          <b>{info_emp_dest.get('razao', '')}</b><br>
                          {info_emp_dest.get('texto', '')}
                        </div>
                        """,
                        unsafe_allow_html=True,
                    )

            cep_dest_padrao = st.session_state.get("cep_d_val", "")
            
            entrega_diferente = st.checkbox("Endereço de entrega diferente do faturamento")
            if not entrega_diferente:
                cep_destino = st.text_input("CEP Destino", value=cep_dest_padrao, key="cep_d", disabled=True)
            else:
                cep_destino = st.text_input("CEP Destino", value=cep_dest_padrao, key="cep_d")
                
            st.session_state.cep_d_val = cep_destino

            info_d = buscar_endereco_cep(cep_destino)
            if info_d["ok"]:
                st.markdown(
                    f"""
                    <div style="background:#f8fafc;border-left:3px solid #0284c7;border-radius:6px;padding:7px 10px;font-size:0.78rem;color:#0369a1;">
                      <b>{info_d['cidade']}/{info_d['uf']}</b> · {info_d.get('logradouro','')} {info_d.get('bairro','')}
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

    with col_nf:
        with st.container(border=True):
            st.markdown(
                '<div style="font-size:0.8rem;font-weight:800;color:#0f172a;margin-bottom:10px;letter-spacing:0.04em;">'
                'DADOS DA NOTA FISCAL</div>',
                unsafe_allow_html=True,
            )
            valor_nf_str = st.text_input("Valor Total da NF (R$)", placeholder="Ex: 6.908,13")
            valor_nf = None
            if valor_nf_str:
                try:
                    import re
                    # Mantém apenas números, vírgula e ponto
                    v_str = re.sub(r'[^\d,\.]', '', valor_nf_str)
                    if ',' in v_str:
                        # Se tem vírgula, a última é o separador decimal
                        parts = v_str.rsplit(',', 1)
                        inteiro = parts[0].replace('.', '')
                        decimal = parts[1]
                        valor_nf = float(f"{inteiro}.{decimal}")
                    else:
                        # Se não tem vírgula, trata o ponto (se houver) como decimal
                        valor_nf = float(v_str)
                except ValueError:
                    st.error("Formato inválido. Use um formato como 6.908,13")
            obs_carga = st.text_area("OBSERVAÇÃO (OPCIONAL)")

    # 2. ESPECIFICAÇÃO DE VOLUMES (LARGURA TOTAL 100% — SEM CORTES)
    with st.container(border=True):
        col_vol_h, col_vol_act = st.columns([3, 1])
        with col_vol_h:
            st.markdown(
                '<div style="font-size:0.84rem;font-weight:800;color:#0f172a;margin-bottom:4px;letter-spacing:0.04em;">'
                'ESPECIFICAÇÃO DE VOLUMES</div>'
                '<div style="font-size:0.75rem;color:#64748b;margin-bottom:8px;">Informe a quantidade e as dimensões de cada volume. Os pesos cubado e tarifado são calculados em tempo real.</div>',
                unsafe_allow_html=True,
            )
        with col_vol_act:
            if st.button("Adicionar Volume", use_container_width=True):
                if "volumes_data" not in st.session_state or not st.session_state.volumes_data:
                    st.session_state.volumes_data = [
                        {"Qtd": 1, "Alt (cm)": 30.0, "Larg (cm)": 30.0, "Comp (cm)": 40.0, "Peso Un. (kg)": 10.0}
                    ]
                else:
                    st.session_state.volumes_data.append(
                        {"Qtd": 1, "Alt (cm)": 30.0, "Larg (cm)": 30.0, "Comp (cm)": 40.0, "Peso Un. (kg)": 10.0}
                    )
                st.rerun()

        if "volumes_data" not in st.session_state:
            st.session_state.volumes_data = []

        if len(st.session_state.volumes_data) == 0:
            df_volumes_input = pd.DataFrame(columns=["Qtd", "Alt (cm)", "Larg (cm)", "Comp (cm)", "Peso Un. (kg)"])
        else:
            df_volumes_input = pd.DataFrame(st.session_state.volumes_data)

        edited_volumes = st.data_editor(
            df_volumes_input,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            column_config={
                "Qtd": st.column_config.NumberColumn("Qtd", min_value=1, step=1, default=1),
                "Alt (cm)": st.column_config.NumberColumn("Alt (cm)", min_value=1.0, step=5.0, default=30.0),
                "Larg (cm)": st.column_config.NumberColumn("Larg (cm)", min_value=1.0, step=5.0, default=30.0),
                "Comp (cm)": st.column_config.NumberColumn("Comp (cm)", min_value=1.0, step=5.0, default=40.0),
                "Peso Un. (kg)": st.column_config.NumberColumn("Peso Un. (kg)", min_value=0.1, step=0.5, default=10.0, format="%.2f"),
            },
        )

        if edited_volumes is not None:
            st.session_state.volumes_data = edited_volumes.to_dict(orient="records")

        # Cálculo automático em tempo real
        total_volumes_qtd = 0
        peso_real_total = 0.0
        volume_total_m3 = 0.0

        for _, row in edited_volumes.iterrows():
            try:
                q = int(row.get("Qtd", 1))
                a = float(row.get("Alt (cm)", 30)) / 100
                l = float(row.get("Larg (cm)", 30)) / 100
                c = float(row.get("Comp (cm)", 40)) / 100
                pu = float(row.get("Peso Un. (kg)", 10))

                total_volumes_qtd += q
                peso_real_total += (q * pu)
                volume_total_m3 += (q * a * l * c)
            except Exception:
                pass

        fator_cubagem = 300
        peso_cubado_total = volume_total_m3 * fator_cubagem
        peso_tarifado = max(peso_real_total, peso_cubado_total)
        e_cubado = peso_cubado_total > peso_real_total

        cor_tag = "#991b1b" if e_cubado else "#166534"
        bg_tag = "#fee2e2" if e_cubado else "#dcfce7"
        texto_regra = "Cobrança por Cubagem (Volumosa)" if e_cubado else "Cobrança por Peso Real"

        st.markdown(
            f"""
            <div style="background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:12px 18px;margin-top:10px;">
              <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:12px;">
                <span style="font-size:0.84rem;color:#334155;">Volumes: <b>{total_volumes_qtd}</b> · Cubagem: <b>{volume_total_m3:.4f} m³</b></span>
                <span style="font-size:0.84rem;color:#334155;">Peso Real: <b>{peso_real_total:.2f} kg</b> · Peso Cubado: <b>{peso_cubado_total:.2f} kg</b></span>
                <span style="background:{bg_tag};color:{cor_tag};padding:4px 10px;border-radius:6px;font-size:0.78rem;font-weight:700;">
                  PESO TARIFADO: {peso_tarifado:.2f} kg ({texto_regra})
                </span>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.write("")
    # Botão de Ação Primária
    st.markdown('<div class="btn-cta">', unsafe_allow_html=True)
    btn_processar = st.button("PROCESSAR COTAÇÕES", use_container_width=True, type="primary")
    st.markdown('</div>', unsafe_allow_html=True)

    # Placeholder para Skeleton Loader ou Resultados
    placeholder_resultados = st.empty()

    @st.dialog("Confirmação de Dados")
    def confirm_quote(cep_origem, cep_destino, total_volumes_qtd, volume_total_m3, peso_real_total, peso_tarifado):
        st.markdown(f"**Remetente (CEP):** {cep_origem} ➔ **Destinatário (CEP):** {cep_destino}")
        st.markdown(f"**Volumes:** {total_volumes_qtd} vol(s) | **Cubagem:** {volume_total_m3:.4f} m³")
        st.markdown(f"**Peso Real:** {peso_real_total:.2f} kg | **Peso Tarifado:** {peso_tarifado:.2f} kg")
        st.write("")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Confirmar e Cotar", type="primary", use_container_width=True):
                st.session_state.processar_confirmado = True
                st.rerun()
        with col2:
            if st.button("Voltar para Editar", use_container_width=True):
                st.rerun()

    if btn_processar:
        if not cep_origem or str(cep_origem).strip() == "":
            st.warning('Por favor, preencha o CEP de Origem.')
            st.stop()
        if not cep_destino or str(cep_destino).strip() == "":
            st.warning('Por favor, preencha o CEP de Destino.')
            st.stop()
        if valor_nf is None or valor_nf <= 0:
            st.warning('Por favor, preencha o Valor Total da Nota Fiscal.')
            st.stop()
        if peso_real_total is None or peso_real_total <= 0:
            st.warning('Por favor, preencha as Dimensões e o Peso Total da Carga.')
            st.stop()
            
        erros = []
        cep_o_num = re.sub(r"\D", "", str(cep_origem))
        cep_d_num = re.sub(r"\D", "", str(cep_destino))

        if len(cep_o_num) != 8:
            erros.append("O CEP de Origem deve conter 8 dígitos válidos.")
        elif not info_o.get("ok"):
            erros.append(f"CEP de Origem '{cep_origem}' não localizado.")

        if len(cep_d_num) != 8:
            erros.append("O CEP de Destino deve conter 8 dígitos válidos.")
        elif not info_d.get("ok"):
            erros.append(f"CEP de Destino '{cep_destino}' não localizado.")

        if total_volumes_qtd is None or total_volumes_qtd <= 0 or volume_total_m3 is None or volume_total_m3 <= 0:
            erros.append("Adicione ao menos um volume com dimensões válidas.")

        if erros:
            for erro in erros:
                st.error(f"⚠️ {erro}")
        else:
            confirm_quote(cep_origem, cep_destino, total_volumes_qtd, volume_total_m3, peso_real_total, peso_tarifado)

    if st.session_state.get("processar_confirmado"):
        st.session_state.processar_confirmado = False
        # Exibir Skeleton Loaders imediatamente (Feedback de alta performance)
        with placeholder_resultados.container():
                st.markdown(
                    """
                    <div style="margin:20px 0;">
                      <div style="color:#4f46e5;font-weight:700;font-size:0.9rem;margin-bottom:12px;display:flex;align-items:center;gap:8px;">
                        <span class="badge-triage badge-rapida">⚡ PROCESSANDO EM PARALELO</span> Consultando APIs oficiais e tabelas simultaneamente...
                      </div>
                      <div style="display:grid;grid-template-columns:repeat(3, 1fr);gap:14px;">
                        <div class="skeleton-card"></div>
                        <div class="skeleton-card"></div>
                        <div class="skeleton-card"></div>
                      </div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        cubagem_lista = []
        for _, row in edited_volumes.iterrows():
            try:
                a = float(row["Alt (cm)"]) / 100
                l = float(row["Larg (cm)"]) / 100
                c = float(row["Comp (cm)"]) / 100
                q = int(row["Qtd"])
                cubagem_lista.append({"Altura": a, "Largura": l, "Comprimento": c, "Volumes": q})
            except Exception:
                pass

        uf_dest = info_d.get("uf", "")
        cidade_dest = info_d.get("cidade", "")

        # 1. Consultas a APIs Oficiais em Paralelo
        with st.spinner('Buscando os melhores preços...'):
            futures = {}
            with ThreadPoolExecutor(max_workers=5) as executor:
                if USUARIO_BRASPRESS and SENHA_BRASPRESS:
                    futures["braspress"] = executor.submit(
                        cotar_braspress, cep_origem, cep_destino, peso_real_total, valor_nf,
                        cubagem_lista, cnpj_remetente, doc_destinatario
                    )

                if DOMINIO_COPEX and USUARIO_COPEX:
                    futures["coopex"] = executor.submit(
                        cotar_ssw_wsdl, "Coopex", DOMINIO_COPEX, USUARIO_COPEX, SENHA_COPEX,
                        cep_origem, cep_destino, peso_real_total, valor_nf, cubagem_lista,
                        cnpj_remetente, doc_destinatario, 1
                    )

            # Processamento Braspress
            if "braspress" in futures:
                try:
                    r_braspress = futures["braspress"].result(timeout=20)
                    if r_braspress.get("Valor Frete (R$)") is not None:
                        r_braspress["Atendida"] = True
                        r_braspress["Tags"] = ["BRASIL"]
                        r_braspress["Base Tarifária"] = "API Oficial REST v1 (Token Oauth2)"
                        r_braspress["Motivo"] = "Cobertura Nacional Confirmada"
                    else:
                        r_braspress["Atendida"] = False
                        r_braspress["Status"] = "Fora do ar / Indisponível"
                        r_braspress["Motivo"] = "API Oficial Braspress retornou erro ou indisponibilidade temporária"
                        r_braspress["Tags"] = ["BRASIL"]
                        r_braspress["Base Tarifária"] = "API Oficial REST v1"
                except Exception as e:
                    r_braspress = {"Transportadora": "Braspress", "Nº Cotação": "-", "Valor Frete (R$)": None, "Prazo (Dias Úteis)": 0, "Status": "Fora do ar / Indisponível", "Atendida": False, "Motivo": f"Timeout ou erro de conexão: {str(e)[:30]}", "Tags": ["BRASIL"], "Base Tarifária": "API Oficial REST v1"}
            else:
                r_braspress = {"Transportadora": "Braspress", "Nº Cotação": "-", "Valor Frete (R$)": None, "Prazo (Dias Úteis)": 0, "Status": "Fora do ar / Indisponível", "Atendida": False, "Motivo": "Credenciais da API Braspress não configuradas", "Tags": ["BRASIL"], "Base Tarifária": "API Oficial REST v1"}

            # Processamento Coopex
            if "coopex" in futures:
                try:
                    r_coopex = futures["coopex"].result(timeout=20)
                    if r_coopex.get("Valor Frete (R$)") is not None:
                        r_coopex["Atendida"] = True
                        r_coopex["Tags"] = ["PR", "SC", "RS"]
                        r_coopex["Base Tarifária"] = "Webservice SSW WSDL Oficial"
                        r_coopex["Motivo"] = "Webservice SSW Confirmado"
                    else:
                        r_coopex = te.calcular_frete_estrito("COOPEX", peso_tarifado, peso_real_total, valor_nf, uf_dest, cidade_dest, cep_destino)
                except Exception:
                    r_coopex = te.calcular_frete_estrito("COOPEX", peso_tarifado, peso_real_total, valor_nf, uf_dest, cidade_dest, cep_destino)
            else:
                r_coopex = te.calcular_frete_estrito("COOPEX", peso_tarifado, peso_real_total, valor_nf, uf_dest, cidade_dest, cep_destino)

        # 2. Avaliação Rigorosa das 14 Demais Tabelas Contratuais Homologadas
        chaves_tabelas = [
            "PRINCESA", "ALFA", "TW", "ENVIA_RAPIDO", "GARCIA", "SUDOESTE",
            "CARRION", "EXPRESSO_SAO_MIGUEL", "LOGDI", "OURO_NEGRO",
            "RODONAVES", "TECMAR", "AGEX", "VIP"
        ]

        resultados_tabelas = [
            te.calcular_frete_estrito(k, peso_tarifado, peso_real_total, valor_nf, uf_dest, cidade_dest, cep_destino)
            for k in chaves_tabelas
        ]

        resultados = [r_braspress, r_coopex] + resultados_tabelas

        # Salvar na sessão
        st.session_state.ultima_cotacao = {
            "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            "remetente": cnpj_remetente,
            "destinatario": doc_destinatario,
            "origem_cep": cep_origem,
            "origem_texto": info_o.get("texto", ""),
            "destino_cep": cep_destino,
            "destino_texto": info_d.get("texto", ""),
            "uf_dest": uf_dest,
            "cidade_dest": cidade_dest,
            "peso": peso_real_total,
            "volume_m3": volume_total_m3,
            "peso_cubado": peso_cubado_total,
            "peso_tarifado": peso_tarifado,
            "valor_nf": valor_nf,
            "volumes_qtd": total_volumes_qtd,
            "volumes_data": st.session_state.get("volumes_data", []),
            "modalidade": st.session_state.get("tipo_frete_input", "FOB (Pago pelo Destinatário)"),
            "resultados": resultados,
            "obs": obs_carga,
        }

        # Histórico com a melhor oferta válida
        opcoes_validas_hist = [r for r in resultados if r.get("Atendida") and r.get("Valor Frete (R$)") is not None]
        if opcoes_validas_hist:
            df_temp_v = pd.DataFrame(opcoes_validas_hist).sort_values("Valor Frete (R$)").reset_index(drop=True)
            melhor_item = df_temp_v.iloc[0]
            salvar_historico_item({
                "id": f"COT-{datetime.now().strftime('%Y%m%d-%H%M%S')}",
                "data_hora": datetime.now().strftime("%d/%m/%Y %H:%M"),
                "origem": f"{info_o.get('cidade','')}/{info_o.get('uf','')}".strip("/"),
                "destino": f"{cidade_dest}/{uf_dest}".strip("/"),
                "cep_origem": cep_origem,
                "cep_destino": cep_destino,
                "peso_kg": peso_real_total,
                "peso_tarifado_kg": peso_tarifado,
                "valor_nf": valor_nf,
                "melhor_transportadora": melhor_item["Transportadora"],
                "melhor_valor": melhor_item["Valor Frete (R$)"],
                "melhor_prazo": int(melhor_item.get("Prazo (Dias Úteis)") or 0),
                "melhor_tipo": melhor_item.get("Base Tarifária", "Tabela Contratual"),
            })

        placeholder_resultados.empty()

    # ================= APRESENTAÇÃO DOS RESULTADOS (ESTRITA E SEGREGADA) =================
    if "ultima_cotacao" in st.session_state and st.session_state.ultima_cotacao:
        cot_info = st.session_state.ultima_cotacao
        resultados = cot_info.get("resultados", [])

        opcoes_atendidas = [r for r in resultados if r.get("Atendida") and r.get("Valor Frete (R$)") is not None]
        opcoes_indisponiveis = [r for r in resultados if not r.get("Atendida") or r.get("Valor Frete (R$)") is None]

        df_valid = pd.DataFrame(opcoes_atendidas)
        if not df_valid.empty:
            df_valid = df_valid.sort_values("Valor Frete (R$)").reset_index(drop=True)

        st.markdown("---")

        # Cabeçalho da Seção de Resultados com Ações Rápidas
        col_res_header, col_res_excel, col_res_wa = st.columns([2.2, 1, 1.3])
        with col_res_header:
            st.markdown(f"#### Comparativo de Ofertas — Destino: {cot_info.get('cidade_dest','')}/{cot_info.get('uf_dest','')}")
            st.caption(f"CEP {cot_info.get('destino_cep')} · Carga: {cot_info.get('peso',0):.2f} kg ({cot_info.get('volumes_qtd',1)} vol) · NF: {formatar_moeda(cot_info.get('valor_nf',0))}")
        with col_res_excel:
            df_all = pd.DataFrame(resultados)
            df_all["_ordem"] = df_all["Valor Frete (R$)"].isna().astype(int)
            df_all = df_all.sort_values(["_ordem", "Valor Frete (R$)"]).drop(columns=["_ordem"]).reset_index(drop=True)
            excel_bytes = gerar_excel_cotacao(cot_info, df_all)
            st.download_button(
                label="Baixar em Excel (.xlsx)",
                data=excel_bytes,
                file_name=f"cotacao_{cot_info.get('destino_cep')}_{datetime.now():%Y%m%d_%H%M}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with col_res_wa:
            if len(df_valid) > 0:
                v_data = cot_info.get('volumes_data', [])
                if v_data:
                    dim_list = [f"{v.get('Comp (cm)', 0)}x{v.get('Larg (cm)', 0)}x{v.get('Alt (cm)', 0)}" for v in v_data]
                    dimensoes_resumidas = "\n".join(dim_list)
                else:
                    dimensoes_resumidas = "N/A"
                
                melhor_op = df_valid.iloc[0]
                
                obs_text = f"\n*Observação:* {cot_info.get('obs')}" if cot_info.get('obs') else ""
                msg_completa = (
                    f"Olá!\n"
                    f"Por gentileza, peço a validação desta cotação de frete:\n\n"
                    f"CNPJ remetente: {cot_info.get('remetente', 'N/A')}\n"
                    f"CEP remetente: {cot_info.get('origem_cep', 'N/A')}\n"
                    f"CNPJ destinatário: {cot_info.get('destinatario', 'N/A')}\n"
                    f"CEP destinatário: {cot_info.get('destino_cep', 'N/A')}\n"
                    f"Volumes: {cot_info.get('volumes_qtd', 0)}\n"
                    f"Medidas:\n"
                    f"{dimensoes_resumidas}\n"
                    f"Peso: {cot_info.get('peso', 0):.2f}kg\n"
                    f"Valor da NF: R${cot_info.get('valor_nf', 0):,.2f}{obs_text}\n\n"
                    f"*Estimativa via Tabela :* *{formatar_moeda(melhor_op['Valor Frete (R$)'])}*\n\n"
                    f"Pode confirmar se o valor bate com a tabela que temos com vocês?"
                )
                numero_wa = obter_numero_wa(melhor_op['Transportadora'])
                link_wa_geral = f"https://wa.me/{numero_wa}?text={urllib.parse.quote(msg_completa)}"
                st.link_button("Gerar Proposta p/ WhatsApp", link_wa_geral, use_container_width=True)

        if len(df_valid) == 0:
            st.error("Nenhuma transportadora atende este destino nas condições contratuais informadas.")
        else:
            # ================= TOP 3 RANKING (OPÇÕES DISTINTAS POR MENOR PREÇO) =================
            # Filtra para obter até 3 transportadoras distintas ordenadas estritamente pelo Menor Preço
            top_candidatos = []
            transp_vistas = set()
            for _, r_cand in df_valid.iterrows():
                nome_t = r_cand["Transportadora"]
                if nome_t not in transp_vistas:
                    transp_vistas.add(nome_t)
                    top_candidatos.append(r_cand)
                if len(top_candidatos) == 3:
                    break

            # Identificar o menor prazo global de entrega entre todas as opções válidas
            df_com_prazo = df_valid[df_valid["Prazo (Dias Úteis)"] > 0]
            min_prazo_global = int(df_com_prazo["Prazo (Dias Úteis)"].min()) if not df_com_prazo.empty else None

            prazo_op1 = int(top_candidatos[0]["Prazo (Dias Úteis)"]) if (
                len(top_candidatos) > 0 and pd.notna(top_candidatos[0]["Prazo (Dias Úteis)"]) and int(top_candidatos[0]["Prazo (Dias Úteis)"]) > 0
            ) else 0

            st.markdown(
                '<div style="font-size:0.84rem;font-weight:800;color:#0f172a;letter-spacing:0.04em;text-transform:uppercase;margin:12px 0 8px 0;">'
                'TOP 3 OPÇÕES (RANKING POR MENOR PREÇO)'
                '</div>',
                unsafe_allow_html=True
            )

            cols_ranking = st.columns(len(top_candidatos))

            for idx, r_item in enumerate(top_candidatos):
                pos = idx + 1
                col_atual = cols_ranking[idx]

                with col_atual:
                    prazo_val = int(r_item["Prazo (Dias Úteis)"]) if pd.notna(r_item["Prazo (Dias Úteis)"]) and int(r_item["Prazo (Dias Úteis)"]) > 0 else 0
                    p_txt = f"{prazo_val} dias úteis" if prazo_val > 0 else "A confirmar"
                    
                    # Verificação de Smart Tags (Regras de Negócio de Desempate e Agilidade)
                    is_mais_rapida = (min_prazo_global is not None and prazo_val == min_prazo_global)
                    entrega_antes_do_primeiro = (pos > 1 and prazo_op1 > 0 and 0 < prazo_val < prazo_op1)

                    # Estilização por posição
                    if pos == 1:
                        card_class = "card-ranking-pos1"
                        badge_pos_html = '<span class="badge-ranking badge-pos1">🏆 1º LUGAR · MELHOR PREÇO</span>'
                        preco_cor = "#047857"
                        if len(top_candidatos) > 1:
                            dif_econ = float(top_candidatos[1]["Valor Frete (R$)"]) - float(r_item["Valor Frete (R$)"])
                            dica_html = f'<div style="font-size:0.72rem;color:#059669;font-weight:700;margin-top:6px;padding-top:4px;border-top:1px dashed #bbf7d0;">Economia de {formatar_moeda(dif_econ)} vs. 2ª opção</div>'
                        else:
                            dica_html = ''
                    elif pos == 2:
                        card_class = "card-ranking-pos2"
                        badge_pos_html = '<span class="badge-ranking badge-pos2">🥈 2ª OPÇÃO</span>'
                        preco_cor = "#0284c7"
                        if entrega_antes_do_primeiro:
                            dias_eco = prazo_op1 - prazo_val
                            dica_html = f'<div style="font-size:0.72rem;color:#b45309;font-weight:700;margin-top:6px;padding-top:4px;border-top:1px dashed #fed7aa;">⚡ Entrega {dias_eco} dia(s) antes da 1ª opção</div>'
                        else:
                            dica_html = ''
                    else:
                        card_class = "card-ranking-pos3"
                        badge_pos_html = '<span class="badge-ranking badge-pos3">🥉 3ª OPÇÃO</span>'
                        preco_cor = "#1e293b"
                        if entrega_antes_do_primeiro:
                            dias_eco = prazo_op1 - prazo_val
                            dica_html = f'<div style="font-size:0.72rem;color:#b45309;font-weight:700;margin-top:6px;padding-top:4px;border-top:1px dashed #fed7aa;">⚡ Entrega {dias_eco} dia(s) antes da 1ª opção</div>'
                        else:
                            dica_html = ''

                    # Smart Tag "Mais Rápida"
                    smart_tag_html = ""
                    if is_mais_rapida:
                        smart_tag_html = f'<span class="badge-ranking badge-fast">⚡ MAIS RÁPIDA ({prazo_val}D)</span>'
                    elif entrega_antes_do_primeiro:
                        smart_tag_html = '<span class="badge-ranking badge-fast">⚡ MAIS RÁPIDA</span>'

                    # Logo oficial da transportadora (SVG ou Fallback)
                    logo_html = gerar_badge_logo_html(r_item["Transportadora"])

                    # Tags de abrangência removidas conforme Ticket 5
                    tags_html = ""

                    # Link WhatsApp da opção
                    v_data = cot_info.get('volumes_data', [])
                    if v_data:
                        dim_list = [f"{v.get('Comp (cm)', 0)}x{v.get('Larg (cm)', 0)}x{v.get('Alt (cm)', 0)}" for v in v_data]
                        dimensoes_resumidas = "\n".join(dim_list)
                    else:
                        dimensoes_resumidas = "N/A"

                    obs_text = f"\n*Observação:* {cot_info.get('obs')}" if cot_info.get('obs') else ""
                    msg_card = (
                        f"Olá!\n"
                        f"Por gentileza, peço a validação desta cotação de frete:\n\n"
                        f"CNPJ remetente: {cot_info.get('remetente', 'N/A')}\n"
                        f"CEP remetente: {cot_info.get('origem_cep', 'N/A')}\n"
                        f"CNPJ destinatário: {cot_info.get('destinatario', 'N/A')}\n"
                        f"CEP destinatário: {cot_info.get('destino_cep', 'N/A')}\n"
                        f"Volumes: {cot_info.get('volumes_qtd', 0)}\n"
                        f"Medidas:\n"
                        f"{dimensoes_resumidas}\n"
                        f"Peso: {cot_info.get('peso', 0):.2f}kg\n"
                        f"Valor da NF: R${cot_info.get('valor_nf', 0):,.2f}{obs_text}\n\n"
                        f"*Estimativa via Tabela :* *{formatar_moeda(r_item['Valor Frete (R$)'])}*\n\n"
                        f"Pode confirmar se o valor bate com a tabela que temos com vocês?"
                    )
                    numero_wa = obter_numero_wa(r_item['Transportadora'])
                    link_wa_card = f"https://wa.me/{numero_wa}?text={urllib.parse.quote(msg_card)}"

                    card_html = f"""
                    <div class="{card_class}" style="border: 1px solid #ddd; border-radius: 8px; padding: 15px; text-align: center; box-shadow: 2px 2px 5px rgba(0,0,0,0.1); margin-bottom: 10px; background-color: #fff;">
                      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;gap:6px;flex-wrap:wrap;">
                        {badge_pos_html}
                        {smart_tag_html}
                      </div>
                      <div style="margin-bottom:8px;">
                        {logo_html}
                      </div>
                      <div style="margin:8px 0 6px 0;">
                        <div style="font-size:0.70rem;font-weight:700;color:#64748b;text-transform:uppercase;letter-spacing:0.04em;">Valor Total do Frete</div>
                        <div style="font-size:1.80rem;font-weight:900;color:{preco_cor};line-height:1.1;">
                          {formatar_moeda(r_item['Valor Frete (R$)'])}
                        </div>
                      </div>
                      <div style="background:#ffffff;border:1px solid #e2e8f0;border-radius:8px;padding:8px 10px;margin-bottom:10px;text-align:left;">
                        <div style="display:flex;justify-content:space-between;align-items:center;">
                          <span style="font-size:0.78rem;color:#475569;font-weight:600;">Prazo Previsto:</span>
                          <span style="font-size:0.85rem;color:#0f172a;font-weight:800;">{p_txt}</span>
                        </div>
                        <div style="display:flex;justify-content:space-between;align-items:center;margin-top:4px;">
                          <span style="font-size:0.78rem;color:#475569;font-weight:600;">Peso Tarifado:</span>
                          <span style="font-size:0.85rem;color:#0f172a;font-weight:800;">{r_item.get('Peso Tarifado', f"{cot_info.get('peso_tarifado', 0):.2f} kg")}</span>
                        </div>
                        {dica_html}
                      </div>
                      <a href="{link_wa_card}" target="_blank" style="display:flex;align-items:center;justify-content:center;gap:6px;background:linear-gradient(135deg, #25D366, #128C7E);color:#ffffff;font-weight:700;font-size:0.80rem;padding:7px 12px;border-radius:8px;text-decoration:none;box-shadow:0 2px 6px rgba(37,211,102,0.25);">
                        Proposta WhatsApp
                      </a>
                    </div>
                    """
                    st.markdown(card_html.replace('\n', ''), unsafe_allow_html=True)


            st.write("")
            # ================= TABELA DE OPÇÕES VÁLIDAS =================
            st.markdown("##### Opções de Transportadoras")
            df_t = df_valid.copy()
            df_t["Prazo"] = df_t["Prazo (Dias Úteis)"].apply(
                lambda x: f"{int(x)} dias" if pd.notna(x) and int(x) > 0 else "A confirmar"
            )
            df_t["Abrangência"] = df_t["Tags"].apply(lambda tags: " ".join([f"[{t}]" for t in tags]) if isinstance(tags, list) else "")

            def gerar_link_wa(row):
                valor = formatar_moeda(row["Valor Frete (R$)"])
                prazo_val = int(row["Prazo (Dias Úteis)"]) if pd.notna(row["Prazo (Dias Úteis)"]) else 0
                prazo_str = f"{prazo_val} dias úteis" if prazo_val > 0 else "A confirmar"
                tipo_str = row.get("Base Tarifária", "Tabela Contratual")

                v_data = cot_info.get('volumes_data', [])
                if v_data:
                    dim_list = [f"{v.get('Comp (cm)', 0)}x{v.get('Larg (cm)', 0)}x{v.get('Alt (cm)', 0)}" for v in v_data]
                    dimensoes_resumidas = "\n".join(dim_list)
                else:
                    dimensoes_resumidas = "N/A"

                msg = (
                    f"Olá!\n"
                    f"Por gentileza, peço a validação desta cotação de frete:\n\n"
                    f"CNPJ remetente: {cot_info.get('remetente', 'N/A')}\n"
                    f"CEP remetente: {cot_info.get('origem_cep', 'N/A')}\n"
                    f"CNPJ destinatário: {cot_info.get('destinatario', 'N/A')}\n"
                    f"CEP destinatário: {cot_info.get('destino_cep', 'N/A')}\n"
                    f"Volumes: {cot_info.get('volumes_qtd', 0)}\n"
                    f"Medidas:\n"
                    f"{dimensoes_resumidas}\n"
                    f"Peso: {cot_info.get('peso', 0):.2f}kg\n"
                    f"Valor da NF: R${cot_info.get('valor_nf', 0):,.2f}\n\n"
                    f"*Estimativa via Tabela :* *{valor}*\n\n"
                    f"Pode confirmar se o valor bate com a tabela que temos com vocês?"
                )
                numero_wa = obter_numero_wa(row['Transportadora'])
                return f"https://wa.me/{numero_wa}?text={urllib.parse.quote(msg)}"

            df_t["Ação"] = df_t.apply(gerar_link_wa, axis=1)
            st.dataframe(
                df_t[["Transportadora", "Valor Frete (R$)", "Prazo", "Ação"]],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Transportadora": st.column_config.TextColumn("Transportadora", width="medium"),
                    "Valor Frete (R$)": st.column_config.NumberColumn("Valor Frete", format="R$ %.2f", width="small"),
                    "Prazo": st.column_config.TextColumn("Prazo", width="small"),
                    "Ação": st.column_config.LinkColumn("WhatsApp", display_text="Enviar Proposta", width="small")
                },
            )

        # Seção de Transportadoras Não Atendidas / Indisponíveis (Tolerância Zero a Falsos Positivos)
        if opcoes_indisponiveis:
            st.write("")
            with st.expander(f"Transportadoras Sem Atendimento ou Fora do Ar ({len(opcoes_indisponiveis)})", expanded=False):
                st.caption(
                    "As opções abaixo foram excluídas do comparativo comercial porque a rota não consta nas tabelas "
                    "contratuais vigentes da Next Cable ou porque a API não retornou resposta operacional."
                )
                for item_un in opcoes_indisponiveis:
                    nome_t = item_un.get("Transportadora", "")
                    tags_t = item_un.get("Tags", [])
                    tags_html = " ".join([
                        f"<span style='background:#f1f5f9;color:#64748b;padding:1px 5px;border-radius:4px;font-size:0.7rem;font-weight:700;border:1px solid #e2e8f0;'>[{t}]</span>"
                        for t in tags_t
                    ])
                    motivo_t = item_un.get("Motivo", item_un.get("Status", "Não atende esta localidade"))
                    eh_api = "api" in str(item_un.get("Base Tarifária", "")).lower() or "conexão" in motivo_t.lower() or "credenciais" in motivo_t.lower() or "fora do ar" in str(item_un.get("Status", "")).lower()
                    status_badge = "⚠️ Fora do ar / Indisponível" if eh_api else "🚫 Não atende esta região"
                    badge_style = "background:#fef3c7;color:#92400e;border:1px solid #fde68a;" if eh_api else "background:#f1f5f9;color:#475569;border:1px solid #cbd5e1;"

                    st.markdown(
                        f"""
                        <div style="background:#ffffff;border:1px solid #e2e8f0;border-left:4px solid #94a3b8;border-radius:6px;padding:8px 12px;margin-bottom:8px;display:flex;justify-content:space-between;align-items:center;">
                          <div>
                            <div style="font-weight:700;color:#334155;font-size:0.88rem;">
                              {nome_t} <span style="margin-left:8px;">{tags_html}</span>
                            </div>
                            <div style="color:#64748b;font-size:0.78rem;margin-top:3px;">
                              {motivo_t}
                            </div>
                          </div>
                          <div>
                            <span style="{badge_style}padding:3px 10px;border-radius:12px;font-size:0.72rem;font-weight:700;white-space:nowrap;">
                              {status_badge}
                            </span>
                          </div>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

# ==============================================================================
# MÓDULO 2: RASTREAMENTO DE PEDIDOS & ENTREGAS (entregas.xlsx)
# ==============================================================================
elif aba == "Rastreamento de Pedidos":
    st.markdown("#### Gestão e Rastreamento de Entregas")

    sub_aba = st.radio("Modo de Rastreamento", ["Painel de Entregas em Lote (entregas.xlsx)", "Rastreamento Individual"], horizontal=True, label_visibility="collapsed")

    if sub_aba == "Painel de Entregas em Lote (entregas.xlsx)":
        st.markdown(
            "Carregue, gerencie e atualize o status de todas as cargas expedidas pela **Next Cable**. "
            "O sistema consulta as transportadoras em paralelo e sincroniza automaticamente a planilha."
        )

        if "df_entregas" not in st.session_state:
            if os.path.exists(ENTREGAS_FILE):
                try:
                    st.session_state.df_entregas = pd.read_excel(ENTREGAS_FILE)
                except Exception:
                    st.session_state.df_entregas = pd.DataFrame(columns=[
                        "CNPJ", "Nota Fiscal", "Status", "Previsão de Entrega", "Data de Entrega",
                        "Última Ocorrência", "Data Ocorrência", "Destino", "Valor Frete"
                    ])
            else:
                st.session_state.df_entregas = pd.DataFrame(columns=[
                    "CNPJ", "Nota Fiscal", "Status", "Previsão de Entrega", "Data de Entrega",
                    "Última Ocorrência", "Data Ocorrência", "Destino", "Valor Frete"
                ])

        if st.session_state.df_entregas.empty:
            st.session_state.df_entregas = pd.DataFrame([
                {"CNPJ": "26434839000121", "Nota Fiscal": "10025", "Status": "Aguardando Consulta", "Previsão de Entrega": "", "Data de Entrega": "", "Última Ocorrência": "", "Data Ocorrência": "", "Destino": "São Paulo/SP", "Valor Frete": ""},
                {"CNPJ": "26434839000121", "Nota Fiscal": "10026", "Status": "Aguardando Consulta", "Previsão de Entrega": "", "Data de Entrega": "", "Última Ocorrência": "", "Data Ocorrência": "", "Destino": "Curitiba/PR", "Valor Frete": ""},
            ])

        col_up1, col_up2 = st.columns([2.5, 1])
        with col_up2:
            arq_up = st.file_uploader("Subir outra planilha de entregas", type=["xlsx", "xls", "csv"], key="up_entregas")
            if arq_up is not None:
                try:
                    df_novo = pd.read_csv(arq_up) if arq_up.name.endswith(".csv") else pd.read_excel(arq_up)
                    st.session_state.df_entregas = df_novo
                    st.success("Planilha carregada!")
                except Exception as e:
                    st.error(f"Erro: {e}")

        df_atual = st.session_state.df_entregas
        if not df_atual.empty and "Status" in df_atual.columns:
            total_nfs = len(df_atual)
            status_series = df_atual["Status"].astype(str).str.lower()
            n_entregue = status_series.str.contains("entregue").sum()
            n_transito = (status_series.str.contains("trânsito") | status_series.str.contains("transito") | status_series.str.contains("saiu")).sum()
            n_ocorr = (status_series.str.contains("ocorrência") | status_series.str.contains("retido") | status_series.str.contains("erro")).sum()

            k1, k2, k3, k4 = st.columns(4)
            with k1:
                st.metric("Total de Entregas", total_nfs)
            with k2:
                st.metric("Entregues", n_entregue)
            with k3:
                st.metric("Em Trânsito", n_transito)
            with k4:
                st.metric("Ocorrências / Atrasos", n_ocorr)

        df_entregas_editado = st.data_editor(
            st.session_state.df_entregas,
            num_rows="dynamic",
            width="stretch",
            hide_index=True,
            column_config={
                "CNPJ": st.column_config.TextColumn("CNPJ Remetente"),
                "Nota Fiscal": st.column_config.TextColumn("Nº NF", required=True),
                "Status": st.column_config.TextColumn("Status"),
                "Previsão de Entrega": st.column_config.TextColumn("Previsão"),
                "Data de Entrega": st.column_config.TextColumn("Data Entrega"),
                "Última Ocorrência": st.column_config.TextColumn("Última Ocorrência"),
                "Data Ocorrência": st.column_config.TextColumn("Data Ocorrência"),
                "Destino": st.column_config.TextColumn("Destino"),
                "Valor Frete": st.column_config.TextColumn("Valor Frete"),
            }
        )
        st.session_state.df_entregas = df_entregas_editado

        st.write("")
        b1, b2, b3 = st.columns([1.5, 1, 1])
        with b1:
            btn_atualizar_lote = st.button("Atualizar Status das Entregas", use_container_width=True, type="primary")
        with b2:
            if st.button("Salvar entregas.xlsx", use_container_width=True):
                try:
                    df_entregas_editado.to_excel(ENTREGAS_FILE, index=False, engine="openpyxl")
                    st.success(f"Arquivo '{ENTREGAS_FILE}' salvo com sucesso!")
                except Exception as e:
                    st.error(f"Erro ao salvar arquivo: {e}")
        with b3:
            excel_entregas_bytes = gerar_excel_entregas(df_entregas_editado)
            st.download_button(
                "Exportar Planilha Atualizada",
                data=excel_entregas_bytes,
                file_name=f"entregas_atualizadas_{datetime.now():%Y%m%d_%H%M}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        if btn_atualizar_lote:
            if df_entregas_editado.empty:
                st.warning("Nenhuma nota fiscal para atualizar.")
            else:
                with st.spinner("Consultando status de todas as remessas em paralelo..."):
                    df_proc = df_entregas_editado.copy()

                    def rastrear_linha(idx, r):
                        nf = str(r.get("Nota Fiscal", "")).strip()
                        cnpj = str(r.get("CNPJ", "")).strip() or "26434839000121"
                        if not nf:
                            return idx, None
                        res_track = rastrear_braspress(nf, cnpj)
                        return idx, res_track

                    with ThreadPoolExecutor(max_workers=8) as executor:
                        futs = [executor.submit(rastrear_linha, i, r) for i, r in df_proc.iterrows()]
                        for f in as_completed(futs):
                            i_idx, res = f.result()
                            if res:
                                df_proc.at[i_idx, "Status"] = res["status"]
                                if res["previsao"]:
                                    df_proc.at[i_idx, "Previsão de Entrega"] = res["previsao"]
                                if res["data_entrega"]:
                                    df_proc.at[i_idx, "Data de Entrega"] = res["data_entrega"]
                                if res["ultima_ocorrencia"]:
                                    df_proc.at[i_idx, "Última Ocorrência"] = res["ultima_ocorrencia"]
                                if res["data_ocorrencia"]:
                                    df_proc.at[i_idx, "Data Ocorrência"] = res["data_ocorrencia"]
                                if res["destino"] and not df_proc.at[i_idx, "Destino"]:
                                    df_proc.at[i_idx, "Destino"] = res["destino"]
                                if res["valor_frete"]:
                                    df_proc.at[i_idx, "Valor Frete"] = formatar_moeda(res["valor_frete"])

                    st.session_state.df_entregas = df_proc
                    try:
                        df_proc.to_excel(ENTREGAS_FILE, index=False, engine="openpyxl")
                    except Exception:
                        pass
                    st.success("Atualização em lote concluída e salva no arquivo!")
                    st.rerun()

    elif sub_aba == "Rastreamento Individual":
        st.markdown("Consulta individual detalhada de carga com histórico de ocorrências.")
        c_ri1, c_ri2 = st.columns([1, 2], gap="medium")
        with c_ri1:
            with st.container(border=True):
                nf_ind = st.text_input("Número da Nota Fiscal (NF)")
                cnpj_ind = st.text_input("CNPJ Remetente", value="26434839000121")
                btn_rastrear_ind = st.button("Rastrear NF Agora", use_container_width=True, type="primary")

        with c_ri2:
            if btn_rastrear_ind and nf_ind:
                with st.spinner("Consultando sistema de tracking..."):
                    res_ind = rastrear_braspress(nf_ind, cnpj_ind)
                    with st.container(border=True):
                        if res_ind["sucesso"]:
                            st.success(f"Status Atual da NF {nf_ind}: {res_ind['status']}")
                            col_info1, col_info2 = st.columns(2)
                            with col_info1:
                                if res_ind["previsao"]:
                                    st.write(f"**Previsão de Entrega:** {res_ind['previsao']}")
                                if res_ind["data_entrega"]:
                                    st.write(f"**Entregue em:** {res_ind['data_entrega']}")
                                if res_ind["destino"]:
                                    st.write(f"**Destino:** {res_ind['destino']}")
                            with col_info2:
                                if res_ind["ultima_ocorrencia"]:
                                    st.write(f"**Última Ocorrência:** {res_ind['ultima_ocorrencia']}")
                                if res_ind["data_ocorrencia"]:
                                    st.write(f"**Data do Evento:** {res_ind['data_ocorrencia']}")

                            st.markdown("---")
                            st.markdown("**Histórico de Ocorrências:**")
                            eventos = res_ind.get("eventos", [])
                            if eventos:
                                for e in reversed(eventos):
                                    d_txt = e.get("data", e.get("dataHora", "Data N/D"))
                                    desc_txt = e.get("descricao", e.get("status", ""))
                                    st.markdown(
                                        f'<div style="border-left:3px solid #10b981;padding-left:12px;margin-bottom:10px;">'
                                        f'<div style="font-weight:700;color:#1e3a8a;font-size:0.85rem;">{d_txt}</div>'
                                        f'<div style="color:#334155;font-size:0.82rem;">{desc_txt}</div>'
                                        f'</div>',
                                        unsafe_allow_html=True,
                                    )
                            else:
                                st.info("Sem eventos detalhados cadastrados até o momento.")
                        else:
                            st.error(f"{res_ind['status']}")
                            st.caption("Verifique o número da nota fiscal e o CNPJ do remetente.")


# ==============================================================================
# MÓDULO 4: HISTÓRICO DE COTAÇÕES
# ==============================================================================
elif aba == "Histórico de Cotações (Em Breve)":
    st.markdown("#### Histórico de Cotações Realizadas (Em Breve)")

    historico = carregar_historico()

    if not historico:
        st.info("Nenhuma cotação registrada nesta base de dados ainda. Faça uma cotação na aba 'Nova Cotação'.")
    else:
        df_hist = pd.DataFrame(historico)

        m1, m2, m3, m4 = st.columns(4)
        with m1:
            st.metric("Total de Cotações", len(df_hist))
        with m2:
            st.metric("Destinos Únicos", df_hist["destino"].nunique())
        with m3:
            mais_frequente = df_hist["melhor_transportadora"].mode().iloc[0] if not df_hist.empty else "-"
            st.metric("Mais Competitiva", mais_frequente)
        with m4:
            media_peso = df_hist["peso_kg"].mean() if not df_hist.empty else 0
            st.metric("Média de Peso", f"{media_peso:.1f} kg")

        st.write("")
        col_acoes1, col_acoes2, col_acoes3 = st.columns([2, 1, 1])
        with col_acoes2:
            buf_h = io.BytesIO()
            df_hist.to_excel(buf_h, index=False, engine="openpyxl")
            st.download_button(
                "Exportar Histórico (.xlsx)",
                data=buf_h.getvalue(),
                file_name=f"historico_cotacoes_{datetime.now():%Y%m%d}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with col_acoes3:
            if st.button("Limpar Histórico", use_container_width=True):
                limpar_historico_arquivo()
                st.rerun()

        df_exibir = df_hist.copy()
        df_exibir["Valor"] = df_exibir["melhor_valor"].apply(formatar_moeda)
        df_exibir["Prazo"] = df_exibir["melhor_prazo"].apply(lambda x: f"{int(x)} dias" if x > 0 else "—")
        df_exibir["Peso Real"] = df_exibir["peso_kg"].apply(lambda x: f"{x:.2f} kg")
        df_exibir["Peso Tarifado"] = df_exibir["peso_tarifado_kg"].apply(lambda x: f"{x:.2f} kg")
        df_exibir["NF"] = df_exibir["valor_nf"].apply(formatar_moeda)

        cols_mostrar = ["data_hora", "origem", "destino", "melhor_transportadora", "Valor", "Prazo", "Peso Real", "Peso Tarifado", "NF"]
        cols_existentes = [c for c in cols_mostrar if c in df_exibir.columns]

        st.dataframe(
            df_exibir[cols_existentes].rename(columns={
                "data_hora": "Data/Hora",
                "origem": "Origem",
                "destino": "Destino",
                "melhor_transportadora": "Melhor Oferta",
            }),
            width="stretch",
            hide_index=True,
        )
        