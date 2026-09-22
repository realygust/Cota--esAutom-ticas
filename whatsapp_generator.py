import urllib.parse
import streamlit as st
import tabelas_engine as te

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

