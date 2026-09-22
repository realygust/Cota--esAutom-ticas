import ast

with open('main.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

with open('main.py', 'r', encoding='utf-8') as f:
    code = f.read()

tree = ast.parse(code)

funcs_to_remove = []
for node in tree.body:
    if isinstance(node, ast.FunctionDef):
        funcs_to_remove.append((node.lineno, node.end_lineno))

# Create a boolean array to mark lines for deletion
to_delete = [False] * (len(lines) + 1)
for start, end in funcs_to_remove:
    for i in range(start, end + 1):
        to_delete[i] = True

new_lines = []
for i, line in enumerate(lines, start=1):
    if not to_delete[i]:
        new_lines.append(line)

new_code = "".join(new_lines)

# Remove the CONTATOS_TRANSPORTADORAS dictionary as well, if we can find it
import re
new_code = re.sub(r'CONTATOS_TRANSPORTADORAS\s*=\s*\{.*?\n\}', '', new_code, flags=re.DOTALL)
new_code = re.sub(r'TABELAS_INFO\s*=\s*\{.*?\n\}', '', new_code, flags=re.DOTALL)
new_code = re.sub(r'COBERTURA\s*=\s*\{.*?\n\}', '', new_code, flags=re.DOTALL)
new_code = re.sub(r'TABELAS\s*=\s*\{.*?\}\n\}', '', new_code, flags=re.DOTALL) # Might be tricky, let's keep it in config.py instead

imports = """
# ================= IMPORTAÇÕES REFATORADAS =================
import utils
import api_services
import excel_generator
import whatsapp_generator
from config import *
import html
# ============================================================
"""

# add aliases to keep code working without prefixing everything
aliases = """
_secret = utils._secret
carregar_historico = utils.carregar_historico
salvar_historico_item = utils.salvar_historico_item
limpar_historico_arquivo = utils.limpar_historico_arquivo
formatar_moeda = utils.formatar_moeda
limpar_mensagem_ssw = utils.limpar_mensagem_ssw
limpar_documento = utils.limpar_documento
limpar_cep = utils.limpar_cep
xml_escape = utils.xml_escape
buscar_endereco_cep = utils.buscar_endereco_cep
buscar_empresa_cnpj = utils.buscar_empresa_cnpj
mapear_regiao = utils.mapear_regiao

calcular_frete_tabela = api_services.calcular_frete_tabela
cotar_braspress = api_services.cotar_braspress
rastrear_braspress = api_services.rastrear_braspress
cotar_ssw_wsdl = api_services.cotar_ssw_wsdl

gerar_excel_cotacao = excel_generator.gerar_excel_cotacao
gerar_excel_entregas = excel_generator.gerar_excel_entregas

obter_numero_wa = whatsapp_generator.obter_numero_wa
gerar_badge_logo_html = whatsapp_generator.gerar_badge_logo_html

# Alias the dictionary constants in case they are used in main
CONTATOS_TRANSPORTADORAS = whatsapp_generator.CONTATOS_TRANSPORTADORAS
"""

final_code = imports + aliases + new_code

with open('main_refatorado.py', 'w', encoding='utf-8') as f:
    f.write(final_code)

print("main_refatorado.py criado!")
