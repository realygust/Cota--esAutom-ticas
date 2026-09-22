import db_historico as db
import uuid
import json
from datetime import datetime

db.init_db()
db.execute_query('DELETE FROM usuarios WHERE email IN ("admin@test.com", "venda@test.com", "vendb@test.com")')

admin_id = db.create_user('admin@test.com', 'Admin Test', 'admin123', 'ADMIN')
vend_a_id = db.create_user('venda@test.com', 'Vendedor A', 'venda123', 'VENDEDOR')
vend_b_id = db.create_user('vendb@test.com', 'Vendedor B', 'vendb123', 'VENDEDOR')

cotacao_id = str(uuid.uuid4())
dados = {'id': cotacao_id, 'cliente': 'Cliente A', 'status': 'CRIADA'}
db.salvar_cotacao(dados, vend_a_id)
print('Criada cotacao por Vendedor A.')

b_fetch = db.obter_cotacao_por_id(cotacao_id, vend_b_id, 'VENDEDOR')
print('B consegue ver a cotacao de A?', b_fetch is not None)

a_fetch = db.obter_cotacao_por_id(cotacao_id, vend_a_id, 'VENDEDOR')
print('A consegue ver a propria cotacao?', a_fetch is not None)

admin_fetch = db.obter_cotacao_por_id(cotacao_id, admin_id, 'ADMIN')
print('Admin consegue ver a cotacao de A?', admin_fetch is not None)

try:
    db.excluir_cotacao(cotacao_id, vend_b_id, 'VENDEDOR')
except Exception as e:
    print('Erro esperado ao excluir:', str(e))
    
a_fetch_after = db.obter_cotacao_por_id(cotacao_id, vend_a_id, 'VENDEDOR')
print('Cotacao ainda existe para A?', a_fetch_after is not None)

new_id = db.duplicar_cotacao(cotacao_id, vend_a_id, 'VENDEDOR')
print('A duplicou a cotacao. Novo id existe?', db.obter_cotacao_por_id(new_id, vend_a_id, 'VENDEDOR') is not None)

try:
    db.duplicar_cotacao(cotacao_id, vend_b_id, 'VENDEDOR')
    print('B duplicou (NAO DEVERIA)')
except Exception as e:
    print('Erro esperado ao duplicar por B:', str(e))
