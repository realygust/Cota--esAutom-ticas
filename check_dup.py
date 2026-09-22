import db_historico as db
db.init_db()
venda = db.execute_query("SELECT id FROM usuarios WHERE email='venda@test.com'", fetchone=True)
vendb = db.execute_query("SELECT id FROM usuarios WHERE email='vendb@test.com'", fetchone=True)
if venda and vendb:
    cotacoes_a = db.listar_historico(venda['id'], 'VENDEDOR')
    if cotacoes_a:
        cotacao_id = cotacoes_a[0]['id']
        new_id_b = db.duplicar_cotacao(cotacao_id, vendb['id'], 'VENDEDOR')
        print('B tentou duplicar. Retornou:', new_id_b)
