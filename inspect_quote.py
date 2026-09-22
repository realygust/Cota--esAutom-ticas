import db_historico, json
q = db_historico.execute_query("SELECT * FROM cotacoes WHERE id = 'COT-20260921-171014'", fetchone=True)
if q:
    dados = json.loads(q.get('dados_json', '{}'))
    print('resultados_transportadoras type:', type(dados.get('resultados_transportadoras')))
    if isinstance(dados.get('resultados_transportadoras'), dict):
        print('resultados_transportadoras keys:', dados.get('resultados_transportadoras').keys())
    print('resultados_transportadoras value:', json.dumps(dados.get('resultados_transportadoras'), indent=2))
    print('melhor_oferta:', json.dumps(dados.get('melhor_oferta'), indent=2))
else:
    print('Quote not found')
