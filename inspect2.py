import db_historico, json
q = db_historico.execute_query("SELECT * FROM cotacoes WHERE id = 'COT-20260921-171014'", fetchone=True)
if q:
    dados = json.loads(q.get('dados_json', '{}'))
    print('dados_json keys:', list(dados.keys()))
    if 'resultados_transportadoras' in dados:
        print('resultados_transportadoras keys:', dados['resultados_transportadoras'])
    if 'fretes' in dados:
        print('fretes:', dados['fretes'])
    if 'melhor_oferta' in dados:
        print('melhor_oferta:', dados['melhor_oferta'])
    print(json.dumps(dados, indent=2)[:500])
