import re

with open('main.py', 'r', encoding='utf-8') as f:
    content = f.read()

replacement2 = """                        if manter_conectado and cookie_controller:
                            session_token = db_historico.create_session_token(user_db['id'])
                            cookie_controller.set('nextcable_session_token', session_token, max_age=60*60*24*30) # 30 dias
"""
content = re.sub(
    r'\s+if manter_conectado and cookie_controller:\n\s+cookie_controller\.set\(\'nextcable_user_email\', user_db\[\'email\'\], max_age=60\*60\*24\*30\) # 30 dias\n\s+cookie_controller\.set\(\'nextcable_user_hash\', user_db\[\'password_hash\'\], max_age=60\*60\*24\*30\)\n',
    "\n" + replacement2,
    content,
    flags=re.DOTALL
)

with open('main.py', 'w', encoding='utf-8') as f:
    f.write(content)
