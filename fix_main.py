import re

with open('main.py', 'r', encoding='utf-8') as f:
    content = f.read()

replacement = """                                        st.session_state.user = user_db
                                        if manter_conectado and cookie_controller:
                                            session_token = db_historico.create_session_token(user_db['id'])
                                            cookie_controller.set('nextcable_session_token', session_token, max_age=60*60*24*30) # 30 dias
                                        st.rerun()"""

content = re.sub(
    r'                                        st\.session_state\.user = user_db\n                        if manter_conectado and cookie_controller:\n                            session_token = db_historico\.create_session_token\(user_db\[\'id\'\]\)\n                            cookie_controller\.set\(\'nextcable_session_token\', session_token, max_age=60\*60\*24\*30\) # 30 dias\n                                        st\.rerun\(\)',
    replacement,
    content,
    flags=re.DOTALL
)

with open('main.py', 'w', encoding='utf-8') as f:
    f.write(content)
