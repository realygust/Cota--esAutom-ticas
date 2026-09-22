import re

with open('main.py', 'r', encoding='utf-8') as f:
    content = f.read()

replacement1 = """    # Tentativa de login automático via cookie
    if cookie_controller and st.session_state.user is None:
        saved_token = cookie_controller.get('nextcable_session_token')
        if saved_token:
            user_db = db_historico.authenticate_session_token(saved_token)
            if user_db and user_db.get("status") == "ATIVO":
                st.session_state.user = user_db
            else:
                cookie_controller.remove('nextcable_session_token')
"""

content = re.sub(
    r'    # Tentativa de login autom[áa]tico via cookie\n    if cookie_controller and st\.session_state\.user is None:.*?                cookie_controller\.remove\(\'nextcable_user_hash\'\)\n',
    replacement1,
    content,
    flags=re.DOTALL
)


replacement2 = """                            if manter_conectado and cookie_controller:
                                session_token = db_historico.create_session_token(user_db['id'])
                                cookie_controller.set('nextcable_session_token', session_token, max_age=60*60*24*30) # 30 dias
"""
content = re.sub(
    r'                            if manter_conectado and cookie_controller:\n                                cookie_controller\.set\(\'nextcable_user_email\', user_db\[\'email\'\], max_age=60\*60\*24\*30\) # 30 dias\n                                cookie_controller\.set\(\'nextcable_user_hash\', user_db\[\'password_hash\'\], max_age=60\*60\*24\*30\)\n',
    replacement2,
    content,
    flags=re.DOTALL
)

replacement3 = """            if cookie_controller_logout:
                if st.session_state.user:
                    db_historico.clear_session_token(st.session_state.user['id'])
                cookie_controller_logout.remove('nextcable_session_token')
"""
content = re.sub(
    r'            if cookie_controller_logout:\n                cookie_controller_logout\.remove\(\'nextcable_user_email\'\)\n                cookie_controller_logout\.remove\(\'nextcable_user_hash\'\)\n',
    replacement3,
    content,
    flags=re.DOTALL
)

with open('main.py', 'w', encoding='utf-8') as f:
    f.write(content)
