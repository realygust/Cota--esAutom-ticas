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
    r'    # Tentativa de login automtico via cookie\n    if cookie_controller and st.session_state.user is None:.*?                cookie_controller.remove\(\'nextcable_user_hash\'\)\n',
    replacement1,
    content,
    flags=re.DOTALL
)

with open('main.py', 'w', encoding='utf-8') as f:
    f.write(content)
