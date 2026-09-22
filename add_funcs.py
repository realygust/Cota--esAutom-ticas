import secrets
import datetime

with open('db_historico.py', 'r', encoding='utf-8') as f:
    content = f.read()

if "import secrets" not in content:
    content = "import secrets\nimport datetime\n" + content

if "def create_session_token" not in content:
    content += """
def create_session_token(user_id: str) -> str:
    token = secrets.token_hex(32)
    expires_at = (datetime.datetime.now() + datetime.timedelta(days=30)).isoformat()
    query = "UPDATE usuarios SET session_token = ?, session_expires_at = ? WHERE id = ?"
    execute_query(query, (token, expires_at, user_id), commit=True)
    return token

def authenticate_session_token(token: str):
    if not token:
        return None
    query = "SELECT * FROM usuarios WHERE session_token = ?"
    user = execute_query(query, (token,), fetchone=True)
    if user and user.get("session_expires_at"):
        try:
            expires_at = datetime.datetime.fromisoformat(user["session_expires_at"])
            if datetime.datetime.now() < expires_at:
                return user
        except Exception:
            pass
    return None

def clear_session_token(user_id: str):
    query = "UPDATE usuarios SET session_token = NULL, session_expires_at = NULL WHERE id = ?"
    execute_query(query, (user_id,), commit=True)
"""

with open('db_historico.py', 'w', encoding='utf-8') as f:
    f.write(content)
