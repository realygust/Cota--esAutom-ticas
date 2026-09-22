import re

with open('db_historico.py', 'r', encoding='utf-8') as f:
    content = f.read()

# Fix the broken replacement from earlier
content = re.sub(
    r'    try:\n        execute_query\("ALTER TABLE usuarios ADD COLUMN session_token " \+ text_type\)\n        execute_query\("ALTER TABLE usuarios ADD COLUMN session_expires_at " \+ datetime_type\)\n    except Exception:\n        pass\n\n    execute_query\(f"""\n        CREATE TABLE IF NOT EXISTS cotacoes \(d_at \{datetime_type\}\n        \)\n    """\)',
    '',
    content
)

# Insert the session tokens columns to the user table
content = re.sub(
    r'password_salt \{text_type\},\n\s*created_at \{datetime_type\}',
    r'password_salt {text_type},\n            session_token {text_type},\n            session_expires_at {datetime_type},\n            created_at {datetime_type}',
    content
)

# Insert the alter table logic before the cotacoes table creation
content = re.sub(
    r'    execute_query\(f"""\n        CREATE TABLE IF NOT EXISTS cotacoes \(',
    r'    try:\n        execute_query("ALTER TABLE usuarios ADD COLUMN session_token " + text_type)\n        execute_query("ALTER TABLE usuarios ADD COLUMN session_expires_at " + datetime_type)\n    except Exception:\n        pass\n\n    execute_query(f"""\n        CREATE TABLE IF NOT EXISTS cotacoes (',
    content
)


with open('db_historico.py', 'w', encoding='utf-8') as f:
    f.write(content)
