import sqlite3
from pathlib import Path
from .utils import hash_password  # Adicione este import

DB_PATH = Path(__file__).parent.parent / "users.db"

def get_connection():
    return sqlite3.connect(DB_PATH)

def init_db():
    conn = get_connection()
    c = conn.cursor()
    c.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        email TEXT NOT NULL UNIQUE,
        company TEXT,
        role TEXT,
        password TEXT NOT NULL,
        prev_passwords TEXT DEFAULT '',
        access TEXT DEFAULT 'all'
    )
    """)
    # Usuário padrão
    c.execute("SELECT * FROM users WHERE email = ?", ("tabitha.pessoa@philips.com",))
    if not c.fetchone():
        c.execute("""
            INSERT INTO users (name, email, company, role, password, prev_passwords, access)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            "Tabitha Pessôa",
            "tabitha.pessoa@philips.com",
            "Philips",
            "Especialista de testes",
            hash_password("Hed@1411"),  # Gera o hash corretamente
            "",
            "all"
        ))
    conn.commit()
    conn.close()