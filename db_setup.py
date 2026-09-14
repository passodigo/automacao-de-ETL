"""
Setup do banco de dados (SQLite)
----------------------------------
Cria o arquivo do banco e o schema automaticamente na primeira execução.
Não exige instalar nem configurar nenhum servidor de banco - o arquivo
`comites.db` é criado ao lado deste script na primeira vez que rodar.

Uso:
    from db_setup import get_connection, inicializar_banco

    inicializar_banco()              # roda uma vez, cria tabelas se não existirem
    conn = get_connection()          # conexão pra usar no resto do pipeline
"""

import sqlite3
from pathlib import Path

# Caminho do arquivo do banco (fica ao lado deste script)
DB_PATH = Path(__file__).parent / "comites.db"


SCHEMA = """
CREATE TABLE IF NOT EXISTS comite (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT UNIQUE NOT NULL
);

CREATE TABLE IF NOT EXISTS comite_alias (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    comite_id INTEGER NOT NULL REFERENCES comite(id),
    alias TEXT UNIQUE NOT NULL,
    origem TEXT,
    confianca REAL,
    revisado_manualmente INTEGER DEFAULT 0,   -- 0 = false, 1 = true (SQLite não tem BOOLEAN nativo)
    criado_em TEXT DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS programa (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    codigo TEXT UNIQUE NOT NULL,
    descricao TEXT
);

CREATE TABLE IF NOT EXISTS exercicio (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ano INTEGER UNIQUE NOT NULL
);

CREATE TABLE IF NOT EXISTS carteira_previa (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    comite_id INTEGER NOT NULL REFERENCES comite(id),
    exercicio_id INTEGER NOT NULL REFERENCES exercicio(id),
    valor REAL NOT NULL,
    UNIQUE (comite_id, exercicio_id)
);

CREATE TABLE IF NOT EXISTS investimento (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    comite_id INTEGER NOT NULL REFERENCES comite(id),
    programa_id INTEGER NOT NULL REFERENCES programa(id),
    exercicio_id INTEGER NOT NULL REFERENCES exercicio(id),
    valor REAL NOT NULL,
    UNIQUE (comite_id, programa_id, exercicio_id)
);
"""


def get_connection() -> sqlite3.Connection:
    """Abre (ou cria) a conexão com o arquivo do banco."""
    conn = sqlite3.connect(DB_PATH)
    conn.execute("PRAGMA foreign_keys = ON")  # SQLite exige habilitar FK explicitamente
    return conn


def inicializar_banco():
    """Cria o arquivo do banco (se não existir) e aplica o schema."""
    conn = get_connection()
    try:
        conn.executescript(SCHEMA)
        conn.commit()
        print(f"Banco pronto em: {DB_PATH}")
    finally:
        conn.close()


if __name__ == "__main__":
    inicializar_banco()