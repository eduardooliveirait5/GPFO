# ═══════════════════════════════════════════════════════════
#  BASE DE DADOS
#  Responsável pela comunicação com o SQLite.
# ═══════════════════════════════════════════════════════════

import sqlite3
from app.config import DATABASE_PATH

def init_db():
    DATABASE_PATH.parent.mkdir(parents=True, exist_ok=True)

    conn = sqlite3.connect(DATABASE_PATH)
    conn.row_factory = sqlite3.Row

    conn.executescript("""
        CREATE TABLE IF NOT EXISTS formacoes (
            codigo      TEXT PRIMARY KEY,
            nome        TEXT NOT NULL DEFAULT '',
            criada_em   TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS sessoes (
            id                   INTEGER PRIMARY KEY AUTOINCREMENT,
            formacao_codigo      TEXT NOT NULL,
            titulo               TEXT NOT NULL,
            data                 TEXT NOT NULL,
            hora_inicio_janela   TEXT,
            hora_fim_janela      TEXT,
            duracao_sessao_min   REAL NOT NULL,
            ficheiro_original    TEXT,
            importada_em         TEXT NOT NULL,
            FOREIGN KEY (formacao_codigo) REFERENCES formacoes(codigo)
        );

        CREATE TABLE IF NOT EXISTS presencas (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            sessao_id       INTEGER NOT NULL,
            email           TEXT NOT NULL,
            nome            TEXT NOT NULL,
            funcao          TEXT,
            duracao_min     REAL NOT NULL DEFAULT 0,
            num_reconexoes  INTEGER NOT NULL DEFAULT 1,
            FOREIGN KEY (sessao_id) REFERENCES sessoes(id)
        );
    """)
    conn.commit()
    return conn

def get_conn():
    conn = sqlite3.connect(DATABASE_PATH)
    conn.row_factory = sqlite3.Row
    return conn

