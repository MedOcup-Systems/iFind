"""
database.py — Gerenciamento do banco de dados SQLite.
Tabelas: usuarios, execucoes, resultados
Criado automaticamente na primeira execução.
"""

import sqlite3
import hashlib
import json
from pathlib import Path
from datetime import datetime

DB_PATH = Path(__file__).parent / "clinica.db"


def _conexao():
    """Retorna conexão com row_factory para acesso por nome de coluna."""
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn


def inicializar_banco():
    """
    Cria as tabelas se não existirem e insere o usuário admin padrão.
    Deve ser chamada na inicialização do app.
    Admin padrão: usuário='admin', senha='admin123' (alterar nas Configurações).
    """
    conn = _conexao()
    cur = conn.cursor()

    cur.executescript("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id        INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario   TEXT UNIQUE NOT NULL,
            senha_hash TEXT NOT NULL,
            nome      TEXT NOT NULL,
            email     TEXT,
            admin     INTEGER DEFAULT 0,
            ativo     INTEGER DEFAULT 1,
            criado_em TEXT DEFAULT (datetime('now','localtime'))
        );

        CREATE TABLE IF NOT EXISTS execucoes (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario_id    INTEGER REFERENCES usuarios(id),
            inicio        TEXT,
            fim           TEXT,
            total         INTEGER DEFAULT 0,
            encontrados   INTEGER DEFAULT 0,
            nao_encontrados INTEGER DEFAULT 0,
            drive_raiz    TEXT,
            pasta_destino TEXT
        );

        CREATE TABLE IF NOT EXISTS resultados (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            execucao_id INTEGER REFERENCES execucoes(id),
            nome        TEXT,
            data        TEXT,
            encontrado  INTEGER DEFAULT 0,
            arquivo     TEXT,
            erro        TEXT,
            score_fuzzy REAL DEFAULT 100.0
        );
    """)

    # Cria admin padrão se não existir
    cur.execute("SELECT id FROM usuarios WHERE usuario = 'admin'")
    if not cur.fetchone():
        senha_hash = hashlib.sha256("admin123".encode()).hexdigest()
        cur.execute(
            "INSERT INTO usuarios (usuario, senha_hash, nome, admin) VALUES (?,?,?,?)",
            ("admin", senha_hash, "Administrador", 1)
        )

    conn.commit()
    conn.close()


# ---------------------------------------------------------------------------
# Usuários
# ---------------------------------------------------------------------------

def verificar_login(usuario: str, senha: str) -> dict | None:
    """
    Verifica credenciais. Retorna dict do usuário ou None se inválido.
    """
    senha_hash = hashlib.sha256(senha.encode()).hexdigest()
    conn = _conexao()
    row = conn.execute(
        "SELECT * FROM usuarios WHERE usuario=? AND senha_hash=? AND ativo=1",
        (usuario, senha_hash)
    ).fetchone()
    conn.close()
    return dict(row) if row else None


def listar_usuarios() -> list[dict]:
    conn = _conexao()
    rows = conn.execute("SELECT id, usuario, nome, email, admin, ativo, criado_em FROM usuarios").fetchall()
    conn.close()
    return [dict(r) for r in rows]


def criar_usuario(usuario: str, senha: str, nome: str, email: str = "", admin: int = 0) -> bool:
    """Cria novo usuário. Retorna False se o usuário já existir."""
    try:
        senha_hash = hashlib.sha256(senha.encode()).hexdigest()
        conn = _conexao()
        conn.execute(
            "INSERT INTO usuarios (usuario, senha_hash, nome, email, admin) VALUES (?,?,?,?,?)",
            (usuario, senha_hash, nome, email, admin)
        )
        conn.commit()
        conn.close()
        return True
    except sqlite3.IntegrityError:
        return False


def alterar_senha(usuario_id: int, nova_senha: str):
    senha_hash = hashlib.sha256(nova_senha.encode()).hexdigest()
    conn = _conexao()
    conn.execute("UPDATE usuarios SET senha_hash=? WHERE id=?", (senha_hash, usuario_id))
    conn.commit()
    conn.close()


def desativar_usuario(usuario_id: int):
    conn = _conexao()
    conn.execute("UPDATE usuarios SET ativo=0 WHERE id=?", (usuario_id,))
    conn.commit()
    conn.close()


# ---------------------------------------------------------------------------
# Execuções e resultados
# ---------------------------------------------------------------------------

def iniciar_execucao(usuario_id: int, drive_raiz: str, pasta_destino: str) -> int:
    """Registra início de uma execução. Retorna o ID da execução criada."""
    conn = _conexao()
    cur = conn.execute(
        "INSERT INTO execucoes (usuario_id, inicio, drive_raiz, pasta_destino) VALUES (?,?,?,?)",
        (usuario_id, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), drive_raiz, pasta_destino)
    )
    execucao_id = cur.lastrowid
    conn.commit()
    conn.close()
    return execucao_id


def finalizar_execucao(execucao_id: int, total: int, encontrados: int):
    """Atualiza os totais e o horário de fim da execução."""
    conn = _conexao()
    conn.execute(
        """UPDATE execucoes SET fim=?, total=?, encontrados=?, nao_encontrados=?
           WHERE id=?""",
        (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
         total, encontrados, total - encontrados, execucao_id)
    )
    conn.commit()
    conn.close()


def salvar_resultado(execucao_id: int, resultado: dict):
    """Salva o resultado de um registro individual."""
    conn = _conexao()
    conn.execute(
        """INSERT INTO resultados (execucao_id, nome, data, encontrado, arquivo, erro, score_fuzzy)
           VALUES (?,?,?,?,?,?,?)""",
        (execucao_id,
         resultado.get("nome", ""),
         resultado.get("data", ""),
         1 if resultado.get("encontrado") else 0,
         resultado.get("arquivo", ""),
         resultado.get("erro", ""),
         resultado.get("score_fuzzy", 100.0))
    )
    conn.commit()
    conn.close()


def listar_execucoes(usuario_id: int = None, limite: int = 50) -> list[dict]:
    """Lista execuções, opcionalmente filtradas por usuário."""
    conn = _conexao()
    if usuario_id:
        rows = conn.execute(
            """SELECT e.*, u.nome as usuario_nome
               FROM execucoes e JOIN usuarios u ON e.usuario_id=u.id
               WHERE e.usuario_id=? ORDER BY e.inicio DESC LIMIT ?""",
            (usuario_id, limite)
        ).fetchall()
    else:
        rows = conn.execute(
            """SELECT e.*, u.nome as usuario_nome
               FROM execucoes e JOIN usuarios u ON e.usuario_id=u.id
               ORDER BY e.inicio DESC LIMIT ?""",
            (limite,)
        ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def resultados_da_execucao(execucao_id: int) -> list[dict]:
    conn = _conexao()
    rows = conn.execute(
        "SELECT * FROM resultados WHERE execucao_id=? ORDER BY id",
        (execucao_id,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def estatisticas_gerais() -> dict:
    """
    Retorna um dicionário com estatísticas agregadas para o dashboard.
    Inclui: total de execuções, total de buscas, taxa de sucesso,
    execuções por dia (últimos 30 dias), top 5 usuários mais ativos.
    """
    conn = _conexao()

    totais = conn.execute("""
        SELECT
            COUNT(DISTINCT e.id)        AS total_execucoes,
            COALESCE(SUM(e.total), 0)   AS total_buscas,
            COALESCE(SUM(e.encontrados),0) AS total_encontrados
        FROM execucoes e
    """).fetchone()

    por_dia = conn.execute("""
        SELECT DATE(inicio) AS dia, COUNT(*) AS execucoes,
               SUM(encontrados) AS encontrados, SUM(total) AS total
        FROM execucoes
        WHERE inicio >= DATE('now','-30 days')
        GROUP BY DATE(inicio)
        ORDER BY dia
    """).fetchall()

    top_usuarios = conn.execute("""
        SELECT u.nome, COUNT(e.id) AS execucoes, SUM(e.encontrados) AS encontrados
        FROM execucoes e JOIN usuarios u ON e.usuario_id=u.id
        GROUP BY e.usuario_id
        ORDER BY execucoes DESC
        LIMIT 5
    """).fetchall()

    conn.close()

    taxa = 0.0
    if totais["total_buscas"] and totais["total_buscas"] > 0:
        taxa = round(totais["total_encontrados"] / totais["total_buscas"] * 100, 1)

    return {
        "total_execucoes" : totais["total_execucoes"],
        "total_buscas"    : totais["total_buscas"],
        "total_encontrados": totais["total_encontrados"],
        "taxa_sucesso"    : taxa,
        "por_dia"         : [dict(r) for r in por_dia],
        "top_usuarios"    : [dict(r) for r in top_usuarios],
    }
