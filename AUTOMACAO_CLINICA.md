# PROMPT MASTER V2 — Automação de Documentos para Clínica
## Sistema completo com todas as funcionalidades

---

> **REGRA ABSOLUTA PARA O AGENTE:** Leia este documento **do início ao fim** antes de criar qualquer arquivo. Siga cada parte **na ordem exata**. Não invente bibliotecas alternativas. Não pule etapas. Não crie arquivos com nomes diferentes dos especificados. Quando uma instrução disser "copie exatamente", copie exatamente — sem resumir, sem refatorar, sem "melhorar".

---

## PARTE 1 — VISÃO GERAL DO SISTEMA

### O problema

Uma clínica médica possui PDFs organizados em pastas por mês e dia. Cada PDF pode ter muitas páginas, cada uma correspondendo a um paciente diferente. O funcionário precisa localizar manualmente a página de cada paciente em uma lista Excel — processo lento e propenso a erros.

### O que será construído

Um sistema web local completo com **8 módulos**:

| Módulo | Arquivo | Função |
|---|---|---|
| Instalador autônomo | `setup.py` | Detecta SO, instala Tesseract, configura PATH |
| Lógica de busca | `processor.py` | Lê Excel, navega pastas, busca em PDFs, extrai páginas |
| Banco de dados | `database.py` | SQLite — histórico, usuários, estatísticas |
| Autenticação | `auth.py` | Login/logout, controle de sessão |
| E-mail | `mailer.py` | Envio automático dos PDFs encontrados |
| Busca inteligente | `fuzzy.py` — **integrado no processor.py** | Similaridade fonética com rapidfuzz |
| Interface principal | `app.py` | Streamlit com abas: Busca, Histórico, Estatísticas, Configurações |
| Configurações | `config.json` | Criado automaticamente na primeira execução |

### Arquitetura completa

```
USUÁRIO → navegador → app.py (Streamlit :8501)
                          │
              ┌───────────┼───────────────┐
              │           │               │
         processor.py  database.py    mailer.py
              │           │               │
         setup.py      SQLite DB      SMTP Server
         (Tesseract)   (histórico)    (e-mail)
              │
         fuzzy.py (integrado)
         OCR fallback (pytesseract)
```

---

## PARTE 2 — ESTRUTURA DE ARQUIVOS

Crie exatamente esta estrutura:

```
automacao_clinica/
├── setup.py          ← instalador autônomo Tesseract
├── processor.py      ← lógica de busca + fuzzy integrado
├── database.py       ← SQLite: histórico, usuários, stats
├── auth.py           ← autenticação de usuários
├── mailer.py         ← envio de e-mail com PDFs
├── app.py            ← interface Streamlit (4 abas)
├── requirements.txt  ← dependências
└── README.md         ← instruções
```

Arquivos gerados automaticamente em runtime (NÃO criar manualmente):
- `clinica.db` — banco SQLite
- `config.json` — configurações persistentes
- `config_tesseract.py` — caminho do Tesseract no Windows
- `output/` — pasta de PDFs extraídos

---

## PARTE 3 — `requirements.txt`

Crie exatamente assim:

```
streamlit>=1.32.0
streamlit-authenticator>=0.3.1
openpyxl>=3.1.0
PyMuPDF>=1.23.0
pdfplumber>=0.10.0
pytesseract>=0.3.10
Pillow>=10.0.0
pandas>=2.0.0
rapidfuzz>=3.6.0
plotly>=5.18.0
bcrypt>=4.1.0
```

---

## PARTE 4 — `setup.py`

Copie integralmente:

```python
"""
setup.py — Instalador autônomo do Tesseract OCR.
Suporta: Windows 10/11, Ubuntu/Debian/Fedora/Arch Linux, macOS.
Execute: python setup.py
"""

import os, sys, platform, subprocess, shutil, tempfile, ctypes
from pathlib import Path

TESSERACT_WINDOWS_URL = (
    "https://github.com/UB-Mannheim/tesseract/releases/download/"
    "v5.3.3.20231005/tesseract-ocr-w64-setup-5.3.3.20231005.exe"
)
TESSERACT_WINDOWS_PATHS = [
    Path(r"C:\Program Files\Tesseract-OCR"),
    Path(r"C:\Program Files (x86)\Tesseract-OCR"),
    Path(os.environ.get("LOCALAPPDATA", ""), "Tesseract-OCR"),
]

V = "\033[92m"; A = "\033[93m"; E = "\033[91m"; I = "\033[94m"; R = "\033[0m"; N = "\033[1m"

def info(m): print(f"{I}[INFO]{R} {m}")
def ok(m):   print(f"{V}[OK]{R}  {m}")
def aviso(m):print(f"{A}[AVISO]{R} {m}")
def erro(m): print(f"{E}[ERRO]{R} {m}")
def titulo(m):
    print(f"\n{N}{I}{'='*55}{R}\n{N}{I}  {m}{R}\n{N}{I}{'='*55}{R}\n")

def detectar_sistema():
    s = platform.system().lower()
    bits = "64" if sys.maxsize > 2**32 else "32"
    distro = pkg_mgr = None
    if s == "linux":
        os_release = Path("/etc/os-release")
        if os_release.exists():
            for linha in os_release.read_text().splitlines():
                if linha.startswith("ID="):
                    distro = linha.split("=")[1].strip().strip('"').lower()
                    break
        for mgr, cmd in [("apt","apt-get"),("dnf","dnf"),("yum","yum"),
                          ("pacman","pacman"),("zypper","zypper")]:
            if shutil.which(cmd):
                pkg_mgr = mgr; break
    elif s == "darwin":
        if shutil.which("brew"): pkg_mgr = "brew"
    return {"os": "windows" if s=="windows" else ("macos" if s=="darwin" else "linux"),
            "bits": bits, "distro": distro, "pkg_mgr": pkg_mgr}

def tesseract_instalado():
    if shutil.which("tesseract"):
        return True, _versao("tesseract")
    if platform.system().lower() == "windows":
        for c in TESSERACT_WINDOWS_PATHS:
            t = c / "tesseract.exe"
            if t.exists(): return True, _versao(str(t))
    return False, ""

def _versao(exe):
    try:
        r = subprocess.run([exe,"--version"], capture_output=True, text=True, timeout=10)
        linha = (r.stdout or r.stderr).splitlines()[0] if (r.stdout or r.stderr) else ""
        return linha.strip()
    except: return "versão desconhecida"

def tem_admin():
    if platform.system().lower() == "windows":
        try: return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except: return False
    return os.geteuid() == 0

def pedir_admin():
    if platform.system().lower() == "windows":
        aviso("Solicitando privilégios via UAC...")
        ctypes.windll.shell32.ShellExecuteW(
            None,"runas",sys.executable,f'"{os.path.abspath(sys.argv[0])}"',None,1)
        sys.exit(0)
    else:
        erro("Execute com: sudo python setup.py"); sys.exit(1)

def instalar_windows(bits):
    info("Baixando instalador Tesseract para Windows...")
    import urllib.request, urllib.error
    with tempfile.TemporaryDirectory() as tmp:
        dest = Path(tmp) / "tess.exe"
        try:
            def prog(b,bs,t):
                if t>0: print(f"\r  {min(100,int(b*bs*100/t))}%",end="",flush=True)
            urllib.request.urlretrieve(TESSERACT_WINDOWS_URL, dest, prog)
            print(); ok("Download concluído.")
        except urllib.error.URLError as e:
            erro(f"Falha no download: {e}"); sys.exit(1)
        install_dir = Path(r"C:\Program Files\Tesseract-OCR")
        info(f"Instalando em {install_dir}...")
        r = subprocess.run([str(dest),"/S",f"/D={install_dir}"], check=False)
        if r.returncode not in (0, 3010):
            erro(f"Instalador retornou {r.returncode}"); sys.exit(1)
    ok("Instalado.")
    _path_windows(install_dir)
    _config_pytesseract(install_dir)

def _path_windows(d):
    info("Adicionando ao PATH do sistema...")
    r = subprocess.run(["powershell","-Command",
        '[System.Environment]::GetEnvironmentVariable("Path","Machine")'],
        capture_output=True, text=True)
    atual = r.stdout.strip()
    if str(d) not in atual:
        subprocess.run(["powershell","-Command",
            f'[System.Environment]::SetEnvironmentVariable("Path","{atual};{d}","Machine")'],
            check=True)
        os.environ["PATH"] = os.environ.get("PATH","") + f";{d}"
    ok(f"PATH atualizado: {d}")
    aviso("Feche e reabra o terminal para efeito completo.")

def _config_pytesseract(d):
    p = Path(__file__).parent / "config_tesseract.py"
    p.write_text(f'import pytesseract\npytesseract.pytesseract.tesseract_cmd = r"{d/"tesseract.exe"}"\n',
                 encoding="utf-8")
    ok(f"config_tesseract.py criado em {p}")

def instalar_linux(pkg_mgr, distro):
    if not pkg_mgr:
        erro("Gerenciador de pacotes não detectado.")
        erro("Instale manualmente: sudo apt install tesseract-ocr tesseract-ocr-por"); sys.exit(1)
    cmds = {
        "apt":   [["apt-get","update","-qq"],
                  ["apt-get","install","-y","tesseract-ocr","tesseract-ocr-por"]],
        "dnf":   [["dnf","install","-y","tesseract","tesseract-langpack-por"]],
        "yum":   [["yum","install","-y","tesseract"]],
        "pacman":[["pacman","-Sy","--noconfirm","tesseract","tesseract-data-por"]],
        "zypper":[["zypper","install","-y","tesseract-ocr","tesseract-ocr-traineddata-portuguese"]],
    }
    for cmd in cmds.get(pkg_mgr, []):
        info(f"Executando: {' '.join(cmd)}")
        if subprocess.run(cmd, check=False).returncode != 0:
            erro(f"Falha: {' '.join(cmd)}"); sys.exit(1)
    ok("Tesseract instalado no Linux.")

def instalar_macos(pkg_mgr):
    if not pkg_mgr:
        info("Instalando Homebrew...")
        os.system('/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"')
        if not shutil.which("brew"):
            erro("Falha ao instalar Homebrew. Acesse: https://brew.sh"); sys.exit(1)
    subprocess.run(["brew","install","tesseract"], check=True)
    subprocess.run(["brew","install","tesseract-lang"], check=True)
    ok("Tesseract instalado no macOS.")

def verificar_deps():
    titulo("Verificando dependências Python")
    reqs = Path(__file__).parent / "requirements.txt"
    if not reqs.exists():
        aviso("requirements.txt não encontrado."); return
    info("Instalando dependências...")
    r = subprocess.run([sys.executable,"-m","pip","install","-r",str(reqs),"--quiet"], check=False)
    ok("Dependências OK.") if r.returncode == 0 else aviso("Verifique as dependências manualmente.")

def verificar_final():
    titulo("Verificação final")
    inst, versao = tesseract_instalado()
    if inst: ok(f"Tesseract: {versao}")
    else:
        erro("Tesseract não encontrado após instalação.")
        aviso("Feche o terminal, reabra e execute: python setup.py")
        return False
    try:
        import pytesseract
        cfg = Path(__file__).parent / "config_tesseract.py"
        if cfg.exists(): exec(cfg.read_text(), {})
        pytesseract.get_tesseract_version()
        ok("pytesseract OK.")
    except: aviso("pytesseract não conseguiu verificar — reinicie o terminal.")
    try:
        r = subprocess.run(["tesseract","--list-langs"], capture_output=True, text=True, timeout=10)
        if "por" in r.stdout + r.stderr: ok("Idioma português disponível.")
        else:
            aviso("Português não encontrado.")
            s = platform.system().lower()
            if s=="linux": print("  sudo apt install tesseract-ocr-por")
            elif s=="darwin": print("  brew install tesseract-lang")
    except: aviso("Não foi possível verificar idiomas.")
    return True

def main():
    titulo("Configuração do Sistema — Clínica v2")
    sys_info = detectar_sistema()
    os_nome, bits, distro, pkg_mgr = (sys_info[k] for k in ["os","bits","distro","pkg_mgr"])
    info(f"Sistema: {platform.system()} {platform.release()} ({bits}-bit)")
    if distro: info(f"Distro Linux: {distro}")
    if pkg_mgr: info(f"Gerenciador: {pkg_mgr}")
    verificar_deps()
    titulo("Verificando Tesseract OCR")
    inst, versao = tesseract_instalado()
    if inst:
        ok(f"Tesseract já instalado: {versao}")
    else:
        aviso("Tesseract NÃO encontrado.")
        if not tem_admin(): pedir_admin()
        titulo(f"Instalando — {os_nome.upper()}")
        if os_nome == "windows": instalar_windows(bits)
        elif os_nome == "linux": instalar_linux(pkg_mgr, distro)
        elif os_nome == "macos": instalar_macos(pkg_mgr)
        else: erro(f"SO não suportado: {os_nome}"); sys.exit(1)
    verificar_final()
    titulo("Sistema pronto")
    print(f"  {N}\033[92mstreamlit run app.py\033[0m\n")

if __name__ == "__main__":
    main()
```

---

## PARTE 5 — `database.py`

Copie integralmente. Este arquivo gerencia o banco SQLite com três tabelas: `usuarios`, `execucoes` e `resultados`.

```python
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
```

---

## PARTE 6 — `auth.py`

Copie integralmente. Gerencia sessão de login via `st.session_state`.

```python
"""
auth.py — Controle de autenticação e sessão Streamlit.
Usa st.session_state para persistir o usuário logado.
"""

import streamlit as st
from database import verificar_login, inicializar_banco


def inicializar_sessao():
    """
    Deve ser chamada no início de cada página do app.
    Garante que as chaves de sessão existam.
    """
    if "usuario_logado" not in st.session_state:
        st.session_state["usuario_logado"] = None
    if "usuario_id" not in st.session_state:
        st.session_state["usuario_id"] = None
    if "usuario_admin" not in st.session_state:
        st.session_state["usuario_admin"] = False


def esta_logado() -> bool:
    return st.session_state.get("usuario_logado") is not None


def usuario_atual() -> dict | None:
    """Retorna dict com dados do usuário logado ou None."""
    if not esta_logado():
        return None
    return {
        "id"    : st.session_state.get("usuario_id"),
        "nome"  : st.session_state.get("usuario_logado"),
        "admin" : st.session_state.get("usuario_admin", False),
    }


def fazer_login(usuario: str, senha: str) -> bool:
    """
    Valida credenciais. Se válidas, salva na sessão e retorna True.
    """
    user = verificar_login(usuario, senha)
    if user:
        st.session_state["usuario_logado"] = user["nome"]
        st.session_state["usuario_id"]     = user["id"]
        st.session_state["usuario_admin"]  = bool(user["admin"])
        return True
    return False


def fazer_logout():
    """Limpa a sessão do usuário."""
    for chave in ["usuario_logado", "usuario_id", "usuario_admin"]:
        st.session_state[chave] = None
    st.session_state["usuario_admin"] = False


def tela_login():
    """
    Renderiza a tela de login e bloqueia o app até que o usuário se autentique.
    Chame esta função no início do app.py antes de qualquer conteúdo.
    """
    inicializar_sessao()

    if esta_logado():
        return  # já autenticado, continua normalmente

    st.title("🔍 Buscador de Documentos")
    st.caption("Sistema de busca automática em PDFs — Clínica")
    st.divider()

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.subheader("Entrar")
        usuario = st.text_input("Usuário", placeholder="seu.usuario")
        senha   = st.text_input("Senha",   type="password", placeholder="••••••••")

        if st.button("Entrar", type="primary", use_container_width=True):
            if fazer_login(usuario, senha):
                st.rerun()
            else:
                st.error("Usuário ou senha incorretos.")

        st.caption("Primeiro acesso: usuário `admin`, senha `admin123`")
        st.caption("Altere a senha em **Configurações** após o primeiro login.")

    st.stop()  # bloqueia o restante do app até login
```

---

## PARTE 7 — `mailer.py`

Copie integralmente. Envia e-mails com PDF anexado via SMTP.

```python
"""
mailer.py — Envio de e-mails com PDFs encontrados.
Configuração SMTP salva em config.json pelo app.py.
"""

import smtplib
import json
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text      import MIMEText
from email.mime.base      import MIMEBase
from email                import encoders
from pathlib               import Path


CONFIG_PATH = Path(__file__).parent / "config.json"


def _carregar_config() -> dict:
    """Lê as configurações de e-mail do config.json."""
    if not CONFIG_PATH.exists():
        return {}
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def email_configurado() -> bool:
    """Retorna True se as configurações mínimas de SMTP estão presentes."""
    cfg = _carregar_config()
    smtp = cfg.get("smtp", {})
    return bool(smtp.get("host") and smtp.get("usuario") and smtp.get("senha"))


def enviar_email(
    destinatario: str,
    assunto: str,
    corpo: str,
    caminho_pdf: str = None
) -> tuple[bool, str]:
    """
    Envia um e-mail via SMTP, com PDF opcional como anexo.

    Parâmetros:
        destinatario : endereço de e-mail do destinatário
        assunto      : assunto do e-mail
        corpo        : texto do corpo (suporta HTML básico)
        caminho_pdf  : caminho do PDF a anexar (opcional)

    Retorna:
        (True, "")           em caso de sucesso
        (False, mensagem_erro) em caso de falha
    """
    cfg = _carregar_config()
    smtp_cfg = cfg.get("smtp", {})

    host     = smtp_cfg.get("host", "")
    porta    = int(smtp_cfg.get("porta", 587))
    usuario  = smtp_cfg.get("usuario", "")
    senha    = smtp_cfg.get("senha", "")
    remetente = smtp_cfg.get("remetente", usuario)

    if not all([host, usuario, senha]):
        return False, "Configurações SMTP incompletas. Configure em Configurações > E-mail."

    msg = MIMEMultipart()
    msg["From"]    = remetente
    msg["To"]      = destinatario
    msg["Subject"] = assunto
    msg.attach(MIMEText(corpo, "html"))

    # Anexa o PDF se fornecido e se o arquivo existir
    if caminho_pdf and Path(caminho_pdf).exists():
        with open(caminho_pdf, "rb") as f:
            parte = MIMEBase("application", "octet-stream")
            parte.set_payload(f.read())
        encoders.encode_base64(parte)
        nome_arquivo = Path(caminho_pdf).name
        parte.add_header("Content-Disposition", f'attachment; filename="{nome_arquivo}"')
        msg.attach(parte)

    try:
        with smtplib.SMTP(host, porta, timeout=15) as servidor:
            servidor.ehlo()
            servidor.starttls()
            servidor.login(usuario, senha)
            servidor.sendmail(remetente, destinatario, msg.as_string())
        return True, ""
    except smtplib.SMTPAuthenticationError:
        return False, "Falha na autenticação SMTP. Verifique usuário e senha."
    except smtplib.SMTPConnectError:
        return False, f"Não foi possível conectar ao servidor {host}:{porta}."
    except Exception as e:
        return False, f"Erro ao enviar e-mail: {str(e)}"


def enviar_relatorio_execucao(
    destinatario: str,
    total: int,
    encontrados: int,
    nao_encontrados: int,
    usuario_nome: str
) -> tuple[bool, str]:
    """
    Envia um e-mail resumo ao final de uma execução.
    """
    taxa = round(encontrados / total * 100) if total > 0 else 0
    assunto = f"Relatório de Busca — {encontrados}/{total} encontrados"
    corpo = f"""
    <html><body style="font-family: Arial, sans-serif; color: #333;">
    <h2 style="color: #1D9E75;">Relatório de Busca de Documentos</h2>
    <p>Executado por: <strong>{usuario_nome}</strong></p>
    <hr>
    <table style="border-collapse:collapse; width:100%;">
      <tr><td style="padding:8px;background:#f5f5f5;">Total de registros</td>
          <td style="padding:8px;font-weight:bold;">{total}</td></tr>
      <tr><td style="padding:8px;background:#d4edda;">Encontrados</td>
          <td style="padding:8px;font-weight:bold;color:#155724;">{encontrados}</td></tr>
      <tr><td style="padding:8px;background:#f8d7da;">Não encontrados</td>
          <td style="padding:8px;font-weight:bold;color:#721c24;">{nao_encontrados}</td></tr>
      <tr><td style="padding:8px;background:#f5f5f5;">Taxa de sucesso</td>
          <td style="padding:8px;font-weight:bold;">{taxa}%</td></tr>
    </table>
    <p style="color:#888;font-size:12px;margin-top:20px;">
      Enviado automaticamente pelo Sistema de Busca de Documentos — Clínica
    </p>
    </body></html>
    """
    return enviar_email(destinatario, assunto, corpo)
```

---

## PARTE 8 — `processor.py`

Copie integralmente. Inclui busca fuzzy com `rapidfuzz` integrada.

```python
"""
processor.py — Lógica de busca e extração de documentos PDF.
Inclui busca fuzzy (tolerância a erros de digitação/OCR) via rapidfuzz.
Não renomeie as funções — app.py e database.py dependem delas.
"""

import fitz  # PyMuPDF
import openpyxl
import unicodedata
import re
import io
from pathlib import Path
from datetime import datetime


# ---------------------------------------------------------------------------
# Utilitários de texto
# ---------------------------------------------------------------------------

def normalizar_texto(texto: str) -> str:
    """Remove acentos, converte para minúsculas, colapsa espaços."""
    if not texto:
        return ""
    texto = unicodedata.normalize("NFD", texto)
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", texto.lower()).strip()


def nome_contem(texto_pagina: str, nome_buscado: str, threshold: float = 80.0) -> tuple[bool, float]:
    """
    Verifica se o nome aparece no texto da página.

    Estratégia dupla:
    1. Busca exata (normalizada) — retorna score 100.0
    2. Busca fuzzy com rapidfuzz — retorna score da melhor correspondência

    O threshold define a pontuação mínima para aceitar uma correspondência fuzzy.
    Padrão: 80.0 (tolerante a pequenos erros de OCR).

    Retorna:
        (encontrado: bool, score: float)
    """
    texto_norm = normalizar_texto(texto_pagina)
    nome_norm  = normalizar_texto(nome_buscado)

    # 1. Busca exata — todas as palavras do nome presentes no texto
    palavras = nome_norm.split()
    if all(p in texto_norm for p in palavras):
        return True, 100.0

    # 2. Busca fuzzy
    try:
        from rapidfuzz import fuzz, process

        # Divide o texto em janelas do tamanho do nome para comparação local
        palavras_texto = texto_norm.split()
        tam_nome = len(palavras)

        melhor_score = 0.0
        for i in range(max(1, len(palavras_texto) - tam_nome + 1)):
            janela = " ".join(palavras_texto[i : i + tam_nome + 2])
            score = fuzz.token_set_ratio(nome_norm, janela)
            if score > melhor_score:
                melhor_score = score

        if melhor_score >= threshold:
            return True, float(melhor_score)

    except ImportError:
        pass  # rapidfuzz não instalado — continua com busca exata apenas

    return False, 0.0


# ---------------------------------------------------------------------------
# Leitura da planilha Excel
# ---------------------------------------------------------------------------

def ler_planilha(caminho_excel: str) -> list[dict]:
    """
    Lê o Excel e retorna lista de {'nome': str, 'data': any, 'email': str}.

    Detecção automática de colunas por palavras-chave no cabeçalho.
    Aceita também coluna de e-mail para envio automático (opcional).
    Fallback: coluna A = nome, coluna B = data.
    """
    wb = openpyxl.load_workbook(caminho_excel, data_only=True)
    ws = wb.active

    if ws.max_row < 2:
        raise ValueError("A planilha está vazia ou tem apenas o cabeçalho.")

    cabecalho = [str(c.value).lower().strip() if c.value else "" for c in ws[1]]

    col_nome  = next((i for i,h in enumerate(cabecalho)
                      if any(p in h for p in {"nome","paciente","patient","name"})), 0)
    col_data  = next((i for i,h in enumerate(cabecalho)
                      if any(p in h for p in {"data","date","dt","dia","procedimento","proc"})), 1)
    col_email = next((i for i,h in enumerate(cabecalho)
                      if any(p in h for p in {"email","e-mail","correio","mail"})), None)

    registros = []
    for linha in ws.iter_rows(min_row=2, values_only=True):
        if all(v is None for v in linha):
            continue
        nome = linha[col_nome] if col_nome < len(linha) else None
        if nome is None:
            continue
        data  = linha[col_data]  if col_data  < len(linha) else None
        email = linha[col_email] if col_email is not None and col_email < len(linha) else ""
        registros.append({
            "nome" : str(nome).strip(),
            "data" : data,
            "email": str(email).strip() if email else "",
        })

    if not registros:
        raise ValueError("Nenhum registro válido encontrado na planilha.")
    return registros


# ---------------------------------------------------------------------------
# Navegação nas pastas do drive
# ---------------------------------------------------------------------------

MESES = {
    "01":"janeiro",  "02":"fevereiro", "03":"marco",
    "04":"abril",    "05":"maio",      "06":"junho",
    "07":"julho",    "08":"agosto",    "09":"setembro",
    "10":"outubro",  "11":"novembro",  "12":"dezembro",
}

# Para personalizar os nomes das pastas, edite o dicionário MESES acima.
# Exemplos de variações comuns:
#   "03":"Março"     (com acento — se as pastas foram criadas assim)
#   "03":"03"        (número puro)
#   "03":"march"     (inglês)


def extrair_dia_mes(data) -> tuple[str, str]:
    """
    Converte qualquer formato de data em (dia, mes) com zero à esquerda.
    Aceita: datetime, 'DD/MM', 'DD/MM/AAAA', 'AAAA-MM-DD', serial Excel.
    """
    if isinstance(data, datetime):
        return data.strftime("%d"), data.strftime("%m")
    texto = str(data).strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}", texto):
        p = texto.split("-"); return p[2][:2], p[1][:2]
    if "/" in texto:
        p = texto.split("/"); return p[0].strip().zfill(2), p[1].strip().zfill(2)
    try:
        import datetime as dt_mod
        n = int(float(texto))
        d = dt_mod.datetime(1899,12,30) + dt_mod.timedelta(days=n)
        return d.strftime("%d"), d.strftime("%m")
    except Exception:
        pass
    raise ValueError(f"Formato de data não reconhecido: '{data}'. Use DD/MM, DD/MM/AAAA ou AAAA-MM-DD.")


def montar_caminho_pasta(drive_raiz: str, data) -> Path:
    """Constrói: drive_raiz / nome_do_mes / dia"""
    dia, mes_num = extrair_dia_mes(data)
    mes_nome = MESES.get(mes_num, mes_num)
    return Path(drive_raiz) / mes_nome / dia


# ---------------------------------------------------------------------------
# Busca dentro dos PDFs
# ---------------------------------------------------------------------------

def extrair_texto_pagina(pagina: fitz.Page) -> str:
    """
    Extrai texto de uma página PDF.
    Estratégia: texto embutido primeiro; se vazio, aplica OCR a 200 DPI.
    """
    texto = pagina.get_text()
    if len(texto.strip()) < 20:
        try:
            from PIL import Image
            import pytesseract
            cfg_path = Path(__file__).parent / "config_tesseract.py"
            if cfg_path.exists():
                exec(cfg_path.read_text(), {})
            pix = pagina.get_pixmap(dpi=200)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            try:    texto = pytesseract.image_to_string(img, lang="por")
            except: texto = pytesseract.image_to_string(img)
        except ImportError:
            pass
    return texto


def buscar_nome_em_pdf(caminho_pdf: Path, nome_buscado: str,
                        threshold: float = 80.0) -> list[tuple[int, float]]:
    """
    Busca o nome em todas as páginas do PDF.

    Retorna lista de (numero_pagina, score_fuzzy).
    Score 100.0 = correspondência exata; < 100 = correspondência fuzzy.
    Lista vazia = não encontrado.
    """
    encontradas = []
    doc = fitz.open(str(caminho_pdf))
    for num, pagina in enumerate(doc):
        texto = extrair_texto_pagina(pagina)
        achou, score = nome_contem(texto, nome_buscado, threshold)
        if achou:
            encontradas.append((num, score))
    doc.close()
    return encontradas


# ---------------------------------------------------------------------------
# Extração e salvamento
# ---------------------------------------------------------------------------

def sanitizar_nome_arquivo(nome: str) -> str:
    """Converte nome em string válida para qualquer SO."""
    sem_acento = unicodedata.normalize("NFD", nome)
    sem_acento = "".join(c for c in sem_acento if unicodedata.category(c) != "Mn")
    limpo = re.sub(r"[^\w\s-]", "", sem_acento)
    return re.sub(r"\s+", "_", limpo.strip())


def extrair_e_salvar_paginas(
    caminho_pdf: Path,
    paginas_scores: list[tuple[int, float]],
    pasta_destino: Path,
    nome_paciente: str
) -> Path:
    """
    Extrai as páginas indicadas e salva como novo PDF.
    Evita sobrescrever: adiciona sufixo numérico se o arquivo já existir.
    """
    pasta_destino = Path(pasta_destino)
    pasta_destino.mkdir(parents=True, exist_ok=True)

    nome_base = sanitizar_nome_arquivo(nome_paciente)
    saida = pasta_destino / f"{nome_base}.pdf"
    contador = 2
    while saida.exists():
        saida = pasta_destino / f"{nome_base}_{contador}.pdf"
        contador += 1

    doc_orig = fitz.open(str(caminho_pdf))
    doc_novo = fitz.open()
    for idx, _ in sorted(paginas_scores):
        doc_novo.insert_pdf(doc_orig, from_page=idx, to_page=idx)
    doc_novo.save(str(saida))
    doc_novo.close()
    doc_orig.close()
    return saida


# ---------------------------------------------------------------------------
# Função principal
# ---------------------------------------------------------------------------

def processar_lista(
    caminho_excel: str,
    drive_raiz: str,
    pasta_destino: str,
    threshold_fuzzy: float = 80.0,
    callback=None
) -> list[dict]:
    """
    Processa todos os registros da planilha.

    Parâmetros:
        caminho_excel   : caminho para o .xlsx
        drive_raiz      : raiz do drive com as pastas dos meses
        pasta_destino   : onde salvar os PDFs extraídos
        threshold_fuzzy : pontuação mínima para busca fuzzy (0-100, padrão 80)
        callback        : função(progresso: float, mensagem: str) para atualizar UI

    Retorna:
        Lista de dicts: nome, data, email, encontrado, arquivo, erro, score_fuzzy
    """
    registros  = ler_planilha(caminho_excel)
    resultados = []
    total      = len(registros)

    for i, reg in enumerate(registros):
        nome  = reg["nome"]
        data  = reg["data"]
        email = reg.get("email", "")

        resultado = {
            "nome"       : nome,
            "data"       : str(data),
            "email"      : email,
            "encontrado" : False,
            "arquivo"    : "",
            "erro"       : "",
            "score_fuzzy": 0.0,
        }

        try:
            pasta_dia = montar_caminho_pasta(drive_raiz, data)

            if not pasta_dia.exists():
                resultado["erro"] = f"Pasta não encontrada: {pasta_dia}"
            else:
                pdfs = sorted(pasta_dia.glob("*.pdf"))
                if not pdfs:
                    resultado["erro"] = f"Nenhum PDF em: {pasta_dia}"
                else:
                    melhor_paginas = []
                    melhor_pdf     = None
                    melhor_score   = 0.0

                    for pdf in pdfs:
                        paginas = buscar_nome_em_pdf(pdf, nome, threshold_fuzzy)
                        if paginas:
                            score_max = max(s for _, s in paginas)
                            if score_max > melhor_score:
                                melhor_paginas = paginas
                                melhor_pdf     = pdf
                                melhor_score   = score_max
                            if melhor_score == 100.0:
                                break  # correspondência perfeita, para imediatamente

                    if melhor_paginas:
                        saida = extrair_e_salvar_paginas(
                            melhor_pdf, melhor_paginas,
                            Path(pasta_destino), nome
                        )
                        resultado["encontrado"]  = True
                        resultado["arquivo"]     = str(saida)
                        resultado["score_fuzzy"] = melhor_score
                    else:
                        resultado["erro"] = "Nome não encontrado nos PDFs da pasta."

        except Exception as e:
            resultado["erro"] = str(e)

        resultados.append(resultado)

        if callback:
            icone = "✓" if resultado["encontrado"] else "✗"
            callback((i + 1) / total, f"{icone} {nome}")

    return resultados
```

---

## PARTE 9 — `app.py`

Copie integralmente. Interface Streamlit com 4 abas: **Busca**, **Histórico**, **Estatísticas**, **Configurações**.

```python
"""
app.py — Interface Streamlit completa.
4 abas: Busca, Histórico, Estatísticas, Configurações.

Para rodar na rede da clínica:
    streamlit run app.py --server.address 0.0.0.0 --server.port 8501
"""

import streamlit as st
import tempfile
import os
import platform
import shutil
import subprocess
import sys
import json
import pandas as pd
from pathlib import Path
from datetime import datetime

# ---------------------------------------------------------------------------
# set_page_config — OBRIGATORIAMENTE o primeiro comando Streamlit
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="Buscador de Documentos — Clínica",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ---------------------------------------------------------------------------
# Verificação do Tesseract
# ---------------------------------------------------------------------------

def _tesseract_ok() -> bool:
    if shutil.which("tesseract"):
        return True
    if platform.system().lower() == "windows":
        for p in [Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
                  Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe")]:
            if p.exists():
                os.environ["PATH"] += f";{p.parent}"
                return True
    return False


def _verificar_tesseract():
    if _tesseract_ok():
        return
    st.warning("⚙️ Tesseract OCR não encontrado. Instalando automaticamente...")
    setup = Path(__file__).parent / "setup.py"
    if not setup.exists():
        st.error("setup.py não encontrado. Coloque-o na mesma pasta que o app.py.")
        st.stop()
    with st.status("Instalando Tesseract OCR...", expanded=True) as status:
        st.write("Detectando sistema e iniciando instalação...")
        r = subprocess.run([sys.executable, str(setup)],
                           capture_output=True, text=True, timeout=300)
        if r.returncode == 0:
            status.update(label="Instalado com sucesso!", state="complete")
            st.code(r.stdout, language=None)
            st.info("Recarregue a página (F5) para continuar.")
        else:
            status.update(label="Falha na instalação.", state="error")
            st.error("Não foi possível instalar automaticamente.")
            st.code(r.stdout + r.stderr, language=None)
        st.stop()


_verificar_tesseract()

# ---------------------------------------------------------------------------
# Inicialização
# ---------------------------------------------------------------------------

try:
    from database   import inicializar_banco, listar_execucoes, resultados_da_execucao, \
                           iniciar_execucao, finalizar_execucao, salvar_resultado, \
                           estatisticas_gerais, listar_usuarios, criar_usuario, \
                           alterar_senha, desativar_usuario
    from auth       import tela_login, usuario_atual, fazer_logout, inicializar_sessao
    from processor  import processar_lista
    from mailer     import enviar_email, enviar_relatorio_execucao, email_configurado
except ImportError as e:
    st.error(f"Erro de importação: {e}")
    st.info("Execute: pip install -r requirements.txt")
    st.stop()

inicializar_banco()
tela_login()  # bloqueia até o usuário fazer login

usuario = usuario_atual()

# ---------------------------------------------------------------------------
# Caminho do config.json
# ---------------------------------------------------------------------------

CONFIG_PATH = Path(__file__).parent / "config.json"


def carregar_config() -> dict:
    if not CONFIG_PATH.exists():
        return {"smtp": {}, "drive_raiz": "", "pasta_destino": "",
                "threshold_fuzzy": 80, "enviar_email_auto": False,
                "email_relatorio": ""}
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)


def salvar_config(cfg: dict):
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


cfg = carregar_config()

# ---------------------------------------------------------------------------
# Header global
# ---------------------------------------------------------------------------

col_h1, col_h2 = st.columns([8, 2])
with col_h1:
    st.markdown(f"### 🔍 Buscador de Documentos")
with col_h2:
    st.caption(f"Olá, **{usuario['nome']}**")
    if st.button("Sair", key="btn_logout"):
        fazer_logout()
        st.rerun()

st.divider()

# ---------------------------------------------------------------------------
# Abas principais
# ---------------------------------------------------------------------------

aba_busca, aba_historico, aba_stats, aba_config = st.tabs([
    "🔍 Busca", "📋 Histórico", "📊 Estatísticas", "⚙️ Configurações"
])


# ============================================================
# ABA 1 — BUSCA
# ============================================================

with aba_busca:
    st.subheader("Nova busca")

    col1, col2 = st.columns(2)

    with col1:
        arquivo_excel = st.file_uploader(
            "📄 Planilha Excel com nomes e datas",
            type=["xlsx", "xls"],
            help="Colunas detectadas automaticamente. Aceita coluna de e-mail para envio automático."
        )
        drive_raiz = st.text_input(
            "📁 Caminho raiz do drive",
            value=cfg.get("drive_raiz", ""),
            placeholder="Ex: Z:\\Procedimentos  ou  /mnt/drive/proc"
        )

    with col2:
        pasta_destino = st.text_input(
            "📂 Pasta de destino",
            value=cfg.get("pasta_destino", ""),
            placeholder="Ex: C:\\Resultados  ou  /home/user/resultados"
        )
        threshold = st.slider(
            "🎯 Sensibilidade da busca fuzzy",
            min_value=60, max_value=100,
            value=cfg.get("threshold_fuzzy", 80),
            step=5,
            help="100 = apenas correspondência exata. 80 = tolera pequenos erros de OCR. 60 = muito permissivo."
        )

    enviar_email_auto = st.checkbox(
        "Enviar e-mail automático ao final (requer coluna 'email' na planilha e SMTP configurado)",
        value=cfg.get("enviar_email_auto", False)
    )

    email_relatorio = ""
    if enviar_email_auto:
        email_relatorio = st.text_input(
            "E-mail para relatório de conclusão",
            value=cfg.get("email_relatorio", ""),
            placeholder="gestor@clinica.com.br"
        )

    iniciar = st.button("▶ Iniciar busca", type="primary", use_container_width=True)

    if iniciar:
        erros = []
        if not arquivo_excel: erros.append("Faça upload da planilha Excel.")
        if not drive_raiz.strip(): erros.append("Informe o caminho do drive.")
        if not pasta_destino.strip(): erros.append("Informe a pasta de destino.")
        if erros:
            for e in erros: st.error(f"❌ {e}")
            st.stop()

        # Salva preferências
        cfg.update({"drive_raiz": drive_raiz.strip(), "pasta_destino": pasta_destino.strip(),
                    "threshold_fuzzy": threshold, "enviar_email_auto": enviar_email_auto,
                    "email_relatorio": email_relatorio})
        salvar_config(cfg)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(arquivo_excel.read())
            caminho_tmp = tmp.name

        execucao_id = iniciar_execucao(usuario["id"], drive_raiz.strip(), pasta_destino.strip())

        st.divider()
        st.subheader("Progresso")
        barra   = st.progress(0, text="Iniciando...")
        st_txt  = st.empty()
        st_log  = st.empty()
        log_lines = []

        def cb(prog, msg):
            barra.progress(prog, text=f"{int(prog*100)}% concluído")
            st_txt.markdown(f"Processando: **{msg}**")
            log_lines.append(msg)
            st_log.code("\n".join(log_lines[-20:]), language=None)

        try:
            resultados = processar_lista(
                caminho_excel   = caminho_tmp,
                drive_raiz      = drive_raiz.strip(),
                pasta_destino   = pasta_destino.strip(),
                threshold_fuzzy = threshold,
                callback        = cb,
            )
        except Exception as e:
            st.error(f"❌ Erro crítico: {e}")
            os.unlink(caminho_tmp)
            st.stop()

        os.unlink(caminho_tmp)

        # Salva no banco
        total       = len(resultados)
        encontrados = sum(1 for r in resultados if r["encontrado"])
        for r in resultados:
            salvar_resultado(execucao_id, r)
        finalizar_execucao(execucao_id, total, encontrados)

        barra.progress(1.0, text="Concluído!")
        st_txt.empty()

        # Envio de e-mails automático
        if enviar_email_auto and email_configurado():
            with st.spinner("Enviando e-mails..."):
                enviados = 0
                for r in resultados:
                    if r["encontrado"] and r.get("email"):
                        ok_mail, err_mail = enviar_email(
                            destinatario=r["email"],
                            assunto=f"Documento encontrado — {r['nome']}",
                            corpo=f"<p>Prezado(a), o documento de <strong>{r['nome']}</strong> "
                                  f"referente a <strong>{r['data']}</strong> foi localizado e está anexo.</p>",
                            caminho_pdf=r["arquivo"]
                        )
                        if ok_mail: enviados += 1
                if email_relatorio:
                    enviar_relatorio_execucao(
                        email_relatorio, total, encontrados, total - encontrados, usuario["nome"]
                    )
            st.success(f"📧 {enviados} e-mail(s) enviado(s) com sucesso.")

        # Resultados
        st.divider()
        st.subheader("Resultados")
        nao_enc = total - encontrados
        taxa    = round(encontrados / total * 100) if total > 0 else 0

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total",           total)
        c2.metric("Encontrados",     encontrados)
        c3.metric("Não encontrados", nao_enc)
        c4.metric("Taxa de sucesso", f"{taxa}%")

        df = pd.DataFrame(resultados).rename(columns={
            "nome":"Nome","data":"Data","email":"E-mail",
            "encontrado":"Encontrado","arquivo":"Arquivo gerado",
            "erro":"Observação","score_fuzzy":"Score"
        })

        def colorir(row):
            cor = "#d4edda" if row["Encontrado"] else "#f8d7da"
            return [f"background-color:{cor}"] * len(row)

        st.dataframe(df.style.apply(colorir, axis=1),
                     use_container_width=True, hide_index=True)

        csv = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button("⬇️ Baixar relatório CSV", csv,
                           "relatorio_busca.csv", "text/csv",
                           use_container_width=True)

        if nao_enc > 0:
            st.warning(f"⚠️ {nao_enc} registro(s) não encontrado(s). "
                       "Verifique a coluna 'Observação'. "
                       "Tente reduzir a sensibilidade fuzzy se houver falsos positivos.")


# ============================================================
# ABA 2 — HISTÓRICO
# ============================================================

with aba_historico:
    st.subheader("Histórico de execuções")

    # Admin vê todas; usuário comum vê só as suas
    uid = None if usuario["admin"] else usuario["id"]
    execucoes = listar_execucoes(usuario_id=uid, limite=100)

    if not execucoes:
        st.info("Nenhuma execução registrada ainda.")
    else:
        df_exec = pd.DataFrame(execucoes)
        df_exec = df_exec.rename(columns={
            "id":"ID","usuario_nome":"Usuário","inicio":"Início","fim":"Fim",
            "total":"Total","encontrados":"Encontrados",
            "nao_encontrados":"Não enc.","drive_raiz":"Drive","pasta_destino":"Destino"
        })
        df_exec = df_exec[["ID","Usuário","Início","Fim","Total","Encontrados","Não enc."]]

        st.dataframe(df_exec, use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("Detalhes de uma execução")

        ids = [e["id"] for e in execucoes]
        exec_sel = st.selectbox("Selecione o ID da execução", ids, format_func=lambda x: f"#{x}")

        if exec_sel:
            res = resultados_da_execucao(exec_sel)
            if res:
                df_res = pd.DataFrame(res).rename(columns={
                    "nome":"Nome","data":"Data","encontrado":"Encontrado",
                    "arquivo":"Arquivo","erro":"Observação","score_fuzzy":"Score"
                })

                def colorir_hist(row):
                    cor = "#d4edda" if row["Encontrado"] else "#f8d7da"
                    return [f"background-color:{cor}"] * len(row)

                st.dataframe(df_res[["Nome","Data","Encontrado","Score","Arquivo","Observação"]]
                             .style.apply(colorir_hist, axis=1),
                             use_container_width=True, hide_index=True)

                csv_hist = df_res.to_csv(index=False).encode("utf-8-sig")
                st.download_button("⬇️ Exportar execução CSV", csv_hist,
                                   f"execucao_{exec_sel}.csv", "text/csv")
            else:
                st.info("Nenhum resultado detalhado para esta execução.")


# ============================================================
# ABA 3 — ESTATÍSTICAS
# ============================================================

with aba_stats:
    st.subheader("Estatísticas de uso")

    try:
        import plotly.graph_objects as go
        import plotly.express as px
        PLOTLY = True
    except ImportError:
        PLOTLY = False
        st.warning("Instale plotly para gráficos: pip install plotly")

    stats = estatisticas_gerais()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total de execuções",  stats["total_execucoes"])
    c2.metric("Total de buscas",     stats["total_buscas"])
    c3.metric("Total encontrados",   stats["total_encontrados"])
    c4.metric("Taxa média de sucesso", f"{stats['taxa_sucesso']}%")

    if PLOTLY and stats["por_dia"]:
        st.divider()
        df_dia = pd.DataFrame(stats["por_dia"])

        fig1 = px.bar(df_dia, x="dia", y=["encontrados","total"],
                      barmode="group",
                      labels={"dia":"Data","value":"Quantidade","variable":"Tipo"},
                      color_discrete_map={"encontrados":"#1D9E75","total":"#B5D4F4"},
                      title="Buscas por dia — últimos 30 dias")
        fig1.update_layout(plot_bgcolor="rgba(0,0,0,0)",
                           paper_bgcolor="rgba(0,0,0,0)",
                           legend_title_text="")
        st.plotly_chart(fig1, use_container_width=True)

    if PLOTLY and stats["top_usuarios"]:
        st.divider()
        df_usr = pd.DataFrame(stats["top_usuarios"])
        fig2 = px.bar(df_usr, x="nome", y="execucoes",
                      labels={"nome":"Usuário","execucoes":"Execuções"},
                      color="execucoes", color_continuous_scale="Teal",
                      title="Top usuários mais ativos")
        fig2.update_layout(plot_bgcolor="rgba(0,0,0,0)",
                           paper_bgcolor="rgba(0,0,0,0)",
                           showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)

    if not stats["por_dia"] and not stats["top_usuarios"]:
        st.info("Nenhuma execução registrada ainda. As estatísticas aparecerão após a primeira busca.")


# ============================================================
# ABA 4 — CONFIGURAÇÕES
# ============================================================

with aba_config:
    st.subheader("Configurações")

    # --- Minha conta ---
    st.markdown("#### Minha conta")
    with st.expander("Alterar senha"):
        s1 = st.text_input("Senha atual",   type="password", key="s_atual")
        s2 = st.text_input("Nova senha",    type="password", key="s_nova")
        s3 = st.text_input("Confirmar",     type="password", key="s_conf")
        if st.button("Salvar senha"):
            from database import verificar_login as vl
            if not vl(usuario.get("nome",""), s1):
                st.error("Senha atual incorreta.")
            elif s2 != s3:
                st.error("As senhas não coincidem.")
            elif len(s2) < 6:
                st.error("A nova senha deve ter ao menos 6 caracteres.")
            else:
                alterar_senha(usuario["id"], s2)
                st.success("Senha alterada com sucesso!")

    # --- E-mail SMTP ---
    st.markdown("#### Configurações de e-mail (SMTP)")
    with st.expander("Configurar servidor SMTP"):
        smtp = cfg.get("smtp", {})
        smtp_host   = st.text_input("Servidor SMTP", value=smtp.get("host",""),
                                    placeholder="smtp.gmail.com")
        smtp_porta  = st.number_input("Porta", value=int(smtp.get("porta",587)),
                                      min_value=1, max_value=65535, step=1)
        smtp_user   = st.text_input("Usuário/E-mail", value=smtp.get("usuario",""))
        smtp_senha  = st.text_input("Senha do e-mail", type="password",
                                    value=smtp.get("senha",""))
        smtp_rem    = st.text_input("Nome/e-mail remetente",
                                    value=smtp.get("remetente", smtp.get("usuario","")))

        st.caption(
            "Para Gmail: use smtp.gmail.com porta 587 e crie uma "
            "[Senha de App](https://myaccount.google.com/apppasswords) "
            "(não use sua senha normal do Gmail)."
        )

        if st.button("Salvar configurações de e-mail"):
            cfg["smtp"] = {
                "host": smtp_host, "porta": smtp_porta,
                "usuario": smtp_user, "senha": smtp_senha,
                "remetente": smtp_rem or smtp_user
            }
            salvar_config(cfg)
            st.success("Configurações de e-mail salvas.")

        if st.button("Testar envio de e-mail"):
            if not smtp_host or not smtp_user or not smtp_senha:
                st.error("Preencha os campos antes de testar.")
            else:
                from mailer import enviar_email as _env
                ok_t, err_t = _env(
                    smtp_user,
                    "Teste do sistema de busca de documentos",
                    "<p>Se você recebeu este e-mail, as configurações SMTP estão corretas.</p>"
                )
                if ok_t: st.success("E-mail de teste enviado com sucesso!")
                else:    st.error(f"Falha no envio: {err_t}")

    # --- Gerenciar usuários (somente admin) ---
    if usuario["admin"]:
        st.markdown("#### Gerenciar usuários")
        with st.expander("Usuários cadastrados"):
            usuarios = listar_usuarios()
            df_u = pd.DataFrame(usuarios)[["id","usuario","nome","email","admin","ativo","criado_em"]]
            df_u.columns = ["ID","Login","Nome","E-mail","Admin","Ativo","Criado em"]
            st.dataframe(df_u, use_container_width=True, hide_index=True)

        with st.expander("Criar novo usuário"):
            nu = st.text_input("Login do novo usuário")
            nn = st.text_input("Nome completo")
            ne = st.text_input("E-mail (opcional)")
            ns = st.text_input("Senha", type="password")
            na = st.checkbox("Administrador")
            if st.button("Criar usuário"):
                if not nu or not nn or not ns:
                    st.error("Preencha login, nome e senha.")
                elif len(ns) < 6:
                    st.error("Senha deve ter ao menos 6 caracteres.")
                else:
                    if criar_usuario(nu, ns, nn, ne, int(na)):
                        st.success(f"Usuário '{nu}' criado com sucesso.")
                    else:
                        st.error(f"Já existe um usuário com o login '{nu}'.")

    # --- Informações do sistema ---
    st.markdown("#### Informações do sistema")
    with st.expander("Diagnóstico"):
        import platform as _pl
        st.code(
            f"Sistema:    {_pl.system()} {_pl.release()}\n"
            f"Python:     {_pl.python_version()}\n"
            f"Streamlit:  {st.__version__}\n"
            f"Tesseract:  {'OK — ' + shutil.which('tesseract') if shutil.which('tesseract') else 'não encontrado'}\n"
            f"Banco:      {Path('clinica.db').resolve()}\n"
            f"Config:     {CONFIG_PATH.resolve()}\n",
            language=None
        )
```

---

## PARTE 10 — `README.md`

Crie com este conteúdo:

````markdown
# Automação de Busca de Documentos — Clínica v2

Sistema completo com busca inteligente em PDFs, histórico, estatísticas e envio de e-mail.

## Instalação

```bash
pip install -r requirements.txt
python setup.py          # instala Tesseract automaticamente
streamlit run app.py
```

## Acesso pela rede da clínica

```bash
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```
Acesse de qualquer PC da rede: `http://<IP-DO-SERVIDOR>:8501`

**IP no Windows:** abra o terminal → `ipconfig` → procure "Endereço IPv4"

## Primeiro acesso

- Usuário: `admin`
- Senha: `admin123`
- **Altere a senha imediatamente** em Configurações → Minha conta

## Estrutura de pastas esperada no drive

```
<raiz>/
├── marco/
│   ├── 02/
│   │   └── procedimentos.pdf
│   └── 15/
│       └── atendimentos.pdf
└── abril/
    └── ...
```

## Personalização dos nomes de mês

Se as pastas usam nomes diferentes, edite o dicionário `MESES` no início do `processor.py`:

```python
MESES = {
    "01": "janeiro",
    "02": "fevereiro",
    "03": "marco",   # ← altere aqui se necessário (ex: "Março", "03", "march")
    ...
}
```

## Configurar e-mail automático (Gmail)

1. Acesse: Configurações → Configurações de e-mail
2. Servidor: `smtp.gmail.com`, Porta: `587`
3. Crie uma **Senha de App** em: myaccount.google.com/apppasswords
4. Use a Senha de App (não sua senha normal do Gmail)
5. Adicione uma coluna `email` na planilha para envio por paciente

## Solução de problemas

| Problema | Solução |
|---|---|
| "Pasta não encontrada" | Verifique o caminho raiz e os nomes dos meses no `processor.py` |
| "Nome não encontrado" | Reduza a sensibilidade fuzzy (ex: 70) ou instale o Tesseract (`python setup.py`) |
| Falsos positivos | Aumente a sensibilidade fuzzy (ex: 90 ou 100) |
| Tesseract não instala | Execute `python setup.py` como administrador |
| E-mail não envia | Verifique as credenciais e use Senha de App para Gmail |
| Encoding errado no CSV | No Excel: Dados → De texto/CSV → Codificação UTF-8 |

## Arquivos gerados automaticamente

| Arquivo | Conteúdo |
|---|---|
| `clinica.db` | Banco SQLite com histórico completo |
| `config.json` | Preferências salvas (caminhos, SMTP, threshold) |
| `config_tesseract.py` | Caminho do Tesseract (somente Windows) |
| `output/` | PDFs extraídos (se não especificar outro destino) |
````

---

## PARTE 11 — ORDEM DE CRIAÇÃO E CHECKLIST FINAL

### Ordem obrigatória

1. `requirements.txt`
2. `setup.py`
3. `database.py`
4. `auth.py`
5. `mailer.py`
6. `processor.py`
7. `app.py`
8. `README.md`

### Checklist — verifique cada item antes de concluir

**setup.py**
- [ ] Contém `detectar_sistema`, `tesseract_instalado`, `tem_admin`, `pedir_admin`
- [ ] Contém `instalar_windows`, `instalar_linux`, `instalar_macos`
- [ ] Contém `_path_windows` e `_config_pytesseract`
- [ ] Contém `verificar_deps` e `verificar_final`

**database.py**
- [ ] Contém as 3 tabelas: `usuarios`, `execucoes`, `resultados`
- [ ] Contém `inicializar_banco` com criação do admin padrão
- [ ] Contém `verificar_login`, `criar_usuario`, `alterar_senha`
- [ ] Contém `iniciar_execucao`, `finalizar_execucao`, `salvar_resultado`
- [ ] Contém `listar_execucoes`, `resultados_da_execucao`, `estatisticas_gerais`

**auth.py**
- [ ] Contém `tela_login`, `esta_logado`, `fazer_login`, `fazer_logout`
- [ ] `tela_login` chama `st.stop()` se não estiver logado

**mailer.py**
- [ ] Contém `enviar_email`, `email_configurado`, `enviar_relatorio_execucao`
- [ ] Usa configurações do `config.json`

**processor.py**
- [ ] Contém `nome_contem` com busca dupla (exata + fuzzy)
- [ ] Contém `ler_planilha` com detecção de coluna `email`
- [ ] `processar_lista` aceita parâmetro `threshold_fuzzy`
- [ ] `processar_lista` retorna `score_fuzzy` em cada resultado

**app.py**
- [ ] `set_page_config` é o PRIMEIRO comando Streamlit do arquivo
- [ ] Contém 4 abas: Busca, Histórico, Estatísticas, Configurações
- [ ] Aba Busca salva resultados no banco via `salvar_resultado`
- [ ] Aba Busca envia e-mails automáticos se configurado
- [ ] Aba Histórico permite selecionar execução e ver detalhes
- [ ] Aba Estatísticas usa Plotly para gráficos
- [ ] Aba Configurações tem: alterar senha, SMTP, gerenciar usuários (admin), diagnóstico

---

## PARTE 12 — GUIA DE AJUSTES PARA O DONO DO SISTEMA

### Ajuste 1 — Estrutura de pastas diferente

Se as pastas no drive não seguem o padrão `nome_do_mes/dia`, edite o dicionário `MESES` no `processor.py`:

```python
# Exemplo: pastas nomeadas como "03", "04" etc. (número puro)
MESES = {
    "01": "01", "02": "02", "03": "03", "04": "04",
    "05": "05", "06": "06", "07": "07", "08": "08",
    "09": "09", "10": "10", "11": "11", "12": "12",
}

# Exemplo: pastas com acento ("Março", "Abril"...)
MESES = {
    "01": "Janeiro",  "02": "Fevereiro", "03": "Março",
    "04": "Abril",    "05": "Maio",      "06": "Junho",
    "07": "Julho",    "08": "Agosto",    "09": "Setembro",
    "10": "Outubro",  "11": "Novembro",  "12": "Dezembro",
}
```

Se a estrutura for `ano/mes/dia` (ex: `2024/03/02`), substitua a função `montar_caminho_pasta` no `processor.py`:

```python
def montar_caminho_pasta(drive_raiz: str, data) -> Path:
    dia, mes_num = extrair_dia_mes(data)
    # Descomentar e ajustar conforme necessário:
    # return Path(drive_raiz) / "2024" / mes_num / dia   # ano fixo
    # return Path(drive_raiz) / ano / mes_num / dia      # se tiver coluna de ano
    return Path(drive_raiz) / MESES.get(mes_num, mes_num) / dia  # padrão atual
```

### Ajuste 2 — Sensibilidade da busca fuzzy

Na interface, o slider "Sensibilidade da busca fuzzy" controla o threshold:
- **100** = apenas correspondência exata (mais seguro, pode perder nomes com OCR ruim)
- **80** = padrão recomendado (tolera pequenos erros)
- **60** = muito permissivo (pode gerar falsos positivos)

### Ajuste 3 — Porta da rede diferente

Se a porta 8501 estiver ocupada:
```bash
streamlit run app.py --server.address 0.0.0.0 --server.port 8888
```

### Ajuste 4 — Iniciar o sistema automaticamente no Windows

Crie um arquivo `iniciar.bat` na pasta do projeto:
```bat
@echo off
cd /d %~dp0
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
pause
```
Coloque um atalho deste arquivo na pasta Inicialização do Windows para rodar automaticamente.

---

## PARTE 13 — RESTRIÇÕES ABSOLUTAS

Não faça nada da lista abaixo:

- Substituir `PyMuPDF (fitz)` por `PyPDF2`, `pypdf` ou `pdfminer`
- Usar `glob.glob()` como string em vez de `Path.glob()`
- Colocar `set_page_config` em qualquer posição que não seja a **primeira linha** de código Streamlit do `app.py`
- Mudar os nomes das funções em `processor.py`, `database.py`, `auth.py` ou `mailer.py`
- Usar `threading` ou `asyncio`
- Criar banco de dados que não seja SQLite (não use PostgreSQL, MySQL, MongoDB)
- Usar caminhos de arquivo hardcoded — todos vêm do `config.json` ou da interface
- Colocar senhas ou credenciais diretamente no código — ficam no `config.json`
- Usar `st.experimental_rerun()` — use `st.rerun()` (API atual do Streamlit)

---

*PROMPT MASTER V2 — Sistema Completo com Histórico, Login, E-mail, Fuzzy e Estatísticas.*
