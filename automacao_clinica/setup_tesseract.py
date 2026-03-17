"""
setup_tesseract.py
==================
Instalador autônomo do Tesseract OCR com 4 estratégias em cascata.
Se uma falhar, tenta a próxima automaticamente.

Estratégias (Windows):
  1. winget  — gerenciador nativo do Windows 10/11
  2. choco   — Chocolatey (se instalado)
  3. Download direto do instalador UB-Mannheim + execução silenciosa
  4. Download do binário portátil (sem instalador, sem admin)

Estratégias (Linux):
  1. apt-get / dnf / yum / pacman / zypper

Estratégias (macOS):
  1. brew

Execute:
    python setup_tesseract.py

Ou com privilégios explícitos no Linux/Mac:
    sudo python setup_tesseract.py
"""

import sys as _sys
import io as _io

# Força UTF-8 no terminal Windows (evita UnicodeEncodeError com cp1252)
if hasattr(_sys.stdout, 'buffer') and _sys.stdout.encoding.lower() not in ('utf-8','utf8'):
    _sys.stdout = _io.TextIOWrapper(_sys.stdout.buffer, encoding='utf-8', errors='replace')
if hasattr(_sys.stderr, 'buffer') and _sys.stderr.encoding.lower() not in ('utf-8','utf8'):
    _sys.stderr = _io.TextIOWrapper(_sys.stderr.buffer, encoding='utf-8', errors='replace')


import os
import sys
import platform
import subprocess
import shutil
import tempfile
import ctypes
import time
from pathlib import Path

# ---------------------------------------------------------------------------
# URLs e caminhos
# ---------------------------------------------------------------------------

# Instalador completo — requer admin
WINDOWS_INSTALLER_URL = (
    "https://github.com/UB-Mannheim/tesseract/releases/download/"
    "v5.3.3.20231005/tesseract-ocr-w64-setup-5.3.3.20231005.exe"
)

# Portátil — NÃO requer admin, extrai direto numa pasta local
WINDOWS_PORTABLE_URL = (
    "https://github.com/UB-Mannheim/tesseract/releases/download/"
    "v5.3.3.20231005/tesseract-ocr-w64-setup-5.3.3.20231005.exe"
)

WINDOWS_DEFAULT_PATHS = [
    Path(r"C:\Program Files\Tesseract-OCR"),
    Path(r"C:\Program Files (x86)\Tesseract-OCR"),
    Path(os.environ.get("LOCALAPPDATA", "C:\\Users\\Default\\AppData\\Local")) / "Tesseract-OCR",
    Path(os.environ.get("APPDATA",      "C:\\Users\\Default\\AppData\\Roaming")) / "Tesseract-OCR",
]

# Pasta portátil — sem admin necessário
PASTA_PORTATIL = Path(__file__).parent / "tesseract_bin"

# ---------------------------------------------------------------------------
# Cores no terminal
# ---------------------------------------------------------------------------

G = "\033[92m"   # verde
Y = "\033[93m"   # amarelo
R = "\033[91m"   # vermelho
B = "\033[94m"   # azul
N = "\033[1m"    # negrito
X = "\033[0m"    # reset

def ok(m):    print(f"{G}  [OK]{X}    {m}")
def info(m):  print(f"{B}  [INFO]{X}  {m}")
def aviso(m): print(f"{Y}  [AVISO]{X} {m}")
def erro(m):  print(f"{R}  [ERRO]{X}  {m}")
def titulo(m):
    print(f"\n{N}{B}{'-'*52}{X}")
    print(f"{N}{B}  {m}{X}")
    print(f"{N}{B}{'-'*52}{X}\n")
def tentativa(n, total, m):
    print(f"\n{N}  [{n}/{total}] {m}...{X}")


# ---------------------------------------------------------------------------
# Utilitários gerais
# ---------------------------------------------------------------------------

def eh_windows():  return platform.system().lower() == "windows"
def eh_linux():    return platform.system().lower() == "linux"
def eh_macos():    return platform.system().lower() == "darwin"

def tem_admin() -> bool:
    if eh_windows():
        try: return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except: return False
    return os.geteuid() == 0

def run(cmd: list, timeout=120, check=False) -> subprocess.CompletedProcess:
    """Executa um comando e retorna o resultado sem levantar exceção."""
    try:
        return subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=check
        )
    except subprocess.TimeoutExpired:
        aviso(f"Timeout ao executar: {' '.join(cmd)}")
        return subprocess.CompletedProcess(cmd, 1, "", "timeout")
    except FileNotFoundError:
        return subprocess.CompletedProcess(cmd, 1, "", "comando não encontrado")
    except Exception as e:
        return subprocess.CompletedProcess(cmd, 1, "", str(e))

def tesseract_executavel() -> str | None:
    """
    Procura o executável do Tesseract em todas as localizações conhecidas.
    Retorna o caminho completo ou None.
    """
    # 1. PATH do sistema
    exe = shutil.which("tesseract")
    if exe:
        return exe

    if eh_windows():
        # 2. Caminhos padrão do Windows
        for base in WINDOWS_DEFAULT_PATHS:
            t = base / "tesseract.exe"
            if t.exists():
                return str(t)

        # 3. Pasta portátil local do projeto
        t = PASTA_PORTATIL / "tesseract.exe"
        if t.exists():
            return str(t)

        # 4. Busca no registro do Windows
        try:
            import winreg
            chave = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                r"SOFTWARE\Tesseract-OCR"
            )
            pasta, _ = winreg.QueryValueEx(chave, "InstallDir")
            t = Path(pasta) / "tesseract.exe"
            if t.exists():
                return str(t)
        except Exception:
            pass

    return None

def tesseract_ok() -> tuple[bool, str]:
    """Verifica se o Tesseract está funcional. Retorna (ok, versão_ou_erro)."""
    exe = tesseract_executavel()
    if not exe:
        return False, "executável não encontrado"
    r = run([exe, "--version"], timeout=10)
    if r.returncode == 0:
        versao = (r.stdout or r.stderr).splitlines()[0].strip()
        return True, versao
    return False, f"erro ao executar: {r.stderr.strip()}"

def configurar_pytesseract(caminho_exe: str):
    """
    Cria config_tesseract.py apontando para o executável encontrado.
    Também adiciona o diretório ao PATH da sessão atual.
    """
    pasta_exe = str(Path(caminho_exe).parent)

    # Adiciona ao PATH da sessão atual imediatamente
    if pasta_exe not in os.environ.get("PATH", ""):
        os.environ["PATH"] = pasta_exe + os.pathsep + os.environ.get("PATH", "")

    # Cria arquivo de configuração para o pytesseract
    config_path = Path(__file__).parent / "config_tesseract.py"
    config_path.write_text(
        f'# Gerado automaticamente por setup_tesseract.py\n'
        f'# Não edite manualmente.\n'
        f'import pytesseract\n'
        f'pytesseract.pytesseract.tesseract_cmd = r"{caminho_exe}"\n',
        encoding="utf-8"
    )
    ok(f"config_tesseract.py criado -> {caminho_exe}")


# ---------------------------------------------------------------------------
# ESTRATÉGIAS WINDOWS
# ---------------------------------------------------------------------------

def estrategia_winget() -> bool:
    """
    Estratégia 1 — winget (Windows Package Manager nativo do Win10/11).
    Não requer admin se a instalação for para o usuário atual.
    """
    if not eh_windows():
        return False
    if not shutil.which("winget"):
        aviso("winget não disponível neste sistema.")
        return False

    info("Tentando instalar via winget...")
    r = run(
        ["winget", "install", "--id", "UB-Mannheim.TesseractOCR",
         "--silent", "--accept-package-agreements",
         "--accept-source-agreements"],
        timeout=180
    )
    if r.returncode == 0:
        ok("Instalado via winget.")
        time.sleep(2)  # aguarda o PATH ser atualizado pelo winget
        return True
    erro(f"winget falhou (código {r.returncode}): {r.stderr.strip()[:200]}")
    return False


def estrategia_chocolatey() -> bool:
    """
    Estratégia 2 — Chocolatey.
    Requer admin e Chocolatey instalado.
    """
    if not eh_windows():
        return False
    if not shutil.which("choco"):
        aviso("Chocolatey não está instalado.")
        return False
    if not tem_admin():
        aviso("Chocolatey requer admin — pulando.")
        return False

    info("Tentando instalar via Chocolatey...")
    r = run(["choco", "install", "tesseract", "-y", "--no-progress"], timeout=180)
    if r.returncode == 0:
        ok("Instalado via Chocolatey.")
        return True
    erro(f"choco falhou: {r.stderr.strip()[:200]}")
    return False


def estrategia_instalador_direto() -> bool:
    """
    Estratégia 3 — Download do instalador UB-Mannheim + execução silenciosa.
    Requer admin.
    """
    if not eh_windows():
        return False
    if not tem_admin():
        aviso("Instalador direto requer privilégios de admin — pulando.")
        return False

    info(f"Baixando instalador de:\n    {WINDOWS_INSTALLER_URL}")

    import urllib.request
    import urllib.error

    with tempfile.TemporaryDirectory() as tmp:
        dest = Path(tmp) / "tess_setup.exe"

        try:
            def progresso(blocos, tam_bloco, total):
                if total > 0:
                    pct = min(100, int(blocos * tam_bloco * 100 / total))
                    mb  = blocos * tam_bloco / 1024 / 1024
                    print(f"\r  Baixando... {pct}% ({mb:.1f} MB)", end="", flush=True)

            urllib.request.urlretrieve(WINDOWS_INSTALLER_URL, dest, progresso)
            print()
            ok("Download completo.")
        except urllib.error.URLError as e:
            erro(f"Falha no download: {e}")
            return False
        except Exception as e:
            erro(f"Erro inesperado no download: {e}")
            return False

        install_dir = Path(r"C:\Program Files\Tesseract-OCR")
        info(f"Instalando em {install_dir} (modo silencioso)...")

        r = run(
            [str(dest), "/S", f"/D={install_dir}"],
            timeout=300
        )

        # 0 = sucesso, 3010 = requer reinício (mas instalado)
        if r.returncode not in (0, 3010):
            erro(f"Instalador retornou código {r.returncode}")
            if r.stderr: erro(r.stderr.strip()[:300])
            return False

    ok(f"Instalado em {install_dir}")
    _adicionar_path_windows_permanente(install_dir)
    return True


def _adicionar_path_windows_permanente(pasta: Path):
    """Adiciona pasta ao PATH permanente do sistema via PowerShell + registro."""
    info("Adicionando ao PATH permanente do sistema...")

    ps_ler = '[System.Environment]::GetEnvironmentVariable("Path","Machine")'
    r = run(["powershell", "-Command", ps_ler])
    path_atual = r.stdout.strip()

    if str(pasta) not in path_atual:
        novo = f"{path_atual};{pasta}"
        ps_gravar = (
            f'[System.Environment]::SetEnvironmentVariable('
            f'"Path","{novo}","Machine")'
        )
        run(["powershell", "-Command", ps_gravar])
        ok(f"PATH do sistema atualizado: {pasta}")
    else:
        info("Já estava no PATH do sistema.")

    # Atualiza PATH da sessão atual também
    os.environ["PATH"] = str(pasta) + os.pathsep + os.environ.get("PATH", "")


def estrategia_portatil() -> bool:
    """
    Estratégia 4 — Instalação portátil sem admin.
    Extrai os arquivos do Tesseract diretamente na pasta do projeto.
    Usa o próprio instalador NSIS em modo de extração silenciosa.
    NÃO modifica o registro nem o PATH do sistema.
    O pytesseract é configurado via config_tesseract.py.
    """
    if not eh_windows():
        return False

    info("Tentando instalação portátil (sem admin)...")
    info(f"Destino: {PASTA_PORTATIL}")

    import urllib.request
    import urllib.error

    PASTA_PORTATIL.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmp:
        dest = Path(tmp) / "tess_setup.exe"

        try:
            def progresso(blocos, tam_bloco, total):
                if total > 0:
                    pct = min(100, int(blocos * tam_bloco * 100 / total))
                    mb  = blocos * tam_bloco / 1024 / 1024
                    print(f"\r  Baixando... {pct}% ({mb:.1f} MB)", end="", flush=True)

            urllib.request.urlretrieve(WINDOWS_PORTABLE_URL, dest, progresso)
            print()
            ok("Download completo.")
        except urllib.error.URLError as e:
            erro(f"Falha no download: {e}")
            return False

        # NSIS aceita /D= para extração sem instalar
        info("Extraindo arquivos portáteis...")
        r = run(
            [str(dest), "/S", f"/D={PASTA_PORTATIL}"],
            timeout=300
        )
        if r.returncode not in (0, 3010):
            # Tenta extração alternativa com 7zip se disponível
            sz = shutil.which("7z") or shutil.which("7za")
            if sz:
                info("Tentando extração com 7-Zip...")
                r2 = run([sz, "x", str(dest), f"-o{PASTA_PORTATIL}", "-y"], timeout=120)
                if r2.returncode != 0:
                    erro("Extração falhou.")
                    return False
            else:
                erro(f"Extração falhou (código {r.returncode})")
                return False

    # Verifica se o executável foi criado
    tess_exe = PASTA_PORTATIL / "tesseract.exe"
    if not tess_exe.exists():
        # Busca recursiva caso tenha sido extraído numa subpasta
        encontrados = list(PASTA_PORTATIL.rglob("tesseract.exe"))
        if encontrados:
            tess_exe = encontrados[0]
        else:
            erro("tesseract.exe não encontrado após extração.")
            return False

    ok(f"Tesseract portátil disponível em: {tess_exe}")
    return True


# ---------------------------------------------------------------------------
# ESTRATÉGIAS LINUX
# ---------------------------------------------------------------------------

def estrategia_linux() -> bool:
    """Detecta o gerenciador e instala tesseract + idioma português."""
    if not eh_linux():
        return False

    gerenciadores = {
        "apt-get": [
            ["apt-get", "update", "-qq"],
            ["apt-get", "install", "-y", "tesseract-ocr", "tesseract-ocr-por"],
        ],
        "dnf": [["dnf", "install", "-y", "tesseract", "tesseract-langpack-por"]],
        "yum": [["yum", "install", "-y", "tesseract"]],
        "pacman": [["pacman", "-Sy", "--noconfirm", "tesseract", "tesseract-data-por"]],
        "zypper": [
            ["zypper", "install", "-y", "tesseract-ocr",
             "tesseract-ocr-traineddata-portuguese"]
        ],
    }

    for mgr, cmds in gerenciadores.items():
        if not shutil.which(mgr):
            continue
        info(f"Gerenciador detectado: {mgr}")
        for cmd in cmds:
            info(f"Executando: {' '.join(cmd)}")
            r = run(cmd, timeout=180)
            if r.returncode != 0:
                erro(f"Falhou: {r.stderr.strip()[:200]}")
                return False
        ok("Tesseract instalado no Linux.")
        return True

    erro("Nenhum gerenciador de pacotes reconhecido encontrado.")
    return False


# ---------------------------------------------------------------------------
# ESTRATÉGIAS MACOS
# ---------------------------------------------------------------------------

def estrategia_macos() -> bool:
    if not eh_macos():
        return False

    if not shutil.which("brew"):
        info("Homebrew não encontrado. Instalando...")
        os.system(
            '/bin/bash -c "$(curl -fsSL '
            'https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"'
        )
        if not shutil.which("brew"):
            erro("Falha ao instalar Homebrew. Acesse: https://brew.sh")
            return False
        ok("Homebrew instalado.")

    info("Instalando tesseract via Homebrew...")
    r1 = run(["brew", "install", "tesseract"], timeout=300)
    r2 = run(["brew", "install", "tesseract-lang"], timeout=300)

    if r1.returncode == 0:
        ok("Tesseract instalado no macOS.")
        return True
    erro(f"brew falhou: {r1.stderr.strip()[:200]}")
    return False


# ---------------------------------------------------------------------------
# VERIFICAÇÃO FINAL
# ---------------------------------------------------------------------------

def verificar_e_configurar() -> bool:
    """
    Localiza o executável, configura o pytesseract e verifica se está funcional.
    """
    titulo("Verificação e configuração final")

    exe = tesseract_executavel()
    if not exe:
        erro("Executável não encontrado em nenhum local conhecido.")
        return False

    ok(f"Executável encontrado: {exe}")
    configurar_pytesseract(exe)

    # Teste funcional
    r = run([exe, "--version"], timeout=10)
    versao = (r.stdout or r.stderr).splitlines()[0].strip() if r.returncode == 0 else ""
    if r.returncode == 0:
        ok(f"Versão: {versao}")
    else:
        erro("Tesseract encontrado mas não executou corretamente.")
        return False

    # Verifica idiomas
    r2 = run([exe, "--list-langs"], timeout=10)
    langs = r2.stdout + r2.stderr
    if "por" in langs:
        ok("Idioma português (por) disponível.")
    else:
        aviso("Idioma português não encontrado.")
        aviso("Para adicionar: baixe 'por.traineddata' de")
        aviso("https://github.com/tesseract-ocr/tessdata/raw/main/por.traineddata")
        aviso(f"e coloque em: {Path(exe).parent / 'tessdata'}")

    # Testa o pytesseract
    try:
        import pytesseract
        cfg = Path(__file__).parent / "config_tesseract.py"
        if cfg.exists():
            exec(cfg.read_text(encoding="utf-8"), {})
        pytesseract.get_tesseract_version()
        ok("pytesseract comunicando com o Tesseract corretamente.")
    except ImportError:
        aviso("pytesseract não instalado. Execute: pip install pytesseract")
    except Exception as e:
        aviso(f"pytesseract não conseguiu comunicar: {e}")
        aviso("Execute: pip install pytesseract  e reinicie o terminal.")

    return True


# ---------------------------------------------------------------------------
# ORQUESTRADOR PRINCIPAL
# ---------------------------------------------------------------------------

def instalar_windows_cascata() -> bool:
    """
    Tenta instalar no Windows em ordem de prioridade.
    Para na primeira estratégia bem-sucedida.
    """
    estrategias = [
        (1, "winget (gerenciador nativo)",          estrategia_winget),
        (2, "Chocolatey",                           estrategia_chocolatey),
        (3, "Instalador direto UB-Mannheim (admin)",estrategia_instalador_direto),
        (4, "Instalação portátil (sem admin)",      estrategia_portatil),
    ]

    for num, nome, func in estrategias:
        tentativa(num, len(estrategias), nome)
        try:
            if func():
                time.sleep(1)  # pequena pausa para o SO registrar os arquivos
                exe = tesseract_executavel()
                if exe:
                    ok(f"Tesseract encontrado após estratégia {num}: {exe}")
                    return True
                else:
                    aviso(f"Estratégia {num} reportou sucesso mas executável não encontrado. Tentando próxima...")
        except Exception as e:
            erro(f"Exceção na estratégia {num}: {e}")

    return False


def elevar_para_admin():
    """
    Reabre este mesmo script com privilégios de administrador via UAC.
    O processo atual encerra; o novo processo elevado assume.
    Só funciona no Windows.
    """
    import ctypes
    script = os.path.abspath(sys.argv[0])
    parametros = " ".join(f'"{a}"' for a in sys.argv[1:])
    print("[INFO] Solicitando privilégios de administrador via UAC...")
    print("[INFO] Uma janela de confirmação deve aparecer.")
    # ShellExecuteW com "runas" abre o UAC e relança o script como admin
    resultado = ctypes.windll.shell32.ShellExecuteW(
        None,        # hwnd
        "runas",     # operação — força UAC
        sys.executable,  # programa: o próprio Python
        f'"{script}" {parametros}',  # argumentos
        None,        # diretório
        1            # SW_SHOWNORMAL — abre janela visível
    )
    # ShellExecuteW retorna > 32 em caso de sucesso
    if resultado <= 32:
        print(f"[ERRO] Falha ao solicitar admin (código {resultado}).")
        print("[ERRO] Execute o script manualmente como Administrador:")
        print(f"       Clique com botão direito em setup_tesseract.py -> Executar como administrador")
        sys.exit(1)
    sys.exit(0)  # encerra o processo sem admin; o elevado continua


def main():
    titulo("Instalador Autônomo do Tesseract OCR — v3")
    info(f"Sistema: {platform.system()} {platform.release()}")
    info(f"Python:  {platform.python_version()}")
    info(f"Admin:   {'Sim' if tem_admin() else 'Nao'}")

    # -- Elevação automática para admin (Windows) --------------------------------
    # Verifica ANTES de qualquer outra coisa. Se não tiver admin no Windows,
    # reabre o script via UAC automaticamente — sem precisar do usuário fazer nada.
    if eh_windows() and not tem_admin():
        aviso("Este script requer privilégios de administrador.")
        aviso("Solicitando elevação via UAC automaticamente...")
        elevar_para_admin()
        # Se elevar_para_admin() retornar (não deveria), para aqui
        return

    # -- Já está instalado? --------------------------------------------------
    titulo("Verificando instalação existente")
    esta_ok, versao_ou_erro = tesseract_ok()
    if esta_ok:
        ok(f"Tesseract já está instalado e funcional: {versao_ou_erro}")
        verificar_e_configurar()
        titulo("Nada a fazer — sistema já está pronto")
        print(f"  {N}{G}streamlit run app.py{X}\n")
        return

    aviso(f"Tesseract não encontrado: {versao_ou_erro}")

    # -- Instala conforme o SO -----------------------------------------------
    titulo("Iniciando instalação")

    sucesso = False

    if eh_windows():
        sucesso = instalar_windows_cascata()
    elif eh_linux():
        tentativa(1, 1, "Gerenciador de pacotes Linux")
        sucesso = estrategia_linux()
    elif eh_macos():
        tentativa(1, 1, "Homebrew (macOS)")
        sucesso = estrategia_macos()
    else:
        erro(f"Sistema operacional não suportado: {platform.system()}")
        sys.exit(1)

    # -- Resultado -----------------------------------------------------------
    if not sucesso:
        titulo("Todas as estratégias automáticas falharam")
        print(f"{Y}  Instale manualmente:{X}\n")
        if eh_windows():
            print(f"  1. Acesse: https://github.com/UB-Mannheim/tesseract/wiki")
            print(f"  2. Baixe o instalador para Windows 64-bit")
            print(f"  3. Execute como administrador")
            print(f"  4. Marque 'Add to PATH' durante a instalacao")
            print(f"  5. Reinicie o terminal e execute este script novamente")
        elif eh_linux():
            print(f"  sudo apt install tesseract-ocr tesseract-ocr-por")
        elif eh_macos():
            print(f"  brew install tesseract tesseract-lang")
        print()
        sys.exit(1)

    # -- Configura e valida --------------------------------------------------
    if not verificar_e_configurar():
        aviso("Instalação concluída mas validação falhou.")
        aviso("Feche o terminal, abra um novo e execute novamente.")
        sys.exit(1)

    titulo("Instalação concluída com sucesso")
    print(f"  {N}{G}streamlit run app.py{X}\n")


if __name__ == "__main__":
    main()