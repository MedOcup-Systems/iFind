"""
launcher.py v3 — iFind Clinica
================================
Tela de loading + WebView embutido (sem precisar de Chrome/Edge aberto).

CORRECOES v3:
  - webview.start() agora roda na thread PRINCIPAL (obrigatorio no Windows)
  - Splash nao e destruida antes do webview abrir (evitava abertura)
  - Fallback para browser tambem esta na thread principal
  - Thread de background apenas sinaliza — thread principal executa

Requer: Python 3.9+ com tkinter (incluso no Python padrao do Windows)
"""

import os, sys, subprocess, threading, socket, time, webbrowser
from pathlib import Path
import tkinter as tk

# ---------------------------------------------------------------------------
# Configuracoes
# ---------------------------------------------------------------------------
PASTA         = Path(__file__).parent
VENV          = PASTA / ".venv"
PY_VENV       = VENV / "Scripts" / "python.exe"
STREAMLIT_EXE = VENV / "Scripts" / "streamlit.exe"
APP_PY        = PASTA / "app.py"
PORTA_FILE    = PASTA / ".porta_local"
ICO_FILE      = PASTA / "ifind.ico"

PACKAGES = [
    "streamlit>=1.32.0", "PyMuPDF>=1.23.0", "Pillow>=10.0.0",
    "pytesseract>=0.3.10", "openpyxl>=3.1.0", "pandas>=2.0.0",
    "rapidfuzz>=3.6.0", "plotly>=5.18.0", "pywebview>=4.4.0",
]

CF  = "#0F1117"
CV  = "#1D9E75"
CT  = "#FFFFFF"
CT2 = "#9A9A9A"
CBG = "#252836"
CE  = "#E24B4A"
CA  = "#EF9F27"

W = 520; H = 320

# ---------------------------------------------------------------------------
# Utilitarios de porta
# ---------------------------------------------------------------------------
def porta_livre(p):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.3)
        return s.connect_ex(("127.0.0.1", p)) != 0

def encontrar_porta():
    p = 8501
    if PORTA_FILE.exists():
        try:
            s = int(PORTA_FILE.read_text().strip())
            if 8501 <= s <= 8599:
                p = s
        except:
            pass
    for porta in list(range(p, 8600)) + list(range(8501, p)):
        if porta_livre(porta):
            PORTA_FILE.write_text(str(porta))
            return porta
    return 8501

def streamlit_ok(porta, timeout=90):
    fim = time.time() + timeout
    while time.time() < fim:
        if not porta_livre(porta):
            return True
        time.sleep(0.4)
    return False

# ---------------------------------------------------------------------------
# Splash Screen
# ---------------------------------------------------------------------------
class Splash:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("iFind Clinica")
        self.root.configure(bg=CF)
        self.root.geometry(f"{W}x{H}")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self._fechar_hard)
        self.proc = None
        self._fechando = False

        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth()  - W) // 2
        y = (self.root.winfo_screenheight() - H) // 2
        self.root.geometry(f"{W}x{H}+{x}+{y}")

        if ICO_FILE.exists():
            try:
                self.root.iconbitmap(str(ICO_FILE))
            except:
                pass

        self._build()

    def _build(self):
        tk.Frame(self.root, bg=CV, height=5).pack(fill="x")

        corpo = tk.Frame(self.root, bg=CF)
        corpo.pack(fill="both", expand=True, padx=36, pady=22)

        tk.Label(corpo, text="iFind", font=("Segoe UI", 38, "bold"),
                 fg=CV, bg=CF).pack(anchor="w")
        tk.Label(corpo, text="Sistema de Busca em Documentos",
                 font=("Segoe UI", 11), fg=CT2, bg=CF).pack(anchor="w")

        tk.Frame(corpo, bg="#2A2D3A", height=1).pack(fill="x", pady=12)

        self._sv = tk.StringVar(value="Iniciando...")
        tk.Label(corpo, textvariable=self._sv, font=("Segoe UI", 10, "bold"),
                 fg=CT, bg=CF, anchor="w").pack(fill="x")

        self._dv = tk.StringVar(value="")
        self._dlb = tk.Label(corpo, textvariable=self._dv,
                              font=("Segoe UI", 9), fg=CT2, bg=CF, anchor="w")
        self._dlb.pack(fill="x", pady=(2, 10))

        self._canvas = tk.Canvas(corpo, bg=CF, height=8,
                                  highlightthickness=0, bd=0)
        self._canvas.pack(fill="x")
        self._canvas.create_rectangle(0, 0, W - 72, 8, fill=CBG, outline="",
                                       tags="bg")
        self._canvas.create_rectangle(0, 0, 0, 8, fill=CV, outline="",
                                       tags="fg")
        self._prog_atual = 0.0
        self._prog_alvo  = 0.0

        self._pv = tk.StringVar(value="")
        tk.Label(corpo, textvariable=self._pv, font=("Segoe UI", 8),
                 fg=CT2, bg=CF, anchor="e").pack(fill="x")

        self._log_frame = tk.Frame(corpo, bg=CF)
        self._log_frame.pack(fill="x", pady=(6, 0))
        self._log_labels = []
        for _ in range(3):
            lbl = tk.Label(self._log_frame, text="", font=("Consolas", 8),
                           fg=CT2, bg=CF, anchor="w")
            lbl.pack(fill="x")
            self._log_labels.append(lbl)
        self._log_lines = []

        rod = tk.Frame(self.root, bg="#0A0C12", height=30)
        rod.pack(fill="x", side="bottom")
        rod.pack_propagate(False)
        tk.Label(rod, text="Nova Medicina e Seguranca do Trabalho",
                 font=("Segoe UI", 8), fg="#555", bg="#0A0C12").pack(
                 side="left", padx=14, pady=7)
        tk.Label(rod, text=f"Python {sys.version_info.major}.{sys.version_info.minor}",
                 font=("Segoe UI", 8), fg="#555", bg="#0A0C12").pack(
                 side="right", padx=14, pady=7)

        self._animar()

    def _animar(self):
        if self._fechando:
            return
        diff = self._prog_alvo - self._prog_atual
        if abs(diff) > 0.001:
            self._prog_atual += diff * 0.12
        w = self._canvas.winfo_width()
        if w > 1:
            fill_w = int(self._prog_atual * w)
            self._canvas.coords("bg", 0, 0, w, 8)
            self._canvas.coords("fg", 0, 0, fill_w, 8)
        self.root.after(16, self._animar)

    # ── helpers thread-safe ──────────────────────────────────────

    def status(self, txt, cor=None):
        def _u():
            self._sv.set(txt)
            self._dlb.config(fg=cor or CT2)
        self.root.after(0, _u)

    def detalhe(self, txt, cor=None):
        def _u():
            self._dv.set(txt)
            self._dlb.config(fg=cor or CT2)
        self.root.after(0, _u)

    def prog(self, v):
        def _u():
            self._prog_alvo = max(0.0, min(1.0, v))
            self._pv.set(f"{int(self._prog_alvo * 100)}%")
        self.root.after(0, _u)

    def log(self, txt, cor=None):
        def _u():
            self._log_lines.append((txt, cor or CT2))
            if len(self._log_lines) > 3:
                self._log_lines = self._log_lines[-3:]
            for i, lbl in enumerate(self._log_labels):
                if i < len(self._log_lines):
                    t, c = self._log_lines[-(len(self._log_lines) - i)]
                    lbl.config(text=t, fg=c)
                else:
                    lbl.config(text="")
        self.root.after(0, _u)

    def erro(self, txt):
        self.status("ERRO: " + txt, CE)
        self.detalhe("Feche e tente novamente.", CE)
        self.log("!! " + txt, CE)

    def esconder(self):
        """Esconde a janela SEM destruir — mantém o loop tkinter vivo."""
        def _u():
            self.root.withdraw()
        self.root.after(0, _u)

    def encerrar_tudo(self):
        """Encerra o splash E o processo Streamlit."""
        def _u():
            self._fechando = True
            if self.proc and self.proc.poll() is None:
                try:
                    self.proc.terminate()
                except:
                    pass
            self.root.destroy()
        self.root.after(0, _u)

    def _fechar_hard(self):
        """Botao X: encerra tudo."""
        self._fechando = True
        if self.proc and self.proc.poll() is None:
            try:
                self.proc.terminate()
            except:
                pass
        self.root.destroy()

    def agendar_na_principal(self, fn):
        """
        Agenda fn() para rodar NA THREAD PRINCIPAL do tkinter.
        Usado para abrir o webview — obrigatorio no Windows.
        """
        self.root.after(0, fn)

    def bg(self, fn):
        threading.Thread(target=fn, daemon=True).start()

    def loop(self):
        self.root.mainloop()


# ---------------------------------------------------------------------------
# Abertura do WebView / browser — DEVE rodar na thread principal
# ---------------------------------------------------------------------------

def _abrir_webview_principal(porta, proc, sp):
    """
    Chamada pela thread principal via sp.agendar_na_principal().
    Tenta abrir pywebview; fallback para browser se falhar.
    """
    # Esconde o splash antes de abrir o webview
    sp.root.withdraw()

    try:
        import webview

        def _ao_fechar():
            # Webview fechou — encerra Streamlit e destroi o tkinter
            if proc and proc.poll() is None:
                try:
                    proc.terminate()
                except:
                    pass
            try:
                sp.root.after(0, sp.root.destroy)
            except:
                pass

        janela = webview.create_window(
            title            = "iFind Clinica",
            url              = f"http://localhost:{porta}",
            width            = 1280,
            height           = 800,
            min_size         = (900, 600),
            resizable        = True,
            background_color = "#FFFFFF",
        )
        janela.events.closed += _ao_fechar

        # start() e BLOQUEANTE — mantém a thread principal ocupada
        # (correto: e assim que o pywebview funciona no Windows)
        webview.start(gui="edgechromium", debug=False)

    except ImportError:
        # pywebview nao instalado — abre no browser e aguarda
        webbrowser.open(f"http://localhost:{porta}")
        # Aguarda o Streamlit encerrar em background,
        # mostra mini janela para o usuario poder fechar
        _mostrar_mini_janela(porta, proc, sp)

    except Exception as e:
        # Qualquer erro no webview — fallback para browser
        _log_erro(e)
        webbrowser.open(f"http://localhost:{porta}")
        _mostrar_mini_janela(porta, proc, sp)


def _log_erro(e):
    """Salva erro em arquivo de log para diagnostico."""
    try:
        log_path = PASTA / "launcher_error.log"
        with open(log_path, "a", encoding="utf-8") as f:
            import traceback
            f.write(f"\n[{time.strftime('%Y-%m-%d %H:%M:%S')}] WebView error:\n")
            f.write(traceback.format_exc())
    except:
        pass


def _mostrar_mini_janela(porta, proc, sp):
    """
    Mini janela que aparece quando o webview nao esta disponivel.
    Mostra o link e botoes Abrir / Encerrar.
    Fica aberta enquanto o Streamlit rodar.
    """
    # Reaproveit a janela do splash (ja existe, root esta vivo)
    root = sp.root

    # Limpa e reconstrói como janela mini
    for w in root.winfo_children():
        w.destroy()

    root.geometry("340x110")
    root.resizable(False, False)
    root.deiconify()

    # Recentraliza
    root.update_idletasks()
    x = (root.winfo_screenwidth()  - 340) // 2
    y = (root.winfo_screenheight() - 110) // 2
    root.geometry(f"340x110+{x}+{y}")

    tk.Frame(root, bg=CV, height=4).pack(fill="x")

    corpo = tk.Frame(root, bg=CF)
    corpo.pack(fill="both", expand=True, padx=16, pady=10)

    tk.Label(corpo, text="iFind esta rodando", font=("Segoe UI", 11, "bold"),
             fg=CV, bg=CF).pack(anchor="w")

    url = f"http://localhost:{porta}"
    tk.Label(corpo, text=url, font=("Segoe UI", 9), fg=CT2, bg=CF,
             cursor="hand2").pack(anchor="w")

    bts = tk.Frame(corpo, bg=CF)
    bts.pack(anchor="e", pady=(8, 0))

    tk.Button(
        bts, text="Abrir no navegador",
        command=lambda: webbrowser.open(url),
        bg=CV, fg=CT, font=("Segoe UI", 9),
        relief="flat", padx=12, cursor="hand2",
    ).pack(side="left", padx=(0, 6))

    def _enc():
        if proc and proc.poll() is None:
            try:
                proc.terminate()
            except:
                pass
        root.destroy()

    tk.Button(
        bts, text="Encerrar",
        command=_enc,
        bg="#2A2D3A", fg=CT2, font=("Segoe UI", 9),
        relief="flat", padx=12, cursor="hand2",
    ).pack(side="left")

    # Monitora o processo em background — fecha a janela se Streamlit morrer
    def _vigiar():
        if proc:
            proc.wait()
        try:
            root.after(0, root.destroy)
        except:
            pass

    threading.Thread(target=_vigiar, daemon=True).start()

    # Volta para o mainloop (janela mini fica aberta)


# ---------------------------------------------------------------------------
# Logica de inicializacao (thread de background)
# ---------------------------------------------------------------------------

def init(sp):
    try:
        # 1. Python
        sp.status("Verificando Python...")
        sp.prog(0.04)
        if sys.version_info < (3, 9):
            sp.erro("Python 3.9+ necessario")
            time.sleep(6)
            sp.encerrar_tudo()
            return
        sp.log(f"  Python {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro} OK", CV)

        # 2. Venv
        sp.status("Preparando ambiente...")
        sp.prog(0.10)
        py = str(PY_VENV) if PY_VENV.exists() else sys.executable

        if not PY_VENV.exists():
            sp.status("Criando ambiente virtual...")
            sp.log("  Criando .venv...", CA)
            r = subprocess.run(
                [sys.executable, "-m", "venv", str(VENV)],
                capture_output=True, timeout=120
            )
            if r.returncode != 0:
                sp.erro("Falha ao criar .venv")
                sp.log("  " + r.stderr.decode(errors="replace")[:80], CE)
                time.sleep(6)
                sp.encerrar_tudo()
                return
            py = str(PY_VENV)
            sp.log("  .venv criado", CV)
        else:
            sp.log("  .venv OK", CV)

        sp.prog(0.18)
        time.sleep(0.1)

        # 3. Dependencias
        sp.status("Verificando dependencias...")
        sp.prog(0.22)
        chk = subprocess.run(
            [py, "-c", "import streamlit, fitz, rapidfuzz"],
            capture_output=True
        )

        if chk.returncode != 0:
            sp.log("  Instalando pacotes (aguarde)...", CA)
            for i, pkg in enumerate(PACKAGES):
                nm = pkg.split(">=")[0].split("==")[0]
                sp.status(f"Instalando ({i+1}/{len(PACKAGES)}): {nm}...")
                sp.log(f"  {nm}...", CA)
                r = subprocess.run(
                    [py, "-m", "pip", "install", "--quiet",
                     "--disable-pip-version-check", pkg],
                    capture_output=True, text=True, timeout=300
                )
                sp.prog(0.22 + (i + 1) / len(PACKAGES) * 0.32)
                if r.returncode != 0:
                    sp.erro(f"Falha ao instalar: {nm}")
                    time.sleep(6)
                    sp.encerrar_tudo()
                    return
            sp.log("  Pacotes instalados", CV)
        else:
            sp.log("  Pacotes OK", CV)

        sp.prog(0.57)
        time.sleep(0.1)

        # 4. Tesseract
        sp.status("Verificando Tesseract OCR...")
        sp.prog(0.60)
        import shutil as sh
        cfg_t = PASTA / "config_tesseract.py"
        tess_ok = any([
            sh.which("tesseract"),
            Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe").exists(),
            Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe").exists(),
            (PASTA / "tesseract_bin" / "tesseract.exe").exists(),
            cfg_t.exists() and cfg_t.stat().st_size > 80,
        ])

        if tess_ok:
            sp.log("  Tesseract OK", CV)
        else:
            sp.log("  Instalando Tesseract...", CA)
            setup = PASTA / "setup_tesseract.py"
            if setup.exists():
                sp.status("Instalando Tesseract (pode demorar)...")
                r = subprocess.run(
                    [py, str(setup)],
                    capture_output=True, text=True, timeout=360
                )
                sp.log(
                    "  Instalado" if r.returncode == 0 else "  Aviso: instalar manualmente",
                    CV if r.returncode == 0 else CA
                )

        sp.prog(0.70)
        time.sleep(0.1)

        # 5. Updates
        sp.status("Verificando atualizacoes...")
        sp.prog(0.73)
        try:
            upd = PASTA / "updater.py"
            if upd.exists():
                r = subprocess.run(
                    [py, "-c",
                     f"import sys; sys.path.insert(0, r'{PASTA}'); "
                     "from updater import verificar_versao_disponivel; "
                     "d = verificar_versao_disponivel(timeout=4); "
                     "print(d['version'] if d else 'ok')"],
                    capture_output=True, text=True, timeout=8
                )
                s = r.stdout.strip()
                sp.log(f"  Nova versao: {s}" if (s and s != "ok") else "  Atualizado",
                       CA if (s and s != "ok") else CV)
        except:
            sp.log("  Sem conexao", CT2)

        sp.prog(0.78)
        time.sleep(0.1)

        # 6. Porta
        sp.status("Detectando porta...")
        sp.prog(0.82)
        porta = encontrar_porta()
        sp.log(f"  Porta: {porta}", CV)
        sp.prog(0.86)
        time.sleep(0.1)

        # 7. Streamlit
        sp.status("Iniciando sistema...")
        sp.prog(0.88)

        if not STREAMLIT_EXE.exists():
            sp.erro("streamlit.exe nao encontrado")
            sp.log("  Apague .venv e reinicie", CA)
            time.sleep(6)
            sp.encerrar_tudo()
            return

        cmd = [
            str(STREAMLIT_EXE), "run", str(APP_PY),
            f"--server.port={porta}",
            "--server.address=0.0.0.0",
            "--server.headless=true",
            "--server.fileWatcherType=none",
            "--browser.gatherUsageStats=false",
            "--theme.base=light",
            "--theme.primaryColor=#1D9E75",
            "--theme.backgroundColor=#FFFFFF",
            "--theme.secondaryBackgroundColor=#F1EFE8",
            "--theme.textColor=#2C2C2A",
            "--theme.font=sans serif",
        ]

        fl = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        proc = subprocess.Popen(
            cmd, cwd=str(PASTA), creationflags=fl,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        sp.proc = proc
        sp.prog(0.92)

        # 8. Healthcheck
        sp.status("Aguardando sistema ficar pronto...")
        sp.log("  Aguardando resposta HTTP...")

        if not streamlit_ok(porta, timeout=90):
            sp.erro("Sistema nao respondeu em 90s")
            sp.log(f"  Tente: http://localhost:{porta}", CA)
            time.sleep(6)
            sp.encerrar_tudo()
            return

        sp.prog(1.0)
        sp.log(f"  Pronto! http://localhost:{porta}", CV)
        sp.status("Abrindo iFind...", CV)
        time.sleep(0.6)

        # 9. ── PONTO CRITICO ──────────────────────────────────────
        # Agenda a abertura do webview NA THREAD PRINCIPAL (obrigatorio).
        # NAO chama sp.fechar() aqui — o tkinter precisa continuar vivo
        # para o root.mainloop() servir de loop de eventos para o webview.
        sp.agendar_na_principal(
            lambda: _abrir_webview_principal(porta, proc, sp)
        )

    except Exception as ex:
        import traceback
        _log_erro(ex)
        sp.erro(str(ex)[:60])
        sp.log("  " + str(ex)[:80], CE)
        time.sleep(6)
        sp.encerrar_tudo()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    os.chdir(str(PASTA))
    sp = Splash()
    sp.bg(lambda: init(sp))
    sp.loop()  # mainloop roda aqui — nunca sai antes do webview fechar


if __name__ == "__main__":
    main()