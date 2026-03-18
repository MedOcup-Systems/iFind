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
# Verificação e instalação automática do Tesseract
# ---------------------------------------------------------------------------

def _tesseract_ok() -> bool:
    if shutil.which("tesseract"):
        return True

    if platform.system().lower() == "windows":
        candidatos = [
            Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe"),
            Path(r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"),
            Path(os.environ.get("LOCALAPPDATA", "")) / "Tesseract-OCR" / "tesseract.exe",
            Path(os.environ.get("APPDATA", ""))      / "Tesseract-OCR" / "tesseract.exe",
        ]
        for exe in candidatos:
            if exe.exists():
                os.environ["PATH"] = str(exe.parent) + os.pathsep + os.environ.get("PATH", "")
                return True

    portatil = Path(__file__).parent / "tesseract_bin" / "tesseract.exe"
    if portatil.exists():
        os.environ["PATH"] = str(portatil.parent) + os.pathsep + os.environ.get("PATH", "")
        return True

    cfg_tess = Path(__file__).parent / "config_tesseract.py"
    if cfg_tess.exists():
        try:
            exec(cfg_tess.read_text(encoding="utf-8"), {})
            import pytesseract as _pt
            _pt.get_tesseract_version()
            return True
        except Exception:
            pass

    return False

def _mostrar_instrucoes_manuais():
    sistema = platform.system().lower()
    st.markdown("**Instalação manual:**")
    if sistema == "windows":
        st.markdown(
            "1. Acesse: https://github.com/UB-Mannheim/tesseract/wiki\n"
            "2. Baixe o instalador para Windows 64-bit\n"
            "3. Execute **como Administrador**\n"
            "4. Marque a option **Add to PATH** durante a instalação\n"
            "5. Reinicie o terminal e execute `streamlit run app.py` novamente"
        )
    elif sistema == "linux":
        st.code("sudo apt install tesseract-ocr")
    elif sistema == "darwin":
        st.code("brew install tesseract")

def _verificar_tesseract():
    if _tesseract_ok():
        return

    st.warning(
        "⚙️ **Tesseract OCR não encontrado.** "
        "Iniciando instalação automática — aguarde."
    )

    setup_path = Path(__file__).parent / "setup_tesseract.py"
    if not setup_path.exists():
        setup_path = Path(__file__).parent / "setup.py"

    if not setup_path.exists():
        st.error(
            "❌ Arquivo de instalação não encontrado. "
            "Coloque o `setup.py` na mesma pasta que o `app.py` e recarregue."
        )
        st.stop()

    st.info(f"🔧 Instalando Tesseract OCR via `{setup_path.name}`...")
    st.caption("O script de instalação será executado. Uma janela UAC pode aparecer pedindo permissão de administrador — clique em Sim.")

    import ctypes

    def _tem_admin():
        try:
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except Exception:
            return False

    if platform.system().lower() == "windows" and not _tem_admin():
        st.warning("⚙️ Solicitando permissão de administrador via UAC. Confirme a janela que aparecerá.")
        script = str(setup_path.resolve())
        res = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, f'"{script}"', None, 1
        )
        if res <= 32:
            st.error(f"❌ Falha ao solicitar admin (código {res}). Execute o setup_tesseract.py manualmente como Administrador.")
            _mostrar_instrucoes_manuais()
            st.stop()

        import time
        st.info("Aguardando a instalação concluir (até 3 minutos)...")
        for _ in range(36):
            time.sleep(5)
            if _tesseract_ok():
                break
    else:
        placeholder_log = st.empty()
        resultado = subprocess.run(
            [sys.executable, str(setup_path)],
            capture_output=True,
            text=True,
            timeout=360
        )
        saida = resultado.stdout + resultado.stderr
        linhas = [l for l in saida.splitlines() if l.strip()]
        placeholder_log.code("\n".join(linhas[-20:]), language=None)

        if resultado.returncode != 0:
            st.error(
                "❌ Não foi possível instalar o Tesseract automaticamente. "
                "Veja o log acima e instale manualmente."
            )
            _mostrar_instrucoes_manuais()
            st.stop()
        else:
            st.success("✅ Instalação concluída!")

    if _tesseract_ok():
        st.success("✅ Tesseract configurado. Recarregando o sistema...")
        import time
        time.sleep(1)
        st.rerun()
    else:
        st.warning(
            "⚠️ Instalação concluída, mas o Tesseract ainda não foi detectado. "
            "Isso pode ocorrer porque o PATH do sistema ainda não foi atualizado. "
            "**Feche este terminal, abra um novo e execute `streamlit run app.py` novamente.**"
        )
        st.stop()

# ---------------------------------------------------------------------------
# CHAMADA DA VERIFICAÇÃO
# ---------------------------------------------------------------------------
_verificar_tesseract()

# ---------------------------------------------------------------------------
# Inicialização dos módulos
# ---------------------------------------------------------------------------
try:
    from database import (
        inicializar_banco, listar_execucoes, resultados_da_execucao,
        iniciar_execucao, finalizar_execucao, salvar_resultado,
        estatisticas_gerais, listar_usuarios, criar_usuario,
        alterar_senha, desativar_usuario
    )
    from auth      import tela_login, usuario_atual, fazer_logout, inicializar_sessao
    from processor import processar_lista
    from mailer    import enviar_email, enviar_relatorio_execucao, email_configurado
except ImportError as e:
    st.error(f"Erro de importação: {e}")
    st.info("Execute: pip install -r requirements.txt")
    st.stop()

def _main():
    inicializar_banco()
    tela_login()

    usuario = usuario_atual()
    if usuario is None:
        return

    # ---------------------------------------------------------------------------
    # Caminho do config.json
    # ---------------------------------------------------------------------------
    CONFIG_PATH = Path(__file__).parent / "config.json"

    def carregar_config() -> dict:
        if not CONFIG_PATH.exists():
            return {
                "smtp": {}, "drive_raiz": "", "pasta_destino": "",
                "threshold_fuzzy": 80, "enviar_email_auto": False,
                "email_relatorio": ""
            }
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)

    def salvar_config(cfg: dict):
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2, ensure_ascii=False)

    cfg = carregar_config()

    if "input_drive_raiz" not in st.session_state:
        st.session_state["input_drive_raiz"] = cfg.get("drive_raiz", "")

    if "input_pasta_destino" not in st.session_state:
        st.session_state["input_pasta_destino"] = cfg.get("pasta_destino", "")

    # ---------------------------------------------------------------------------
    # Header global
    # ---------------------------------------------------------------------------
    col_h1, col_h2 = st.columns([8, 2])
    with col_h1:
        st.markdown("### 🔍 Buscador de Documentos")
    with col_h2:
        st.caption(f"Olá, **{usuario.get('nome', '')}**")
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

        def cb_selecionar_pasta_drive():
            from processor import abrir_seletor_pasta
            pasta_sel = abrir_seletor_pasta("Selecionar pasta raiz do drive")
            if pasta_sel:
                st.session_state["input_drive_raiz"] = pasta_sel

        def cb_selecionar_pasta_destino():
            from processor import abrir_seletor_pasta
            pasta_sel = abrir_seletor_pasta("Selecionar pasta de destino")
            if pasta_sel:
                st.session_state["input_pasta_destino"] = pasta_sel

        col1, col2 = st.columns(2)

        with col1:
            arquivo_excel = st.file_uploader(
                "📄 Planilha Excel com nomes e datas",
                type=["xlsx", "xls"],
                help="Colunas detectadas automaticamente. Aceita coluna de e-mail para envio automático."
            )

            _dr_col1, _dr_col2 = st.columns([5, 1])
            with _dr_col1:
                drive_raiz = st.text_input(
                    "📁 Caminho raiz do drive",
                    placeholder="Ex: Z:\\Procedimentos  ou  /mnt/drive/proc",
                    key="input_drive_raiz"
                )
            with _dr_col2:
                st.markdown("<br>", unsafe_allow_html=True)
                st.button("🗂️", key="btn_drive", help="Selecionar pasta pelo explorador", on_click=cb_selecionar_pasta_drive)

        with col2:
            _pd_col1, _pd_col2 = st.columns([5, 1])
            with _pd_col1:
                pasta_destino = st.text_input(
                    "📂 Pasta de destino",
                    placeholder="Ex: C:\\Resultados  ou  /home/user/resultados",
                    key="input_pasta_destino"
                )
            with _pd_col2:
                st.markdown("<br>", unsafe_allow_html=True)
                st.button("🗂️", key="btn_destino", help="Selecionar pasta pelo explorador", on_click=cb_selecionar_pasta_destino)

            threshold = st.slider(
                "🎯 Sensibilidade da busca fuzzy",
                min_value=60, max_value=100,
                value=cfg.get("threshold_fuzzy", 80),
                step=5,
                help="100 = apenas correspondência exata. 80 = tolera erros de OCR. 60 = muito permissivo."
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

        iniciar = st.button("▶ Iniciar busca", type="primary")

        if iniciar:
            erros = []
            if not arquivo_excel:         erros.append("Faça upload da planilha Excel.")
            if not drive_raiz.strip():    erros.append("Informe o caminho do drive.")
            if not pasta_destino.strip(): erros.append("Informe a pasta de destino.")
            if erros:
                for e in erros:
                    st.error(f"❌ {e}")
                st.stop()

            cfg.update({
                "drive_raiz": drive_raiz.strip(),
                "pasta_destino": pasta_destino.strip(),
                "threshold_fuzzy": threshold,
                "enviar_email_auto": enviar_email_auto,
                "email_relatorio": email_relatorio,
            })
            salvar_config(cfg)

            with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                tmp.write(arquivo_excel.read())
                caminho_tmp = tmp.name

            execucao_id = iniciar_execucao(usuario.get('id'), drive_raiz.strip(), pasta_destino.strip())

            st.divider()
            st.subheader("Progresso")
            barra     = st.progress(0, text="Iniciando...")
            st_txt    = st.empty()
            st_log    = st.empty()
            log_lines: list[str] = []

            def cb(prog: float, msg: str):
                barra.progress(prog, text=f"{int(prog * 100)}% concluído")
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

            total       = len(resultados)
            encontrados = sum(1 for r in resultados if r["encontrado"])
            for r in resultados:
                salvar_resultado(execucao_id, r)
            finalizar_execucao(execucao_id, total, encontrados)

            barra.progress(1.0, text="Concluído!")
            st_txt.empty()

            if enviar_email_auto and email_configurado():
                with st.spinner("Enviando e-mails..."):
                    enviados = 0
                    for r in resultados:
                        if r["encontrado"] and r.get("email"):
                            ok_mail, _ = enviar_email(
                                destinatario=r["email"],
                                assunto=f"Documento encontrado — {r['nome']}",
                                corpo=(
                                    f"<p>Prezado(a), o documento de <strong>{r['nome']}</strong> "
                                    f"referente a <strong>{r['data']}</strong> foi localizado e está anexo.</p>"
                                ),
                                caminho_pdf=r["arquivo"]
                            )
                            if ok_mail:
                                enviados += 1
                    if email_relatorio:
                        enviar_relatorio_execucao(
                            email_relatorio, total, encontrados,
                            total - encontrados, usuario.get('nome', '')
                        )
                st.success(f"📧 {enviados} e-mail(s) enviado(s) com sucesso.")

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
                "nome": "Nome", "data": "Data", "email": "E-mail",
                "encontrado": "Encontrado", "arquivo": "Arquivo gerado",
                "erro": "Observação", "score_fuzzy": "Score",
            })

            def colorir(row):
                cor = "#d4edda" if row["Encontrado"] else "#f8d7da"
                return [f"background-color:{cor}"] * len(row)

            st.dataframe(
                df.style.apply(colorir, axis=1),
                width='stretch',
                hide_index=True,
            )

            csv = df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "⬇️ Baixar relatório CSV", csv,
                "relatorio_busca.csv", "text/csv",
                width='stretch',
            )

            if nao_enc > 0:
                st.warning(
                    f"⚠️ {nao_enc} registro(s) não encontrado(s). "
                    "Verifique a coluna 'Observação'. "
                    "Tente reduzir a sensibilidade fuzzy se houver falsos positivos."
                )

    # ============================================================
    # ABA 2 — HISTÓRICO
    # ============================================================
    with aba_historico:
        st.subheader("Histórico de execuções")

        uid = None if usuario.get("admin", False) else usuario.get("id")
        execucoes = listar_execucoes(usuario_id=uid, limite=100)

        if not execucoes:
            st.info("Nenhuma execução registrada ainda.")
        else:
            df_exec = pd.DataFrame(execucoes).rename(columns={
                "id": "ID", "usuario_nome": "Usuário", "inicio": "Início", "fim": "Fim",
                "total": "Total", "encontrados": "Encontrados",
                "nao_encontrados": "Não enc.", "drive_raiz": "Drive",
                "pasta_destino": "Destino",
            })
            st.dataframe(
                df_exec[["ID", "Usuário", "Início", "Fim", "Total", "Encontrados", "Não enc."]],
                width='stretch',
                hide_index=True,
            )

            st.divider()
            st.subheader("Detalhes de uma execução")

            ids = [e["id"] for e in execucoes]
            exec_sel = st.selectbox(
                "Selecione o ID da execução", ids,
                format_func=lambda x: f"#{x}"
            )

            if exec_sel:
                res = resultados_da_execucao(exec_sel)
                if res:
                    df_res = pd.DataFrame(res).rename(columns={
                        "nome": "Nome", "data": "Data", "encontrado": "Encontrado",
                        "arquivo": "Arquivo", "erro": "Observação", "score_fuzzy": "Score",
                    })

                    def colorir_hist(row):
                        cor = "#d4edda" if row["Encontrado"] else "#f8d7da"
                        return [f"background-color:{cor}"] * len(row)

                    st.dataframe(
                        df_res[["Nome", "Data", "Encontrado", "Score", "Arquivo", "Observação"]]
                        .style.apply(colorir_hist, axis=1),
                        width='stretch',
                        hide_index=True,
                    )

                    csv_hist = df_res.to_csv(index=False).encode("utf-8-sig")
                    st.download_button(
                        "⬇️ Exportar execução CSV", csv_hist,
                        f"execucao_{exec_sel}.csv", "text/csv",
                        width='stretch',
                    )
                else:
                    st.info("Nenhum resultado detalhado para esta execução.")

    # ============================================================
    # ABA 3 — ESTATÍSTICAS
    # ============================================================
    with aba_stats:
        st.subheader("Estatísticas de uso")

        try:
            import plotly.express as px
            PLOTLY = True
        except ImportError:
            PLOTLY = False
            st.warning("Instale plotly para gráficos: pip install plotly")

        stats = estatisticas_gerais()

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total de execuções",    stats["total_execucoes"])
        c2.metric("Total de buscas",       stats["total_buscas"])
        c3.metric("Total encontrados",     stats["total_encontrados"])
        c4.metric("Taxa média de sucesso", f"{stats['taxa_sucesso']}%")

        if PLOTLY and stats["por_dia"]:
            st.divider()
            df_dia = pd.DataFrame(stats["por_dia"])
            fig1 = px.bar(
                df_dia, x="dia", y=["encontrados", "total"],
                barmode="group",
                labels={"dia": "Data", "value": "Quantidade", "variable": "Tipo"},
                color_discrete_map={"encontrados": "#1D9E75", "total": "#B5D4F4"},
                title="Buscas por dia — últimos 30 dias",
            )
            fig1.update_layout(
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                legend_title_text="",
            )
            st.plotly_chart(fig1, use_container_width=True)

        if PLOTLY and stats["top_usuarios"]:
            st.divider()
            df_usr = pd.DataFrame(stats["top_usuarios"])
            fig2 = px.bar(
                df_usr, x="nome", y="execucoes",
                labels={"nome": "Usuário", "execucoes": "Execuções"},
                color="execucoes", color_continuous_scale="Teal",
                title="Top usuários mais ativos",
            )
            fig2.update_layout(
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
                showlegend=False,
            )
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
            s1 = st.text_input("Senha atual",  type="password", key="s_atual")
            s2 = st.text_input("Nova senha",   type="password", key="s_nova")
            s3 = st.text_input("Confirmar",    type="password", key="s_conf")
            if st.button("Salvar senha"):
                from database import verificar_login as vl
                if not vl(usuario.get("usuario", usuario.get("nome", "")), s1):
                    st.error("Senha atual incorreta.")
                elif s2 != s3:
                    st.error("As senhas não coincidem.")
                elif len(s2) < 6:
                    st.error("A nova senha deve ter ao menos 6 caracteres.")
                else:
                    alterar_senha(usuario.get('id'), s2)
                    st.success("Senha alterada com sucesso!")

        # --- E-mail SMTP ---
        st.markdown("#### Configurações de e-mail (SMTP)")
        with st.expander("Configurar servidor SMTP"):
            smtp = cfg.get("smtp", {})
            smtp_host  = st.text_input("Servidor SMTP",         value=smtp.get("host", ""),    placeholder="smtp.gmail.com")
            smtp_porta = st.number_input("Porta",               value=int(smtp.get("porta", 587)), min_value=1, max_value=65535, step=1)
            smtp_user  = st.text_input("Usuário/E-mail",        value=smtp.get("usuario", ""))
            smtp_senha = st.text_input("Senha do e-mail",       type="password", value=smtp.get("senha", ""))
            smtp_rem   = st.text_input("Nome/e-mail remetente", value=smtp.get("remetente", smtp.get("usuario", "")))

            st.caption(
                "Para Gmail: use smtp.gmail.com porta 587 e crie uma "
                "[Senha de App](https://myaccount.google.com/apppasswords) "
                "(não use sua senha normal do Gmail)."
            )

            if st.button("Salvar configurações de e-mail"):
                cfg["smtp"] = {
                    "host": smtp_host, "porta": smtp_porta,
                    "usuario": smtp_user, "senha": smtp_senha,
                    "remetente": smtp_rem or smtp_user,
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
                        "<p>Se você recebeu este e-mail, as configurações SMTP estão corretas.</p>",
                    )
                    if ok_t:
                        st.success("E-mail de teste enviado com sucesso!")
                    else:
                        st.error(f"Falha no envio: {err_t}")

        # --- Gerenciar usuários (somente admin) ---
        if usuario.get("admin", False):
            st.markdown("#### Gerenciar usuários")
            with st.expander("Usuários cadastrados"):
                usuarios_lista = listar_usuarios()
                df_u = pd.DataFrame(usuarios_lista)[["id", "usuario", "nome", "email", "admin", "ativo", "criado_em"]]
                df_u.columns = ["ID", "Login", "Nome", "E-mail", "Admin", "Ativo", "Criado em"]
                st.dataframe(df_u, width='stretch', hide_index=True)

            with st.expander("Criar novo usuário"):
                nu = st.text_input("Login do novo usuário", key="nu")
                nn = st.text_input("Nome completo",         key="nn")
                ne = st.text_input("E-mail (opcional)",     key="ne")
                ns = st.text_input("Senha", type="password", key="ns")
                na = st.checkbox("Administrador",           key="na")
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

        # --- Diagnóstico do sistema ---
        st.markdown("#### Informações do sistema")
        with st.expander("Diagnóstico"):
            import platform as _pl
            tess_path = shutil.which("tesseract") or "não encontrado no PATH"
            portatil_exe = Path(__file__).parent / "tesseract_bin" / "tesseract.exe"
            if portatil_exe.exists():
                tess_path += f"  |  portátil: {portatil_exe}"
            st.code(
                f"Sistema:    {_pl.system()} {_pl.release()}\n"
                f"Python:     {_pl.python_version()}\n"
                f"Streamlit:  {st.__version__}\n"
                f"Tesseract:  {tess_path}\n"
                f"Banco:      {(Path(__file__).parent / 'clinica.db').resolve()}\n"
                f"Config:     {CONFIG_PATH.resolve()}\n",
                language=None,
            )

_main()