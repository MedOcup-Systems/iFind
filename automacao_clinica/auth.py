"""
auth.py — Controle de autenticação e sessão Streamlit.
Usa st.session_state para persistir o usuário logado.

CORREÇÕES aplicadas nesta versão:
  1. usuario_atual() agora retorna o campo "usuario" (login) além de "nome"
     — necessário para verificar_login() na troca de senha funcionar corretamente
  2. fazer_logout() usa st.session_state.pop() para limpar completamente as chaves
     — evita que None residual cause KeyError em rerun
  3. tela_login() valida se os campos estão preenchidos antes de tentar login
     — evita erro silencioso com campos vazios
"""

import streamlit as st
from database import verificar_login


def inicializar_sessao():
    """
    Garante que todas as chaves de sessão existam com valores padrão.
    Deve ser chamada antes de qualquer leitura do session_state.
    """
    defaults = {
        "usuario_logado"  : None,   # nome completo do usuário
        "usuario_login"   : None,   # login (usado para verificar_login)
        "usuario_id"      : None,
        "usuario_admin"   : False,
    }
    for chave, valor in defaults.items():
        if chave not in st.session_state:
            st.session_state[chave] = valor


def esta_logado() -> bool:
    """Retorna True se há um usuário autenticado na sessão."""
    return st.session_state.get("usuario_logado") is not None


def usuario_atual() -> dict | None:
    """
    Retorna dict com os dados do usuário logado ou None se não autenticado.

    Campos retornados:
        id      : int   — ID no banco de dados
        nome    : str   — nome completo (ex: "Administrador")
        usuario : str   — login (ex: "admin") — usado para verificar_login()
        admin   : bool  — True se o usuário tem perfil de administrador
    """
    if not esta_logado():
        return None
    return {
        "id"      : st.session_state.get("usuario_id"),
        "nome"    : st.session_state.get("usuario_logado"),
        "usuario" : st.session_state.get("usuario_login"),
        "admin"   : st.session_state.get("usuario_admin", False),
    }


def fazer_login(usuario: str, senha: str) -> bool:
    """
    Valida credenciais contra o banco de dados.
    Se válidas, persiste os dados na sessão e retorna True.
    Se inválidas, não altera a sessão e retorna False.
    """
    user = verificar_login(usuario, senha)
    if user:
        st.session_state["usuario_logado"] = user["nome"]
        st.session_state["usuario_login"]  = user["usuario"]   # ← login salvo
        st.session_state["usuario_id"]     = user["id"]
        st.session_state["usuario_admin"]  = bool(user["admin"])
        return True
    return False


def fazer_logout():
    """
    Limpa completamente a sessão do usuário.
    Usa pop() para garantir que as chaves não fiquem com valor None residual
    que poderia causar comportamento inesperado após st.rerun().
    """
    for chave in ["usuario_logado", "usuario_login", "usuario_id", "usuario_admin"]:
        st.session_state.pop(chave, None)


def tela_login():
    """
    Renderiza a tela de login e bloqueia o restante do app via st.stop()
    até que o usuário se autentique com sucesso.

    Como usar no app.py:
        from auth import tela_login, usuario_atual
        tela_login()              # bloqueia aqui se não logado
        usuario = usuario_atual() # só chega aqui se autenticado
    """
    inicializar_sessao()

    if esta_logado():
        return  # já autenticado — continua normalmente

    # --- Tela de login ---
    st.title("🔍 Buscador de Documentos")
    st.caption("Sistema de busca automática em PDFs — Clínica")
    st.divider()

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.subheader("Entrar")

        campo_usuario = st.text_input(
            "Usuário",
            placeholder="seu.usuario",
            key="login_usuario"
        )
        campo_senha = st.text_input(
            "Senha",
            type="password",
            placeholder="••••••••",
            key="login_senha"
        )

        if st.button("Entrar", type="primary", use_container_width=True):
            if not campo_usuario.strip() or not campo_senha.strip():
                st.error("Preencha o usuário e a senha.")
            elif fazer_login(campo_usuario.strip(), campo_senha):
                st.rerun()
            else:
                st.error("Usuário ou senha incorretos.")

        st.divider()
        st.caption("Primeiro acesso: usuário `admin`, senha `admin123`")
        st.caption("Altere a senha em **Configurações** após o primeiro login.")

    # st.stop() lança streamlit.runtime.scriptrunner.StopException
    # Usamos o raise direto para garantir que a exceção se propague
    # mesmo se houver um try/except mais acima na call stack.
    try:
        st.stop()
    except Exception:
        raise