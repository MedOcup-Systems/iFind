"""
updater.py — Sistema de Auto-Update para iFind Clínica
=======================================================

Arquitetura adaptada para o projeto real:
  - NÃO é PyInstaller. O sistema roda como arquivos .py + .venv.
  - O que é atualizado: arquivos .py, .bat, .toml — não um .exe monolítico.
  - Isso é mais simples, mais seguro e sem risco de "exe em uso".

Fluxo:
  1. Ao iniciar o app, verifica versão no servidor (JSON público).
  2. Se há versão nova: baixa um .zip com os arquivos atualizados.
  3. Verifica SHA256 do zip baixado.
  4. Extrai e substitui os arquivos (com backup automático).
  5. Reinicia o Streamlit com os novos arquivos.

Servidor necessário:
  Qualquer hospedagem que sirva dois arquivos estáticos:
    - version.json   → metadados da versão atual
    - ifind_vX.Y.Z.zip → arquivos do update

  Sugestão gratuita: GitHub Releases (recomendado) ou qualquer CDN.

Formato do version.json:
  {
    "version": "2.1.0",
    "url": "https://github.com/seu-usuario/ifind/releases/download/v2.1.0/ifind_v2.1.0.zip",
    "hash_sha256": "abc123...",
    "notas": "Correção de caminhos, modo varredura total, detecção de porta.",
    "obrigatorio": false
  }

  Como funciona:
  # 1. Faz as alterações nos arquivos
    # 2. Incrementa VERSION em updater.py (ex: "2.1.0")
    # 3. Gera o pacote:
    python gerar_release.py 2.1.0 "Descrição do que mudou"
    # 4. Sobe dist_releases/version.json e dist_releases/ifind_v2.1.0.zip no GitHub
    # 5. Pronto — todos os PCs com o sistema vão receber a notificação automaticamente

  obrigatorio: true  → força update sem perguntar ao usuário
  obrigatorio: false → pergunta ao usuário antes de atualizar
"""

import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path
from datetime import datetime

# ── Versão atual do sistema ────────────────────────────────────
VERSION = "1.0.1"

# ── URL do arquivo de versão no servidor ──────────────────────
VERSION_URL = "https://raw.githubusercontent.com/novamedicinasoftwares/iFind/main/version.json"

# ── Arquivos que NÃO devem ser sobrescritos no update ─────────
# Preserva configurações e banco de dados do usuário
ARQUIVOS_PROTEGIDOS = {
    "config.json",       # configurações SMTP, caminhos, threshold
    "clinica.db",        # histórico de execuções
    "config_tesseract.py",  # configuração do Tesseract
    ".auth_token",       # token de login salvo
    ".porta_local",      # porta preferida deste PC
}

# ── Pasta raiz do projeto ──────────────────────────────────────
PASTA_PROJETO = Path(__file__).parent

# ── Arquivo de log de updates ─────────────────────────────────
LOG_UPDATE = PASTA_PROJETO / "update.log"


def _log(msg: str):
    """Registra mensagem no log de updates com timestamp."""
    linha = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    print(linha)
    try:
        with open(LOG_UPDATE, "a", encoding="utf-8") as f:
            f.write(linha + "\n")
    except Exception:
        pass


def _versao_maior(v_nova: str, v_atual: str) -> bool:
    """
    Compara versões semânticas (MAJOR.MINOR.PATCH).
    Retorna True se v_nova > v_atual.
    """
    try:
        nova  = tuple(int(x) for x in v_nova.strip().lstrip("v").split("."))
        atual = tuple(int(x) for x in v_atual.strip().lstrip("v").split("."))
        return nova > atual
    except Exception:
        return False


def verificar_versao_disponivel(timeout: int = 8) -> dict | None:
    """
    Consulta o servidor e retorna os metadados da versão disponível.
    Retorna None se não conseguir acessar ou se já está na versão mais recente.

    Retorna dict com: version, url, hash_sha256, notas, obrigatorio
    """
    try:
        import urllib.request
        import urllib.error

        req = urllib.request.Request(
            VERSION_URL,
            headers={"User-Agent": "iFind-Clinica-Updater/1.0", "Cache-Control": "no-cache"},
        )
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            dados = json.loads(resp.read().decode("utf-8"))

        versao_servidor = dados.get("version", "0.0.0")

        if not _versao_maior(versao_servidor, VERSION):
            _log(f"Sistema atualizado. Versão atual: {VERSION}")
            return None

        _log(f"Nova versão disponível: {versao_servidor} (atual: {VERSION})")
        return dados

    except Exception as e:
        _log(f"Não foi possível verificar atualizações: {e}")
        return None


def _calcular_sha256(caminho: Path) -> str:
    """Calcula SHA256 de um arquivo."""
    sha256 = hashlib.sha256()
    with open(caminho, "rb") as f:
        for bloco in iter(lambda: f.read(65536), b""):
            sha256.update(bloco)
    return sha256.hexdigest()


def _baixar_arquivo(url: str, destino: Path, timeout: int = 120) -> bool:
    """
    Baixa arquivo com barra de progresso simples no terminal.
    Usa urllib da stdlib — sem dependência de requests.
    """
    import urllib.request

    try:
        _log(f"Baixando: {url}")

        def _progresso(count, block_size, total):
            if total > 0:
                pct = min(int(count * block_size * 100 / total), 100)
                print(f"\r  Baixando... {pct}%", end="", flush=True)

        urllib.request.urlretrieve(url, str(destino), reporthook=_progresso)
        print()  # quebra linha após o progresso
        _log(f"Download concluído: {destino.name} ({destino.stat().st_size // 1024} KB)")
        return True

    except Exception as e:
        _log(f"Erro no download: {e}")
        return False


def _fazer_backup(pasta_projeto: Path) -> Path | None:
    """
    Cria backup dos arquivos .py e .bat atuais antes de atualizar.
    Salva em pasta_projeto/.backups/YYYY-MM-DD_HH-MM/
    """
    try:
        ts = datetime.now().strftime("%Y-%m-%d_%H-%M")
        pasta_backup = pasta_projeto / ".backups" / ts
        pasta_backup.mkdir(parents=True, exist_ok=True)

        extensoes_backup = {".py", ".bat", ".toml", ".txt", ".md"}
        for arq in pasta_projeto.iterdir():
            if arq.is_file() and arq.suffix.lower() in extensoes_backup:
                shutil.copy2(arq, pasta_backup / arq.name)

        # Backup do .streamlit/config.toml
        config_toml = pasta_projeto / ".streamlit" / "config.toml"
        if config_toml.exists():
            (pasta_backup / ".streamlit").mkdir(exist_ok=True)
            shutil.copy2(config_toml, pasta_backup / ".streamlit" / "config.toml")

        _log(f"Backup criado em: {pasta_backup}")
        return pasta_backup

    except Exception as e:
        _log(f"Aviso: não foi possível criar backup: {e}")
        return None


def _aplicar_update(zip_path: Path, pasta_projeto: Path) -> bool:
    """
    Extrai o zip do update e sobrescreve os arquivos do projeto.
    Respeita ARQUIVOS_PROTEGIDOS — nunca sobrescreve dados do usuário.
    """
    try:
        with zipfile.ZipFile(zip_path, "r") as zf:
            arquivos_no_zip = zf.namelist()
            _log(f"Arquivos no update: {arquivos_no_zip}")

            for nome_arq in arquivos_no_zip:
                # Ignora diretórios
                if nome_arq.endswith("/"):
                    continue

                # Nunca sobrescreve arquivos protegidos
                nome_base = Path(nome_arq).name
                if nome_base in ARQUIVOS_PROTEGIDOS:
                    _log(f"  Protegido (mantido): {nome_base}")
                    continue

                # Extrai respeitando subpastas (ex: .streamlit/config.toml)
                destino = pasta_projeto / nome_arq
                destino.parent.mkdir(parents=True, exist_ok=True)

                with zf.open(nome_arq) as src, open(destino, "wb") as dst:
                    shutil.copyfileobj(src, dst)

                _log(f"  Atualizado: {nome_arq}")

        return True

    except Exception as e:
        _log(f"Erro ao aplicar update: {e}")
        return False


def _limpar_backups_antigos(pasta_projeto: Path, manter: int = 5):
    """Mantém apenas os N backups mais recentes para não acumular espaço."""
    pasta_backups = pasta_projeto / ".backups"
    if not pasta_backups.exists():
        return
    try:
        backups = sorted(pasta_backups.iterdir(), key=lambda p: p.name, reverse=True)
        for backup_antigo in backups[manter:]:
            shutil.rmtree(backup_antigo, ignore_errors=True)
            _log(f"Backup antigo removido: {backup_antigo.name}")
    except Exception:
        pass


def executar_update(dados_versao: dict, silencioso: bool = False) -> bool:
    """
    Executa o processo completo de update.

    Parâmetros:
        dados_versao : dict retornado por verificar_versao_disponivel()
        silencioso   : True = não pede confirmação (obrigatorio=True força isso)

    Retorna True se o update foi aplicado com sucesso.
    """
    versao_nova  = dados_versao.get("version", "?")
    url_zip      = dados_versao.get("url", "")
    hash_esperado= dados_versao.get("hash_sha256", "").lower()
    notas        = dados_versao.get("notas", "Sem notas de versão.")
    obrigatorio  = dados_versao.get("obrigatorio", False)

    if not url_zip:
        _log("ERRO: URL do update não informada no version.json")
        return False

    _log(f"Iniciando update para versão {versao_nova}")
    _log(f"Notas: {notas}")

    # Baixa o zip em pasta temporária
    with tempfile.TemporaryDirectory() as tmp_dir:
        zip_path = Path(tmp_dir) / f"ifind_v{versao_nova}.zip"

        if not _baixar_arquivo(url_zip, zip_path):
            _log("ERRO: Falha no download. Update abortado.")
            return False

        # Verifica integridade SHA256
        if hash_esperado:
            hash_real = _calcular_sha256(zip_path)
            if hash_real != hash_esperado:
                _log(f"ERRO: Hash inválido! Esperado: {hash_esperado} | Recebido: {hash_real}")
                _log("Update abortado por segurança.")
                return False
            _log("Integridade SHA256 verificada ✓")
        else:
            _log("Aviso: hash não informado no version.json — pulando verificação.")

        # Faz backup antes de sobrescrever
        _fazer_backup(PASTA_PROJETO)
        _limpar_backups_antigos(PASTA_PROJETO)

        # Aplica o update
        if not _aplicar_update(zip_path, PASTA_PROJETO):
            _log("ERRO: Falha ao aplicar update. Verifique o backup.")
            return False

    _log(f"Update para {versao_nova} aplicado com sucesso!")
    return True


def verificar_e_atualizar(silencioso: bool = False) -> bool:
    """
    Ponto de entrada principal.
    Chame esta função no início do app.

    Retorna True se houve update (app deve reiniciar).
    Retorna False se não há update ou ocorreu erro.
    """
    _log(f"Verificando atualizações... (versão atual: {VERSION})")
    dados = verificar_versao_disponivel()

    if not dados:
        return False  # Já atualizado ou sem conexão

    return executar_update(dados, silencioso=silencioso)


# ── Execução direta (para testes) ─────────────────────────────
if __name__ == "__main__":
    print(f"iFind Clínica — Verificador de Updates")
    print(f"Versão instalada: {VERSION}")
    print(f"Servidor: {VERSION_URL}")
    print()
    resultado = verificar_e_atualizar(silencioso=False)
    if resultado:
        print("\nUpdate aplicado! Reinicie o sistema.")
    else:
        print("\nNenhum update necessário.")
