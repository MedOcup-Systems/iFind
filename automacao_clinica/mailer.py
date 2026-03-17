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
