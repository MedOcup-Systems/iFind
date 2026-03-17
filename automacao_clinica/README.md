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
