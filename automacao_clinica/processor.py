"""
processor.py — Lógica de busca e extração de documentos PDF.
Inclui busca fuzzy (tolerância a erros de digitação/OCR) via rapidfuzz.

MELHORIAS NESTA VERSÃO:
  1. ler_planilha() muito mais inteligente:
     - Detecta cabeçalho mesmo que não esteja na linha 1
     - Aceita planilhas sem cabeçalho (usa heurística de conteúdo)
     - Ignora linhas mescladas, vazias e totalizadores
     - Detecta colunas por múltiplos critérios, não só pelo cabeçalho
     - Tenta todas as abas da planilha se a ativa estiver vazia
     - Normaliza datas em qualquer formato encontrado
     - Nunca lança erro por "planilha vazia" sem tentar muito mais

  2. Seletor de pasta nativo do Windows via tkinter (Ctrl+O style)
     — função abrir_seletor_pasta() disponível para o app.py usar

Não renomeie as funções — app.py depende delas.
"""

import fitz  # PyMuPDF
import openpyxl
import unicodedata
import re
import io
from pathlib import Path
from datetime import datetime, date


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


def nome_contem(texto_pagina: str, nome_buscado: str,
                threshold: float = 80.0) -> tuple[bool, float]:
    """
    Verifica se o nome aparece no texto da página.
    Estratégia dupla: busca exata normalizada → busca fuzzy com rapidfuzz.
    Retorna (encontrado, score).
    """
    texto_norm = normalizar_texto(texto_pagina)
    nome_norm  = normalizar_texto(nome_buscado)

    # 1. Busca exata
    palavras = nome_norm.split()
    if palavras and all(p in texto_norm for p in palavras):
        return True, 100.0

    # 2. Busca fuzzy
    try:
        from rapidfuzz import fuzz
        palavras_texto = texto_norm.split()
        tam = len(palavras)
        melhor = 0.0
        for i in range(max(1, len(palavras_texto) - tam + 1)):
            janela = " ".join(palavras_texto[i: i + tam + 2])
            score  = fuzz.token_set_ratio(nome_norm, janela)
            if score > melhor:
                melhor = score
        if melhor >= threshold:
            return True, float(melhor)
    except ImportError:
        pass

    return False, 0.0


# ---------------------------------------------------------------------------
# Seletor de pasta nativo do Windows / macOS / Linux
# ---------------------------------------------------------------------------

def abrir_seletor_pasta(titulo: str = "Selecionar pasta") -> str:
    """
    Abre o seletor de pasta nativo do sistema operacional.
    Windows: usa tkinter (incluso no Python padrão).
    macOS  : usa osascript.
    Linux  : tenta zenity, yad ou kdialog.

    Retorna o caminho escolhido como string, ou "" se cancelado.
    """
    import platform
    sistema = platform.system().lower()

    if sistema == "windows":
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()           # esconde a janela principal
            root.wm_attributes("-topmost", True)  # aparece na frente
            pasta = filedialog.askdirectory(title=titulo)
            root.destroy()
            return pasta or ""
        except Exception:
            return ""

    elif sistema == "darwin":
        try:
            import subprocess
            script = f'choose folder with prompt "{titulo}"'
            r = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True, text=True
            )
            if r.returncode == 0:
                # Converte formato macOS (colon-separated) para POSIX
                caminho = r.stdout.strip().replace(":", "/")
                if caminho.startswith("/"):
                    return caminho
                return "/" + caminho
        except Exception:
            return ""

    else:  # Linux
        for cmd in [
            ["zenity", "--file-selection", "--directory", f"--title={titulo}"],
            ["yad",    "--file", "--directory", f"--title={titulo}"],
            ["kdialog","--getexistingdirectory", "."],
        ]:
            try:
                import subprocess, shutil
                if not shutil.which(cmd[0]):
                    continue
                r = subprocess.run(cmd, capture_output=True, text=True)
                if r.returncode == 0:
                    return r.stdout.strip()
            except Exception:
                continue

    return ""


# ---------------------------------------------------------------------------
# Leitura inteligente da planilha Excel
# ---------------------------------------------------------------------------

# Palavras-chave para detectar colunas automaticamente
_PALAVRAS_NOME  = {"nome", "paciente", "patient", "name", "beneficiario",
                   "segurado", "cliente", "pessoa", "titular"}
_PALAVRAS_DATA  = {"data", "date", "dt", "dia", "procedimento", "proc",
                   "atendimento", "consulta", "exame", "realizacao", "competencia"}
_PALAVRAS_EMAIL = {"email", "e-mail", "correio", "mail"}

# Palavras que indicam que uma linha é cabeçalho, total ou lixo
_PALAVRAS_LIXO  = {"total", "subtotal", "soma", "media", "count",
                   "grand total", "totais", "#", "n/a", "n.a."}


def _celula_para_str(valor) -> str:
    """Converte qualquer valor de célula para string limpa."""
    if valor is None:
        return ""
    if isinstance(valor, (datetime, date)):
        return str(valor)
    return str(valor).strip()


def _parece_data(valor) -> bool:
    """Heurística: o valor parece uma data?"""
    if isinstance(valor, (datetime, date)):
        return True
    s = str(valor).strip()
    padroes = [
        r"^\d{1,2}/\d{1,2}(/\d{2,4})?$",   # DD/MM ou DD/MM/AAAA
        r"^\d{4}-\d{2}-\d{2}",               # AAAA-MM-DD
        r"^\d{1,2}-\d{1,2}-\d{2,4}$",        # DD-MM-AAAA
    ]
    return any(re.match(p, s) for p in padroes)


def _parece_nome(valor) -> bool:
    """Heurística: o valor parece um nome de pessoa?"""
    s = str(valor).strip()
    if not s or len(s) < 3:
        return False
    # Nome: contém letras, sem números, sem muitos caracteres especiais
    tem_letra = any(c.isalpha() for c in s)
    tem_numero = any(c.isdigit() for c in s)
    return tem_letra and not tem_numero and len(s.split()) >= 1


def _linha_e_lixo(linha: tuple) -> bool:
    """Retorna True se a linha parece ser cabeçalho repetido, total ou linha vazia."""
    valores = [_celula_para_str(v).lower() for v in linha if v is not None]
    if not valores:
        return True
    # Linha com palavras de lixo
    for v in valores:
        if any(lixo in v for lixo in _PALAVRAS_LIXO):
            return True
    return False


def _detectar_coluna(cabecalho: list[str], palavras_chave: set,
                     fallback: int) -> int:
    """Detecta o índice da coluna pelo cabeçalho. Retorna fallback se não achar."""
    for i, h in enumerate(cabecalho):
        h_norm = normalizar_texto(h)
        if any(p in h_norm for p in palavras_chave):
            return i
    return fallback


def _score_celula_cabecalho(celula_str: str) -> int:
    """
    Retorna quanto esta célula parece ser um RÓTULO de cabeçalho.
    Usa comparação exata (ou prefixo/sufixo) com palavras-chave conhecidas,
    não 'contains' — assim 'pedro@email.com' não pontua por ter 'email'.
    """
    s = celula_str.lower().strip()
    if not s:
        return 0
    score = 0
    # Valor exatamente igual a uma palavra-chave
    if s in _PALAVRAS_NOME:  score += 4
    if s in _PALAVRAS_DATA:  score += 4
    if s in _PALAVRAS_EMAIL: score += 2
    # Valor que começa ou termina com palavra-chave (ex: "data procedimento")
    for p in _PALAVRAS_NOME | _PALAVRAS_DATA | _PALAVRAS_EMAIL:
        if s.startswith(p + " ") or s.endswith(" " + p):
            score += 2
    return score


def _encontrar_linha_cabecalho(ws) -> tuple[int, bool]:
    """
    Detecta a linha do cabeçalho e se a planilha TEM cabeçalho.

    Retorna:
        (numero_linha, tem_cabecalho)
        - numero_linha : linha onde está o cabeçalho (1-indexed)
        - tem_cabecalho: False se nenhuma linha claramente tem cabeçalho
                         (planilha começa direto nos dados)

    Usa comparação exata com palavras-chave para evitar falsos positivos
    (ex: "pedro@email.com" não deve ser detectado como cabeçalho por ter "email").
    """
    melhor_linha  = 1
    melhor_score  = 0

    for num_linha in range(1, min(11, ws.max_row + 1)):
        score = sum(
            _score_celula_cabecalho(_celula_para_str(c.value))
            for c in ws[num_linha]
        )
        if score > melhor_score:
            melhor_score = score
            melhor_linha = num_linha

    # Threshold mínimo: score >= 4 para considerar que tem cabeçalho real
    # (evita que qualquer linha com um valor parcialmente parecido vire cabeçalho)
    tem_cabecalho = melhor_score >= 4
    return melhor_linha, tem_cabecalho


def _detectar_colunas_por_conteudo(ws, primeira_linha: int) -> tuple[int, int, int]:
    """
    Quando não há cabeçalho, detecta as colunas analisando o conteúdo
    das primeiras linhas de dados.

    Retorna (col_nome, col_data, col_email) — índices 0-based.
    """
    amostras_por_col: dict[int, list] = {}

    for num_linha in range(primeira_linha, min(primeira_linha + 10, ws.max_row + 1)):
        for col_idx, cell in enumerate(ws[num_linha]):
            amostras_por_col.setdefault(col_idx, []).append(cell.value)

    col_nome  = 0  # fallback: coluna A
    col_data  = 1  # fallback: coluna B
    col_email = -1

    score_nome  = {i: 0 for i in amostras_por_col}
    score_data  = {i: 0 for i in amostras_por_col}
    score_email = {i: 0 for i in amostras_por_col}

    for col_idx, valores in amostras_por_col.items():
        for v in valores:
            if v is None:
                continue
            s = str(v).strip()
            # Parece nome: só letras e espaços, 2+ palavras, comprimento razoável
            if re.match(r'^[A-Za-zÀ-ÿ\s]{4,60}$', s) and len(s.split()) >= 2:
                score_nome[col_idx] += 2
            # Parece nome simples (1 palavra longa)
            elif re.match(r'^[A-Za-zÀ-ÿ]{4,30}$', s):
                score_nome[col_idx] += 1
            # Parece data
            if _parece_data(v):
                score_data[col_idx] += 3
            # Parece email
            if '@' in s and '.' in s:
                score_email[col_idx] += 3

    # Escolhe a coluna com maior score para cada tipo
    # (sem conflito: cada tipo pega a melhor coluna diferente)
    usadas: set[int] = set()

    best_nome = max(score_nome, key=lambda i: score_nome[i])
    if score_nome[best_nome] > 0:
        col_nome = best_nome
        usadas.add(best_nome)

    candidatos_data = {i: s for i, s in score_data.items() if i not in usadas}
    if candidatos_data:
        best_data = max(candidatos_data, key=lambda i: candidatos_data[i])
        if score_data[best_data] > 0:
            col_data = best_data
            usadas.add(best_data)

    candidatos_email = {i: s for i, s in score_email.items() if i not in usadas}
    if candidatos_email:
        best_email = max(candidatos_email, key=lambda i: candidatos_email[i])
        if score_email[best_email] > 0:
            col_email = best_email

    return col_nome, col_data, col_email


def _ler_aba(ws) -> list[dict]:
    """
    Extrai registros de uma aba de planilha com total tolerância a variações:
    - Com ou sem cabeçalho
    - Cabeçalho em qualquer linha (não só a primeira)
    - Colunas em qualquer ordem
    - Linhas vazias no meio
    - Células mescladas (tratadas como vazias)
    - Linhas de total/subtotal (ignoradas)
    Retorna lista de dicts {nome, data, email} ou lista vazia.
    """
    if ws.max_row < 1:
        return []

    # 1. Detecta se há cabeçalho e em que linha está
    linha_cab, tem_cabecalho = _encontrar_linha_cabecalho(ws)

    if tem_cabecalho:
        # Planilha tem cabeçalho: usa palavras-chave para detectar colunas
        cabecalho  = [_celula_para_str(c.value) for c in ws[linha_cab]]
        col_nome   = _detectar_coluna(cabecalho, _PALAVRAS_NOME,  0)
        col_data   = _detectar_coluna(cabecalho, _PALAVRAS_DATA,  1)
        col_email  = _detectar_coluna(cabecalho, _PALAVRAS_EMAIL, -1)
        primeira_linha_dados = linha_cab + 1
        nomes_cabecalho = {normalizar_texto(h) for h in cabecalho if h}
    else:
        # Planilha SEM cabeçalho: detecta colunas pelo conteúdo dos dados
        cabecalho  = []
        col_nome, col_data, col_email = _detectar_colunas_por_conteudo(ws, 1)
        primeira_linha_dados = 1  # dados começam na linha 1
        nomes_cabecalho = set()

    registros = []

    for num_linha in range(primeira_linha_dados, ws.max_row + 1):
        linha = tuple(c.value for c in ws[num_linha])

        # Ignora linhas completamente vazias
        if all(v is None for v in linha):
            continue

        # Ignora linhas de lixo/total
        if _linha_e_lixo(linha):
            continue

        nome  = linha[col_nome]  if col_nome  < len(linha) else None
        data  = linha[col_data]  if col_data  < len(linha) else None
        email = linha[col_email] if 0 <= col_email < len(linha) else ""

        nome_str = _celula_para_str(nome)

        # Ignora linha se nome for vazio
        if not nome_str:
            continue

        # Ignora linha se nome for igual a alguma célula do cabeçalho
        if nome_str and normalizar_texto(nome_str) in nomes_cabecalho:
            continue

        # Se a data da coluna detectada não parece uma data,
        # vasculha todas as outras colunas da linha procurando uma data válida
        if not _parece_data(data):
            for idx, val in enumerate(linha):
                if idx != col_nome and _parece_data(val):
                    data = val
                    break

        # Só adiciona se tiver pelo menos nome (data pode estar em branco)
        registros.append({
            "nome" : nome_str,
            "data" : data,
            "email": _celula_para_str(email),
        })

    return registros


def ler_planilha(caminho_excel: str) -> list[dict]:
    """
    Lê o arquivo Excel de forma inteligente e tolerante a variações.

    Estratégias aplicadas em ordem:
    1. Tenta a aba ativa primeiro
    2. Se vazia ou com poucos registros, tenta todas as outras abas
    3. Usa a aba com mais registros válidos
    4. Dentro de cada aba, detecta automaticamente onde começa o cabeçalho
    5. Detecta colunas por palavras-chave (tolerante a variações de nome)
    6. Ignora linhas de total, cabeçalhos repetidos e células mescladas
    7. Se a coluna de data detectada estiver vazia, busca data em outras colunas

    Lança ValueError apenas se nenhuma aba tiver nenhum registro válido.
    """
    wb = openpyxl.load_workbook(caminho_excel, data_only=True)

    # Tenta todas as abas e pega a que tiver mais registros
    melhor_registros: list[dict] = []
    erros_por_aba: list[str]     = []

    abas_para_tentar = [wb.active] + [
        wb[nome] for nome in wb.sheetnames
        if wb[nome] != wb.active
    ]

    for ws in abas_para_tentar:
        if ws is None:
            continue
        try:
            registros = _ler_aba(ws)
            if len(registros) > len(melhor_registros):
                melhor_registros = registros
        except Exception as e:
            erros_por_aba.append(f"Aba '{ws.title}': {e}")

    if not melhor_registros:
        detalhe = "; ".join(erros_por_aba) if erros_por_aba else "nenhum dado encontrado"
        raise ValueError(
            f"Não foi possível extrair registros válidos da planilha. "
            f"Detalhes: {detalhe}. "
            f"Verifique se a planilha tem colunas de nome e data com dados preenchidos."
        )

    return melhor_registros


# ---------------------------------------------------------------------------
# Navegação nas pastas do drive
# ---------------------------------------------------------------------------

MESES = {
    "01": "janeiro",  "02": "fevereiro", "03": "marco",
    "04": "abril",    "05": "maio",      "06": "junho",
    "07": "julho",    "08": "agosto",    "09": "setembro",
    "10": "outubro",  "11": "novembro",  "12": "dezembro",
}

# Para personalizar os nomes das pastas do drive, edite o dicionário acima.
# Exemplos:
#   "03": "Março"    (com acento)
#   "03": "03"       (número puro)
#   "03": "march"    (inglês)


def extrair_dia_mes(data) -> tuple[str, str]:
    """
    Converte qualquer formato de data em (dia, mes) com zero à esquerda.
    Aceita: datetime, date, 'DD/MM', 'DD/MM/AAAA', 'AAAA-MM-DD', serial Excel.
    """
    if isinstance(data, (datetime, date)):
        if isinstance(data, datetime):
            return data.strftime("%d"), data.strftime("%m")
        return data.strftime("%d"), data.strftime("%m")

    texto = str(data).strip()

    if re.match(r"^\d{4}-\d{2}-\d{2}", texto):
        p = texto.split("-")
        return p[2][:2], p[1][:2]

    if "/" in texto:
        p = texto.split("/")
        return p[0].strip().zfill(2), p[1].strip().zfill(2)

    if "-" in texto and not texto.startswith("-"):
        p = texto.split("-")
        if len(p) >= 2:
            return p[0].strip().zfill(2), p[1].strip().zfill(2)

    # Serial numérico do Excel
    try:
        import datetime as dt_mod
        n = int(float(texto))
        d = dt_mod.datetime(1899, 12, 30) + dt_mod.timedelta(days=n)
        return d.strftime("%d"), d.strftime("%m")
    except Exception:
        pass

    raise ValueError(
        f"Formato de data não reconhecido: '{data}'. "
        "Aceitos: DD/MM, DD/MM/AAAA, AAAA-MM-DD ou objeto datetime."
    )


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
    Tenta texto embutido primeiro; se vazio, aplica OCR a 200 DPI.
    """
    texto = pagina.get_text()

    if len(texto.strip()) < 20:
        try:
            from PIL import Image
            import pytesseract

            cfg_path = Path(__file__).parent / "config_tesseract.py"
            if cfg_path.exists():
                exec(cfg_path.read_text(encoding="utf-8"), {})

            pix = pagina.get_pixmap(dpi=200)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            try:
                texto = pytesseract.image_to_string(img, lang="por")
            except Exception:
                texto = pytesseract.image_to_string(img)
        except ImportError:
            pass

    return texto


def buscar_nome_em_pdf(caminho_pdf: Path, nome_buscado: str,
                        threshold: float = 80.0) -> list[tuple[int, float]]:
    """
    Busca o nome em todas as páginas do PDF.
    Retorna lista de (numero_pagina, score_fuzzy).
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
    limpo      = re.sub(r"[^\w\s-]", "", sem_acento)
    return re.sub(r"\s+", "_", limpo.strip())


def extrair_e_salvar_paginas(
    caminho_pdf: Path,
    paginas_scores: list[tuple[int, float]],
    pasta_destino: Path,
    nome_paciente: str,
) -> Path:
    """
    Extrai as páginas encontradas e salva como novo PDF.
    Evita sobrescrever arquivos existentes adicionando sufixo numérico.
    """
    pasta_destino = Path(pasta_destino)
    pasta_destino.mkdir(parents=True, exist_ok=True)

    nome_base = sanitizar_nome_arquivo(nome_paciente)
    saida     = pasta_destino / f"{nome_base}.pdf"
    contador  = 2
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
    callback=None,
) -> list[dict]:
    """
    Processa todos os registros da planilha.

    Parâmetros:
        caminho_excel   : caminho para o .xlsx
        drive_raiz      : raiz do drive com as pastas dos meses
        pasta_destino   : onde salvar os PDFs extraídos
        threshold_fuzzy : pontuação mínima para busca fuzzy (0-100)
        callback        : função(progresso: float, mensagem: str)

    Retorna lista de dicts: nome, data, email, encontrado, arquivo, erro, score_fuzzy
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
                    melhor_paginas: list[tuple[int, float]] = []
                    melhor_pdf    = None
                    melhor_score  = 0.0

                    for pdf in pdfs:
                        paginas = buscar_nome_em_pdf(pdf, nome, threshold_fuzzy)
                        if paginas:
                            score_max = max(s for _, s in paginas)
                            if score_max > melhor_score:
                                melhor_paginas = paginas
                                melhor_pdf     = pdf
                                melhor_score   = score_max
                            if melhor_score == 100.0:
                                break

                    if melhor_paginas:
                        saida = extrair_e_salvar_paginas(
                            melhor_pdf, melhor_paginas,
                            Path(pasta_destino), nome,
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
            icone = "[OK]" if resultado["encontrado"] else "[X]"
            callback((i + 1) / total, f"{icone} {nome}")

    return resultados