"""
gerar_release.py — Gera o pacote de update para publicar no GitHub
====================================================================

USO:
    python gerar_release.py 2.1.0
    python gerar_release.py 2.1.0 "Correcao de bugs no filtro ASO"
    python gerar_release.py 2.1.0 "Atualizacao critica" --obrigatorio

O script faz tudo automaticamente:
  1. Empacota os arquivos em ifind_v2.1.0.zip
  2. Calcula SHA256 para verificacao de integridade
  3. Gera o version.json com URL e hash corretos
  4. Atualiza o VERSION no updater.py para a nova versao
  5. Copia o version.json para a raiz do projeto (para o Git)
  6. Mostra instrucoes exatas de publicacao no GitHub

Repositorio configurado: novamedicinasoftwares/iFind
"""

import hashlib
import json
import re
import sys
import zipfile
from pathlib import Path
from datetime import datetime

# ================================================================
# CONFIGURACAO — ajuste apenas estas variaveis
# ================================================================

# Usuario e repositorio no GitHub
GITHUB_USER = "novamedicinasoftwares"
GITHUB_REPO = "iFind"

# Arquivos incluidos no pacote de update (sem dados do usuario)
# IMPORTANTE: nao incluir config.json, clinica.db, .auth_token
ARQUIVOS_INCLUIR = [
    "launcher.py",        # ponto de entrada — OBRIGATORIO
    "app.py",
    "auth.py",
    "database.py",
    "processor.py",
    "mailer.py",
    "setup_tesseract.py",
    "updater.py",
    "requirements.txt",
    "iniciar.bat",
]

# Pastas incluidas (apenas arquivos especificos dentro delas)
PASTAS_INCLUIR = [
    (".streamlit", ["config.toml"]),
]

# Arquivos opcionais — incluidos se existirem, ignorados se nao
ARQUIVOS_OPCIONAIS = [
    "README.md",
    "ifind.ico",
]

# ================================================================


def calcular_sha256(caminho: Path) -> str:
    sha256 = hashlib.sha256()
    with open(caminho, "rb") as f:
        for bloco in iter(lambda: f.read(65536), b""):
            sha256.update(bloco)
    return sha256.hexdigest()


def atualizar_version_no_updater(pasta: Path, nova_versao: str) -> bool:
    """
    Atualiza a constante VERSION no updater.py para a nova versao.
    Retorna True se atualizou, False se nao encontrou.
    """
    updater = pasta / "updater.py"
    if not updater.exists():
        return False

    conteudo = updater.read_text(encoding="utf-8")
    novo = re.sub(
        r'^VERSION\s*=\s*["\'][\d.]+["\']',
        f'VERSION = "{nova_versao}"',
        conteudo,
        flags=re.MULTILINE
    )

    if novo == conteudo:
        return False  # nao encontrou ou ja estava correto

    updater.write_text(novo, encoding="utf-8")
    return True


def gerar_release(versao: str, notas: str = "", obrigatorio: bool = False):
    pasta = Path(__file__).parent

    # URLs corretas do GitHub
    # ZIP fica nas Releases:  github.com/USER/REPO/releases/download/vX/arquivo.zip
    # JSON fica no raw:       raw.githubusercontent.com/USER/REPO/main/version.json
    url_zip_base  = f"https://github.com/{GITHUB_USER}/{GITHUB_REPO}/releases/download"
    url_json_raw  = f"https://raw.githubusercontent.com/{GITHUB_USER}/{GITHUB_REPO}/main/version.json"

    pasta_dist = pasta / "dist_releases"
    pasta_dist.mkdir(exist_ok=True)

    nome_zip    = f"ifind_v{versao}.zip"
    caminho_zip = pasta_dist / nome_zip
    caminho_json_dist = pasta_dist / "version.json"
    caminho_json_raiz = pasta / "version.json"  # copia na raiz para o Git

    sep = "=" * 56
    print(f"\n{sep}")
    print(f"  iFind Clinica — Gerando release v{versao}")
    print(f"{sep}\n")

    # ── Empacota os arquivos ─────────────────────────────────────
    incluidos = []
    ausentes  = []

    with zipfile.ZipFile(caminho_zip, "w", zipfile.ZIP_DEFLATED) as zf:

        # Arquivos obrigatorios
        for nome in ARQUIVOS_INCLUIR:
            caminho = pasta / nome
            if caminho.exists():
                zf.write(caminho, nome)
                incluidos.append(nome)
                print(f"  [OK] {nome}")
            else:
                ausentes.append(nome)
                print(f"  [!!] {nome} -- NAO ENCONTRADO (verifique!)")

        # Arquivos opcionais
        for nome in ARQUIVOS_OPCIONAIS:
            caminho = pasta / nome
            if caminho.exists():
                zf.write(caminho, nome)
                incluidos.append(nome)
                print(f"  [OK] {nome}")
            # sem aviso se opcional nao existe

        # Pastas
        for pasta_rel, arquivos in PASTAS_INCLUIR:
            for arq in arquivos:
                caminho = pasta / pasta_rel / arq
                nome_zip_interno = f"{pasta_rel}/{arq}"
                if caminho.exists():
                    zf.write(caminho, nome_zip_interno)
                    incluidos.append(nome_zip_interno)
                    print(f"  [OK] {nome_zip_interno}")
                else:
                    print(f"  [--] {nome_zip_interno} -- nao encontrado (pulado)")

    tamanho_kb = caminho_zip.stat().st_size // 1024
    print(f"\n  ZIP: {caminho_zip.name} ({tamanho_kb} KB)")

    # ── SHA256 ───────────────────────────────────────────────────
    sha256 = calcular_sha256(caminho_zip)
    print(f"  SHA256: {sha256}")

    # ── version.json ─────────────────────────────────────────────
    url_zip_final = f"{url_zip_base}/v{versao}/{nome_zip}"

    dados = {
        "version"    : versao,
        "url"        : url_zip_final,
        "hash_sha256": sha256,
        "notas"      : notas or f"Atualizacao v{versao}",
        "obrigatorio": obrigatorio,
        "gerado_em"  : datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "arquivos"   : incluidos,
    }

    # Salva em dist_releases/ (para upload manual) e na raiz (para o Git)
    for dest in [caminho_json_dist, caminho_json_raiz]:
        with open(dest, "w", encoding="utf-8") as f:
            json.dump(dados, f, indent=2, ensure_ascii=False)

    print(f"  version.json: {caminho_json_dist.name} (+ copia na raiz para o Git)")

    # ── Atualiza VERSION no updater.py ───────────────────────────
    if atualizar_version_no_updater(pasta, versao):
        print(f"  updater.py: VERSION atualizado para {versao}")
    else:
        print(f"  [AVISO] Nao foi possivel atualizar VERSION no updater.py")
        print(f"          Atualize manualmente: VERSION = \"{versao}\"")

    # ── Instrucoes de publicacao ─────────────────────────────────
    print(f"""
{sep}
  PUBLICAR NO GITHUB — passo a passo:
{sep}

  PASSO 1 — Commit do version.json (para o updater.py dos clientes):
  ------------------------------------------------------------------
  git add version.json updater.py
  git commit -m "release: v{versao}"
  git push origin main

  PASSO 2 — Criar a Release no GitHub:
  ------------------------------------------------------------------
  1. Acesse:
     https://github.com/{GITHUB_USER}/{GITHUB_REPO}/releases/new

  2. Em "Choose a tag" digite:  v{versao}
     Clique em "Create new tag: v{versao} on publish"

  3. Em "Release title" coloque:  iFind v{versao}

  4. Em "Describe this release" cole as notas:
     {notas or f"Atualizacao v{versao}"}

  5. Arraste os arquivos para "Attach binaries":
     - dist_releases/{nome_zip}
       (este e o arquivo baixado pelo updater.py dos clientes)

  6. Se quiser distribuir o instalador tambem:
     - dist/ifind_clinica_v{versao}_setup.exe
       (se ja compilou o Inno Setup)

  7. Clique em "Publish release"

  PASSO 3 — Verificar que o updater.py vai encontrar:
  ------------------------------------------------------------------
  URL do ZIP (verificar apos publicar):
  {url_zip_final}

  URL do version.json (ja esta no ar apos o git push):
  {url_json_raw}

{sep}
""")

    if ausentes:
        print(f"  [!!] ATENCAO: {len(ausentes)} arquivo(s) NAO encontrado(s):")
        for a in ausentes:
            print(f"       - {a}")
        print()

    return caminho_zip, caminho_json_dist


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso:  python gerar_release.py <versao> [notas] [--obrigatorio]")
        print()
        print("Exemplos:")
        print("  python gerar_release.py 2.1.0")
        print("  python gerar_release.py 2.1.0 \"Correcao no filtro ASO\"")
        print("  python gerar_release.py 2.1.0 \"Atualizacao critica\" --obrigatorio")
        sys.exit(1)

    versao = sys.argv[1]
    notas  = sys.argv[2] if len(sys.argv) > 2 and not sys.argv[2].startswith("--") else ""
    obrig  = "--obrigatorio" in sys.argv

    gerar_release(versao, notas, obrig)