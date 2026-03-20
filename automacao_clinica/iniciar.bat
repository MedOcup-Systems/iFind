@echo off
setlocal EnableDelayedExpansion
title iFind Clinica

:: ============================================================
::  Pasta raiz
:: ============================================================
set "PROJETO=%~dp0"
if "%PROJETO:~-1%"=="\" set "PROJETO=%PROJETO:~0,-1%"
cd /d "%PROJETO%"

set "STREAMLIT=%PROJETO%\.venv\Scripts\streamlit.exe"
set "PY=%PROJETO%\.venv\Scripts\python.exe"
set "PYPIP=%PROJETO%\.venv\Scripts\python.exe"
set "PORTA_FILE=%PROJETO%\.porta_local"

cls
echo.
echo  +----------------------------------------------------------+
echo  ^|      iFIND CLINICA v1                                   ^|
echo  ^|      Sistema de busca automatica em PDFs                ^|
echo  +----------------------------------------------------------+
echo.
echo  Pasta: %PROJETO%
echo.

:: ============================================================
::  [1/6]  Python
:: ============================================================
echo  [1/6]  Verificando Python...

where python >nul 2>&1
if errorlevel 1 goto :SEM_PYTHON

python -c "import sys; sys.exit(0 if sys.version_info>=(3,9) else 1)" >nul 2>&1
if errorlevel 1 goto :PYTHON_VELHO

echo         Python OK
goto :STEP2

:SEM_PYTHON
echo.
echo  ERRO: Python nao encontrado.
echo  Baixe em: https://www.python.org/downloads/
echo  Marque "Add Python to PATH" ao instalar.
goto :ERRO_FATAL

:PYTHON_VELHO
echo.
echo  ERRO: Python 3.9 ou superior e necessario.
echo  Baixe em: https://www.python.org
goto :ERRO_FATAL

:: ============================================================
::  [2/6]  Ambiente virtual
:: ============================================================
:STEP2
echo  [2/6]  Ambiente virtual...

if exist "%PROJETO%\.venv\Scripts\activate.bat" goto :VENV_EXISTS

echo         Criando pela primeira vez...
python -m venv "%PROJETO%\.venv" >nul 2>&1
if errorlevel 1 goto :VENV_ERRO
echo         Criado.
goto :VENV_ATIVA

:VENV_EXISTS
echo         Ja existe - reutilizando.

:VENV_ATIVA
call "%PROJETO%\.venv\Scripts\activate.bat" >nul 2>&1
if errorlevel 1 goto :VENV_ERRO
echo         Ativado.
goto :STEP3

:VENV_ERRO
echo.
echo  ERRO: Falha no ambiente virtual.
echo  Apague a pasta .venv e tente novamente.
goto :ERRO_FATAL

:: ============================================================
::  [3/6]  Dependencias
:: ============================================================
:STEP3
echo  [3/6]  Dependencias Python...

"%PY%" -c "import streamlit" >nul 2>&1
if not errorlevel 1 goto :DEPS_OK

echo         Instalando pacotes (primeira vez - aguarde)...
echo.

"%PYPIP%" -m pip install --quiet --disable-pip-version-check streamlit
"%PYPIP%" -m pip install --quiet --disable-pip-version-check PyMuPDF
"%PYPIP%" -m pip install --quiet --disable-pip-version-check Pillow
"%PYPIP%" -m pip install --quiet --disable-pip-version-check pytesseract
"%PYPIP%" -m pip install --quiet --disable-pip-version-check openpyxl
"%PYPIP%" -m pip install --quiet --disable-pip-version-check pandas
"%PYPIP%" -m pip install --quiet --disable-pip-version-check rapidfuzz
"%PYPIP%" -m pip install --quiet --disable-pip-version-check plotly

echo.
"%PY%" -c "import streamlit" >nul 2>&1
if errorlevel 1 goto :DEPS_ERRO
echo         Pacotes instalados.
goto :DEPS_VERIFICA

:DEPS_ERRO
echo.
echo  ERRO: Falha ao instalar dependencias.
echo  Verifique sua conexao com a internet e tente novamente.
goto :ERRO_FATAL

:DEPS_OK
echo         Ja instalados.

:DEPS_VERIFICA
if not exist "%STREAMLIT%" goto :STREAM_ERRO
echo         streamlit.exe OK
goto :STEP4

:STREAM_ERRO
echo.
echo  ERRO: streamlit.exe nao encontrado em .venv\Scripts\
echo  Apague a pasta .venv e rode o iniciar.bat novamente.
goto :ERRO_FATAL

:: ============================================================
::  [4/6]  Arquivos
:: ============================================================
:STEP4
echo  [4/6]  Arquivos do sistema...

set "FALTA=0"
for %%F in (app.py processor.py database.py auth.py mailer.py) do (
    if exist "%PROJETO%\%%F" (
        echo         %%F - OK
    ) else (
        echo         %%F - AUSENTE
        set "FALTA=1"
    )
)
for %%F in (setup_tesseract.py config_tesseract.py config.json clinica.db) do (
    if exist "%PROJETO%\%%F" (
        echo         %%F - presente
    ) else (
        echo         %%F - sera criado automaticamente
    )
)
if "!FALTA!"=="1" goto :ARQUIVO_ERRO
echo         Todos presentes.
goto :STEP5

:ARQUIVO_ERRO
echo.
echo  ERRO: Arquivos essenciais ausentes.
echo  Coloque todos os .py na mesma pasta que o iniciar.bat
goto :ERRO_FATAL

:: ============================================================
::  [5/6]  Tesseract
:: ============================================================
:STEP5
echo  [5/6]  Tesseract OCR...

set "TESS_OK=0"

where tesseract >nul 2>&1
if not errorlevel 1 ( set "TESS_OK=1" & goto :TESS_FIM )

if exist "C:\Program Files\Tesseract-OCR\tesseract.exe" (
    set "PATH=C:\Program Files\Tesseract-OCR;!PATH!"
    set "TESS_OK=1"
    goto :TESS_FIM
)
if exist "C:\Program Files (x86)\Tesseract-OCR\tesseract.exe" (
    set "PATH=C:\Program Files (x86)\Tesseract-OCR;!PATH!"
    set "TESS_OK=1"
    goto :TESS_FIM
)
if exist "%PROJETO%\config_tesseract.py" (
    set "TESS_OK=1"
    goto :TESS_FIM
)
if exist "%PROJETO%\tesseract_bin\tesseract.exe" (
    set "PATH=%PROJETO%\tesseract_bin;!PATH!"
    set "TESS_OK=1"
)

:TESS_FIM
if "!TESS_OK!"=="1" (
    echo         Encontrado.
) else (
    echo         Nao encontrado - sera instalado automaticamente ao iniciar.
)

:: ============================================================
::  [6/6]  Porta — deteccao inteligente por PC
::
::  Logica:
::    1. Le a porta salva em .porta_local (especifica deste PC)
::    2. Verifica se essa porta ainda esta livre
::    3. Se estiver ocupada, busca a proxima livre no range 8501-8599
::    4. Salva a porta escolhida em .porta_local para proxima vez
::
::  Cada PC da rede usa sua propria porta — sem conflito.
:: ============================================================
:STEP6
echo  [6/6]  Porta de rede...

set "PORTA=8501"

:: Tenta ler porta salva anteriormente neste PC
if exist "%PORTA_FILE%" (
    set /p PORTA_SALVA=<"%PORTA_FILE%"
    :: Valida se e um numero entre 8501 e 8599
    echo !PORTA_SALVA! | findstr /R "^85[0-9][0-9]$" >nul 2>&1
    if not errorlevel 1 (
        set "PORTA=!PORTA_SALVA!"
        echo         Porta preferida deste PC: !PORTA_SALVA!
    )
)

:: Verifica se a porta escolhida esta realmente livre
:: Se nao estiver, busca a proxima livre no range 8501-8599
set "TENTATIVAS=0"
:BUSCA_PORTA
netstat -ano 2>nul | findstr ":!PORTA! " | findstr "LISTENING" >nul 2>&1
if not errorlevel 1 (
    :: Porta ocupada — tenta a proxima
    set /a PORTA+=1
    set /a TENTATIVAS+=1

    :: Nao passou de 8599
    if !PORTA! GTR 8599 set "PORTA=8501"

    :: Seguranca: nao tenta mais de 99 vezes
    if !TENTATIVAS! LSS 99 goto :BUSCA_PORTA

    :: Chegou aqui sem achar porta livre — usa 8501 mesmo
    echo         AVISO: nenhuma porta livre encontrada no range 8501-8599.
    echo         Tentando com 8501 assim mesmo...
    set "PORTA=8501"
    goto :PORTA_OK
)

:PORTA_OK
:: Salva a porta escolhida para este PC reutilizar na proxima vez
echo !PORTA!>"%PORTA_FILE%"
echo         Porta escolhida: !PORTA! (salva para este PC)

:: ============================================================
::  Detecta IP local para acesso pela rede
:: ============================================================
set "IP_LOCAL=seu-ip"
for /f "tokens=2 delims=:" %%I in ('ipconfig 2^>nul ^| findstr "IPv4"') do (
    set "IP_TMP=%%I"
    set "IP_TMP=!IP_TMP: =!"
    echo !IP_TMP! | findstr /R "^192\. ^10\. ^172\." >nul 2>&1
    if not errorlevel 1 (
        set "IP_LOCAL=!IP_TMP!"
        goto :IP_OK
    )
)
:IP_OK

:: ============================================================
::  Pronto - inicia
:: ============================================================
echo.
echo  +----------------------------------------------------------+
echo  ^|  Tudo pronto. Abrindo o sistema...                      ^|
echo  ^|                                                          ^|
echo  ^|  Acesso local  : http://localhost:!PORTA!                ^|
echo  ^|  Acesso na rede: http://!IP_LOCAL!:!PORTA!               ^|
echo  ^|                                                          ^|
echo  ^|  Login padrao  : admin / admin123                        ^|
echo  ^|  Para encerrar : feche esta janela ou Ctrl+C             ^|
echo  +----------------------------------------------------------+
echo.

start "" /b cmd /c "timeout /t 3 >nul 2>&1 && start http://localhost:!PORTA!"

"%STREAMLIT%" run "%PROJETO%\app.py" ^
    --server.port=!PORTA! ^
    --server.address=0.0.0.0 ^
    --server.headless=true ^
    --server.fileWatcherType=none ^
    --browser.gatherUsageStats=false ^
    --theme.base=light ^
    --theme.primaryColor="#1D9E75" ^
    --theme.backgroundColor="#FFFFFF" ^
    --theme.secondaryBackgroundColor="#F1EFE8" ^
    --theme.textColor="#2C2C2A" ^
    --theme.font="sans serif"

echo.
echo  Sistema encerrado.
goto :FIM

:: ============================================================
::  Erro fatal
:: ============================================================
:ERRO_FATAL
echo.
echo  +----------------------------------------------------------+
echo  ^|  ERRO - sistema nao iniciado.                           ^|
echo  ^|  Leia as mensagens acima e corrija antes de continuar.  ^|
echo  ^|  Em caso de duvida, consulte o README.md                ^|
echo  +----------------------------------------------------------+

:FIM
echo.
pause
endlocal