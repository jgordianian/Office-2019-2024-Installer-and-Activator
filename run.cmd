@echo off
setlocal EnableExtensions EnableDelayedExpansion

call :ENSURE_ADMIN
if errorlevel 1 exit /b 0

set "BASEPATH=%~dp0"
pushd "%BASEPATH%" >nul 2>&1

set "OFFICE_SETUP=%BASEPATH%setup.exe"
set "LOGFILE=%BASEPATH%log.txt"
set "INSTALLER_VERSION=1.1.0"
set "CREATOR=Jonah Ian Gordian"

call :INIT_LOG
call :SELECT_UI

:MENU
cls
if /i "!LANG_UI!"=="EN" (
    echo ===============================================
    echo    OFFICE VOLUME + LANGUAGES INSTALLER
    echo              Version !INSTALLER_VERSION!
    echo              Author !CREATOR!
    echo ===============================================
    echo.
    echo 1 - Install Office 2019 Volume
    echo 2 - Install Office 2021 Volume
    echo 3 - Install Office 2024 Volume
    echo 4 - Install additional language
    echo 5 - Activate Office via KMS host
    echo 6 - Activate Windows via KMS host
    echo 7 - Exit
    echo.
    choice /c 1234567 /n /m "Select an option (1-7): "
    echo.
) else (
    echo ===============================================
    echo    INSTALADOR OFFICE VOLUME + IDIOMAS
    echo              Version !INSTALLER_VERSION!
    echo              Autor !CREATOR!
    echo ===============================================
    echo.
    echo 1 - Instalar Office 2019 Volume
    echo 2 - Instalar Office 2021 Volume
    echo 3 - Instalar Office 2024 Volume
    echo 4 - Instalar idioma adicional
    echo 5 - Activar Office via host KMS
    echo 6 - Activar Windows via host KMS
    echo 7 - Salir
    echo.
    choice /c 1234567 /n /m "Seleccione una opcion (1-7): "
    echo.
)

if errorlevel 7 goto SALIR
if errorlevel 6 goto ACTIVATE_WINDOWS_ONLY
if errorlevel 5 goto ACTIVATE_ONLY
if errorlevel 4 goto INSTALAR_IDIOMA
if errorlevel 3 goto INSTALAR_2024
if errorlevel 2 goto INSTALAR_2021
if errorlevel 1 goto INSTALAR_2019
goto MENU

:INSTALAR_2019
set "PRODUCTID=ProPlus2019Volume"
set "CHANNEL=PerpetualVL2019"
set "NEWVER=2019"
call :LOG "Selected Office 2019 Volume"
goto INSTALL_SELECTED

:INSTALAR_2021
set "PRODUCTID=ProPlus2021Volume"
set "CHANNEL=PerpetualVL2021"
set "NEWVER=2021"
call :LOG "Selected Office 2021 Volume"
goto INSTALL_SELECTED

:INSTALAR_2024
set "PRODUCTID=ProPlus2024Volume"
set "CHANNEL=PerpetualVL2024"
set "NEWVER=2024"
call :LOG "Selected Office 2024 Volume"
goto INSTALL_SELECTED

:INSTALL_SELECTED
call :REQUIRE_SETUP
if errorlevel 1 goto FIN_ERROR

call :LOG "!TXT_VERSION_SEL!: !NEWVER!"
call :DETECT_INSTALLED_OFFICE

if defined PRODUCTLINE call :LOG "!TXT_PRODUCT_DET!: !PRODUCTLINE!"

if defined PRODUCTLINE if not defined INSTALLED_PRODUCT_ID (
    call :LOG "!TXT_UNSUPPORTED_INSTALLED!"
    goto FIN_ERROR
)

if defined INSTALLEDVER (
    call :LOG "!TXT_INSTALLED_DET!: !INSTALLEDVER!"
    if !INSTALLEDVER! GEQ !NEWVER! (
        call :LOG "!TXT_VER_EQUAL!"
        goto FIN_SUCCESS
    )
)

if defined INSTALLED_PRODUCT_ID (
    call :LOG "!TXT_UNINSTALL!"
    call :WRITE_REMOVE_XML "%BASEPATH%remove.xml"
    call :RUN_SETUP /configure "%BASEPATH%remove.xml"
    if errorlevel 1 (
        call :LOG "!TXT_UNINST_FAIL!"
        goto FIN_ERROR
    )
)

call :PROMPT_PRIMARY_LANGUAGE
if errorlevel 1 goto FIN_ERROR

call :ENSURE_LOCAL_CACHE "!PRODUCTID!" "!CHANNEL!" "!LANG!"
if errorlevel 1 goto FIN_ERROR

call :WRITE_LANG_XML "%BASEPATH%install.xml" "!PRODUCTID!" "!CHANNEL!" "!LANG!" configure
call :LOG "!TXT_INSTALL_START!"
call :RUN_SETUP /configure "%BASEPATH%install.xml"
if errorlevel 1 (
    call :LOG "!TXT_INST_FAIL!"
    goto FIN_ERROR
)

call :PROMPT_POST_INSTALL_ACTIVATION
if errorlevel 1 goto FIN_ERROR

goto FIN_SUCCESS

:INSTALAR_IDIOMA
call :REQUIRE_SETUP
if errorlevel 1 goto FIN_ERROR

call :DETECT_INSTALLED_OFFICE
if not defined PRODUCTLINE (
    call :LOG "!TXT_NO_OFFICE!"
    pause
    goto MENU
)

call :LOG "!TXT_OFFICE_DET!: !PRODUCTLINE!"

if not defined INSTALLED_PRODUCT_ID (
    call :LOG "!TXT_UNSUPPORTED_INSTALLED!"
    pause
    goto MENU
)

set "PRODUCTID=!INSTALLED_PRODUCT_ID!"
set "CHANNEL=!INSTALLED_CHANNEL!"

call :PROMPT_EXTRA_LANGS
if errorlevel 1 goto MENU

call :ENSURE_LOCAL_CACHE "!PRODUCTID!" "!CHANNEL!" "!NORMALIZED_LANGS!"
if errorlevel 1 goto FIN_ERROR

call :WRITE_LANG_XML "%BASEPATH%addlang.xml" "!PRODUCTID!" "!CHANNEL!" "!NORMALIZED_LANGS!" configure
call :LOG "!TXT_ADDLANG_START!"
call :RUN_SETUP /configure "%BASEPATH%addlang.xml"
if errorlevel 1 (
    call :LOG "!TXT_ADDLANG_FAIL!"
    goto FIN_ERROR
)

goto FIN_SUCCESS

:ACTIVATE_ONLY
call :LOG "!TXT_ACTIVATE_ONLY!"
call :ACTIVATE_OFFICE
if errorlevel 1 goto FIN_ERROR
goto FIN_SUCCESS

:ACTIVATE_WINDOWS_ONLY
call :LOG "!TXT_ACTIVATE_WINDOWS_ONLY!"
call :ACTIVATE_WINDOWS
if errorlevel 1 goto FIN_ERROR
goto FIN_SUCCESS

:SALIR
call :LOG "!TXT_EXIT!"
popd >nul 2>&1
exit /b 0

:FIN_SUCCESS
set "FINAL_MESSAGE=!TXT_DONE!"
goto FIN

:FIN_ERROR
set "FINAL_MESSAGE=!TXT_DONE_ERRORS!"
goto FIN

:FIN
echo ===============================================
echo !FINAL_MESSAGE!
echo !TXT_LOG!: %LOGFILE%
echo !TXT_LICENSE_NOTE!
pause
popd >nul 2>&1
exit /b 0

:SELECT_UI
cls
echo ===============================================
echo   SELECCION DE IDIOMA / LANGUAGE SELECTOR
echo ===============================================
echo 1 - Espanol
echo 2 - English
echo.
choice /c 12 /n /m "Seleccione una opcion / Select an option (1-2): "
if errorlevel 2 (
    set "LANG_UI=EN"
    call :SET_TEXT_EN
) else (
    set "LANG_UI=ES"
    call :SET_TEXT_ES
)
call :LOG "!TXT_UI_LANG!: !LANG_UI!"
exit /b 0

:SET_TEXT_EN
set "TXT_UI_LANG=UI language"
set "TXT_NO_SETUP=ERROR: setup.exe not found"
set "TXT_VERSION_SEL=Selected version"
set "TXT_PRODUCT_DET=Detected product"
set "TXT_INSTALLED_DET=Detected installed version"
set "TXT_VER_EQUAL=Installed version is equal or newer. No reinstall was needed."
set "TXT_UNINSTALL=Removing previous supported Office installation..."
set "TXT_UNINST_FAIL=ERROR: Uninstall failed. Review log.txt."
set "TXT_LANG_BASIC=Examples: es-es, en-us, fr-fr, pt-br, it-it, de-de"
set "TXT_LANG_PROMPT=Enter primary language (example: es-es): "
set "TXT_LANG_INVALID=Invalid language code. Use format xx-xx."
set "TXT_LANG_INVALID_LOG=Primary language validation failed"
set "TXT_LANG_SELECTED=Primary language"
set "TXT_CACHE_CHECK=Checking local Office cache in %BASEPATH%Office\Data"
set "TXT_CACHE_MISS=Local cache not found. Downloading the required Office files..."
set "TXT_CACHE_HIT=Local Office cache found. It will be reused."
set "TXT_DL_FAIL=ERROR: Download failed. Review log.txt."
set "TXT_INSTALL_START=Starting installation... this can take several minutes."
set "TXT_INST_FAIL=ERROR: Installation failed. Review log.txt."
set "TXT_NO_OFFICE=No supported Office ProPlus volume installation was detected."
set "TXT_OFFICE_DET=Supported Office installation detected"
set "TXT_LANG_ADDL=Enter additional language(s), comma-separated: "
set "TXT_LANG_ADDL_INVALID=Invalid language list. Use format xx-xx,yy-yy."
set "TXT_LANG_ADDL_INVALID_LOG=Additional language validation failed"
set "TXT_LANGS_SELECTED=Requested languages"
set "TXT_ADDLANG_START=Adding language(s)..."
set "TXT_ADDLANG_FAIL=ERROR: Language installation failed. Review log.txt."
set "TXT_DONE=PROCESS COMPLETED"
set "TXT_DONE_ERRORS=PROCESS FINISHED WITH ERRORS"
set "TXT_LOG=Log"
set "TXT_EXIT=Exiting by user option"
set "TXT_UNSUPPORTED_INSTALLED=Detected Office product is not a supported ProPlus 2019/2021/2024 volume installation. The script will not modify it automatically."
set "TXT_ACTIVATE_ONLY=Selected Office activation"
set "TXT_ACT_PROMPT=Activate Office now using your organization's KMS host? (Y/N): "
set "TXT_ACT_SKIP=Activation skipped by user"
set "TXT_KMS_PROMPT=Enter authorized KMS host or host:port: "
set "TXT_KMS_INVALID=Invalid KMS host. Use hostname, IPv4, or host:port."
set "TXT_KMS_INVALID_LOG=KMS host validation failed"
set "TXT_NO_OSPP=ERROR: ospp.vbs was not found. Office activation files are missing."
set "TXT_ACTIVATING=Activating Office via KMS host"
set "TXT_ACT_FAIL=ERROR: Activation failed. Review log.txt."
set "TXT_ACT_DONE=Activation completed."
set "TXT_ACTIVATE_WINDOWS_ONLY=Selected Windows activation"
set "TXT_NO_SLMGR=ERROR: slmgr.vbs was not found. Windows activation files are missing."
set "TXT_WIN_ACTIVATING=Activating Windows via KMS host"
set "TXT_WIN_ACT_FAIL=ERROR: Windows activation failed. Review log.txt."
set "TXT_WIN_ACT_DONE=Windows activation completed."
set "TXT_LICENSE_NOTE=License note: activate Microsoft products only against an authorized KMS host."
exit /b 0

:SET_TEXT_ES
set "TXT_UI_LANG=Idioma interfaz"
set "TXT_NO_SETUP=ERROR: No se encontro setup.exe"
set "TXT_VERSION_SEL=Version seleccionada"
set "TXT_PRODUCT_DET=Producto detectado"
set "TXT_INSTALLED_DET=Version instalada detectada"
set "TXT_VER_EQUAL=La version instalada es igual o superior. No hizo falta reinstalar."
set "TXT_UNINSTALL=Quitando la instalacion anterior compatible..."
set "TXT_UNINST_FAIL=ERROR: Fallo la desinstalacion. Revise log.txt."
set "TXT_LANG_BASIC=Ejemplos: es-es, en-us, fr-fr, pt-br, it-it, de-de"
set "TXT_LANG_PROMPT=Escriba el idioma principal (ejemplo: es-es): "
set "TXT_LANG_INVALID=Codigo de idioma invalido. Use formato xx-xx."
set "TXT_LANG_INVALID_LOG=Fallo la validacion del idioma principal"
set "TXT_LANG_SELECTED=Idioma principal"
set "TXT_CACHE_CHECK=Verificando cache local de Office en %BASEPATH%Office\Data"
set "TXT_CACHE_MISS=No se encontro cache local. Descargando los archivos necesarios de Office..."
set "TXT_CACHE_HIT=Se encontro cache local de Office. Se reutilizara."
set "TXT_DL_FAIL=ERROR: Fallo la descarga. Revise log.txt."
set "TXT_INSTALL_START=Iniciando instalacion... puede tardar varios minutos."
set "TXT_INST_FAIL=ERROR: Fallo la instalacion. Revise log.txt."
set "TXT_NO_OFFICE=No se detecto una instalacion Office ProPlus Volume compatible."
set "TXT_OFFICE_DET=Se detecto una instalacion Office compatible"
set "TXT_LANG_ADDL=Escriba idioma(s) adicional(es), separados por coma: "
set "TXT_LANG_ADDL_INVALID=Lista de idiomas invalida. Use formato xx-xx,yy-yy."
set "TXT_LANG_ADDL_INVALID_LOG=Fallo la validacion de idiomas adicionales"
set "TXT_LANGS_SELECTED=Idiomas solicitados"
set "TXT_ADDLANG_START=Agregando idioma(s)..."
set "TXT_ADDLANG_FAIL=ERROR: Fallo la instalacion de idioma(s). Revise log.txt."
set "TXT_DONE=PROCESO COMPLETADO"
set "TXT_DONE_ERRORS=EL PROCESO TERMINO CON ERRORES"
set "TXT_LOG=Log"
set "TXT_EXIT=Saliendo por opcion del usuario"
set "TXT_UNSUPPORTED_INSTALLED=El producto Office detectado no es una instalacion ProPlus 2019/2021/2024 Volume compatible. El script no la modificara automaticamente."
set "TXT_ACTIVATE_ONLY=Se selecciono activar Office"
set "TXT_ACT_PROMPT=Activar Office ahora usando el host KMS de su organizacion? (S/N): "
set "TXT_ACT_SKIP=La activacion fue omitida por el usuario"
set "TXT_KMS_PROMPT=Escriba el host KMS autorizado o host:puerto: "
set "TXT_KMS_INVALID=Host KMS invalido. Use hostname, IPv4 o host:puerto."
set "TXT_KMS_INVALID_LOG=Fallo la validacion del host KMS"
set "TXT_NO_OSPP=ERROR: No se encontro ospp.vbs. Faltan los archivos de activacion de Office."
set "TXT_ACTIVATING=Activando Office via host KMS"
set "TXT_ACT_FAIL=ERROR: Fallo la activacion. Revise log.txt."
set "TXT_ACT_DONE=Activacion completada."
set "TXT_ACTIVATE_WINDOWS_ONLY=Se selecciono activar Windows"
set "TXT_NO_SLMGR=ERROR: No se encontro slmgr.vbs. Faltan los archivos de activacion de Windows."
set "TXT_WIN_ACTIVATING=Activando Windows via host KMS"
set "TXT_WIN_ACT_FAIL=ERROR: Fallo la activacion de Windows. Revise log.txt."
set "TXT_WIN_ACT_DONE=Activacion de Windows completada."
set "TXT_LICENSE_NOTE=Nota de licencia: active productos Microsoft solo contra un host KMS autorizado."
exit /b 0

:PROMPT_PRIMARY_LANGUAGE
:PROMPT_PRIMARY_LANGUAGE_LOOP
echo.
echo !TXT_LANG_BASIC!
set "LANG="
set /p "LANG=!TXT_LANG_PROMPT!"
call :NORMALIZE_LANG "!LANG!"
if errorlevel 1 (
    echo !TXT_LANG_INVALID!
    call :LOG "!TXT_LANG_INVALID_LOG!: !LANG!"
    timeout /t 2 >nul
    goto PROMPT_PRIMARY_LANGUAGE_LOOP
)
set "LANG=!NORMALIZED_LANG!"
call :LOG "!TXT_LANG_SELECTED!: !LANG!"
exit /b 0

:PROMPT_EXTRA_LANGS
:PROMPT_EXTRA_LANGS_LOOP
echo.
echo !TXT_LANG_BASIC!
set "LANGS_RAW="
set "NORMALIZED_LANGS="
set /p "LANGS_RAW=!TXT_LANG_ADDL!"
call :NORMALIZE_LANG_LIST "!LANGS_RAW!"
if errorlevel 1 (
    echo !TXT_LANG_ADDL_INVALID!
    call :LOG "!TXT_LANG_ADDL_INVALID_LOG!: !LANGS_RAW!"
    timeout /t 2 >nul
    goto PROMPT_EXTRA_LANGS_LOOP
)
call :LOG "!TXT_LANGS_SELECTED!: !NORMALIZED_LANGS!"
exit /b 0

:PROMPT_POST_INSTALL_ACTIVATION
if /i "!LANG_UI!"=="EN" (
    choice /c YN /n /m "!TXT_ACT_PROMPT!"
) else (
    choice /c SN /n /m "!TXT_ACT_PROMPT!"
)

if errorlevel 2 (
    call :LOG "!TXT_ACT_SKIP!"
    exit /b 0
)

call :ACTIVATE_OFFICE
exit /b %errorlevel%

:PROMPT_KMS_ENDPOINT
:PROMPT_KMS_ENDPOINT_LOOP
echo.
set "KMS_ENDPOINT="
set /p "KMS_ENDPOINT=!TXT_KMS_PROMPT!"
call :NORMALIZE_KMS_ENDPOINT "!KMS_ENDPOINT!"
if errorlevel 1 (
    echo !TXT_KMS_INVALID!
    call :LOG "!TXT_KMS_INVALID_LOG!: !KMS_ENDPOINT!"
    timeout /t 2 >nul
    goto PROMPT_KMS_ENDPOINT_LOOP
)
exit /b 0

:NORMALIZE_LANG
set "LANG_INPUT=%~1"
set "NORMALIZED_LANG="
for /f "usebackq delims=" %%A in (`powershell -NoProfile -Command "$v=$env:LANG_INPUT; if ($v -match '^[a-z]{2}-[a-z]{2}$') { [Console]::Write($v.ToLowerInvariant()) }"`) do (
    set "NORMALIZED_LANG=%%A"
)
if not defined NORMALIZED_LANG exit /b 1
exit /b 0

:NORMALIZE_LANG_LIST
set "LANG_LIST_INPUT=%~1"
set "NORMALIZED_LANGS="
for /f "usebackq delims=" %%A in (`powershell -NoProfile -Command "$v=$env:LANG_LIST_INPUT -replace '\s+',''; if ($v -match '^[a-z]{2}-[a-z]{2}(,[a-z]{2}-[a-z]{2})*$') { [Console]::Write($v.ToLowerInvariant()) }"`) do (
    set "NORMALIZED_LANGS=%%A"
)
if not defined NORMALIZED_LANGS exit /b 1
exit /b 0

:NORMALIZE_KMS_ENDPOINT
set "KMS_ENDPOINT_INPUT=%~1"
set "NORMALIZED_KMS_ENDPOINT="
set "KMSHOST="
set "KMSPORT="
for /f "usebackq delims=" %%A in (`powershell -NoProfile -Command "$v=[string]$env:KMS_ENDPOINT_INPUT; $v=$v.Trim(); if ($v -match '^[A-Za-z0-9][A-Za-z0-9.-]*(:[0-9]{1,5})?$') { [Console]::Write($v.ToLowerInvariant()) }"`) do (
    set "NORMALIZED_KMS_ENDPOINT=%%A"
)
if not defined NORMALIZED_KMS_ENDPOINT exit /b 1
for /f "tokens=1,2 delims=:" %%A in ("!NORMALIZED_KMS_ENDPOINT!") do (
    set "KMSHOST=%%~A"
    set "KMSPORT=%%~B"
)
if not defined KMSPORT set "KMSPORT=1688"
exit /b 0

:LOCATE_OSPP
set "OFFICEPATH="
if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" set "OFFICEPATH=C:\Program Files\Microsoft Office\Office16"
if not defined OFFICEPATH if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" set "OFFICEPATH=C:\Program Files (x86)\Microsoft Office\Office16"
if not defined OFFICEPATH exit /b 1
exit /b 0

:LOCATE_SLMGR
set "SLMGR_PATH="
if exist "%SystemRoot%\Sysnative\slmgr.vbs" set "SLMGR_PATH=%SystemRoot%\Sysnative\slmgr.vbs"
if not defined SLMGR_PATH if exist "%SystemRoot%\System32\slmgr.vbs" set "SLMGR_PATH=%SystemRoot%\System32\slmgr.vbs"
if not defined SLMGR_PATH exit /b 1
exit /b 0

:ACTIVATE_OFFICE
call :LOCATE_OSPP
if errorlevel 1 (
    call :LOG "!TXT_NO_OSPP!"
    exit /b 1
)

call :PROMPT_KMS_ENDPOINT
if errorlevel 1 (
    exit /b 1
)

cd /d "!OFFICEPATH!"
call :LOG "!TXT_ACTIVATING!: !NORMALIZED_KMS_ENDPOINT!"
cscript //nologo ospp.vbs /sethst:!KMSHOST! >> "%LOGFILE%" 2>&1
if errorlevel 1 (
    call :LOG "!TXT_ACT_FAIL!"
    exit /b 1
)
cscript //nologo ospp.vbs /setprt:!KMSPORT! >> "%LOGFILE%" 2>&1
if errorlevel 1 (
    call :LOG "!TXT_ACT_FAIL!"
    exit /b 1
)
cscript //nologo ospp.vbs /act >> "%LOGFILE%" 2>&1
if errorlevel 1 (
    call :LOG "!TXT_ACT_FAIL!"
    exit /b 1
)
cscript //nologo ospp.vbs /dstatus >> "%LOGFILE%" 2>&1
call :LOG "!TXT_ACT_DONE!"
exit /b 0

:ACTIVATE_WINDOWS
call :LOCATE_SLMGR
if errorlevel 1 (
    call :LOG "!TXT_NO_SLMGR!"
    exit /b 1
)

call :PROMPT_KMS_ENDPOINT
if errorlevel 1 (
    exit /b 1
)

call :LOG "!TXT_WIN_ACTIVATING!: !KMSHOST!:!KMSPORT!"
cscript //nologo "!SLMGR_PATH!" /skms !KMSHOST!:!KMSPORT! >> "%LOGFILE%" 2>&1
if errorlevel 1 (
    call :LOG "!TXT_WIN_ACT_FAIL!"
    exit /b 1
)
cscript //nologo "!SLMGR_PATH!" /ato >> "%LOGFILE%" 2>&1
if errorlevel 1 (
    call :LOG "!TXT_WIN_ACT_FAIL!"
    exit /b 1
)
cscript //nologo "!SLMGR_PATH!" /xpr >> "%LOGFILE%" 2>&1
call :LOG "!TXT_WIN_ACT_DONE!"
exit /b 0

:ENSURE_LOCAL_CACHE
set "REQ_PRODUCTID=%~1"
set "REQ_CHANNEL=%~2"
set "REQ_LANGS=%~3"

call :LOG "!TXT_CACHE_CHECK!"
if exist "%BASEPATH%Office\Data\*" (
    call :LOG "!TXT_CACHE_HIT!"
    exit /b 0
)

call :LOG "!TXT_CACHE_MISS!"
call :WRITE_LANG_XML "%BASEPATH%download.xml" "!REQ_PRODUCTID!" "!REQ_CHANNEL!" "!REQ_LANGS!" download
call :RUN_SETUP /download "%BASEPATH%download.xml"
if errorlevel 1 (
    call :LOG "!TXT_DL_FAIL!"
    exit /b 1
)

exit /b 0

:DETECT_INSTALLED_OFFICE
set "PRODUCTLINE="
set "INSTALLED_PRODUCT_ID="
set "INSTALLEDVER="
set "INSTALLED_CHANNEL="

for /f "skip=2 tokens=1,2,*" %%A in ('reg query "HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v ProductReleaseIds 2^>nul') do (
    if /i "%%A"=="ProductReleaseIds" if /i "%%B"=="REG_SZ" set "PRODUCTLINE=%%C"
)

if not defined PRODUCTLINE exit /b 0

call :MAP_INSTALLED_PRODUCT "!PRODUCTLINE!"
exit /b 0

:MAP_INSTALLED_PRODUCT
set "LINE=%~1"

echo(!LINE! | findstr /i /c:"ProPlus2024Volume" >nul && (
    set "INSTALLED_PRODUCT_ID=ProPlus2024Volume"
    set "INSTALLEDVER=2024"
    set "INSTALLED_CHANNEL=PerpetualVL2024"
    exit /b 0
)

echo(!LINE! | findstr /i /c:"ProPlus2021Volume" >nul && (
    set "INSTALLED_PRODUCT_ID=ProPlus2021Volume"
    set "INSTALLEDVER=2021"
    set "INSTALLED_CHANNEL=PerpetualVL2021"
    exit /b 0
)

echo(!LINE! | findstr /i /c:"ProPlus2019Volume" >nul && (
    set "INSTALLED_PRODUCT_ID=ProPlus2019Volume"
    set "INSTALLEDVER=2019"
    set "INSTALLED_CHANNEL=PerpetualVL2019"
    exit /b 0
)

echo(!LINE! | findstr /i /c:"2024" >nul && (
    set "INSTALLEDVER=2024"
    set "INSTALLED_CHANNEL=PerpetualVL2024"
    exit /b 0
)

echo(!LINE! | findstr /i /c:"2021" >nul && (
    set "INSTALLEDVER=2021"
    set "INSTALLED_CHANNEL=PerpetualVL2021"
    exit /b 0
)

echo(!LINE! | findstr /i /c:"2019" >nul && (
    set "INSTALLEDVER=2019"
    set "INSTALLED_CHANNEL=PerpetualVL2019"
    exit /b 0
)

exit /b 0

:WRITE_REMOVE_XML
set "TARGET=%~1"
> "%TARGET%" (
    echo ^<Configuration^>
    echo   ^<Remove All="TRUE" /^>
    echo   ^<Display Level="None" AcceptEULA="TRUE" /^>
    echo   ^<Property Name="FORCEAPPSHUTDOWN" Value="TRUE" /^>
    echo ^</Configuration^>
)
call :LOG "Created %~nx1"
exit /b 0

:WRITE_LANG_XML
set "TARGET=%~1"
set "XML_PRODUCTID=%~2"
set "XML_CHANNEL=%~3"
set "XML_LANGS=%~4"
set "XML_MODE=%~5"
set "XML_LANG_LIST=!XML_LANGS:,= !"

> "%TARGET%" (
    echo ^<Configuration^>
    echo   ^<Add OfficeClientEdition="64" Channel="!XML_CHANNEL!" SourcePath="!BASEPATH!"^>
    echo     ^<Product ID="!XML_PRODUCTID!"^>
)

for %%L in (!XML_LANG_LIST!) do (
    >> "%TARGET%" echo       ^<Language ID="%%~L" /^>
)

>> "%TARGET%" (
    echo     ^</Product^>
    echo   ^</Add^>
)

if /i "!XML_MODE!"=="configure" (
    >> "%TARGET%" (
        echo   ^<Display Level="None" AcceptEULA="TRUE" /^>
        echo   ^<Property Name="FORCEAPPSHUTDOWN" Value="TRUE" /^>
    )
)

>> "%TARGET%" echo ^</Configuration^>
call :LOG "Created %~nx1 (Channel=!XML_CHANNEL!, ProductID=!XML_PRODUCTID!, Languages=!XML_LANGS!, Mode=!XML_MODE!)"
exit /b 0

:RUN_SETUP
call :LOG "Running: setup.exe %~1 %~2"
"%OFFICE_SETUP%" %~1 "%~2" >> "%LOGFILE%" 2>&1
exit /b %errorlevel%

:REQUIRE_SETUP
if exist "%OFFICE_SETUP%" exit /b 0
call :LOG "!TXT_NO_SETUP!"
exit /b 1

:INIT_LOG
> "%LOGFILE%" echo ============================================================
>> "%LOGFILE%" echo Start: %DATE% %TIME%
>> "%LOGFILE%" echo ============================================================
>> "%LOGFILE%" echo Script: %~nx0
>> "%LOGFILE%" echo Version: %INSTALLER_VERSION%
>> "%LOGFILE%" echo Author: %CREATOR%
>> "%LOGFILE%" echo BasePath: %BASEPATH%
>> "%LOGFILE%" echo.
exit /b 0

:LOG
set "MSG=%~1"
echo %MSG%
>> "%LOGFILE%" echo %DATE% %TIME% - %MSG%
exit /b 0

:ENSURE_ADMIN
net session >nul 2>&1
if not errorlevel 1 exit /b 0
echo Solicitando permisos de administrador / Requesting administrator permissions...
powershell -NoProfile -Command "Start-Process -FilePath $env:ComSpec -ArgumentList '/c','\"\"%~f0\"\"' -WorkingDirectory '%~dp0' -Verb RunAs"
exit /b 1
