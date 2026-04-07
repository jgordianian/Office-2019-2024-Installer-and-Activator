# Office 2019-2024 Installer and Activator

## English

This repository contains a Windows batch script to:

- Install Microsoft Office 2019 Volume
- Install Microsoft Office 2021 Volume
- Install Microsoft Office 2024 Volume
- Add additional Office languages
- Activate Office through an authorized KMS host
- Activate Windows through an authorized KMS host

The main entry point is `run.cmd`.

## Requirements

- Windows with administrator rights
- `setup.exe` from the Office Deployment Tool in the same folder as the script
- Internet access when the local Office cache is missing
- A valid and authorized KMS host if activation is used

## Included Files

- `run.cmd`: Main interactive installer and activation script
- `setup.exe`: Office Deployment Tool executable
- `configuration-Office365-x64.xml`: Additional ODT configuration file
- `install.xml`: Generated or reused Office install configuration
- `addlang.xml`: Generated or reused language install configuration
- `log.txt`: Execution log

## Features

- Bilingual interface: Spanish and English
- Automatic elevation to administrator when needed
- Detection of installed supported Office volume editions
- Removal of an older supported Office volume edition before upgrading
- Local Office cache reuse from `Office\Data` when available
- Language code validation
- KMS host validation with optional custom port
- Centralized logging to `log.txt`

## Supported Office Products

- `ProPlus2019Volume`
- `ProPlus2021Volume`
- `ProPlus2024Volume`

If another Office edition is detected, the script will not modify it automatically.

## How to Use

1. Place `setup.exe` in the same folder as `run.cmd`.
2. Run `run.cmd`.
3. Accept the administrator prompt if Windows requests elevation.
4. Select the interface language.
5. Choose one of the menu options:
   - `1`: Install Office 2019 Volume
   - `2`: Install Office 2021 Volume
   - `3`: Install Office 2024 Volume
   - `4`: Install an additional Office language
   - `5`: Activate Office using a KMS host
   - `6`: Activate Windows using a KMS host
   - `7`: Exit
6. Follow the on-screen prompts.

## Language Format

Use language tags in the format `xx-xx`, for example:

- `es-es`
- `en-us`
- `fr-fr`
- `pt-br`
- `it-it`
- `de-de`

For additional languages, enter a comma-separated list such as:

```text
es-es,en-us,fr-fr
```

## KMS Host Format

You can enter:

- `kms.company.local`
- `kms.company.local:1688`
- `192.168.1.10`
- `192.168.1.10:1688`

If no port is entered, the script uses port `1688`.

## Logging

All operations are written to `log.txt`, including:

- Installation steps
- Cache detection
- Download attempts
- Activation attempts
- Errors returned by `setup.exe`, `ospp.vbs`, or `slmgr.vbs`

## Important Notes

- Use this script only with properly licensed Microsoft products.
- Use KMS activation only against an authorized KMS host.
- Office activation uses `ospp.vbs`.
- Windows activation uses `slmgr.vbs`.
- The script is intended for supported Office volume editions only.

---

## Espanol

Este repositorio contiene un script por lotes para Windows que permite:

- Instalar Microsoft Office 2019 Volume
- Instalar Microsoft Office 2021 Volume
- Instalar Microsoft Office 2024 Volume
- Agregar idiomas adicionales de Office
- Activar Office mediante un host KMS autorizado
- Activar Windows mediante un host KMS autorizado

El archivo principal es `run.cmd`.

## Requisitos

- Windows con permisos de administrador
- `setup.exe` del Office Deployment Tool en la misma carpeta que el script
- Conexion a Internet cuando no exista cache local de Office
- Un host KMS valido y autorizado si se va a activar

## Archivos Incluidos

- `run.cmd`: Script principal interactivo de instalacion y activacion
- `setup.exe`: Ejecutable del Office Deployment Tool
- `configuration-Office365-x64.xml`: Archivo adicional de configuracion de ODT
- `install.xml`: Configuracion generada o reutilizada para instalar Office
- `addlang.xml`: Configuracion generada o reutilizada para instalar idiomas
- `log.txt`: Registro de ejecucion

## Caracteristicas

- Interfaz bilingue: espanol e ingles
- Elevacion automatica a administrador cuando hace falta
- Deteccion de ediciones Volume de Office compatibles ya instaladas
- Eliminacion de una instalacion compatible anterior antes de actualizar
- Reutilizacion de cache local de Office desde `Office\Data` cuando exista
- Validacion de codigos de idioma
- Validacion de host KMS con puerto opcional
- Registro centralizado en `log.txt`

## Productos Office Compatibles

- `ProPlus2019Volume`
- `ProPlus2021Volume`
- `ProPlus2024Volume`

Si se detecta otra edicion de Office, el script no la modificara automaticamente.

## Uso

1. Coloque `setup.exe` en la misma carpeta que `run.cmd`.
2. Ejecute `run.cmd`.
3. Acepte la solicitud de administrador si Windows pide elevacion.
4. Seleccione el idioma de la interfaz.
5. Elija una de las opciones del menu:
   - `1`: Instalar Office 2019 Volume
   - `2`: Instalar Office 2021 Volume
   - `3`: Instalar Office 2024 Volume
   - `4`: Instalar un idioma adicional de Office
   - `5`: Activar Office usando un host KMS
   - `6`: Activar Windows usando un host KMS
   - `7`: Salir
6. Siga las instrucciones en pantalla.

## Formato de Idioma

Use etiquetas de idioma con formato `xx-xx`, por ejemplo:

- `es-es`
- `en-us`
- `fr-fr`
- `pt-br`
- `it-it`
- `de-de`

Para idiomas adicionales, escriba una lista separada por comas como:

```text
es-es,en-us,fr-fr
```

## Formato del Host KMS

Puede escribir:

- `kms.empresa.local`
- `kms.empresa.local:1688`
- `192.168.1.10`
- `192.168.1.10:1688`

Si no se especifica puerto, el script usa el puerto `1688`.

## Registro

Todas las operaciones se escriben en `log.txt`, incluyendo:

- Pasos de instalacion
- Deteccion de cache
- Intentos de descarga
- Intentos de activacion
- Errores devueltos por `setup.exe`, `ospp.vbs` o `slmgr.vbs`

## Notas Importantes

- Use este script solo con productos Microsoft correctamente licenciados.
- Use activacion KMS solo contra un host KMS autorizado.
- La activacion de Office usa `ospp.vbs`.
- La activacion de Windows usa `slmgr.vbs`.
- El script esta pensado solo para ediciones Office Volume compatibles.
