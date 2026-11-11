<#
.SYNOPSIS
    Arranca Planillas en un servidor Windows y la vuelve a levantar si se cae.

.DESCRIPTION
    Lo lanza la tarea programada «Planillas VMA» que crea instalar.ps1. Registros en <Raiz>\logs:
        app.log          lo que hace la app (lotes procesados, tiempos, errores); uno por día, se guardan 60
        arranque.log     cuándo arrancó y cuándo se cayó
        errores.log      lo que Python escribió antes de poder registrar (p. ej. un error al arrancar)
    Los archivos temporales de cada lote van a <Raiz>\temp (la app los borra al terminar).
    Para probar a mano (PowerShell como administrador):
        powershell -ExecutionPolicy Bypass -File C:\Planillas\app\servidor\iniciar.ps1
#>
param(
    [string]$Raiz = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)),
    [int]$Puerto = 8001
)
$ErrorActionPreference = 'Stop'

$app = Join-Path $Raiz 'app'
$python = Join-Path $Raiz 'venv\Scripts\python.exe'
$logs = Join-Path $Raiz 'logs'
$temp = Join-Path $Raiz 'temp'
New-Item -ItemType Directory -Force -Path $logs, $temp | Out-Null

function Anotar([string]$texto) {
    Add-Content -Path (Join-Path $logs 'arranque.log') -Encoding UTF8 -Value ('{0:yyyy-MM-dd HH:mm:ss} {1}' -f (Get-Date), $texto)
}

# Configuración opcional en <Raiz>\.env (p. ej. DEBUG_GUARDAR_SUBIDAS=1). Las vacías no se definen.
$archivoEnv = Join-Path $Raiz '.env'
if (Test-Path $archivoEnv) {
    foreach ($linea in Get-Content $archivoEnv -Encoding UTF8) {
        if ($linea -match '^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=(.*)$') {
            $valor = $Matches[2].Trim().Trim('"')
            if ($valor) { [Environment]::SetEnvironmentVariable($Matches[1], $valor, 'Process') }
        }
    }
}
$env:TEMP = $temp; $env:TMP = $temp   # los lotes temporales quedan en una carpeta propia, fácil de revisar
$env:PYTHONUTF8 = '1'
$env:PYTHONUNBUFFERED = '1'
$env:EN_SERVIDOR = '1'   # sin inicio de sesión la app no arranca (ver seguridad_web.py)

# Restos de lotes cortados (p. ej. por un reinicio a mitad de proceso) y registros viejos.
Get-ChildItem $temp -Force -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-1) } |
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
Get-ChildItem $logs -File | Where-Object { $_.LastWriteTime -lt (Get-Date).AddDays(-60) } | Remove-Item -ErrorAction SilentlyContinue

$argumentos = "-m servidor.arrancar --puerto $Puerto --registros `"$logs`""
$errores = Join-Path $logs 'errores.log'
while ($true) {
    if (Test-Path $errores) { Move-Item $errores (Join-Path $logs 'errores.anterior.log') -Force }
    Anotar "Arranca Planillas en el puerto $Puerto"
    $proceso = Start-Process -FilePath $python -ArgumentList $argumentos -WorkingDirectory $app -NoNewWindow -PassThru -Wait `
        -RedirectStandardError $errores -RedirectStandardOutput (Join-Path $logs 'salida.log')
    Anotar "Planillas terminó (código $($proceso.ExitCode)). Se vuelve a levantar en 30 segundos; ver errores.log"
    Start-Sleep -Seconds 30
}
