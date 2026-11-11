<#
.SYNOPSIS
    Arma el paquete de Planillas para instalarla en un servidor Windows (ver instalar.ps1).

.DESCRIPTION
    Se corre en la PC de desarrollo, desde la carpeta del proyecto:
        powershell -ExecutionPolicy Bypass -File servidor\empaquetar.ps1
    Deja paquetes\planillas-<fecha>.zip con la app, estos scripts y VERSION.txt (commit y fecha, para saber qué
    versión corre en el servidor). No lleva pruebas, diagnósticos ni archivos temporales.
#>
param([string]$Salida)
$ErrorActionPreference = 'Stop'

$repo = Split-Path -Parent $PSScriptRoot
if (-not $Salida) { $Salida = Join-Path $repo 'paquetes' }

function Paso([string]$texto) { Write-Host "`n== $texto" -ForegroundColor Cyan }

# El Tesseract portable (motor + idioma) viaja en el paquete (no en git; pesa ~178 MB). Sin él, los PDF escaneados no se
# pueden leer en el servidor.
$tessExe = Join-Path $repo 'tesseract\tesseract.exe'
$spa = Join-Path $repo 'tesseract\tessdata\spa.traineddata'
if (-not (Test-Path $tessExe) -or -not (Test-Path $spa)) {
    throw "Falta el Tesseract portable en $repo\tesseract (tesseract.exe + tessdata\spa.traineddata). Armalo una vez (ver README/.gitignore) y volvé a empaquetar."
}

Paso 'Armando el paquete'
$fecha = Get-Date -Format 'yyyyMMdd-HHmm'
$temporal = Join-Path ([IO.Path]::GetTempPath()) "planillas-$fecha"
if (Test-Path $temporal) { Remove-Item -Recurse -Force $temporal }
New-Item -ItemType Directory -Force -Path $temporal, $Salida | Out-Null

foreach ($carpeta in 'api', 'servidor', 'tesseract') { Copy-Item -Recurse (Join-Path $repo $carpeta) (Join-Path $temporal $carpeta) }
foreach ($archivo in 'resaltado_pdf.py', 'requirements.txt', 'requirements-servidor.lock') { Copy-Item (Join-Path $repo $archivo) $temporal }
Get-ChildItem $temporal -Recurse -Directory -Filter '__pycache__' | Remove-Item -Recurse -Force

Push-Location $repo
try {
    $commit = (git rev-parse --short HEAD) 2>$null
    $cambios = @(git status --porcelain 2>$null).Count
} finally { Pop-Location }
$sinCommit = if ($cambios) { "sí ($cambios archivos)" } else { 'no' }
@(
    "Paquete: planillas-$fecha"
    "Commit: $commit"
    "Cambios sin commit: $sinCommit"
    "Armado: $(Get-Date -Format 'yyyy-MM-dd HH:mm') por $env:USERNAME en $env:COMPUTERNAME"
) | Set-Content -Path (Join-Path $temporal 'VERSION.txt') -Encoding UTF8

$zip = Join-Path $Salida "planillas-$fecha.zip"
if (Test-Path $zip) { Remove-Item $zip }
Compress-Archive -Path (Join-Path $temporal '*') -DestinationPath $zip
Remove-Item -Recurse -Force $temporal
$tamano = [math]::Round((Get-Item $zip).Length / 1MB, 1)
Paso "Listo: $zip ($tamano MB)"
Write-Host '   Cópialo a C:\Planillas del servidor junto con servidor\instalar.ps1 y corre allá el instalador.'
