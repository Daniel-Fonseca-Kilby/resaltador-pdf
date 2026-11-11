<#
.SYNOPSIS
    Reinicia (o solo detiene) Planillas en el servidor.

.DESCRIPTION
    En PowerShell como administrador:
        powershell -ExecutionPolicy Bypass -File C:\Planillas\app\servidor\reiniciar.ps1
        powershell -ExecutionPolicy Bypass -File C:\Planillas\app\servidor\reiniciar.ps1 -SoloDetener
    Detener la tarea programada no cierra Python (lo lanza iniciar.ps1 y queda vivo ocupando el puerto): este
    script cierra también esos procesos antes de volver a arrancar.
#>
param(
    [int]$Puerto = 8001,
    [switch]$SoloDetener
)
$ErrorActionPreference = 'Stop'
$Tarea = 'Planillas VMA'

if (Get-ScheduledTask -TaskName $Tarea -ErrorAction SilentlyContinue) { Stop-ScheduledTask -TaskName $Tarea }
Get-CimInstance Win32_Process -Filter "Name = 'python.exe'" |
    Where-Object { $_.CommandLine -like '*servidor.arrancar*' -and $_.CommandLine -like "*--puerto $Puerto *" } |
    ForEach-Object { Write-Host "Cierro Python (proceso $($_.ProcessId))"; Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
if ($SoloDetener) { Write-Host 'Planillas detenida.'; return }

Start-Sleep -Seconds 2
Start-ScheduledTask -TaskName $Tarea
Write-Host 'Arrancando Planillas...'
$limite = (Get-Date).AddMinutes(2)
$responde = $false
while (-not $responde -and (Get-Date) -lt $limite) {
    Start-Sleep -Seconds 3
    try { $responde = (Invoke-WebRequest "http://localhost:$Puerto/salud" -UseBasicParsing -TimeoutSec 5).StatusCode -eq 200 }
    catch { $responde = $false }
}
if ($responde) { Write-Host "Planillas responde: http://$($env:COMPUTERNAME):$Puerto" }
else { Write-Warning 'Planillas no respondió en 2 minutos: revisa C:\Planillas\logs (errores.log, app.log, arranque.log).'; exit 1 }
