<#
.SYNOPSIS
    Instala o actualiza Planillas (Resaltador de Planillas VMA) en un servidor Windows.

.DESCRIPTION
    Se corre en PowerShell como administrador, la primera vez y en cada actualización, con el paquete que arma
    empaquetar.ps1. Copiar antes a C:\Planillas el ZIP y este instalar.ps1, y correrlo desde ahí (no desde
    C:\Planillas\app\servidor, que se borra y se reemplaza):
        cd C:\Planillas
        powershell -ExecutionPolicy Bypass -File .\instalar.ps1 -Paquete C:\Planillas\planillas-20261102-1500.zip

    Deja todo en <Raiz> (por defecto C:\Planillas):
        app\     el código (se reemplaza en cada actualización)
        venv\    Python de la app y sus librerías
        logs\    registros (ver iniciar.ps1)
        temp\    archivos temporales de cada lote
    y la tarea programada «Planillas VMA», que arranca la app con el servidor y la vuelve a levantar si se cae. La
    tarea corre con una cuenta local SIN privilegios (svc_vma_planillas), no como SYSTEM: si comprometen la app, no se
    llevan la máquina. El instalador crea la cuenta, le pone una contraseña fuerte (que no ves ni se guarda en texto)
    y le da permiso solo sobre esta carpeta. No guarda datos: no hay base de datos ni configuración obligatoria.

    -SinTarea   para probar sin permisos de administrador: no crea la tarea ni la regla del firewall, ni arranca.
    -SinLogin   solo para la primera instalación, antes de activar el inicio de sesión: sin esto, la app no arranca
                en el servidor sin inicio de sesión. preparar_sesion.ps1 lo apaga solo.
#>
param(
    [Parameter(Mandatory = $true)][string]$Paquete,
    [string]$Raiz = 'C:\Planillas',
    [int]$Puerto = 8001,
    [string]$Python = 'C:\Program Files\Python314\python.exe',
    [switch]$SinTarea,
    [switch]$SinLogin
)
$ErrorActionPreference = 'Stop'
$Tarea = 'Planillas VMA'
$CuentaServicio = 'svc_vma_planillas'   # cuenta local sin privilegios que corre la tarea (en vez de SYSTEM)

function Paso([string]$texto) { Write-Host "`n== $texto" -ForegroundColor Cyan }

function Correr([string]$programa, [string[]]$argumentos, [string]$siFalla) {
    # Los programas de consola escriben avisos en stderr; en PowerShell 5.1, con 'Stop', eso cortaría el script.
    $anterior = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    & $programa @argumentos
    $codigo = $LASTEXITCODE
    $ErrorActionPreference = $anterior
    if ($codigo -ne 0) { throw "$siFalla (código $codigo)." }
}

function Nueva-Clave([int]$largo = 24) {
    # Contraseña fuerte para la cuenta de servicio (mayúsculas, minúsculas, dígitos y símbolos). No se muestra ni se
    # registra: queda solo dentro de la tarea, cifrada por Windows. Se omiten caracteres confusos (O/0, l/1).
    $may = 'ABCDEFGHJKLMNPQRSTUVWXYZ'; $min = 'abcdefghijkmnpqrstuvwxyz'
    $dig = '23456789'; $sim = '!#$%&*+-=?@'
    $todos = "$may$min$dig$sim"
    $rnd = [Security.Cryptography.RandomNumberGenerator]::Create()
    function Uno([string]$conjunto) {
        $tope = [int][math]::Floor(256 / $conjunto.Length) * $conjunto.Length   # evita el sesgo del módulo
        $b = [byte[]]::new(1)
        do { $rnd.GetBytes($b) } while ($b[0] -ge $tope)
        $conjunto[$b[0] % $conjunto.Length]
    }
    $chars = [System.Collections.Generic.List[char]]::new()
    $chars.Add((Uno $may)); $chars.Add((Uno $min)); $chars.Add((Uno $dig)); $chars.Add((Uno $sim))
    while ($chars.Count -lt $largo) { $chars.Add((Uno $todos)) }
    $rnd.Dispose()
    -join ($chars | Sort-Object { Get-Random })   # mezcla: los obligatorios no quedan siempre al inicio
}

function Asegurar-Derechos([string]$cuenta, [string[]]$derechos) {
    # Agrega la cuenta (por SID) a cada derecho de la política local, de forma aditiva (conserva a quien ya lo tuviera).
    # SeBatchLogonRight: para que la tarea con contraseña arranque (sin esto no arranca; lo confirmamos en JUPITER).
    # SeDeny*: una credencial robada no sirve para entrar interactivo, por RDP ni por la red (defensa en profundidad).
    # secedit fija cada derecho al conjunto del INF, por eso se leen los actuales y se les suma el SID.
    $ErrorActionPreference = 'Continue'   # secedit puede escribir en stderr; con 'Stop' cortaría el script (ver Correr)
    $sid = '*' + (New-Object Security.Principal.NTAccount($cuenta)).Translate([Security.Principal.SecurityIdentifier]).Value
    $base = Join-Path $env:TEMP "vma-derechos-$PID"
    $cfg = "$base.cfg"; $inf = "$base.inf"; $db = "$base.sdb"
    try {
        & secedit /export /cfg $cfg /areas USER_RIGHTS /quiet | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "No pude leer los derechos de usuario (secedit $LASTEXITCODE)." }
        $lineas = Get-Content $cfg
        $cambios = @()
        foreach ($d in $derechos) {
            $actual = $lineas | Where-Object { $_ -match "^\s*$d\s*=" } | Select-Object -First 1
            if ($actual -match [regex]::Escape($sid)) { continue }   # ya lo tiene
            $valor = if ($actual) { ($actual -replace "^\s*$d\s*=\s*", '').Trim() + ",$sid" } else { $sid }
            $cambios += "$d = $valor"
        }
        if (-not $cambios) { return }   # ya estaban todos
        Set-Content $inf -Encoding Unicode -Value (@('[Unicode]', 'Unicode=yes', '[Version]', 'signature="$CHICAGO$"',
            'Revision=1', '[Privilege Rights]') + $cambios)
        & secedit /configure /db $db /cfg $inf /areas USER_RIGHTS /quiet | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "No pude ajustar los derechos de la cuenta (secedit $LASTEXITCODE)." }
    } finally {
        Remove-Item $cfg, $inf, $db -Force -ErrorAction SilentlyContinue
    }
}

function Asegurar-CuentaServicio([string]$cuenta, [string]$raiz, [string[]]$escribibles) {
    # Crea (o pone al día) una cuenta local SIN privilegios para correr la tarea, en vez de SYSTEM. Devuelve la
    # contraseña nueva para registrar la tarea. El permiso «iniciar como proceso por lotes» lo da sola la tarea al
    # registrarse con contraseña. Mínimo privilegio: lee su carpeta, escribe solo en logs/temp, nada más.
    $clave = Nueva-Clave
    $segura = ConvertTo-SecureString $clave -AsPlainText -Force
    if (Get-LocalUser -Name $cuenta -ErrorAction SilentlyContinue) {
        Set-LocalUser -Name $cuenta -Password $segura -PasswordNeverExpires $true
    } else {
        New-LocalUser -Name $cuenta -Password $segura -FullName "Servicio web VMA ($cuenta)" `
            -Description 'Servicio web VMA, sin privilegios' -PasswordNeverExpires | Out-Null
    }
    # En «Usuarios» (por SID, sin depender del idioma) para poder leer el Python instalado «para todos». Si la cuenta ya
    # estaba en el grupo (reinstalación), se ignora el error de «ya es miembro».
    $usuarios = (Get-LocalGroup -SID 'S-1-5-32-545').Name
    Add-LocalGroupMember -Group $usuarios -Member $cuenta -ErrorAction SilentlyContinue
    Asegurar-Derechos $cuenta @('SeBatchLogonRight', 'SeDenyInteractiveLogonRight', 'SeDenyRemoteInteractiveLogonRight', 'SeDenyNetworkLogonRight')   # lotes para arrancar; deny para que una credencial robada no sirva para entrar
    $nombre = "$env:COMPUTERNAME\$cuenta"
    function Permisos([string[]]$lista, [string]$siFalla) {
        # OJO: no llamar a esta función «Icacls»: una función con ese nombre gana sobre icacls.exe y «& icacls» de
        # adentro se llamaría a sí misma (recursión infinita, CallDepthOverflow).
        & icacls @lista | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "$siFalla (icacls código $LASTEXITCODE)." }
    }
    # Leer y ejecutar en toda la raíz; escribir solo en las carpetas de datos/registros.
    Permisos @($raiz, '/grant:r', "${nombre}:(OI)(CI)(RX)") "No pude dar permisos de lectura en $raiz"
    foreach ($sub in $escribibles) {
        $ruta = Join-Path $raiz $sub
        New-Item -ItemType Directory -Force -Path $ruta | Out-Null
        Permisos @($ruta, '/grant:r', "${nombre}:(OI)(CI)(M)") "No pude dar permisos de escritura en $ruta"
    }
    # El .env trae la clave de sesión: solo Administradores, SYSTEM y esta cuenta (saca al grupo Users).
    $archivoEnv = Join-Path $raiz '.env'
    if (Test-Path $archivoEnv) {
        if (-not (Test-Path "$archivoEnv.aclbak")) { Get-Acl $archivoEnv | Export-Clixml "$archivoEnv.aclbak" }
        Permisos @($archivoEnv, '/inheritance:r', '/grant:r', '*S-1-5-32-544:(F)', '*S-1-5-18:(R)', "${nombre}:(R)") "No pude proteger $archivoEnv"
        & icacls $archivoEnv '/remove:g' '*S-1-5-32-545' | Out-Null
    }
    return $clave
}

# --- Comprobaciones ------------------------------------------------------------------------------------------------
if (-not $SinTarea) {
    $identidad = [Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    if (-not $identidad.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        throw 'Abre PowerShell con «Ejecutar como administrador» y vuelve a correr el script.'
    }
}
$Paquete = (Resolve-Path $Paquete).Path
if (-not (Test-Path $Python)) { throw "No encuentro Python en $Python. Instala Python 3.14 «para todos los usuarios»." }
if ($Python -like "$env:SystemDrive\Users\*" -and -not $SinTarea) {
    throw "Python está instalado solo para un usuario ($Python). La tarea corre con la cuenta de servicio: reinstálalo «para todos los usuarios»."
}

$app = Join-Path $Raiz 'app'
$venv = Join-Path $Raiz 'venv'
$pythonApp = Join-Path $venv 'Scripts\python.exe'
New-Item -ItemType Directory -Force -Path $Raiz, (Join-Path $Raiz 'logs'), (Join-Path $Raiz 'temp') | Out-Null

# --- Detener lo que esté corriendo ---------------------------------------------------------------------------------
Paso 'Deteniendo Planillas, si estaba corriendo'
if (-not $SinTarea -and (Get-ScheduledTask -TaskName $Tarea -ErrorAction SilentlyContinue)) {
    Stop-ScheduledTask -TaskName $Tarea
}
# Detener la tarea no siempre cierra Python (lo lanza iniciar.ps1): se busca por su línea de comandos.
Get-CimInstance Win32_Process -Filter "Name = 'python.exe'" |
    Where-Object { $_.CommandLine -like '*servidor.arrancar*' -and $_.CommandLine -like "*--puerto $Puerto *" } |
    ForEach-Object { Write-Host "Cierro Python (proceso $($_.ProcessId))"; Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }
Start-Sleep -Seconds 2

# --- Código --------------------------------------------------------------------------------------------------------
Paso "Copiando el código de $(Split-Path -Leaf $Paquete)"
$nuevo = Join-Path $Raiz 'app.nuevo'
if (Test-Path $nuevo) { Remove-Item -Recurse -Force $nuevo }
Expand-Archive -Path $Paquete -DestinationPath $nuevo
Get-ChildItem $nuevo -Recurse -File | Unblock-File   # el ZIP pudo llegar marcado como «bajado de internet»
if (-not (Test-Path (Join-Path $nuevo 'api\index.py'))) { throw 'El paquete no tiene la app (falta api\index.py).' }
if (Test-Path $app) { Remove-Item -Recurse -Force $app }
Rename-Item $nuevo 'app'
Get-Content (Join-Path $app 'VERSION.txt') | ForEach-Object { Write-Host "   $_" }

# --- Python y librerías --------------------------------------------------------------------------------------------
Paso 'Preparando Python y las librerías (la primera vez tarda unos minutos)'
if (-not (Test-Path $pythonApp)) { Correr $Python @('-m', 'venv', $venv) 'No se pudo crear el entorno de Python' }
Correr $pythonApp @('-m', 'pip', 'install', '--disable-pip-version-check', '-q', '--require-hashes', '-r', (Join-Path $app 'requirements-servidor.lock')) `
    'No se pudieron instalar las librerías (¿llega el servidor a pypi.org?)'

Paso 'Probando que la app carga'
# Se exige ver el mensaje, no solo el código de salida: un antivirus que aísla el python.exe del entorno lo deja
# «terminar bien» sin hacer nada (pasó en la PC de desarrollo con Xcitium).
Push-Location $app
try {
    $anterior = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $salida = & $pythonApp -c "from api.index import app; import waitress, pymupdf; print('Carga bien. PyMuPDF', pymupdf.VersionBind)" 2>&1 | Out-String
    $ErrorActionPreference = $anterior
} finally { Pop-Location }
Write-Host ($salida.Trim() -replace '(?m)^', '   ')
if ($salida -notmatch 'Carga bien') {
    throw "La app no cargó. Si arriba no apareció ningún mensaje, el antivirus pudo haber bloqueado $pythonApp (pedir a TI que lo permita)."
}

if ($SinTarea) {
    Paso 'Listo (sin tarea programada). Para arrancar a mano:'
    Write-Host "   powershell -ExecutionPolicy Bypass -File `"$app\servidor\iniciar.ps1`" -Raiz `"$Raiz`" -Puerto $Puerto"
    return
}

# --- Inicio de sesión ------------------------------------------------------------------------------------------------
$archivoEnv = Join-Path $Raiz '.env'
$conLogin = (Test-Path $archivoEnv) -and (Select-String -Path $archivoEnv -Pattern '^\s*SESION_VMA_SECRETO\s*=\s*\S' -Quiet)
if (-not $conLogin) {
    if ($SinLogin) {
        Add-Content -Path $archivoEnv -Value 'PERMITIR_SIN_LOGIN=1'
        Write-Warning 'Planillas queda SIN inicio de sesión hasta correr preparar_sesion.ps1 -ActivarApps del portal.'
    } elseif (-not $SinTarea) {
        throw "En $archivoEnv falta SESION_VMA_SECRETO: sin inicio de sesión Planillas no arranca. Corre preparar_sesion.ps1 -ActivarApps del portal o, solo en la primera instalación, este script con -SinLogin."
    }
}

function Asegurar-VmaHosts([string]$archivoEnv) {
    # Pone VMA_HOSTS con el nombre y las IP de este servidor, si el .env todavía no trae uno con valor. Así el equipo
    # entra por el nombre o por cualquier IP sin que un Host falso desvíe el login (ver seguridad_web.host_permitido).
    if (-not (Test-Path $archivoEnv)) { return }
    $lineas = [Collections.Generic.List[string]](Get-Content $archivoEnv -Encoding UTF8)
    if ($lineas | Where-Object { $_ -match '^\s*VMA_HOSTS\s*=\s*\S' }) { return }   # ya configurado a mano: respetar
    $ips = Get-NetIPAddress -AddressFamily IPv4 |
        Where-Object { $_.IPAddress -notlike '127.*' -and $_.IPAddress -notlike '169.254.*' } |
        Select-Object -ExpandProperty IPAddress
    $valor = 'VMA_HOSTS=' + ((@([System.Net.Dns]::GetHostName()) + @($ips)) -join ',')
    $puesto = $false
    for ($i = 0; $i -lt $lineas.Count; $i++) { if ($lineas[$i] -match '^\s*VMA_HOSTS\s*=') { $lineas[$i] = $valor; $puesto = $true } }
    if (-not $puesto) { $lineas.Add($valor) }
    [IO.File]::WriteAllLines($archivoEnv, $lineas, (New-Object Text.UTF8Encoding $false))
    Write-Host "   VMA_HOSTS puesto solo: $($valor.Substring(10))"
}

Asegurar-VmaHosts $archivoEnv

# --- Cuenta de servicio, tarea programada y firewall ---------------------------------------------------------------
Paso "Cuenta de servicio «$CuentaServicio» (sin privilegios, en vez de SYSTEM)"
$clave = Asegurar-CuentaServicio $CuentaServicio $Raiz @('logs', 'temp')

Paso "Tarea programada «$Tarea»"
$accion = New-ScheduledTaskAction -Execute 'powershell.exe' -WorkingDirectory $app `
    -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$app\servidor\iniciar.ps1`" -Raiz `"$Raiz`" -Puerto $Puerto"
$disparador = New-ScheduledTaskTrigger -AtStartup
$ajustes = New-ScheduledTaskSettingsSet -ExecutionTimeLimit ([TimeSpan]::Zero) -RestartCount 999 `
    -RestartInterval (New-TimeSpan -Minutes 1) -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -MultipleInstances IgnoreNew
# Corre «aunque no haya sesión iniciada» con la contraseña guardada (cifrada por Windows); nivel limitado, sin elevar.
# Al registrar la tarea con contraseña, Windows le da sola a la cuenta el permiso «iniciar como proceso por lotes».
Register-ScheduledTask -TaskName $Tarea -Action $accion -Trigger $disparador -Settings $ajustes `
    -User "$env:COMPUTERNAME\$CuentaServicio" -Password $clave -RunLevel Limited `
    -Description 'Resaltador de Planillas de Grupo VMA (instalado con servidor\instalar.ps1)' -Force | Out-Null
Remove-Variable clave   # no dejar la contraseña en memoria del script

$regla = "Planillas VMA (puerto $Puerto)"
if (-not (Get-NetFirewallRule -DisplayName $regla -ErrorAction SilentlyContinue)) {
    New-NetFirewallRule -DisplayName $regla -Direction Inbound -Protocol TCP -LocalPort $Puerto -Action Allow -Profile Any `
        -RemoteAddress '10.0.0.0/8', '172.16.0.0/12', '192.168.0.0/16' | Out-Null
    Write-Host "   Regla de firewall creada: $regla (solo red interna)"
}

# --- Arrancar y comprobar ------------------------------------------------------------------------------------------
Paso 'Arrancando Planillas'
Start-ScheduledTask -TaskName $Tarea
$limite = (Get-Date).AddMinutes(2)
$responde = $false
while (-not $responde -and (Get-Date) -lt $limite) {
    Start-Sleep -Seconds 3
    try { $responde = (Invoke-WebRequest "http://localhost:$Puerto/salud" -UseBasicParsing -TimeoutSec 5).StatusCode -eq 200 }
    catch { $responde = $false }
}
if ($responde) {
    Paso "Planillas responde: http://$($env:COMPUTERNAME):$Puerto"
} else {
    Write-Warning "Planillas no respondió en 2 minutos. Últimas líneas de los registros ($Raiz\logs):"
    foreach ($nombre in 'errores.log', 'app.log', 'arranque.log') {
        $ruta = Join-Path $Raiz "logs\$nombre"
        if (Test-Path $ruta) { Write-Host "--- $nombre"; Get-Content $ruta -Tail 15 }
    }
    exit 1
}
