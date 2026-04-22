# ============================================================
#  Setup-PC-LAT.ps1
#  Configuracion estandar de notebooks - IT LAT
#  Autor: IT Support LAT
#  Uso: Ejecutar como Administrador en PowerShell
# ============================================================

#region --- VERIFICACION DE PRIVILEGIOS ---
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "`n[ERROR] Este script debe ejecutarse como Administrador." -ForegroundColor Red
    Write-Host "        Clic derecho en PowerShell > 'Ejecutar como administrador'" -ForegroundColor Yellow
    pause
    exit 1
}
#endregion

#region --- CONFIGURACION GENERAL ---
$LogFile = "$env:TEMP\Setup-PC-LAT_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$ErrorCount = 0

function Write-Log {
    param([string]$Message, [string]$Color = "White")
    $timestamp = Get-Date -Format "HH:mm:ss"
    $line = "[$timestamp] $Message"
    Write-Host $line -ForegroundColor $Color
    Add-Content -Path $LogFile -Value $line
}

function Show-Banner {
    Clear-Host
    Write-Host ""
    Write-Host "  ██╗████████╗    ███████╗███████╗████████╗██╗   ██╗██████╗ " -ForegroundColor Cyan
    Write-Host "  ██║╚══██╔══╝    ██╔════╝██╔════╝╚══██╔══╝██║   ██║██╔══██╗" -ForegroundColor Cyan
    Write-Host "  ██║   ██║       ███████╗█████╗     ██║   ██║   ██║██████╔╝" -ForegroundColor Cyan
    Write-Host "  ██║   ██║       ╚════██║██╔══╝     ██║   ██║   ██║██╔═══╝ " -ForegroundColor Cyan
    Write-Host "  ██║   ██║       ███████║███████╗   ██║   ╚██████╔╝██║     " -ForegroundColor Cyan
    Write-Host "  ╚═╝   ╚═╝       ╚══════╝╚══════╝   ╚═╝    ╚═════╝ ╚═╝     " -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Configuracion Estandar de PC - IT LAT" -ForegroundColor White
    Write-Host "  Log: $LogFile" -ForegroundColor DarkGray
    Write-Host "  ─────────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
}
#endregion

#region --- PASO 1: CREAR USUARIOS ---
function Create-Users {
    Write-Log "=== PASO 1: CREACION DE USUARIOS ===" "Cyan"

    $users = @(
        @{ Name = "msfadmin10";  Password = '$oporteMSFLAT';  IsAdmin = $true  },
        @{ Name = "msfadminLAT"; Password = '$oporteMSFLAT!'; IsAdmin = $true  },
        @{ Name = "msf";         Password = 'usuarioBOA24*';  IsAdmin = $false }
    )

    # Obtener nombre REAL del grupo Administradores via SID (funciona en cualquier idioma)
    $adminSID   = New-Object System.Security.Principal.SecurityIdentifier("S-1-5-32-544")
    $adminGroup = $adminSID.Translate([System.Security.Principal.NTAccount]).Value.Split("\")[1]

    foreach ($u in $users) {

        # FIX: usar Get-LocalUser es mas confiable que el exit code de net user
        $existing = Get-LocalUser -Name $u.Name -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Log "  [SKIP] Usuario '$($u.Name)' ya existe." "Yellow"
            continue
        }

        try {
            net user $u.Name $u.Password /add /y | Out-Null
            net user $u.Name /passwordchg:no  | Out-Null

            # Password nunca expira via flags ADSI
            $userADSI = [ADSI]"WinNT://./$($u.Name),user"
            $userADSI.UserFlags.Value = $userADSI.UserFlags.Value -bor 0x10000
            $userADSI.SetInfo()

            Write-Log "  [OK] Usuario '$($u.Name)' creado." "Green"

        } catch {
            Write-Log "  [ERROR] Error creando '$($u.Name)': $_" "Red"
            $script:ErrorCount++
            continue
        }

        try {
            if ($u.IsAdmin) {
                net localgroup "$adminGroup" $u.Name /add | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    Write-Log "  [OK] '$($u.Name)' agregado a $adminGroup." "Green"
                } else {
                    Write-Log "  [ERROR] No se pudo agregar '$($u.Name)' a $adminGroup." "Red"
                    $script:ErrorCount++
                }
            } else {
                net localgroup Users $u.Name /add | Out-Null
                Write-Log "  [OK] '$($u.Name)' agregado a Users." "Green"
            }
        } catch {
            Write-Log "  [ERROR] Fallo agregando '$($u.Name)' al grupo: $_" "Red"
            $script:ErrorCount++
        }
    }

    Write-Host ""
}
#endregion

#region --- PASO 2: VERIFICAR / INSTALAR WINGET ---
function Install-Winget {
    Write-Log "  [INFO] Descargando e instalando App Installer desde GitHub..." "Yellow"
    try {
        $releases = Invoke-RestMethod "https://api.github.com/repos/microsoft/winget-cli/releases/latest" `
                                      -UseBasicParsing
        $msixUrl  = ($releases.assets | Where-Object { $_.name -like "*.msixbundle" })[0].browser_download_url
        $msixPath = "$env:TEMP\AppInstaller.msixbundle"
        Invoke-WebRequest -Uri $msixUrl -OutFile $msixPath -UseBasicParsing
        Add-AppxPackage -Path $msixPath -ErrorAction Stop
        Write-Log "  [OK] winget instalado/reinstalado correctamente." "Green"
        return $true
    } catch {
        Write-Log "  [ERROR] No se pudo instalar winget: $_" "Red"
        Write-Log "  [INFO] Instala manualmente 'App Installer' desde Microsoft Store." "Yellow"
        $script:ErrorCount++
        return $false
    }
}

function Ensure-Winget {
    Write-Log "=== PASO 2: VERIFICANDO WINGET ===" "Cyan"

    # 1. Verificar que el comando existe
    $wingetCmd = Get-Command winget -ErrorAction SilentlyContinue
    if (-not $wingetCmd) {
        Write-Log "  [INFO] winget no encontrado en el sistema." "Yellow"
        return (Install-Winget)
    }

    # 2. Verificar que realmente funciona (no solo que existe)
    try {
        $ver = & winget --version 2>&1
        if ($LASTEXITCODE -eq 0 -and "$ver" -match "v[\d\.]+") {
            Write-Log "  [OK] winget operativo: $ver" "Green"
            return $true
        } else {
            Write-Log "  [WARN] winget existe pero no responde correctamente (exit: $LASTEXITCODE). Reinstalando..." "Yellow"
            return (Install-Winget)
        }
    } catch {
        Write-Log "  [WARN] winget lanzo excepcion al ejecutarse. Reinstalando..." "Yellow"
        return (Install-Winget)
    }
}
#endregion

#region --- PASO 3: INSTALAR / ACTUALIZAR APLICACIONES ---
function Install-App-Direct {
    # Instalacion directa sin verificacion previa - rapido y sin riesgo de cuelgue
    param([string]$Name, [string]$Id)
    Write-Log "  Instalando: $Name..." "White"
    try {
        winget install --id $Id `
                       --silent `
                       --accept-package-agreements `
                       --accept-source-agreements `
                       2>&1 | Out-Null

        if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335189) {
            # -1978335189 = ya instalado segun winget, se considera OK
            Write-Log "  [OK] $Name instalado / ya presente." "Green"
        } else {
            Write-Log "  [WARN] $Name - codigo salida: $LASTEXITCODE" "Yellow"
        }
    } catch {
        Write-Log "  [ERROR] Excepcion instalando $Name`: $_" "Red"
        $script:ErrorCount++
    }
}

function Install-TeamViewer {
    # Logica especial: verificar -> upgrade -> force si falla
    $tvName = "TeamViewer Host"
    $tvId   = "TeamViewer.TeamViewer.Host"

    Write-Log "  Procesando: $tvName..." "White"
    try {
        # Verificar si esta instalado
        winget list --id $tvId -e 2>&1 | Out-Null
        $isInstalled = ($LASTEXITCODE -eq 0)

        if (-not $isInstalled) {
            Write-Log "  TeamViewer no encontrado. Instalando..." "White"
            winget install --id $tvId `
                           --silent `
                           --accept-package-agreements `
                           --accept-source-agreements `
                           2>&1 | Out-Null

            if ($LASTEXITCODE -eq 0) {
                Write-Log "  [OK] $tvName instalado correctamente." "Green"
            } else {
                Write-Log "  [ERROR] $tvName fallo al instalar (exit: $LASTEXITCODE)." "Red"
                $script:ErrorCount++
            }

        } else {
            # Esta instalado: intentar upgrade
            Write-Log "  TeamViewer encontrado. Actualizando..." "White"
            winget upgrade --id $tvId `
                           --silent `
                           --accept-package-agreements `
                           --accept-source-agreements `
                           2>&1 | Out-Null

            if ($LASTEXITCODE -eq 0) {
                Write-Log "  [OK] $tvName actualizado correctamente." "Green"
            } else {
                # Upgrade fallo: reinstalar con --force
                Write-Log "  [WARN] Upgrade fallo (exit: $LASTEXITCODE). Reinstalando con --force..." "Yellow"
                winget install --id $tvId `
                               --silent `
                               --force `
                               --accept-package-agreements `
                               --accept-source-agreements `
                               2>&1 | Out-Null

                if ($LASTEXITCODE -eq 0) {
                    Write-Log "  [OK] $tvName reinstalado correctamente (--force)." "Green"
                } else {
                    Write-Log "  [ERROR] $tvName fallo incluso con --force (exit: $LASTEXITCODE)." "Red"
                    $script:ErrorCount++
                }
            }
        }
    } catch {
        Write-Log "  [ERROR] Excepcion procesando $tvName`: $_" "Red"
        $script:ErrorCount++
    }
}

function Install-Apps {
    Write-Log "=== PASO 3: INSTALACION DE APLICACIONES ===" "Cyan"

    # Instalacion directa (sin verificacion de version - rapido, sin cuelgues)
    Install-App-Direct "VLC Media Player"     "VideoLAN.VLC"
    Install-App-Direct "7-Zip"                "7zip.7zip"
    Install-App-Direct "IrfanView"            "IrfanSkiljan.IrfanView"
    Install-App-Direct "Microsoft Office 365" "Microsoft.Office"

    # TeamViewer: logica especial con upgrade y fallback --force
    Install-TeamViewer

    Write-Host ""
}
#endregion

#region --- PASO 4: CONFIGURACIONES DE WINDOWS ---
function Apply-WindowsConfig {
    Write-Log "=== PASO 4: CONFIGURACIONES DE WINDOWS ===" "Cyan"

    # Deshabilitar hibernacion
    try {
        powercfg /hibernate off 2>&1 | Out-Null
        Write-Log "  [OK] Hibernacion deshabilitada." "Green"
    } catch {
        Write-Log "  [WARN] No se pudo deshabilitar hibernacion." "Yellow"
    }

    # Deshabilitar arranque rapido (Fast Startup)
    try {
        $fsKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power"
        Set-ItemProperty -Path $fsKey -Name "HiberbootEnabled" -Value 0 -Type DWord -ErrorAction Stop
        Write-Log "  [OK] Arranque rapido (Fast Startup) deshabilitado." "Green"
    } catch {
        Write-Log "  [WARN] No se pudo deshabilitar Fast Startup: $_" "Yellow"
    }

    # Habilitar escritorio remoto (RDP)
    try {
        Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" `
                         -Name "fDenyTSConnections" -Value 0 -ErrorAction Stop
        Enable-NetFirewallRule -DisplayGroup "Remote Desktop" -ErrorAction SilentlyContinue
        Write-Log "  [OK] Escritorio Remoto (RDP) habilitado." "Green"
    } catch {
        Write-Log "  [WARN] No se pudo habilitar RDP: $_" "Yellow"
    }

    # Windows Update - solo notificar
    try {
        $wuKey = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
        if (-not (Test-Path $wuKey)) { New-Item -Path $wuKey -Force | Out-Null }
        Set-ItemProperty -Path $wuKey -Name "AUOptions" -Value 2 -Type DWord
        Write-Log "  [OK] Windows Update: modo 'solo notificar'." "Green"
    } catch {
        Write-Log "  [WARN] No se pudo configurar Windows Update: $_" "Yellow"
    }

    # Deshabilitar cuenta Administrador integrada de Windows (SID *-500)
    try {
        $builtinAdmin = (Get-LocalUser | Where-Object { $_.SID -like "S-1-5-*-500" }).Name
        if ($builtinAdmin) {
            Disable-LocalUser -Name $builtinAdmin -ErrorAction Stop
            Write-Log "  [OK] Cuenta integrada '$builtinAdmin' deshabilitada." "Green"
        }
    } catch {
        Write-Log "  [WARN] No se pudo deshabilitar cuenta integrada: $_" "Yellow"
    }

    Write-Host ""
}
#endregion

#region --- RESUMEN FINAL ---
function Show-Summary {

    $adminUsers  = @("msfadmin10", "msfadminLAT")
    $normalUsers = @("msf")
    $appList     = "VLC  |  7-Zip  |  IrfanView  |  TeamViewer Host  |  Office 365"

    Write-Host ""
    Write-Host "  ─────────────────────────────────────────────────────────" -ForegroundColor DarkGray
    if ($ErrorCount -eq 0) {
        Write-Log "  CONFIGURACION COMPLETADA SIN ERRORES." "Green"
    } else {
        Write-Log "  CONFIGURACION COMPLETADA CON $ErrorCount ERROR(ES). Revisa el log." "Yellow"
    }
    Write-Host "  Log guardado en: $LogFile" -ForegroundColor DarkGray
    Write-Host "  ─────────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host ""
    Write-Host "  Usuarios creados:" -ForegroundColor White
    Write-Host "    [ADMIN] $($adminUsers -join '  /  ')" -ForegroundColor Cyan
    Write-Host "    [USER]  $($normalUsers -join '  /  ')" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Apps instaladas:" -ForegroundColor White
    Write-Host "    $appList" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Configuraciones aplicadas:" -ForegroundColor White
    Write-Host "    Hibernacion OFF  |  Fast Startup OFF  |  RDP ON  |  Win Update notifica  |  Admin integrado OFF" -ForegroundColor Cyan
    Write-Host ""
    pause
}
#endregion

#region --- EJECUCION PRINCIPAL ---
Show-Banner
Create-Users
$wingetOk = Ensure-Winget
if ($wingetOk) { Install-Apps }
Apply-WindowsConfig
Show-Summary
#endregion
