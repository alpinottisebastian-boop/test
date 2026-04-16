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

    foreach ($u in $users) {

        net user $u.Name > $null 2>&1

        if ($LASTEXITCODE -eq 0) {
            Write-Log "  [SKIP] Usuario '$($u.Name)' ya existe." "Yellow"
        } else {
            try {
                # Crear usuario
                net user $u.Name $u.Password /add /y | Out-Null

                # No puede cambiar password
                net user $u.Name /passwordchg:no | Out-Null

                # Password nunca expira (FIX PowerShell viejo)
                $user = [ADSI]"WinNT://./$($u.Name),user"
                $flags = $user.UserFlags.Value
                $user.UserFlags.Value = $flags -bor 0x10000
                $user.SetInfo()

                Write-Log "  [OK] Usuario '$($u.Name)' creado." "Green"

            } catch {
                Write-Log "  [ERROR] Error creando '$($u.Name)': $_" "Red"
                $script:ErrorCount++
                continue
            }

            # 🔹 Manejo de grupo separado (clave)
            try {
                if ($u.IsAdmin) {
                    $adminGroup = [ADSI]"WinNT://./S-1-5-32-544,group"
                    $adminGroup.Add("WinNT://./$($u.Name),user")
                    Write-Log "  [OK] '$($u.Name)' agregado a Administradores." "Green"
                } else {
                    net localgroup Users $u.Name /add | Out-Null
                    Write-Log "  [OK] '$($u.Name)' agregado a Users." "Green"
                }
            } catch {
                Write-Log "  [WARN] No se pudo agregar '$($u.Name)' al grupo: $_" "Yellow"
            }
        }
    }

    Write-Host ""
}
#endregion
#region --- PASO 2: VERIFICAR / INSTALAR WINGET ---
function Ensure-Winget {
    Write-Log "=== PASO 2: VERIFICANDO WINGET ===" "Cyan"

    $winget = Get-Command winget -ErrorAction SilentlyContinue
    if ($winget) {
        Write-Log "  [OK] winget disponible: $($winget.Source)" "Green"
        return $true
    }

    Write-Log "  [INFO] winget no encontrado. Intentando instalar App Installer..." "Yellow"
    try {
        $releases = Invoke-RestMethod "https://api.github.com/repos/microsoft/winget-cli/releases/latest"
        $msixUrl  = ($releases.assets | Where-Object { $_.name -like "*.msixbundle" })[0].browser_download_url
        $msixPath = "$env:TEMP\AppInstaller.msixbundle"
        Invoke-WebRequest -Uri $msixUrl -OutFile $msixPath -UseBasicParsing
        Add-AppxPackage -Path $msixPath
        Write-Log "  [OK] winget instalado correctamente." "Green"
        return $true
    } catch {
        Write-Log "  [ERROR] No se pudo instalar winget: $_" "Red"
        Write-Log "  [INFO] Instala manualmente 'App Installer' desde Microsoft Store." "Yellow"
        $script:ErrorCount++
        return $false
    }
}
#endregion

#region --- PASO 3: INSTALAR APLICACIONES ---
function Install-Apps {
    Write-Log "=== PASO 3: INSTALACION DE APLICACIONES ===" "Cyan"

    $apps = @(
        @{ Name = "VLC Media Player";    Id = "VideoLAN.VLC"              },
        @{ Name = "7-Zip";               Id = "7zip.7zip"                  },
        @{ Name = "IrfanView";           Id = "IrfanSkiljan.IrfanView"        },
        @{ Name = "Microsoft Office 365";Id = "Microsoft.Office"           }
    )

    foreach ($app in $apps) {
        Write-Log "  Instalando: $($app.Name)..." "White"
        try {
            $result = winget install --id $app.Id `
                                     --silent `
                                     --accept-package-agreements `
                                     --accept-source-agreements `
                                     2>&1

            if ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq -1978335189) {
                # -1978335189 = ya instalado (codigo winget)
                Write-Log "  [OK] $($app.Name) instalado / ya presente." "Green"
            } else {
                Write-Log "  [WARN] $($app.Name) - codigo salida: $LASTEXITCODE" "Yellow"
            }
        } catch {
            Write-Log "  [ERROR] Fallo instalando $($app.Name): $_" "Red"
            $script:ErrorCount++
        }
    }
    Write-Host ""
}
#endregion

#region --- PASO 4: CONFIGURACIONES ADICIONALES DE WINDOWS ---
function Apply-WindowsConfig {
    Write-Log "=== PASO 4: CONFIGURACIONES DE WINDOWS ===" "Cyan"

    # Deshabilitar hibernacion (recomendado en laptops corporativas)
    try {
        powercfg /hibernate off 2>&1 | Out-Null
        Write-Log "  [OK] Hibernacion deshabilitada." "Green"
    } catch {
        Write-Log "  [WARN] No se pudo deshabilitar hibernacion." "Yellow"
    }

    # Deshabilitar arranque rapido (Fast Startup)
    # Puede causar problemas al apagar/encender en entornos corporativos
    try {
        $fsKey = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Power"
        Set-ItemProperty -Path $fsKey -Name "HiberbootEnabled" -Value 0 -Type DWord -ErrorAction Stop
        Write-Log "  [OK] Arranque rapido (Fast Startup) deshabilitado." "Green"
    } catch {
        Write-Log "  [WARN] No se pudo deshabilitar Fast Startup: $_" "Yellow"
    }

    # Windows Update - solo notificar, no instalar automatico
    try {
        $wuKey = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
        if (-not (Test-Path $wuKey)) { New-Item -Path $wuKey -Force | Out-Null }
        Set-ItemProperty -Path $wuKey -Name "AUOptions" -Value 2 -Type DWord
        Write-Log "  [OK] Windows Update configurado en modo 'solo notificar'." "Green"
    } catch {
        Write-Log "  [WARN] No se pudo configurar Windows Update: $_" "Yellow"
    }

    Write-Host ""
}
#endregion

#region --- RESUMEN FINAL ---
function Show-Summary {
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
    Write-Host "    [ADMIN] admin10   /  adminLAT" -ForegroundColor Cyan
    Write-Host "    [USER]  hola" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  Apps instaladas:" -ForegroundColor White
    Write-Host "    VLC  |  7-Zip  |  IrfanView  |  Office 365" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Configuraciones aplicadas:" -ForegroundColor White
    Write-Host "    Hibernacion OFF  |  Fast Startup OFF  |  RDP ON  |  Win Update notifica" -ForegroundColor Cyan
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
