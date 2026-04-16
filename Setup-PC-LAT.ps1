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

        # Verificar si existe
        net user $u.Name > $null 2>&1

        if ($LASTEXITCODE -eq 0) {
            Write-Log "  [SKIP] Usuario '$($u.Name)' ya existe." "Yellow"
        } else {
            try {
                # Crear usuario
                net user $u.Name $u.Password /add /y | Out-Null

                # Password nunca expira
                wmic UserAccount where "Name='$($u.Name)'" set PasswordExpires=FALSE | Out-Null

                # Usuario no puede cambiar password
                net user $u.Name /passwordchg:no | Out-Null

                # Agregar a grupo
                if ($u.IsAdmin) {
                    net localgroup Administrators $u.Name /add | Out-Null
                    Write-Log "  [OK] Usuario ADMIN '$($u.Name)' creado." "Green"
                } else {
                    net localgroup Users $u.Name /add | Out-Null
                    Write-Log "  [OK] Usuario ESTANDAR '$($u.Name)' creado." "Green"
                }

            } catch {
                Write-Log "  [ERROR] No se pudo crear '$($u.Name)': $_" "Red"
                $script:ErrorCount++
            }
        }
    }

    Write-Host ""
}
#endregion

