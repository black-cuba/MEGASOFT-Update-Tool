# Forzar salida en UTF-8 para mostrar acentos correctamente 
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# Presentación profesional al ejecutar
Write-Host "=================================================================="
Write-Host "  __  __ ______  _____   ______    _____   ____   ______  _______ "
Write-Host " |  \/  |  ____|/ ____| |  __  |  / ____| / __ \ |  ____||__   __|"
Write-Host " | \  / | |__  | (__) | | |__| | | (___  | |  | || |__      | |   "
Write-Host " | |\/| |  __|  \___  | |  __  |  \___ \ | |  | ||  __|     | |   "
Write-Host " | |  | | |____ ____) | | |  | |  ____) || |  | ||  |       | |   "  
Write-Host " |_|  |_|______|_____/  |_|  |_| |_____/  \____/ |__|       |_|   "
Write-Host "=================================================================="
Write-Host "=================================================================="
Write-Host "   Script de Actualización de Windows"
Write-Host "   Marca Oficial: MEGASOFT"
Write-Host "   Autor: Miguel Ángel Piñeiro"
Write-Host "   Ciudad: Ciego de Ávila, Cuba"
Write-Host "==================================================================`n"

# Verificar ExecutionPolicy
$currentPolicy = Get-ExecutionPolicy
if ($currentPolicy -ne "RemoteSigned") {
    Write-Host "Configurando ExecutionPolicy en RemoteSigned..."
    Set-ExecutionPolicy RemoteSigned -Scope CurrentUser -Force
} else {
    Write-Host "ExecutionPolicy ya está configurado en RemoteSigned.`n"
}

# Verificar módulo PSWindowsUpdate
if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
    Write-Host "Instalando módulo PSWindowsUpdate..."
    Install-Module PSWindowsUpdate -Force -Confirm:$false
}
Import-Module PSWindowsUpdate

# Buscar actualizaciones
Write-Host "Buscando actualizaciones disponibles..."
$updates = Get-WindowsUpdate

if ($updates.Count -eq 0) {
    Write-Host "No hay actualizaciones pendientes."
    exit
}

# Mostrar lista en tabla organizada
Write-Host "`n=============================================================="
Write-Host ("{0,-4} {1,-12} {2,-40}" -f "Nº", "KB", "Título")
Write-Host "--------------------------------------------------------------"

for ($i=0; $i -lt $updates.Count; $i++) {
    $num = $i + 1
    $kb = "KB" + $updates[$i].KBArticleID
    $title = $updates[$i].Title
    Write-Host ("{0,-4} {1,-12} {2,-40}" -f $num, $kb, $title)
}

Write-Host "==============================================================`n"

# Bucle interactivo: no termina hasta que el usuario decida
while ($true) {
    $choice = Read-Host "Escribe el/los número(s) de las actualizaciones que quieres instalar (ej: 1,3,5). 
Si quieres todas, escribe ALL. 
Si quieres ver detalles en la web, escribe INFO y el número (ej: INFO 2). 
Si no quieres instalar nada, escribe EXIT"

    if ($choice -eq "ALL") {
        Install-WindowsUpdate -AcceptAll -IgnoreReboot
        
        # Preguntar reinicio SOLO después de instalar
        $reboot = Read-Host "`n¿Quieres reiniciar ahora? (S/N)"
        if ($reboot -eq "S") {
            Restart-Computer
        }
        break  # salir del bucle tras instalar

    } elseif ($choice -like "INFO*") {
        $num = $choice.Split(" ")[1]
        $index = [int]$num - 1
        if ($index -ge 0 -and $index -lt $updates.Count) {
            $kb = $updates[$index].KBArticleID
            Write-Host "Abriendo página oficial de KB$kb..."
            Start-Process "https://support.microsoft.com/help/$kb"
        }
        # Aquí NO se pregunta reinicio y el script sigue activo

    } elseif ($choice -eq "EXIT") {
        Write-Host "Has decidido no instalar actualizaciones. Cerrando MEGASOFT..."
        break  # salir del bucle

    } else {
        $nums = $choice -split ","
        $kbList = @()
        foreach ($n in $nums) {
            $index = [int]$n.Trim() - 1
            if ($index -ge 0 -and $index -lt $updates.Count) {
                $kbList += $updates[$index].KBArticleID
            }
        }
        if ($kbList.Count -gt 0) {
            Install-WindowsUpdate -KBArticleID $kbList -AcceptAll -IgnoreReboot
            
            # Preguntar reinicio SOLO si hubo instalación
            $reboot = Read-Host "`n¿Quieres reiniciar ahora? (S/N)"
            if ($reboot -eq "S") {
                Restart-Computer
            }
            break  # salir del bucle tras instalar
        } else {
            Write-Host "No se seleccionaron actualizaciones válidas."
        }
    }
}
