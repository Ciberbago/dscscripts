function PruebaDeVelocidad {

    [CmdletBinding()]
    param ()

    # URL del webhook de Teams para enviar la tarjeta adaptativa
    # Este valor debe ser provisto por el entorno de Scalefusion o la configuración del script.
    $TeamsWebhookUrl = "https://default5561a42719684db2aecc8b1dfedb5c.72.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/feeaac27140147ffb1777d5973372031/triggers/manual/paths/invoke/?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=vt4xHaNu7IpvTGfQVC09V_khqt4nBr7OsnYNhldwFaU"

    try {
        Write-Host "Iniciando recopilación de información mínima del sistema..."

        # --- Recopilación de datos básicos ---
        $Hostname = $env:COMPUTERNAME
        $CurrentUser = $env:USERNAME

        # --- Detección del tipo de conexión de red ---
        $ConnectionType = "Desconocido"
        $NetworkName = "N/A"
        
        try {
            # Obtiene el perfil de conexión de red activa
            $ConnectionProfile = Get-NetConnectionProfile -ErrorAction SilentlyContinue
            if ($ConnectionProfile) {
                $NetworkName = $ConnectionProfile.Name
                # Obtiene el adaptador de red asociado para saber si es Wi-Fi o Ethernet
                $NetAdapter = Get-NetAdapter -InterfaceIndex $ConnectionProfile.InterfaceIndex -ErrorAction SilentlyContinue
                if ($NetAdapter) {
                    if ($NetAdapter.MediaType -eq '802.11') {
                        $ConnectionType = "Wi-Fi"
                    } else {
                        $ConnectionType = "Cable (Ethernet)"
                    }
                }
            } else {
                $ConnectionType = "Sin conexión"
            }
        } catch {
            Write-Host "Error al detectar la conexión de red: $($_.Exception.Message)"
        }

        # --- Verificar e instalar Speedtest CLI con Winget ---
        $speedtestPath = Get-Command speedtest.exe -ErrorAction SilentlyContinue
        if (-not $speedtestPath) {
            Write-Host "Speedtest CLI no encontrado. Intentando instalar con Winget..."
            try {
                winget install --id Ookla.Speedtest.CLI -e --accept-package-agreements --accept-source-agreements
                Write-Host "Speedtest CLI instalado correctamente."
            } catch {
                Write-Host "Error al instalar Speedtest CLI: $($_.Exception.Message)"
            }
        }

        # --- Ejecutar Speedtest y obtener resultados ---
        $Ping = "N/A"
        $DownloadSpeed = "N/A"
        $UploadSpeed = "N/A"
        
        $speedtestPath = Get-Command speedtest.exe -ErrorAction SilentlyContinue

        if ($speedtestPath) {
            Write-Host "Ejecutando prueba de velocidad..."
            try {
                # Se agrega el parámetro --accept-license para ejecución no interactiva
                $SpeedtestResult = speedtest --accept-license -f json | ConvertFrom-Json -ErrorAction Stop
                $Ping = "{0:N2} ms" -f $SpeedtestResult.ping.latency
                # Convertir de bits/s a Mbps
                $DownloadSpeed = "{0:N2} Mbps" -f ($SpeedtestResult.download.bandwidth * 8 / 1MB)
                $UploadSpeed = "{0:N2} Mbps" -f ($SpeedtestResult.upload.bandwidth * 8 / 1MB)
                Write-Host "Prueba de velocidad completada."
            } catch {
                Write-Host "Error al ejecutar Speedtest: $($_.Exception.Message)"
            }
        } else {
            Write-Host "Speedtest CLI no está disponible. Saltando la prueba de velocidad."
        }

        $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        # --- Construcción del contenido de la tarjeta (JSON directo) ---
        $AdaptiveCardContent = @{
            type = "AdaptiveCard"
            version = "1.0"
            body = @(
                @{
                    type = "TextBlock"
                    text = "## Reporte de Equipo Simplificado"
                    wrap = $true
                },
                @{
                    type = "FactSet"
                    facts = @(
                        @{ title = "Fecha y Hora:"; value = "$($Timestamp)" },
                        @{ title = "Equipo:"; value = "$($Hostname)" },
                        @{ title = "Usuario:"; value = "$($CurrentUser)" },
                        @{ title = "Tipo de Conexión:"; value = "$($ConnectionType)" },
                        @{ title = "Red Wi-Fi (si aplica):"; value = "$($NetworkName)" },
                        @{ title = "Ping:"; value = "$($Ping)" },
                        @{ title = "Velocidad de Descarga:"; value = "$($DownloadSpeed)" },
                        @{ title = "Velocidad de Subida:"; value = "$($UploadSpeed)" }
                    )
                }
            )
        }

        # --- Crear estructura para envío a Teams ---
        $TeamsMessageBody = @{
            type = "message"
            attachments = @(
                @{
                    contentType = "application/vnd.microsoft.card.adaptive"
                    contentUrl = $null
                    content = $AdaptiveCardContent
                }
            )
        } | ConvertTo-Json -Depth 10

        # --- Envío al webhook ---
        Write-Host "Enviando al webhook de Teams..."
        Invoke-RestMethod -Uri $TeamsWebhookUrl -Method Post -Body $TeamsMessageBody -ContentType "application/json" -ErrorAction Stop
        Write-Host "Tarjeta enviada correctamente."

        return $true
    }
    catch {
        Write-Host "Error grave: $($_.Exception.Message)"
        if ($_.Exception.Response) {
            $reader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
            $responseText = $reader.ReadToEnd()
            Write-Host "Respuesta del servidor:"
            Write-Host $responseText
        }
        return $false
    }
}

function Add-WinGetPath {
  $WinGetPath = (Get-ChildItem -Path "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller*_x64*\winget.exe").DirectoryName
  If (-Not(($Env:Path -split ';') -contains $WinGetPath))
    { 
    If ($env:path -match ";$")
      {
        $env:path +=  $WinGetPath + ";"
      }
    else
      {
        $env:path +=  ";" + $WinGetPath + ";"			
      }
    write-host "Winget path $WinGetPath added to environment variable"
    }
  else
    { 
      write-host "Winget path already exists in registry."
    }
}

Add-WinGetPath


PruebaDeVelocidad
