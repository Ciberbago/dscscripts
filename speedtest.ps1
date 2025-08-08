function Get-SpeedtestExecutable {
    [CmdletBinding()]
    param ()

    $basePath = "$env:ProgramFiles\winget\Packages"
    Write-Host "🔍 Buscando 'speedtest.exe' en: $basePath"

    try {
        $exe = Get-ChildItem -Path $basePath -Recurse -Include "speedtest.exe" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1

        if ($exe -and $exe.FullName) {
            Write-Host "✅ Ejecutable encontrado:"
            Write-Host $exe.FullName
            return $exe.FullName
        } else {
            Write-Warning "⚠️ No se encontró 'speedtest.exe' en el directorio esperado."
            return $null
        }
    } catch {
        Write-Error "❌ Error al buscar el ejecutable: $($_.Exception.Message)"
        return $null
    }
}

function Add-WinGetPath {
    $WinGetPath = (Get-ChildItem -Path "$env:ProgramFiles\WindowsApps\Microsoft.DesktopAppInstaller*_x64*\winget.exe").DirectoryName
    If (-Not(($Env:Path -split ';') -contains $WinGetPath)) {
        $env:path += if ($env:path -match ";$") { $WinGetPath + ";" } else { ";" + $WinGetPath + ";" }
        Write-Host "Winget path $WinGetPath added to environment variable"
    } else {
        Write-Host "Winget path already exists in registry."
    }
}

function PruebaDeVelocidad {
    [CmdletBinding()]
    param ()

    $TeamsWebhookUrl = "https://default5561a42719684db2aecc8b1dfedb5c.72.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/feeaac27140147ffb1777d5973372031/triggers/manual/paths/invoke/?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=vt4xHaNu7IpvTGfQVC09V_khqt4nBr7OsnYNhldwFaU"

    try {
        Write-Host "Iniciando recopilación de información mínima del sistema..."

        $Hostname = $env:COMPUTERNAME
        $CurrentUser = $env:USERNAME
        if ($CurrentUser -eq "SYSTEM") {
            Write-Warning "⚠️ El script se está ejecutando como SYSTEM. Esto puede causar rechazo en el webhook."
        }

        $ConnectionType = "Desconocido"
        $NetworkName = @("N/A")

        try {
            $ConnectionProfile = Get-NetConnectionProfile -ErrorAction SilentlyContinue
            if ($ConnectionProfile) {
                $NetAdapter = Get-NetAdapter -InterfaceIndex $ConnectionProfile.InterfaceIndex -ErrorAction SilentlyContinue
                $NetworkName = @($ConnectionProfile.Name)
                if ($NetAdapter) {
                    $ConnectionType = if ($NetAdapter.MediaType -eq '802.11') { "Wi-Fi" } else { "Cable (Ethernet)" }
                }
            } else {
                $ConnectionType = "Sin conexión"
            }
        } catch {
            Write-Host "Error al detectar la conexión de red: $($_.Exception.Message)"
        }

        $speedtestCmd = Get-Command speedtest.exe -ErrorAction SilentlyContinue
        if (-not $speedtestCmd) {
            Write-Host "Speedtest CLI no encontrado. Intentando instalar con Winget..."
            try {
                winget install --id Ookla.Speedtest.CLI -e --accept-package-agreements --accept-source-agreements --scope machine -ErrorAction Stop | Out-Null
                Write-Host "Speedtest CLI instalado correctamente."
            } catch {
                Write-Host "Error al instalar Speedtest CLI: $($_.Exception.Message)"
            }
        }

        $Ping = "N/A"
        $DownloadSpeed = "N/A"
        $UploadSpeed = "N/A"

        $speedtestCmd = Get-Command speedtest.exe -ErrorAction SilentlyContinue
        $speedtestPath = if ($speedtestCmd) { $speedtestCmd.Source } else { Get-SpeedtestExecutable }

        if ($speedtestPath) {
            Write-Host "Ejecutando prueba de velocidad..."
            try {
                $SpeedtestResult = & $speedtestPath --accept-license -f json | ConvertFrom-Json -ErrorAction Stop
                $Ping = "{0:N2} ms" -f $SpeedtestResult.ping.latency
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

        $LogFolder = "C:\ProgramData\SpeedtestLogs"
        if (-not (Test-Path $LogFolder)) {
            New-Item -Path $LogFolder -ItemType Directory -Force | Out-Null
        }

        $SafeHostname = $Hostname -replace '[^a-zA-Z0-9\-]', '_'
        $SafeTimestamp = $Timestamp -replace '[: ]', '_'
        $LogFile = "$LogFolder\speedtest_${SafeHostname}_$SafeTimestamp.json"

        $Payload = @{
            hostname = $Hostname
            usuario = $CurrentUser
            tipoConexion = $ConnectionType
            red = $NetworkName
            ping = $Ping
            descarga = $DownloadSpeed
            subida = $UploadSpeed
            timestamp = $Timestamp
        }

        $PayloadJson = $Payload | ConvertTo-Json -Depth 5 -Compress
        Set-Content -Path $LogFile -Value $PayloadJson -Encoding UTF8
        Write-Host "📁 Resultado guardado en: $LogFile"

        Write-Host "📤 Subiendo a 0x0.st..."
        try {
            $UploadResponse = Invoke-RestMethod -Uri "https://0x0.st" -Method Post -Form @{ file = Get-Item $LogFile }
            Write-Host "✅ Archivo subido: $UploadResponse"

            $TeamsMessage = @{ text = "📊 Resultado de prueba de velocidad para `$Hostname`: $UploadResponse" } | ConvertTo-Json -Compress
            Invoke-RestMethod -Uri $TeamsWebhookUrl -Method Post -Body $TeamsMessage -ContentType "application/json"
            Write-Host "✅ Enlace enviado a Teams."
        } catch {
            Write-Host "❌ Error al subir o enviar: $($_.Exception.Message)"
        }

        return $true
    } catch {
        Write-Host "❌ Error grave: $($_.Exception.Message)"
        if ($_.Exception.Response) {
            $reader = [System.IO.StreamReader]::new($_.Exception.Response.GetResponseStream())
            $responseText = $reader.ReadToEnd()
            Write-Host "Respuesta del servidor:"
            Write-Host $responseText
        }
        return $false
    }
}

Add-WinGetPath
PruebaDeVelocidad
