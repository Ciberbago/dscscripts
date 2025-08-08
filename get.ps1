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

Get-SpeedtestExecutable
