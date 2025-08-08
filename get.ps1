function Get-SpeedtestExecutable {
    $basePath = "$env:ProgramFiles\winget\Packages"
    $exe = Get-ChildItem -Path $basePath -Recurse -Include "speedtest.exe" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if ($exe -and $exe.FullName) {
        return $exe.FullName
    } else {
        return $null
    }
}

Get-SpeedtestExecutable
