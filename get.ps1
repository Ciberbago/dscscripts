function Get-SpeedtestExecutable {
    $basePath = "$env:ProgramFiles\winget\Packages"
    $exe = Get-ChildItem -Path $basePath -Recurse -Include "speedtest.exe" -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -match "Ookla\.Speedtest\.CLI" } |
        Select-Object -First 1

    return $exe?.FullName
}

Get-SpeedtestExecutable
