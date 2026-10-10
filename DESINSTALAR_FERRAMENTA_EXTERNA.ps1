$ErrorActionPreference = "SilentlyContinue"

$CandidateDirs = @()
if (${env:CommonProgramFiles(x86)}) {
    $CandidateDirs += (Join-Path ${env:CommonProgramFiles(x86)} "Microsoft Shared\Power BI Desktop\External Tools")
}
if ($env:CommonProgramFiles) {
    $CandidateDirs += (Join-Path $env:CommonProgramFiles "Microsoft Shared\Power BI Desktop\External Tools")
}

$principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Start-Process powershell.exe -Verb RunAs -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$($MyInvocation.MyCommand.Path)`"" -Wait
    exit
}

foreach ($d in ($CandidateDirs | Select-Object -Unique)) {
    $p = Join-Path $d "Power BI Analyzer V2.pbitool.json"
    if (Test-Path $p) {
        Remove-Item $p -Force
        Write-Host "Removido: $p" -ForegroundColor Green
    }
}
Write-Host "Concluido. Reinicie o Power BI Desktop."
