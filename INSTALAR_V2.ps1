$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

Write-Host ""
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host " POWER BI ANALYZER V2 - INSTALACAO" -ForegroundColor Cyan
Write-Host "==============================================" -ForegroundColor Cyan
Write-Host ""

# 1) Python
$Py = Get-Command py -ErrorAction SilentlyContinue
if (-not $Py) {
    $Py = Get-Command python -ErrorAction SilentlyContinue
}
if (-not $Py) {
    Write-Host "Python nao encontrado." -ForegroundColor Red
    Write-Host "Instale Python 3.11+ 64 bits e execute novamente."
    Read-Host "Pressione ENTER para sair"
    exit 1
}

if (-not (Test-Path ".venv\Scripts\python.exe")) {
    Write-Host "[1/4] Criando ambiente Python..."
    & $Py.Source -m venv .venv
} else {
    Write-Host "[1/4] Ambiente Python ja existe."
}

Write-Host "[2/4] Instalando dependencias..."
& ".venv\Scripts\python.exe" -m pip install --upgrade pip
& ".venv\Scripts\python.exe" -m pip install -r requirements.txt
if ($LASTEXITCODE -ne 0) {
    throw "Falha ao instalar dependencias Python."
}

# 2) Verifica MSOLAP / ADO
Write-Host "[3/4] Verificando provider MSOLAP..."
$TestScript = @'
import win32com.client, sys
try:
    c = win32com.client.Dispatch("ADODB.Connection")
    print("ADO OK")
except Exception as e:
    print(e)
    sys.exit(1)
'@
$TestScript | & ".venv\Scripts\python.exe" -
if ($LASTEXITCODE -ne 0) {
    throw "ADO/pywin32 nao esta funcional neste Windows."
}

# 3) Registrar External Tool requer pasta protegida. Elevamos somente esta etapa.
$SelfElevated = $args -contains "-ElevatedRegister"
if (-not $SelfElevated) {
    $principal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $isAdmin) {
        Write-Host "[4/4] Solicitando permissao de administrador para registrar no Power BI..."
        $argList = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", "`"$($MyInvocation.MyCommand.Path)`"",
            "-ElevatedRegister"
        )
        Start-Process powershell.exe -Verb RunAs -ArgumentList $argList -Wait
        Write-Host ""
        Write-Host "Instalacao concluida. Feche e reabra o Power BI Desktop." -ForegroundColor Green
        Read-Host "Pressione ENTER"
        exit 0
    }
}

Write-Host "[4/4] Registrando ferramenta externa..."

$CandidateDirs = @()
if (${env:CommonProgramFiles(x86)}) {
    $CandidateDirs += (Join-Path ${env:CommonProgramFiles(x86)} "Microsoft Shared\Power BI Desktop\External Tools")
}
if ($env:CommonProgramFiles) {
    $CandidateDirs += (Join-Path $env:CommonProgramFiles "Microsoft Shared\Power BI Desktop\External Tools")
}
$CandidateDirs = $CandidateDirs | Select-Object -Unique

$TargetDir = $null
foreach ($d in $CandidateDirs) {
    try {
        New-Item -ItemType Directory -Force -Path $d | Out-Null
        $TargetDir = $d
        break
    } catch {}
}
if (-not $TargetDir) {
    throw "Nao foi possivel localizar/criar a pasta External Tools do Power BI."
}

$PowerShellExe = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$Runner = Join-Path $Root "RUN_EXTERNAL_TOOL.ps1"

$Manifest = @{
    version = "1.0.0"
    name = "Power BI Analyzer V2"
    description = "Analisa modelo, DAX, VertiPaq, cardinalidade, memoria e Power Query."
    path = $PowerShellExe
    arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$Runner`" -Server `"%server%`" -Database `"%database%`""
} | ConvertTo-Json -Depth 5

$ManifestPath = Join-Path $TargetDir "Power BI Analyzer V2.pbitool.json"
$Manifest | Set-Content -Path $ManifestPath -Encoding UTF8

Write-Host ""
Write-Host "OK: $ManifestPath" -ForegroundColor Green
Write-Host ""
Write-Host "Feche e reabra o Power BI Desktop."
Write-Host "Depois: Ferramentas Externas > Power BI Analyzer V2" -ForegroundColor Green
