param(
    [Parameter(Mandatory=$true)][string]$Server,
    [Parameter(Mandatory=$true)][string]$Database
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = Join-Path $Root ".venv\Scripts\python.exe"

if (-not (Test-Path $Python)) {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show(
        "O ambiente do Power BI Analyzer V2 ainda não foi instalado.`nExecute INSTALAR_V2.ps1 primeiro.",
        "Power BI Analyzer V2"
    ) | Out-Null
    exit 2
}

$Out = Join-Path $Root "Resultados"
New-Item -ItemType Directory -Force -Path $Out | Out-Null

& $Python (Join-Path $Root "powerbi_analyzer_v2.py") `
    --server $Server `
    --database $Database `
    --out $Out

$Exit = $LASTEXITCODE
if ($Exit -ne 0) {
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show(
        "A análise terminou com erro. Verifique a pasta:`n$Out",
        "Power BI Analyzer V2"
    ) | Out-Null
}
exit $Exit
