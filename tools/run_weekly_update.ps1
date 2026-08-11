# run_weekly_update.ps1 - Ejecutado cada lunes y viernes a las 09:00 por Task Scheduler
# Lanza Claude Code con el prompt de actualizacion semanal Finocam (9 mercados)

$ErrorActionPreference = "Continue"
$ToolsDir = "C:\Users\Daniela\Desktop\Git Finocam\tools"
$LogFile  = "$ToolsDir\logs\update_$(Get-Date -Format 'yyyy-MM-dd').log"

New-Item -ItemType Directory -Force -Path "$ToolsDir\logs" | Out-Null
"=== Finocam Weekly Update - $(Get-Date) ===" | Tee-Object -FilePath $LogFile

# Cargar credenciales desde .env.local (nunca commitear este fichero)
$EnvFile = "$ToolsDir\.env.local"
if (-not (Test-Path $EnvFile)) {
    "ERROR: No se encontro $EnvFile" | Tee-Object -FilePath $LogFile -Append
    exit 1
}

Get-Content $EnvFile | ForEach-Object {
    if ($_ -match '^([^#=]+)=(.+)$') {
        [System.Environment]::SetEnvironmentVariable($matches[1].Trim(), $matches[2].Trim(), 'Process')
    }
}
"Credenciales cargadas desde .env.local" | Add-Content $LogFile

# Escribir prompt expandido (con credenciales) a fichero temporal
$PromptTemplate = Get-Content "$ToolsDir\weekly_update_prompt.md" -Raw -Encoding UTF8
$ExpandedPrompt = $PromptTemplate `
    -replace '\$env:FINOCAM_GITHUB_PAT',   $env:FINOCAM_GITHUB_PAT `
    -replace '\$env:FINOCAM_SENDGRID_KEY', $env:FINOCAM_SENDGRID_KEY
$TempPrompt = "$env:TEMP\finocam_weekly_prompt.md"
[System.IO.File]::WriteAllText($TempPrompt, $ExpandedPrompt, [System.Text.Encoding]::UTF8)

# Ejecutar Claude Code: instruccion corta que lee el fichero de prompt
$Instruction = "Lee el fichero '$TempPrompt' y ejecuta las instrucciones que contiene exactamente, paso a paso."
"Iniciando Claude Code..." | Add-Content $LogFile
& claude --dangerously-skip-permissions -p $Instruction 2>&1 | Tee-Object -FilePath $LogFile -Append

Remove-Item $TempPrompt -Force -ErrorAction SilentlyContinue

"=== Finalizado: $(Get-Date) ===" | Add-Content $LogFile
