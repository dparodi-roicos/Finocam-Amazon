# run_weekly_update.ps1 — Ejecutado cada lunes a las 09:00 por Task Scheduler
# Lanza Claude Code con el prompt de actualización semanal Finocam

$ErrorActionPreference = "Continue"
$ToolsDir  = "C:\Users\Daniela\Desktop\Git Finocam\tools"
$LogFile   = "$ToolsDir\logs\update_$(Get-Date -Format 'yyyy-MM-dd').log"

New-Item -ItemType Directory -Force -Path "$ToolsDir\logs" | Out-Null
"=== Finocam Weekly Update - $(Get-Date) ===" | Tee-Object -FilePath $LogFile

# Cargar credenciales desde .env.local (nunca commitear este fichero)
$EnvFile = "$ToolsDir\.env.local"
if (Test-Path $EnvFile) {
    Get-Content $EnvFile | ForEach-Object {
        if ($_ -match '^([^#=]+)=(.+)$') {
            [System.Environment]::SetEnvironmentVariable($matches[1].Trim(), $matches[2].Trim(), 'Process')
        }
    }
    "Credenciales cargadas desde .env.local" | Add-Content $LogFile
} else {
    "ERROR: No se encontro $EnvFile — abortar" | Tee-Object -FilePath $LogFile -Append
    exit 1
}

# Leer prompt e inyectar variables de entorno
$PromptTemplate = Get-Content "$ToolsDir\weekly_update_prompt.md" -Raw -Encoding UTF8
$Prompt = $PromptTemplate `
    -replace '\$env:FINOCAM_GITHUB_PAT',  $env:FINOCAM_GITHUB_PAT `
    -replace '\$env:FINOCAM_SENDGRID_KEY', $env:FINOCAM_SENDGRID_KEY

# Ejecutar Claude Code con el prompt (sin confirmaciones)
"Iniciando Claude Code..." | Add-Content $LogFile
& claude --dangerously-skip-permissions -p $Prompt 2>&1 | Tee-Object -FilePath $LogFile -Append

"=== Finalizado: $(Get-Date) ===" | Add-Content $LogFile
