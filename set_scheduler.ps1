# set_scheduler.ps1
# Cria/atualiza tarefa: Seg-Sex 09:00 e 14:00, somente com usuario logado

$taskName = "AtualizarBaseExcel"
$ps = "powershell.exe"

# Caminhos (ajuste se necessario)
$scriptPath   = "C:\Users\$env:UserName\downloads\comparador_planilhas_logistica\script_update_sheet.ps1"
$workbookPath = "C:\Users\$env:UserName\downloads\comparador_planilhas_logistica\base_copia.xlsx"
$workingDir   = "C:\Users\$env:UserName\downloads\comparador_planilhas_logistica"

# Valida arquivos
if (-not (Test-Path $scriptPath))   { throw "Script nao encontrado: $scriptPath" }
if (-not (Test-Path $workbookPath)) { throw "Workbook nao encontrado: $workbookPath" }

# Argumentos (com escape correto)
$args = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -WorkbookPath `"$workbookPath`""

# Acao
$action = New-ScheduledTaskAction -Execute $ps -Argument $args -WorkingDirectory $workingDir

# Triggers: Somente dias de semana (Monday-Friday)
$trigger1 = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At 9:00am
$trigger2 = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At 2:00pm

# Principal: somente quando usuario estiver conectado (Interactive)
$principal = New-ScheduledTaskPrincipal `
  -UserId "$env:UserDomain\$env:UserName" `
  -LogonType Interactive `
  -RunLevel Highest

# Settings
$settings = New-ScheduledTaskSettingsSet `
  -StartWhenAvailable `
  -MultipleInstances IgnoreNew `
  -ExecutionTimeLimit (New-TimeSpan -Minutes 60)

# Registra (cria/atualiza)
Register-ScheduledTask `
  -TaskName $taskName `
  -Action $action `
  -Trigger @($trigger1,$trigger2) `
  -Principal $principal `
  -Settings $settings `
  -Force | Out-Null

Write-Host 'Tarefa criada/atualizada: Seg-Sex 09:00 e 14:00 (somente com usuario logado).'