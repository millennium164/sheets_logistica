param(
    [Parameter(Mandatory=$true)]
    [string]$WorkbookPath,

    [string]$LogPath = "C:\Scripts\logs\AtualizaExcel.log",

    # Tempo máximo (em minutos) para esperar o refresh terminar
    [int]$TimeoutMinutes = 10,

    # Se quiser salvar como uma cópia (ex: relatório do dia), informe um caminho.
    # Se vazio, salva o próprio arquivo.
    [string]$SaveAsPath = ""
)

# ============= FUNÇÕES AUXILIARES =============
function Write-Log {
    param([string]$Message)
    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[$ts] $Message"
    Write-Host $line
    $logDir = Split-Path -Parent $LogPath
    if (!(Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    Add-Content -Path $LogPath -Value $line
}

function Release-ComObject {
    param([Object]$com)
    try {
        if ($null -ne $com) {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($com)
        }
    } catch {}
}

# ============= VALIDAÇÕES =============
if (!(Test-Path $WorkbookPath)) {
    throw "Arquivo não encontrado: $WorkbookPath"
}

$WorkbookFullPath = (Resolve-Path $WorkbookPath).Path
Write-Log "Iniciando atualização do arquivo: $WorkbookFullPath"

# ============= EXECUÇÃO =============
$excel = $null
$workbook = $null
$start = Get-Date
$timeout = [TimeSpan]::FromMinutes($TimeoutMinutes)

try {
    # Cria instância do Excel
    $excel = New-Object -ComObject Excel.Application

    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false

    # Dica extra de estabilidade (bloqueia macros ao abrir)
    # 3 = msoAutomationSecurityForceDisable
    try { $excel.AutomationSecurity = 3 } catch {}

    # Abre o arquivo
    $workbook = $excel.Workbooks.Open($WorkbookFullPath, 0, $false)
    Write-Log "Arquivo aberto. Tentando ajustar Calculation..."

    # Tenta colocar cálculo manual (opcional)
    $xlCalculationManual = -4135
    try {
    	$excel.Calculation = $xlCalculationManual
    	Write-Log "Calculation definido como Manual."
    } catch {
    	Write-Log "Aviso: não foi possível definir Calculation (continuando). Detalhe: $($_.Exception.Message)"
    }

    Write-Log "Iniciando RefreshAll..."
    $workbook.RefreshAll() | Out-Null


# Tenta colocar cálculo manual (opcional)
$xlCalculationManual = -4135
try {
    $excel.Calculation = $xlCalculationManual
    Write-Log "Calculation definido como Manual."
} catch {
    Write-Log "Aviso: não foi possível definir Calculation (continuando). Detalhe: $($_.Exception.Message)"
}

Write-Log "Iniciando RefreshAll..."
$workbook.RefreshAll() | Out-Null


    # Aguarda terminar queries assíncronas (Power Query costuma ser async)
    # Método existe em versões mais novas
    try {
        $excel.CalculateUntilAsyncQueriesDone()
        Write-Log "CalculateUntilAsyncQueriesDone concluído."
    } catch {
        Write-Log "CalculateUntilAsyncQueriesDone não disponível nesta versão. Usando espera por Refreshing..."
    }

    # Loop de espera: enquanto houver conexão atualizando OU até timeout
    $stillRefreshing = $true
    while ($stillRefreshing) {
        $elapsed = (Get-Date) - $start
        if ($elapsed -gt $timeout) {
            throw "Timeout: refresh não terminou em $TimeoutMinutes minuto(s)."
        }

        $stillRefreshing = $false

        # Verifica conexões
        foreach ($conn in @($workbook.Connections)) {
            try {
                # Algumas conexões expõem OLEDBConnection/ODBCConnection com propriedade Refreshing
                if ($conn.OLEDBConnection -and $conn.OLEDBConnection.Refreshing) { $stillRefreshing = $true }
                if ($conn.ODBCConnection  -and $conn.ODBCConnection.Refreshing)  { $stillRefreshing = $true }
            } catch {}
        }

        # Verifica QueryTables (caso tenha tabelas externas)
        foreach ($ws in @($workbook.Worksheets)) {
            try {
                foreach ($qt in @($ws.QueryTables)) {
                    if ($qt.Refreshing) { $stillRefreshing = $true }
                }
            } catch {}
        }

        if ($stillRefreshing) {
            Start-Sleep -Seconds 3
        }
    }

    Write-Log "Refresh finalizado. Salvando..."

    if ([string]::IsNullOrWhiteSpace($SaveAsPath)) {
        $workbook.Save()
        Write-Log "Arquivo salvo (overwrite)."
    } else {
        $saveDir = Split-Path -Parent $SaveAsPath
        if (!(Test-Path $saveDir)) { New-Item -ItemType Directory -Path $saveDir -Force | Out-Null }

        # FileFormat 51 = xlsx
        $workbook.SaveAs($SaveAsPath, 51)
        Write-Log "Arquivo salvo como: $SaveAsPath"
    }

    Write-Log "Concluído com sucesso."
}
catch {
    Write-Log "ERRO: $($_.Exception.Message)"
    throw
}
finally {
    # Fecha workbook/Excel com segurança
    try { if ($workbook -ne $null) { $workbook.Close($false) | Out-Null } } catch {}
    try { if ($excel -ne $null) { $excel.Quit() | Out-Null } } catch {}

    # Libera COM
    Release-ComObject $workbook
    Release-ComObject $excel

    # Força coleta para reduzir chance do Excel ficar preso
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    Write-Log "Processo finalizado (Excel fechado)."
}