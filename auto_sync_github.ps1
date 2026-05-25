param(
    [string]$RepoPath = "C:\Users\esther.queiroz\Desktop\PainelProdutividade",
    [string]$Branch = "main",
    [int]$DebounceSeconds = 10
)

function Get-GitCommand {
    $comando = Get-Command git -ErrorAction SilentlyContinue
    if ($comando) {
        return $comando.Source
    }

    $candidatos = @(
        "C:\Users\esther.queiroz\AppData\Local\Programs\Git\cmd\git.exe",
        "C:\Users\esther.queiroz\AppData\Local\Programs\Git\bin\git.exe",
        "C:\Program Files\Git\cmd\git.exe",
        "C:\Program Files\Git\bin\git.exe",
        "C:\Program Files (x86)\Git\cmd\git.exe",
        "C:\Program Files (x86)\Git\bin\git.exe"
    )

    foreach ($candidato in $candidatos) {
        if (Test-Path $candidato) {
            return $candidato
        }
    }

    throw "Git não encontrado. Instale o Git ou ajuste o PATH."
}

function Invoke-Git {
    param(
        [string[]]$Arguments
    )

    & $script:GitExe -C $RepoPath @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "Falha ao executar: git $($Arguments -join ' ')"
    }
}

function Sync-Repo {
    try {
        $status = & $script:GitExe -C $RepoPath status --porcelain
        if (-not $status) {
            Write-Host "Sem alterações para enviar."
            return
        }

        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Write-Host "Sincronizando alterações em $timestamp..."

        Invoke-Git @("add", ".")
        Invoke-Git @("commit", "-m", "auto-sync $timestamp")
        Invoke-Git @("push", "origin", $Branch)

        Write-Host "Push concluído."
    }
    catch {
        Write-Warning $_
    }
}

$script:GitExe = Get-GitCommand

if (-not (Test-Path $RepoPath)) {
    throw "Pasta não encontrada: $RepoPath"
}

Write-Host "Monitorando $RepoPath"
Write-Host "Branch: $Branch"
Write-Host "Debounce: $DebounceSeconds segundos"

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $RepoPath
$watcher.Filter = "*.*"
$watcher.IncludeSubdirectories = $true
$watcher.EnableRaisingEvents = $true

$timer = New-Object System.Timers.Timer
$timer.Interval = $DebounceSeconds * 1000
$timer.AutoReset = $false

Register-ObjectEvent -InputObject $timer -EventName Elapsed -Action {
    Sync-Repo
} | Out-Null

$restartTimer = {
    $timer.Stop()
    $timer.Start()
}

Register-ObjectEvent -InputObject $watcher -EventName Changed -Action $restartTimer | Out-Null
Register-ObjectEvent -InputObject $watcher -EventName Created -Action $restartTimer | Out-Null
Register-ObjectEvent -InputObject $watcher -EventName Deleted -Action $restartTimer | Out-Null
Register-ObjectEvent -InputObject $watcher -EventName Renamed -Action $restartTimer | Out-Null

try {
    while ($true) {
        Start-Sleep -Seconds 1
    }
}
finally {
    $watcher.Dispose()
    $timer.Dispose()
}
