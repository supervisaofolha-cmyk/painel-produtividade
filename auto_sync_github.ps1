param(
    [string]$RepoPath = "C:\Users\esther.queiroz\Desktop\PainelProdutividade",
    [string]$Branch = "main",
    [int]$DebounceSeconds = 45
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

function Write-SyncLog {
    param(
        [string]$Message
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] $Message"
    Write-Host $line
    Add-Content -Path $script:LogPath -Value $line
}

function Should-IgnorePath {
    param(
        [string]$Path
    )

    if (-not $Path) {
        return $true
    }

    $relativePath = $Path.Replace($RepoPath, "").TrimStart("\")
    $normalizedPath = $relativePath.Replace("/", "\")

    $ignoredPatterns = @(
        ".git\*",
        "__pycache__\*",
        ".playwright-plug\*",
        "backups_lista_apoio\*",
        "*.log",
        "*.pyc",
        "*.tmp",
        "~$*"
    )

    foreach ($pattern in $ignoredPatterns) {
        if ($normalizedPath -like $pattern) {
            return $true
        }
    }

    return $false
}

function Sync-Repo {
    if ($script:IsSyncing) {
        $script:PendingSync = $true
        return
    }

    $script:IsSyncing = $true

    try {
        do {
            $script:PendingSync = $false
            $status = & $script:GitExe -C $RepoPath status --porcelain
            if (-not $status) {
                Write-SyncLog "Sem alterações para enviar."
                break
            }

            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            Write-SyncLog "Sincronizando alterações em $timestamp..."

            Invoke-Git @("add", "-A")

            $statusAposAdd = & $script:GitExe -C $RepoPath status --porcelain
            if (-not $statusAposAdd) {
                Write-SyncLog "Nada novo para commit após git add."
                break
            }

            Invoke-Git @("commit", "-m", "auto-sync $timestamp")
            Invoke-Git @("push", "origin", $Branch)

            Write-SyncLog "Push concluído."
        } while ($script:PendingSync)
    }
    catch {
        Write-Warning $_
        Write-SyncLog "Falha no sync: $_"
    }
    finally {
        $script:IsSyncing = $false
    }
}

$script:GitExe = Get-GitCommand
$script:LogPath = Join-Path $RepoPath "auto_sync_github.log"
$script:IsSyncing = $false
$script:PendingSync = $false

if (-not (Test-Path $RepoPath)) {
    throw "Pasta não encontrada: $RepoPath"
}

Write-SyncLog "Monitorando $RepoPath"
Write-SyncLog "Branch: $Branch"
Write-SyncLog "Debounce: $DebounceSeconds segundos"

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
    $path = $Event.SourceEventArgs.FullPath
    if (Should-IgnorePath $path) {
        return
    }

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
