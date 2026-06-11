param(
    [string]$RepoPath = "C:\Users\esther.queiroz\Desktop\PainelProdutividade",
    [string]$Branch = "main",
    [string]$Mensagem = ""
)

$ErrorActionPreference = "Stop"

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

    throw "Git não encontrado."
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

$script:GitExe = Get-GitCommand

if (-not (Test-Path $RepoPath)) {
    throw "Pasta não encontrada: $RepoPath"
}

$status = & $script:GitExe -C $RepoPath status --porcelain
if (-not $status) {
    Write-Host "Sem alterações para publicar."
    exit 0
}

if (-not $Mensagem) {
    $Mensagem = "publicacao manual $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
}

Write-Host "Publicando alterações..."
Invoke-Git @("add", "-A")

$statusAposAdd = & $script:GitExe -C $RepoPath status --porcelain
if (-not $statusAposAdd) {
    Write-Host "Nada novo para commit."
    exit 0
}

Invoke-Git @("commit", "-m", $Mensagem)
Invoke-Git @("push", "origin", $Branch)

Write-Host "Publicação concluída."
