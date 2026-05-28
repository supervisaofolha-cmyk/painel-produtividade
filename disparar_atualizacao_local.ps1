param(
    [string]$Fonte = "todas",
    [string]$Data = ""
)

$body = @{
    fonte = $Fonte
}

if ($Data) {
    $body.data = $Data
}

Invoke-RestMethod `
    -Uri "http://127.0.0.1:8765/atualizar" `
    -Method POST `
    -ContentType "application/json" `
    -Body ($body | ConvertTo-Json -Compress)
