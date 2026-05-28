# painel-produtividade

## Serviço local da planilha

Para atualizar a `produtividade.xlsx` direto no seu computador, inicie:

```powershell
powershell -ExecutionPolicy Bypass -File C:\Users\esther.queiroz\Desktop\PainelProdutividade\iniciar_servico_planilha.ps1
```

O serviço sobe em `http://127.0.0.1:8765`.

Rotas disponíveis:

- `GET /health`
- `POST /atualizar`

Exemplo para atualizar BI, RO e SGD do último dia útil:

```powershell
Invoke-RestMethod `
  -Uri http://127.0.0.1:8765/atualizar `
  -Method POST `
  -ContentType "application/json" `
  -Body '{"fonte":"todas"}'
```

Exemplo para atualizar só o BI de uma data:

```powershell
Invoke-RestMethod `
  -Uri http://127.0.0.1:8765/atualizar `
  -Method POST `
  -ContentType "application/json" `
  -Body '{"fonte":"bi","data":"27/05/2026"}'
```
