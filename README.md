# Desenvolva Lançamentos API

API do motor real para processar extratos bancários e gerar arquivos para importação no Domínio Sistemas.

## Rotas

- `GET /health`
- `POST /processar-extrato`
- `GET /download/{job_id}/{arquivo}`

## Processar Extrato

Envie `multipart/form-data` para `POST /processar-extrato` com:

- `arquivo`: PDF do extrato bancário.
- `plano_contas`: arquivo Excel do plano de contas da empresa. Opcional; se não for enviado, a API usa `Contas.xls`.
- `senha`: senha do PDF, opcional.

A resposta mantém os indicadores, a tabela de lançamentos e os links em `downloads` para conferência XLSX, modelo Domínio XLSM, entradas TXT e pacote ZIP.

## Render

Build Command:

```text
pip install -r requirements.txt
```

Start Command:

```text
uvicorn api_desenvolva_lancamentos:app --host 0.0.0.0 --port $PORT
```

Depois de publicar, teste:

```text
https://SEU-LINK.onrender.com/health
```
