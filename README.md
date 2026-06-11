# Desenvolva Lançamentos API

API do motor real para processar extratos bancários e gerar arquivos para importação no Domínio Sistemas.

## Rotas

- `GET /health`
- `POST /processar-extrato`
- `GET /download/{job_id}/{arquivo}`

## Render

Build Command:

```text
pip install -r requirements.txt
```

Start Command:

```text
python api_desenvolva_lancamentos.py
```

Depois de publicar, teste:

```text
https://SEU-LINK.onrender.com/health
```
