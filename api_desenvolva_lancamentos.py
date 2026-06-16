from __future__ import annotations

import re
import shutil
import tempfile
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from openpyxl import load_workbook
from starlette.requests import Request

from gerar_automacao_extratos import (
    OUT_DIR,
    classify,
    extract_accounts,
    extract_transactions,
)


app = FastAPI(title="Desenvolva Lançamentos API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["GET", "POST", "OPTIONS"],
    allow_headers=["*"],
)

BASE = Path(__file__).resolve().parent
CONTAS = BASE / "Contas.xls"
MODELO_DOMINIO = BASE / "modelo_dominio.xlsm"
JOBS_DIR = OUT_DIR / "api_jobs"
JOBS_DIR.mkdir(parents=True, exist_ok=True)

HISTORICOS = {
    "PAGAMENTO SIMPLES NACIONAL": (61, None),
    "ADIANTAMENTO DE LUCROS JEFFERSON": (62, None),
    "RECEBIMENTO COOPERATIVA": (24, None),
    "PAGAMENTO COMBUSTIVEL": (15, "COMBUSTIVEL"),
}


def slug(text: str) -> str:
    text = text.lower()
    text = re.sub(r"[^a-z0-9]+", "_", text)
    return text.strip("_") or "arquivo"


def save_upload_file(upload: UploadFile, job_dir: Path, prefix: str = "") -> Path:
    if not upload.filename:
        raise HTTPException(status_code=400, detail="Arquivo nao enviado.")

    upload_name = slug(Path(upload.filename).stem) + Path(upload.filename).suffix.lower()
    if prefix:
        upload_name = f"{prefix}_{upload_name}"

    uploaded = job_dir / upload_name
    temp_path = None

    try:
        with tempfile.NamedTemporaryFile(delete=False) as tmp:
            shutil.copyfileobj(upload.file, tmp)
            temp_path = Path(tmp.name)
        shutil.move(str(temp_path), uploaded)
        return uploaded
    except Exception:
        if temp_path and temp_path.exists():
            temp_path.unlink()
        raise


def period_label(period: dict) -> str:
    return f"{period['mes']:02d}_{period['ano']}"


def build_conference_xlsx(entries: list[dict], transactions: list[dict], output: Path) -> None:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Lancamentos"
    headers = [
        "Data",
        "Debitar",
        "Conta Debitar",
        "Creditar",
        "Conta Creditar",
        "Valor",
        "Historico",
        "Descricao",
        "Regra",  
