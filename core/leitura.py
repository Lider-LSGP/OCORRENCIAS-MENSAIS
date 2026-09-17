"""Leitura tolerante de planilhas: xlsx, xls (BIFF) e SYLK da Domínio."""
from __future__ import annotations

import subprocess, shutil, tempfile, glob, os
import pandas as pd


def _sniff(path: str) -> str:
    with open(path, "rb") as f:
        magic = f.read(8)
    if magic[:2] == b"PK":
        return "xlsx"
    if magic.startswith(b"\xd0\xcf\x11\xe0"):
        return "biff"
    return "sylk_ou_texto"


def ler_planilha(path: str, dtype=str, **kw) -> pd.DataFrame:
    """Lê qualquer planilha enviada. SYLK antigo cai no xlrd."""
    fmt = _sniff(path)
    if fmt == "xlsx":
        return pd.read_excel(path, engine="openpyxl", dtype=dtype, **kw)
    try:
        return pd.read_excel(path, engine="xlrd", dtype=dtype, **kw)
    except Exception:
        pass
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        with tempfile.TemporaryDirectory() as td:
            subprocess.run([soffice, "--headless", "--convert-to", "xlsx",
                            "--outdir", td, path], capture_output=True, timeout=180)
            outs = glob.glob(os.path.join(td, "*.xlsx"))
            if outs:
                return pd.read_excel(outs[0], engine="openpyxl", dtype=dtype, **kw)
    raise RuntimeError(f"Não foi possível ler '{os.path.basename(path)}'.")


def ler_arquivo_upload(up_file, tmp_dir: str, dtype=str, **kw) -> pd.DataFrame:
    """Salva um UploadedFile do Streamlit e lê."""
    path = os.path.join(tmp_dir, up_file.name)
    with open(path, "wb") as f:
        f.write(up_file.getbuffer())
    return ler_planilha(path, dtype=dtype, **kw)
