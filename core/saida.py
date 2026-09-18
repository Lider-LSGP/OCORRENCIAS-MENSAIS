"""Geração das saídas: CSV de importação + XLSX de resultado (com OBS) + ZIP.

Para cada tipo processado saem SEMPRE dois arquivos:
  * importacao_<tipo>[_empresa].csv  -> pronto para o EasyApp
  * resultado_<tipo>[_empresa].xlsx  -> conferência com STATUS colorido e OBS
"""
from __future__ import annotations

import io
import re
import zipfile

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

from .empresas import EMPRESAS

_VERDE = PatternFill("solid", start_color="FFD5F5E3")
_LARANJA = PatternFill("solid", start_color="FFFDEBD0")
_CINZA = PatternFill("solid", start_color="FFE5E7E9")
_VERMELHO = PatternFill("solid", start_color="FFFADBD8")
_AMARELO = PatternFill("solid", start_color="FFFCF3CF")
_CAB = PatternFill("solid", start_color="FF0E2C70")


def _fill_por_status(status: str):
    """Escolhe a cor pelo prefixo/conteúdo do STATUS (cobre Faltas, Férias e
    Afastamentos, cujos textos de status são um pouco diferentes entre si)."""
    s = str(status).upper()
    if s.startswith("CADASTRADO") and "ERRO" in s:
        return _VERMELHO
    if s == "VÁLIDO" or s == "CADASTRADO":
        return _VERDE
    if s in ("EM ABERTO", "NÃO CADASTRADO"):
        return _AMARELO
    if s == "JÁ LANÇADO":
        return _CINZA
    return None


def importacao_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def resultado_xlsx_bytes(df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Resultado"
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        ws.append(row)
        if r_idx == 1:
            for c in ws[r_idx]:
                c.font = Font(bold=True, color="FFFFFFFF")
                c.fill = _CAB
                c.alignment = Alignment(horizontal="center", vertical="center")
        else:
            status = str(row[0])
            fill = _fill_por_status(status)
            if fill:
                for c in ws[r_idx]:
                    c.fill = fill
    for i, col in enumerate(df.columns, 1):
        larg = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)])
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(larg + 2, 60)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def _slug(s: str) -> str:
    s = re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")
    return s[:60]


def montar_arquivos(resultados: dict, dividir_por_empresa: bool) -> dict:
    """resultados: {tipo: (df_imp, df_res)} -> {caminho_no_zip: bytes}."""
    arquivos = {}
    for tipo, (df_imp, df_res) in resultados.items():
        if dividir_por_empresa:
            for codi, nome_emp in EMPRESAS.items():
                mask_i = df_imp["nomeempresa"].astype(str).str.upper().str.contains(
                    nome_emp.split()[0] if codi == "1" else _chave_emp(nome_emp), na=False)
                sub_i = df_imp[mask_i]
                sub_r = df_res[df_res["CODI_EMP"].astype(str) == codi] if "CODI_EMP" in df_res.columns else df_res
                if sub_i.empty and sub_r.empty:
                    continue
                pasta = f"{tipo}/{_slug(nome_emp)}"
                arquivos[f"{pasta}/importacao_{tipo.lower()}_{_slug(nome_emp)}.csv"] = importacao_csv_bytes(sub_i)
                arquivos[f"{pasta}/resultado_{tipo.lower()}_{_slug(nome_emp)}.xlsx"] = resultado_xlsx_bytes(sub_r)
        else:
            arquivos[f"{tipo}/importacao_{tipo.lower()}_todas_empresas.csv"] = importacao_csv_bytes(df_imp)
            arquivos[f"{tipo}/resultado_{tipo.lower()}_todas_empresas.xlsx"] = resultado_xlsx_bytes(df_res)
    return arquivos


def _chave_emp(nome_emp: str) -> str:
    # palavra distintiva para casar empresa na coluna nomeempresa
    for chave in ("ATIVA", "MULTISSERVICOS", "LIMPE", "VSP"):
        if chave in nome_emp.upper():
            return "LIDER" if chave in ("MULTISSERVICOS", "LIMPE") else chave
    return nome_emp.split()[0]


def zip_bytes(arquivos: dict) -> bytes:
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as z:
        for caminho, data in arquivos.items():
            z.writestr(caminho, data)
    return bio.getvalue()
