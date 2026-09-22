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

from .empresas import EMPRESAS, APELIDOS, apelido_empresa

_LILAS = PatternFill("solid", start_color="FFE8DAEF")   # p/ desativar (ação pendente)

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
    if "PARA DESATIVAR" in s or s == "CONFERÊNCIA MANUAL" or "DEMITIR NO SISTEMA" in s:
        return _LILAS
    if "DEMISSÃO FUTURA" in s:
        return _LARANJA
    if s.startswith("CADASTRADO") and "ERRO" in s:
        return _VERMELHO
    if s == "VÁLIDO" or s.startswith("CADASTRADO"):
        return _VERDE
    if s.startswith("NÃO CADASTRADO") or s == "EM ABERTO":
        return _AMARELO
    if s == "JÁ LANÇADO":
        return _CINZA
    return None


def importacao_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


_TITULOS_ABA = {
    "FALTAS": "Faltas", "FERIAS": "Férias", "AFASTAMENTOS": "Afastamentos",
    "RESCISOES": "Rescisões-Avisos", "AVISOS": "Rescisões-Avisos",
}


def _escrever_aba(ws, df: pd.DataFrame, colorir_status: bool = True):
    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        ws.append(row)
        if r_idx == 1:
            for c in ws[r_idx]:
                c.font = Font(bold=True, color="FFFFFFFF")
                c.fill = _CAB
                c.alignment = Alignment(horizontal="center", vertical="center")
        elif colorir_status:
            fill = _fill_por_status(row[0])
            if fill:
                for c in ws[r_idx]:
                    c.fill = fill
    for i, col in enumerate(df.columns, 1):
        larg = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)])
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(larg + 2, 60)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def _col_conf(nome):
    for c in nome:
        if re.search(r"(^|[^A-ZÀ-ÿ])NOME(?! NO SISTEMA)", c):
            return c
    return None


def _montar_conferencia(tipo: str, df_res: pd.DataFrame) -> pd.DataFrame:
    """Aba 1 — no formato das planilhas de conferência que a operação preenchia
    manualmente: só as colunas de conferência + uma coluna OBS vazia para
    anotações. STATUS sai pré-preenchido só quando está verdinho na aba
    Resultado (CADASTRADO, ou CADASTRADO - DEMITIDO quando a ficha já mostra
    demissão). Nos demais casos (amarelo/vermelho/lilás...) fica em branco —
    é a operação quem preenche, como fazia antes."""
    if df_res is None or df_res.empty:
        return pd.DataFrame()
    cols = list(df_res.columns)
    def pegar(*nomes):
        for n in nomes:
            if n in cols:
                return n
        return None
    out = {}
    if tipo == "FERIAS":
        mapa = {"NOME": pegar("NOME (DOMÍNIO)"),
                "INÍCIO": pegar("INÍCIO ESPERADO"),
                "FIM": pegar("FIM ESPERADO"),
                "QTD DIAS": pegar("QTD DIAS")}
    elif tipo == "AFASTAMENTOS":
        mapa = {"NOME": pegar("NOME (DOMÍNIO)"),
                "INÍCIO": pegar("INÍCIO (UNIFICADO)"),
                "FIM": pegar("FIM (UNIFICADO)"),
                "MOTIVO": pegar("MOTIVO(S) DOMÍNIO"),
                "QTD DIAS": pegar("QTD DIAS")}
    elif tipo in ("RESCISOES", "AVISOS"):
        mapa = {"NOME": pegar("NOME (DOMÍNIO)"),
                "ADMISSÃO": pegar("ADMISSÃO"),
                "INÍCIO": pegar("DATA AVISO"),
                "FIM": pegar("DATAFIM OCORRÊNCIA"),
                "DEMISSÃO": pegar("DATA DEMISSÃO (FICHA)"),
                "CONCESSOR": pegar("CONCEDIDO POR"),
                "TIPO OCORRÊNCIA": pegar("TIPO OCORRÊNCIA"),
                "MOTIVO": pegar("MOTIVO/CATEGORIA")}
    else:  # FALTAS
        mapa = {"NOME": pegar("NOME (DOMÍNIO)"),
                "REFERÊNCIA": pegar("REFERÊNCIA"),
                "PERÍODO": pegar("PERÍODO LANÇADO"),
                "DIAS": pegar("DIAS DE FALTA")}
    for destino, origem in mapa.items():
        out[destino] = df_res[origem] if origem else ""
    conf = pd.DataFrame(out)
    status = df_res["STATUS"].astype(str)
    obs = df_res["OBS"].astype(str) if "OBS" in df_res.columns else pd.Series([""] * len(df_res))
    def _status_conf(i):
        s = status.iloc[i]
        if s == "CADASTRADO" or s == "VÁLIDO":
            dem = re.search(r"DEMISSÃO no sistema: ([0-9/]+)", obs.iloc[i])
            return "CADASTRADO - DEMITIDO" if dem else s
        return ""  # cores não-verdes: a operação preenche manualmente
    conf["STATUS"] = [_status_conf(i) for i in range(len(conf))]
    conf["OBS"] = ""
    return conf


def _montar_desativar(tipo: str, df_res: pd.DataFrame) -> pd.DataFrame:
    """Aba extra só em Rescisões/Avisos: quem já tem a ocorrência conferida e
    falta só a DESATIVAÇÃO manual da ficha no sistema."""
    if tipo not in ("RESCISOES", "AVISOS") or df_res is None or df_res.empty:
        return pd.DataFrame()
    if "PARA DESATIVAR" not in df_res.columns:
        return pd.DataFrame()
    sub = df_res[df_res["PARA DESATIVAR"].astype(str) == "SIM"].copy()
    if sub.empty:
        return pd.DataFrame()
    keep = [c for c in ("STATUS", "NOME (DOMÍNIO)", "EMPRESA", "TIPO OCORRÊNCIA",
                        "DATA DEMISSÃO (FICHA)", "ADMISSÃO", "CPF") if c in sub.columns]
    out = sub[keep].copy()
    if "STATUS" in out.columns:
        out["STATUS"] = out["STATUS"].astype(str).str.replace(" | PARA DESATIVAR", "", regex=False)
    return out


def resultado_xlsx_bytes(df: pd.DataFrame, titulo_aba: str = "Resultado",
                         tipo: str = None) -> bytes:
    wb = Workbook()
    ws = wb.active
    # nome da aba precisa ser legível e caber no limite de 31 caracteres do Excel
    ws.title = str(titulo_aba)[:31] or "Resultado"
    _escrever_aba(ws, df, colorir_status=True)
    if tipo:
        conf = _montar_conferencia(tipo, df)
        if not conf.empty:
            _escrever_aba(wb.create_sheet("Conferencia"), conf, colorir_status=True)
        desat = _montar_desativar(tipo, df)
        if not desat.empty:
            _escrever_aba(wb.create_sheet("Para Desativar"), desat, colorir_status=True)
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def _slug(s: str) -> str:
    s = re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_")
    return s[:60]


def _nome_bonitinho(tipo: str) -> str:
    """Nome legível do tipo de ocorrência p/ usar em nomes de arquivo/aba."""
    return _TITULOS_ABA.get(tipo, tipo.capitalize())


TIPO_NOME_ARQ = {
    "FALTAS": "FALTAS",
    "FERIAS": "FERIAS",
    "AFASTAMENTOS": "AFASTAMENTOS",
    "RESCISOES": "RESCISOES",
    "AVISOS": "RESCISOES",
}

# apelidos curtos usados nos NOMES de pastas/arquivos do ZIP (como a operação
# organiza as pastas na prática: ATIVA / L COMERCIAL / L MULTISSERVICOS / VSP)
APELIDOS_ARQUIVO = {
    "1": "VSP",
    "2": "ATIVA",
    "3": "L MULTISSERVICOS",
    "4": "L COMERCIAL",
}


def montar_arquivos(resultados: dict, dividir_por_empresa: bool = True,
                    pasta_mes: str = None) -> dict:
    """resultados: {tipo: (df_imp, df_res)} -> {caminho_no_zip: bytes}.

    Estrutura do ZIP (igual à organização manual da operação):
        MES 09/
          ATIVA/
            FALTAS ATIVA - IMPORTACAO.csv
            FALTAS ATIVA - RESULTADO.xlsx
            FERIAS ATIVA - IMPORTACAO.csv
            ...
          L COMERCIAL/ ...
          L MULTISSERVICOS/ ...
          VSP/ ...
    (o arquivo "ALERTA - DEMISSOES FUTURAS.xlsx" na raiz do MES é adicionado
    pelo app.py, não por aqui)
    """
    from datetime import date as _date
    pasta_raiz = pasta_mes or f"MES {_date.today().month:02d}"
    arquivos = {}
    for tipo, (df_imp, df_res) in resultados.items():
        nome_tipo = TIPO_NOME_ARQ.get(tipo, tipo)
        if dividir_por_empresa:
            for codi, nome_emp in EMPRESAS.items():
                apelido = APELIDOS_ARQUIVO.get(codi, APELIDOS.get(codi, codi))
                sub_r = df_res[df_res["CODI_EMP"].astype(str) == codi] \
                    if "CODI_EMP" in df_res.columns else df_res
                # CORREÇÃO: a importação é dividida pelos parceiro_ids que estão
                # no RESULTADO desta empresa — antes filtrava pelo texto de
                # nomeempresa com a palavra "LIDER" para as duas Líder, o que
                # jogava as linhas das duas empresas nas duas pastas (CSV da
                # MULTIS saía com conteúdo da COMERCIAL). Só cai no filtro por
                # nome se não houver parceiro_id para cruzar.
                if not df_imp.empty and "parceiro_id" in df_imp.columns and not sub_r.empty \
                        and "ID SISTEMA (parceiro_id)" in sub_r.columns:
                    ids_emp = set(sub_r["ID SISTEMA (parceiro_id)"].astype(str)) - {""}
                    sub_i = df_imp[df_imp["parceiro_id"].astype(str).isin(ids_emp)]
                elif not df_imp.empty and "nomeempresa" in df_imp.columns:
                    sub_i = df_imp[df_imp["nomeempresa"].astype(str) == nome_emp]
                else:
                    sub_i = df_imp
                if sub_i.empty and sub_r.empty:
                    continue
                pasta = f"{pasta_raiz}/{apelido}"
                arquivos[f"{pasta}/{nome_tipo} {apelido} - IMPORTACAO.csv"] = importacao_csv_bytes(sub_i)
                arquivos[f"{pasta}/{nome_tipo} {apelido} - RESULTADO.xlsx"] = resultado_xlsx_bytes(
                    sub_r, titulo_aba=f"{nome_tipo[:20]} {apelido}"[:31], tipo=tipo)
        else:
            pasta = f"{pasta_raiz}/TODAS EMPRESAS"
            arquivos[f"{pasta}/{nome_tipo} - IMPORTACAO.csv"] = importacao_csv_bytes(df_imp)
            arquivos[f"{pasta}/{nome_tipo} - RESULTADO.xlsx"] = resultado_xlsx_bytes(
                df_res, titulo_aba=nome_tipo, tipo=tipo)
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
