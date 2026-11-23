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

from .empresas import EMPRESAS, APELIDOS, apelido_empresa, codi_por_nome, codigo_empresa

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
    """Conferência resumida: cadastro da ocorrência, validação e OBS manual."""
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
                "DEMISSÃO": pegar("DATA DEMISSÃO (FICHA)", "DATA DEMISSÃO"),
                "CONCESSOR": pegar("CONCEDIDO POR"),
                "TIPO OCORRÊNCIA": pegar("TIPO OCORRÊNCIA"),
                "MOTIVO": pegar("MOTIVO/CATEGORIA")}
    else:  # FALTAS
        mapa = {"NOME": pegar("NOME (DOMÍNIO)"),
                "REFERÊNCIA": pegar("REFERÊNCIA"),
                "PERÍODO": pegar("PERÍODO LANÇADO"),
                "DIAS": pegar("DIAS DE FALTA")}
    vazio = pd.Series([""] * len(df_res), index=df_res.index, dtype=object)
    for destino, origem in mapa.items():
        # coluna ausente tem de virar uma SÉRIE vazia (antes virava escalar "" e
        # o pd.DataFrame estourava "If using all scalar values..." — era uma das
        # formas de o tipo inteiro desaparecer do ZIP)
        out[destino] = df_res[origem].reset_index(drop=True) if origem else vazio.reset_index(drop=True)
    if not out:
        return pd.DataFrame()
    conf = pd.DataFrame(out)
    cadastros, validacoes = [], []
    rescisao = tipo in ("RESCISOES", "AVISOS")
    for _, linha in df_res.fillna("").iterrows():
        status = str(linha.get("STATUS", "")).strip()
        principal = status.split(" | ", 1)[0]
        obs = str(linha.get("OBS", "")).strip()
        cadastro, validacao = "", ""
        if principal.startswith("CADASTRADO"):
            cadastro = re.sub(r"ERRO DATA IN[IÍ]CIO:", "ERRO DATA INICIAL: ", principal)
            cadastro = re.sub(r"(ERRO DATA FINAL:)\s*", r"\1 ", cadastro)
            cadastro = re.sub(r"(\d{2}/\d{2}/)\d{2}(\d{2})\b", r"\1\2", cadastro)
        elif tipo == "FALTAS" and principal == "JÁ LANÇADO":
            tipos = str(linha.get("TIPOS NA DATA", "")).split(",")
            if "13" in tipos:
                cadastro = "CADASTRADO"
            else:
                cadastro = "NÃO CADASTRADO"
                validacao = obs or "OUTRA OCORRÊNCIA NA MESMA DATA — CONFERIR"
        elif principal.startswith("NÃO CADASTRADO") or principal in ("VÁLIDO", "EM ABERTO", "NÃO LANÇAR"):
            cadastro = "NÃO CADASTRADO"
        elif principal == "CONFERÊNCIA MANUAL":
            cadastro = "CONFERÊNCIA MANUAL"

        if rescisao:
            ativo = str(linha.get("ATIVO NO SISTEMA", "")).strip().casefold()
            if "DEMISSÃO FUTURA" in status or linha.get("DEMISSÃO FUTURA?", "") == "SIM":
                validacao = "DEMISSÃO FUTURA"
            elif ativo in {"não", "nao", "n"}:
                validacao = "DESLIGADO"
            elif ativo in {"sim", "s"}:
                validacao = "NÃO DESLIGADO"
            else:
                validacao = "SITUAÇÃO NÃO CONFIRMADA"
        else:
            validacao = validacao or obs
        cadastros.append(cadastro)
        validacoes.append(validacao)
    conf["CADASTRO"] = cadastros
    conf["VALIDAÇÃO"] = validacoes
    conf["OBS"] = ""
    return conf


def _colorir_conferencia(ws, conf):
    cores_cadastro = {"CADASTRADO": _VERDE, "NÃO CADASTRADO": _VERMELHO,
                      "CONFERÊNCIA MANUAL": _AMARELO}
    cores_validacao = {"DESLIGADO": _VERDE, "NÃO DESLIGADO": _VERMELHO,
                       "DEMISSÃO FUTURA": _LILAS, "SITUAÇÃO NÃO CONFIRMADA": _AMARELO}
    for coluna, cores in (("CADASTRO", cores_cadastro), ("VALIDAÇÃO", cores_validacao)):
        indice = conf.columns.get_loc(coluna) + 1
        for i, valor in enumerate(conf[coluna], start=2):
            fill = cores.get(valor)
            if coluna == "VALIDAÇÃO" and "NÃO LANÇAR FALTA 30 DIAS" in valor:
                fill = _VERMELHO
            if coluna == "CADASTRO" and valor.startswith("CADASTRADO (ERRO DATA"):
                fill = _VERMELHO
            cel = ws.cell(row=i, column=indice)
            if fill:
                cel.fill = fill
                cel.font = Font(bold=True)
            cel.alignment = Alignment(vertical="center", wrap_text=True)


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
        ws_conf = wb.create_sheet("Conferencia", 0)
        _escrever_aba(ws_conf, conf, colorir_status=False)
        _colorir_conferencia(ws_conf, conf)
        wb.active = 0
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
          ATIVA/ FALTAS ATIVA - IMPORTACAO.csv / - RESULTADO.xlsx ...
          L COMERCIAL/ ...
          L MULTISSERVICOS/ ...
          VSP/ ...
          A CONFERIR/   <- só se sobrar linha sem empresa identificada
    NADA é descartado em silêncio: toda linha do resultado sai em algum
    arquivo (empresa identificada ou "A CONFERIR").
    """
    from datetime import date as _date
    pasta_raiz = pasta_mes or f"MES {_date.today().month:02d}"
    arquivos = {}

    for tipo, (df_imp, df_res) in resultados.items():
        nome_tipo = TIPO_NOME_ARQ.get(tipo, tipo)
        df_imp = (df_imp if df_imp is not None else pd.DataFrame()).reset_index(drop=True)
        df_res = (df_res if df_res is not None else pd.DataFrame()).reset_index(drop=True)

        if not dividir_por_empresa:
            pasta = f"{pasta_raiz}/TODAS EMPRESAS"
            arquivos[f"{pasta}/{nome_tipo} - IMPORTACAO.csv"] = importacao_csv_bytes(df_imp)
            arquivos[f"{pasta}/{nome_tipo} - RESULTADO.xlsx"] = resultado_xlsx_bytes(
                df_res, titulo_aba=nome_tipo, tipo=tipo)
            continue

        codigos = df_res.get("CODI_EMP", pd.Series("", index=df_res.index)).map(codigo_empresa)
        pids = df_res.get("ID SISTEMA (parceiro_id)", pd.Series("", index=df_res.index)).fillna("").astype(str).str.strip()
        empresas_por_id = {}
        for pid, cod in zip(pids, codigos):
            if pid:
                empresas_por_id.setdefault(pid, set()).add(cod)
        encaminhar = []
        for _, linha in df_imp.iterrows():
            cod = codi_por_nome(linha.get("nomeempresa", ""))
            pid = str(linha.get("parceiro_id", "")).strip()
            encaminhar.append(cod if cod and empresas_por_id.get(pid) == {cod} else "")
        destino = pd.Series(encaminhar, index=df_imp.index, dtype=str)
        for codi in EMPRESAS:
            sub_r = df_res[codigos == codi]
            sub_i = df_imp[destino == codi]
            if sub_r.empty:
                continue
            apelido = APELIDOS_ARQUIVO[codi]
            pasta = f"{pasta_raiz}/{apelido}"
            arquivos[f"{pasta}/{nome_tipo} {apelido} - IMPORTACAO.csv"] = importacao_csv_bytes(sub_i)
            arquivos[f"{pasta}/{nome_tipo} {apelido} - RESULTADO.xlsx"] = resultado_xlsx_bytes(
                sub_r, titulo_aba=f"{nome_tipo[:20]} {apelido}"[:31], tipo=tipo)
        pasta = f"{pasta_raiz}/A CONFERIR"
        if (codigos == "").any():
            arquivos[f"{pasta}/{nome_tipo} - RESULTADO.xlsx"] = resultado_xlsx_bytes(
                df_res[codigos == ""], titulo_aba=nome_tipo, tipo=tipo)
        if (destino == "").any():
            rejeitadas = df_imp[destino == ""].loc[:, [c for c in ("nomefuncionario", "parceiro_id", "nomeempresa") if c in df_imp]]
            arquivos[f"{pasta}/{nome_tipo} - VINCULOS BLOQUEADOS.xlsx"] = resultado_xlsx_bytes(
                rejeitadas, titulo_aba="Vínculos bloqueados")

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
