"""Constantes e montagem do Layout de Importação de Ocorrências (EasyApp)."""
from __future__ import annotations

import calendar
from datetime import date, timedelta

COLUNAS_LAYOUT = [
    "id", "parceiro_id", "nomefuncionario", "codigoresponsavelcobertura",
    "nomeresponsavelcobertura", "tipoocorrencia_id", "descontabeneficio",
    "nomeocorrencia", "nomeempresa", "nomeescala", "nomeposto", "datainicio",
    "qtddiasafastamento", "datafim", "dataprevisaoretorno", "mes", "ano",
    "mes_ano", "usoregcreated_at", "statusocorrencia", "datacadastro",
    "dataocorrencia", "postotrabalhoativo", "system_unit_id",
    "nivelcoberturapostooc", "totaldiasdescontado", "geramapabeneficio",
    "tipooccontrolaferias", "tipoferiasoc",
]

# nunca preenchidos
VAZIAS = {"id", "codigoresponsavelcobertura", "nomeresponsavelcobertura",
          "usoregcreated_at", "system_unit_id", "totaldiasdescontado"}

# defaults que vêm dos registros-modelo do layout (FALTA 30 DIAS)
DEFAULTS_FIXOS = {
    "tipoocorrencia_id": "13",
    "descontabeneficio": "Sim",
    "nomeocorrencia": "FALTA 30 DIAS",
    "statusocorrencia": "Ativo",
    "postotrabalhoativo": "Sim",
    "nivelcoberturapostooc": "Alto",
    "geramapabeneficio": "Sim",
    "tipooccontrolaferias": "Não",
    "tipoferiasoc": "Não",
}


def mes_seguinte(mes: int, ano: int):
    """Mês seguinte ao de referência (falta do mês X é lançada no mês X+1)."""
    return mes % 12 + 1, ano + (1 if mes == 12 else 0)


def periodo_mes(mes: int, ano: int):
    """(início, fim, qtde_dias) do mês — trata 28/29/30/31 e ano bissexto."""
    dias = calendar.monthrange(ano, mes)[1]
    return date(ano, mes, 1), date(ano, mes, dias), dias


def linha_layout(parceiro_id="", nomefuncionario="", nomeempresa="", nomeescala="",
                 nomeposto="", datainicio: date = None, datafim: date = None,
                 hoje: date = None, fixos: dict = None) -> dict:
    """Monta uma linha do layout no formato exato (datas aaaa-mm-dd, mes_ano mm/aaaa)."""
    fixos = fixos or DEFAULTS_FIXOS
    hoje = hoje or date.today()
    row = {c: "" for c in COLUNAS_LAYOUT}
    row.update(fixos)
    row["parceiro_id"] = str(parceiro_id or "")
    row["nomefuncionario"] = nomefuncionario
    row["nomeempresa"] = nomeempresa
    row["nomeescala"] = nomeescala
    row["nomeposto"] = nomeposto
    if datainicio and datafim:
        row["datainicio"] = datainicio.strftime("%Y-%m-%d")
        row["datafim"] = datafim.strftime("%Y-%m-%d")
        row["qtddiasafastamento"] = str((datafim - datainicio).days + 1)
        row["dataprevisaoretorno"] = (datafim + timedelta(days=1)).strftime("%Y-%m-%d")
        row["dataocorrencia"] = datainicio.strftime("%Y-%m-%d")
    row["mes"] = f"{hoje.month:02d}"
    row["ano"] = str(hoje.year)
    row["mes_ano"] = hoje.strftime("%m/%Y")
    row["datacadastro"] = hoje.strftime("%Y-%m-%d")
    for c in VAZIAS:
        row[c] = ""
    return row
