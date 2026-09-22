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

# defaults comuns a qualquer tipo de ocorrência (podem ser sobrescritos por módulo)
BASE_DEFAULTS = {
    "statusocorrencia": "Ativo",
    "postotrabalhoativo": "Sim",
    "nivelcoberturapostooc": "Alto",
    "geramapabeneficio": "Sim",
    "tipooccontrolaferias": "Não",
    "tipoferiasoc": "Não",
}

# ---- FALTAS MÊS TODO ----
DEFAULTS_FALTAS = {
    **BASE_DEFAULTS,
    "tipoocorrencia_id": "13",
    "descontabeneficio": "Sim",
    "nomeocorrencia": "FALTA 30 DIAS",
}

# nomes de exibição por tipoocorrencia_id (para mensagens de "tipo errado")
NOMES_TIPO = {
    "13": "FALTA 30 DIAS",
    "31": "FERIAS GERAL",
    "32": "FERIAS VSP",
    "34": "FERIAS MOTORISTA",
    "18": "AFASTAMENTO INSS",
    "35": "LICENCA MATERNIDADE",
    "41": "LICENCA PATERNA",
    "33": "AVISO PREVIO COLABORADOR",
    "36": "PEDIDO DE DEMISSAO",
    "38": "AVISO PREVIO EMPRESA",
    "39": "ENCERRAMENTO DE CONTRATO",
    "40": "AVISO PREVIO EMPRESA 1 DIA",
}

# tipos adicionais vistos nos relatórios (cancelamentos, faltas, atestados...)
NOMES_TIPO.update({
    "1": "FALTA 1 DIA",
    "2": "ATESTADO MEDICO",
    "17": "CANCELAMENTO DE AFASTAMENTO INSS",
    "19": "CANCELAMENTO DE OCORRENCIA",
    "29": "CANCELAMENTO DE FERIAS",
    "30": "RETORNO",
    "37": "CANCELAMENTO DE AVISO PREVIO",
    "42": "CANCELAMENTO DE AVISO PREVIO EMPRESA",
})


# manter compatibilidade com código antigo que importava DEFAULTS_FIXOS
DEFAULTS_FIXOS = DEFAULTS_FALTAS


def mes_seguinte(mes: int, ano: int):
    """Mês seguinte ao de referência (falta do mês X é lançada no mês X+1)."""
    return mes % 12 + 1, ano + (1 if mes == 12 else 0)


def periodo_mes(mes: int, ano: int):
    """(início, fim, qtde_dias) do mês — trata 28/29/30/31 e ano bissexto."""
    dias = calendar.monthrange(ano, mes)[1]
    return date(ano, mes, 1), date(ano, mes, dias), dias


def linha_layout(parceiro_id="", nomefuncionario="", nomeempresa="", nomeescala="",
                 nomeposto="", datainicio: date = None, datafim: date = None,
                 hoje: date = None, fixos: dict = None,
                 preencher_retorno: bool = True) -> dict:
    """Monta uma linha do layout no formato exato (datas aaaa-mm-dd, mes_ano mm/aaaa).

    preencher_retorno=False deixa `dataprevisaoretorno` vazio (usado em
    Rescisões/Avisos: não faz sentido prever retorno de quem está saindo).
    """
    fixos = fixos if fixos is not None else DEFAULTS_FALTAS
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
        row["dataprevisaoretorno"] = (datafim + timedelta(days=1)).strftime("%Y-%m-%d") \
            if preencher_retorno else ""
        row["dataocorrencia"] = datainicio.strftime("%Y-%m-%d")
    row["mes"] = f"{hoje.month:02d}"
    row["ano"] = str(hoje.year)
    row["mes_ano"] = hoje.strftime("%m/%Y")
    row["datacadastro"] = hoje.strftime("%Y-%m-%d")
    for c in VAZIAS:
        row[c] = ""
    return row
