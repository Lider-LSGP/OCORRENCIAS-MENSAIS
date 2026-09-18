"""Módulo FÉRIAS — processamento completo.

Regras (confirmadas com a operação):
  * Fonte: RELAÇÃO DE FÉRIAS da Domínio -> colunas `inicio_gozo` / `fim_gozo`.
  * Tipo de ocorrência depende do colaborador:
      - 34 FERIAS MOTORISTA  -> função contém "MOTORISTA" na base do sistema
      - 32 FERIAS VSP        -> empresa = VSP
      - 31 FERIAS GERAL      -> demais
  * descontabeneficio = "Sim", EXCETO Férias VSP = "Não".
  * tipooccontrolaferias = "Sim"; tipoferiasoc = "Sim". Resto = padrão do layout.
  * qtddiasafastamento = (fim_gozo - inicio_gozo) + 1.
  * Conferência inteligente:
      - Já cadastrado com as datas corretas -> CADASTRADO.
      - Já cadastrado mas com data(s) erradas -> CADASTRADO (ERRO DATA INICIO:
        dd/mm/aaaa OU/E ERRO DATA FINAL: dd/mm/aaaa).
      - Cadastrado com o tipoocorrencia_id errado (ex.: deveria ser VSP/Motorista
        mas está como Geral) -> aviso na OBS.
      - Pode haver férias cadastrada e depois cancelada/substituída perto do
        período (mês anterior/seguinte) — busca-se o lançamento de férias mais
        próximo do período esperado antes de decidir se falta cadastrar.
      - Sem nada lançado -> NÃO CADASTRADO (vai para a importação).
"""
from __future__ import annotations

from datetime import date
from typing import Optional

import pandas as pd

from .colaboradores import BaseColaboradores, norm_nome
from .common import clean, to_date
from .empresas import nome_empresa
from .layout import linha_layout, COLUNAS_LAYOUT, BASE_DEFAULTS, NOMES_TIPO

TOLERANCIA_PROXIMIDADE_DIAS = 40  # p/ considerar um lançamento "do mesmo período"


def _tipo_ferias(info: dict) -> tuple[str, str]:
    """Decide (tipoocorrencia_id, nomeocorrencia) pela função/empresa do colaborador."""
    funcao = (info.get("funcao") or "").upper()
    empresa = (info.get("empresa") or "").upper()
    if "MOTORISTA" in funcao:
        return "34", "FERIAS MOTORISTA"
    if "VSP" in empresa:
        return "32", "FERIAS VSP"
    return "31", "FERIAS GERAL"


def _fixos_ferias(tipo_id: str, nome_oc: str) -> dict:
    return {
        **BASE_DEFAULTS,
        "tipoocorrencia_id": tipo_id,
        "nomeocorrencia": nome_oc,
        "descontabeneficio": "Não" if tipo_id == "32" else "Sim",
        "tipooccontrolaferias": "Sim",
        "tipoferiasoc": "Sim",
    }


def processar_ferias(df: pd.DataFrame, base: BaseColaboradores,
                     lancadas=None, hoje: date = None, **kw):
    hoje = hoje or date.today()
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    imp_rows, res_rows = [], []
    for _, r in df.iterrows():
        nome_dom = clean(r.get("nome"))
        ini_esp = to_date(r.get("inicio_gozo"))
        fim_esp = to_date(r.get("fim_gozo"))
        if not nome_dom or ini_esp is None or fim_esp is None:
            continue
        codi = clean(r.get("codi_emp"))
        matricula = clean(r.get("i_empregados"))

        c = base.buscar(nome=nome_dom, matricula=matricula, codi_emp=codi)
        info = BaseColaboradores.info(c) if c is not None else {}
        obs = []
        if c is None:
            obs.append("NÃO ENCONTRADO NA BASE DO SISTEMA")
        else:
            if info.get("ativo") and info["ativo"] != "Sim":
                obs.append(f"NÃO ATIVO NO SISTEMA (Ativo: {info['ativo']})")
            if info.get("demissao"):
                obs.append(f"DEMISSÃO no sistema: {info['demissao']}")

        nome_sistema = info.get("nome") or nome_dom
        empresa = info.get("empresa") or nome_empresa(codi)
        tipo_id, nome_oc = _tipo_ferias(info) if info else ("31", "FERIAS GERAL")
        qtd_dias = (fim_esp - ini_esp).days + 1

        # ---- conferência contra o que já está lançado ----
        status = "NÃO CADASTRADO"
        if lancadas is not None and info:
            recs = lancadas.periodos_colaborador(info.get("parceiro_id", ""), nome_sistema,
                                                 tipos=("31", "32", "34"))
            melhor, melhor_dist = None, None
            for rec in recs:
                dist = abs((rec["ini"] - ini_esp).days) + abs((rec["fim"] - fim_esp).days)
                if melhor_dist is None or dist < melhor_dist:
                    melhor, melhor_dist = rec, dist
            if melhor is not None and melhor_dist is not None and \
                    (melhor["ini"] == ini_esp or melhor["fim"] == fim_esp or
                     melhor_dist <= TOLERANCIA_PROXIMIDADE_DIAS):
                erros = []
                if melhor["ini"] != ini_esp:
                    erros.append(f"ERRO DATA INICIO:{melhor['ini'].strftime('%d/%m/%Y')}")
                if melhor["fim"] != fim_esp:
                    erros.append(f"ERRO DATA FINAL:{melhor['fim'].strftime('%d/%m/%Y')}")
                if erros:
                    status = f"CADASTRADO (" + " OU/E ".join(erros) + ")"
                else:
                    status = "CADASTRADO"
                if melhor["tipo"] != tipo_id:
                    nome_lancado = NOMES_TIPO.get(melhor["tipo"], melhor["tipo"])
                    obs.append(f"TIPO LANÇADO DIVERGENTE: sistema tem '{nome_lancado}' "
                               f"(esperado '{nome_oc}')")

        if status == "NÃO CADASTRADO":
            imp_rows.append(linha_layout(
                parceiro_id=info.get("parceiro_id", ""),
                nomefuncionario=nome_sistema,
                nomeempresa=empresa,
                nomeescala=info.get("escala", ""),
                nomeposto=info.get("posto", ""),
                datainicio=ini_esp, datafim=fim_esp, hoje=hoje,
                fixos=_fixos_ferias(tipo_id, nome_oc),
            ))

        res_rows.append({
            "STATUS": status,
            "NOME (DOMÍNIO)": nome_dom,
            "EMPRESA": empresa,
            "CODI_EMP": codi,
            "ID SISTEMA (parceiro_id)": info.get("parceiro_id", ""),
            "NOME NO SISTEMA": info.get("nome", ""),
            "TIPO DE FÉRIAS": nome_oc,
            "INÍCIO ESPERADO": ini_esp.strftime("%d/%m/%Y"),
            "FIM ESPERADO": fim_esp.strftime("%d/%m/%Y"),
            "QTD DIAS": qtd_dias,
            "OBS": " | ".join(obs),
        })

    df_imp = pd.DataFrame(imp_rows, columns=COLUNAS_LAYOUT)
    df_res = pd.DataFrame(res_rows)
    return df_imp, df_res
