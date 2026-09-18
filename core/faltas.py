"""Módulo FALTAS MÊS TODO — processamento completo.

Regras (confirmadas com o operação):
  * Filtro: valor_inf >= LIMIAR (padrão 20) dias de falta.
  * Referência: coluna `data` do relatório (dd/mm/aaaa) -> mês de referência.
    A ocorrência é lançada no MÊS SEGUINTE: datainicio = 1º dia, datafim =
    último dia (28/29/30/31 — bissexto tratado por calendar.monthrange),
    qtddiasafastamento = nº de dias do mês, retorno = datafim + 1.
  * tipoocorrencia_id = 13 (FALTA 30 DIAS) e demais fixos do layout.
  * Casamento com a base: matrícula (matricula_esocial / i_empregados) ->
    nome+empresa -> nome. OBS marca divergências, inativos, PCD etc.
  * Conferência: se já existe QUALQUER ocorrência no sistema com a mesma
    data inicial para o colaborador -> JÁ LANÇADO (não sai na importação).
"""
from __future__ import annotations

import re
from datetime import date, datetime
from typing import Optional, Tuple

import pandas as pd

from .colaboradores import BaseColaboradores, norm_nome, clean
from .empresas import nome_empresa, EMPRESAS
from .layout import linha_layout, COLUNAS_LAYOUT, mes_seguinte, periodo_mes

LIMIAR_DIAS_PADRAO = 20


def detectar_tipo_relatorio(df: pd.DataFrame) -> Optional[str]:
    cols = {str(c).strip().lower() for c in df.columns}
    if "foempregados_nome" in cols or ("valor_inf" in cols and "i_eventos" in cols):
        return "FALTAS"
    if "inicio_gozo" in cols:
        return "FERIAS"
    if "data_real" in cols and "i_afastamentos" in cols:
        return "AFASTAMENTOS"
    if "data_aviso" in cols and "salario" not in cols:
        return "AVISOS"
    if "demissao" in cols or "salario" in cols:
        return "RESCISOES"
    return None


def _to_float(v) -> Optional[float]:
    if v is None or (isinstance(v, float) and v != v):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(".", "").replace(",", ".") if "," in str(v) else str(v).strip()
    try:
        return float(s)
    except ValueError:
        return None


def _mes_ref(v) -> Optional[Tuple[int, int]]:
    """Extrai (mes, ano) da coluna `data` do relatório de faltas."""
    if v is None or (isinstance(v, float) and v != v):
        return None
    if isinstance(v, datetime):
        return v.month, v.year
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%Y"):
        try:
            d = datetime.strptime(s[:10], fmt)
            return d.month, d.year
        except ValueError:
            continue
    return None


def processar_faltas(df: pd.DataFrame, base: BaseColaboradores,
                     lancadas=None, limiar: float = LIMIAR_DIAS_PADRAO,
                     hoje: date = None):
    """Processa o relatório Domínio de faltas.

    Retorna (df_importacao, df_resultado).
    """
    hoje = hoje or date.today()
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    col_nome = "foempregados_nome" if "foempregados_nome" in df.columns else "nome"

    imp_rows, res_rows = [], []
    for _, r in df.iterrows():
        dias = _to_float(r.get("valor_inf"))
        if dias is None or dias < limiar:
            continue
        ref = _mes_ref(r.get("data"))
        nome_dom = clean(r.get(col_nome))
        if not nome_dom or ref is None:
            continue
        codi = clean(r.get("codi_emp"))
        matricula = clean(r.get("matricula_esocial")) or clean(r.get("i_empregados"))

        mes_l, ano_l = mes_seguinte(*ref)
        ini, fim, ndias = periodo_mes(mes_l, ano_l)
        ref_txt = f"{ref[0]:02d}/{ref[1]}"
        periodo_txt = f"{ini.strftime('%d/%m/%Y')} a {fim.strftime('%d/%m/%Y')}"

        # ---- casamento com a base ----
        c = base.buscar(nome=nome_dom, matricula=matricula, codi_emp=codi)
        info = BaseColaboradores.info(c) if c is not None else {}
        obs, matched_por = [], ""
        if c is None:
            obs.append("NÃO ENCONTRADO NA BASE DO SISTEMA")
            matched_por = "—"
        else:
            mat_ok = matricula and re.sub(r"\D", "", matricula) == re.sub(r"\D", "", info.get("matricula", ""))
            matched_por = "matrícula" if mat_ok else "nome"
            if mat_ok and norm_nome(info.get("nome")) != norm_nome(nome_dom):
                obs.append(f"NOME DIVERGENTE (sistema: {info.get('nome')})")
            if info.get("ativo") and info["ativo"] != "Sim":
                obs.append(f"NÃO ATIVO NO SISTEMA (Ativo: {info['ativo']})")
            if info.get("demissao"):
                obs.append(f"DEMISSÃO no sistema: {info['demissao']}")
            if info.get("pcd"):
                obs.append(f"PCD ({info['pcd']})")

        nome_sistema = info.get("nome") or nome_dom
        empresa = info.get("empresa") or nome_empresa(codi)

        # ---- conferência de já lançados ----
        status = "VÁLIDO"
        if lancadas is not None:
            if lancadas.ja_lancado(ini, info.get("parceiro_id", ""), nome_sistema):
                tipos = lancadas.tipos_do_colaborador(ini, info.get("parceiro_id", ""), nome_sistema)
                obs.append(f"Já existe ocorrência tipo(s) {', '.join(tipos)} com início em {ini.strftime('%d/%m/%Y')}")
                status = "JÁ LANÇADO"
        bloqueios = [o for o in obs if o.startswith(("NÃO ENCONTRADO", "NOME DIVERGENTE", "NÃO ATIVO"))]
        if status != "JÁ LANÇADO" and bloqueios:
            status = "EM ABERTO"

        # ---- linha de importação (só o que não está lançado) ----
        if status != "JÁ LANÇADO":
            imp_rows.append(linha_layout(
                parceiro_id=info.get("parceiro_id", ""),
                nomefuncionario=nome_sistema,
                nomeempresa=empresa,
                nomeescala=info.get("escala", ""),
                nomeposto=info.get("posto", ""),
                datainicio=ini, datafim=fim, hoje=hoje,
            ))

        res_rows.append({
            "STATUS": status,
            "NOME (DOMÍNIO)": nome_dom,
            "EMPRESA": empresa,
            "CODI_EMP": codi,
            "MATRÍCULA (DOMÍNIO)": matricula,
            "ID SISTEMA (parceiro_id)": info.get("parceiro_id", ""),
            "NOME NO SISTEMA": info.get("nome", ""),
            "DIAS DE FALTA": int(dias),
            "REFERÊNCIA": ref_txt,
            "PERÍODO LANÇADO": periodo_txt,
            "CASADO POR": matched_por,
            "OBS": " | ".join(obs),
        })

    df_imp = pd.DataFrame(imp_rows, columns=COLUNAS_LAYOUT)
    df_res = pd.DataFrame(res_rows)
    return df_imp, df_res


# ---------------------------------------------------------------------------
# FÉRIAS e AFASTAMENTOS moram em módulos próprios (core/ferias.py e
# core/afastamentos.py) — reexportados aqui só por compatibilidade com quem
# ainda importa `from core.faltas import processar_ferias` etc.
# ---------------------------------------------------------------------------
from .ferias import processar_ferias  # noqa: E402,F401
from .afastamentos import processar_afastamentos  # noqa: E402,F401
from .rescisoes import processar_rescisoes_avisos  # noqa: E402,F401
