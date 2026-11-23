"""Módulo AFASTAMENTOS — processamento completo.

Regras (confirmadas com a operação):
  * Fonte: RELAÇÃO DE AFASTAMENTOS da Domínio -> colunas `data_real` (início),
    `data_fim`, `descricao` (motivo), `i_afastamentos` (dias do trecho).
  * Só interessam afastamentos SUPERIORES A 15 DIAS ou LICENÇA MATERNIDADE
    (paternidade também, raro). Atestados de até 15 dias (ex.: "Doença
    período igual ou inferior a 15 dias") são ignorados aqui — não geram
    ocorrência de afastamento.
  * Um mesmo afastamento pode aparecer em VÁRIOS registros na Domínio porque
    foi complementado (ex.: 15 dias de atestado + depois mais um trecho a
    partir do dia 16 até o fim real). Por isso os trechos do MESMO
    colaborador com motivo "superior a 15 dias" / "novo afast. mesma doença"
    / licença maternidade/paternidade são UNIDOS em uma única faixa contínua
    antes de comparar com o sistema.
  * tipoocorrencia_id:
      - 18 AFASTAMENTO INSS     -> afastamento > 15 dias (doença)
      - 35 LICENCA MATERNIDADE  -> licença maternidade
      - 41 LICENCA PATERNA      -> licença paternidade (raro)
    Ambos: descontabeneficio=Sim, geramapabeneficio=Sim,
    tipooccontrolaferias=Não, tipoferiasoc=Não.
  * Data de início/fim: SEMPRE a data_real / data_fim exatas informadas pela
    Domínio (nenhum ajuste automático de +15 dias é aplicado — confirmado
    com a operação: quem decide se cabe separar atestado+INSS é quem confere
    manualmente no sistema, não o app).
  * Conferência:
      - Já lançado com o MESMO início e fim (ou a união dos trechos do
        sistema cobre o mesmo período) -> CADASTRADO.
      - Já lançado mas com datas diferentes -> CADASTRADO (ERRO DATA INICIO:
        dd/mm/aaaa OU/E ERRO DATA FINAL: dd/mm/aaaa) — apenas um aviso p/
        conferência manual, a data da Domínio NUNCA é alterada pelo app.
      - Nada lançado -> NÃO CADASTRADO (vai para a importação, com a data
        exata da Domínio).
      - Quando o motivo for "Doença período superior a 15 dias" (início de
        doença nova, sem atestado prévio no mesmo relatório) a OBS recebe um
        aviso sugerindo conferir se não falta lançar os 15 dias de atestado
        antes — mas isso NUNCA muda a data usada na importação.
"""
from __future__ import annotations

from datetime import date, timedelta
from typing import Optional

import pandas as pd

from .colaboradores import BaseColaboradores
from .common import clean, to_date, merge_intervalos, melhor_faixa
from .empresas import nome_empresa
from .layout import linha_layout, COLUNAS_LAYOUT, BASE_DEFAULTS, NOMES_TIPO

DIAS_MIN_AFASTAMENTO = 16  # "superior a 15 dias"

FIXOS_AFASTAMENTO = {
    **BASE_DEFAULTS,
    "descontabeneficio": "Sim",
    "geramapabeneficio": "Sim",
    "tipooccontrolaferias": "Não",
    "tipoferiasoc": "Não",
}


def _classificar_motivo(desc: str) -> Optional[str]:
    """Retorna tipoocorrencia_id relevante ou None (ignorar o trecho)."""
    d = (desc or "").upper()
    if "MATERNIDADE" in d:
        return "35"
    if "PATERN" in d:
        return "41"
    if "SUPERIOR A 15" in d or "SUPERIOR A15" in d:
        return "18"
    if "NOVO AFAST" in d:  # continuação de afastamento — herda tipo INSS
        return "18"
    if "IGUAL OU INFERIOR A 15" in d or "INFERIOR A 15" in d:
        return None  # atestado curto — não é ocorrência de afastamento
    return None


def processar_afastamentos(df: pd.DataFrame, base: BaseColaboradores,
                           lancadas=None, hoje: date = None, **kw):
    hoje = hoje or date.today()
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # ---- agrupa trechos relevantes por colaborador (nome+empresa) ----
    # também guarda os atestados CURTOS (<=15 dias, normalmente ignorados)
    # apenas para checar se um afastamento "superior a 15 dias" tem um
    # atestado imediatamente anterior no MESMO relatório (não altera datas,
    # só decide se cabe um aviso de conferência na OBS).
    grupos: dict = {}
    atestados_curtos: dict = {}
    for _, r in df.iterrows():
        desc = clean(r.get("descricao"))
        ini = to_date(r.get("data_real"))
        fim = to_date(r.get("data_fim"))
        if ini is None or fim is None:
            continue
        if fim < ini:
            ini, fim = fim, ini
        nome_dom = clean(r.get("nome"))
        codi = clean(r.get("codi_emp") or r.get("codi_emp.1"))
        chave = (nome_dom.upper(), codi)

        tipo_id = _classificar_motivo(desc)
        if tipo_id is None:
            if "INFERIOR A 15" in desc.upper() or "IGUAL OU INFERIOR" in desc.upper():
                atestados_curtos.setdefault(chave, []).append((ini, fim))
            continue

        matricula = clean(r.get("i_empregados"))
        grupos.setdefault(chave, {"nome": nome_dom, "codi": codi, "matricula": matricula,
                                  "trechos": [], "tipos": set(), "descricoes": set()})
        grupos[chave]["trechos"].append((ini, fim))
        grupos[chave]["tipos"].add(tipo_id)
        grupos[chave]["descricoes"].add(desc)

    imp_rows, res_rows = [], []
    for chave, g in grupos.items():
        merged = merge_intervalos(g["trechos"], gap_dias=1)
        if not merged:
            continue
        ini_esp, fim_esp = min(m[0] for m in merged), max(m[1] for m in merged)
        dias_totais = (fim_esp - ini_esp).days + 1
        if dias_totais < DIAS_MIN_AFASTAMENTO and "35" not in g["tipos"] and "41" not in g["tipos"]:
            continue  # menos de 16 dias e não é licença

        tipo_id = "35" if "35" in g["tipos"] else ("41" if "41" in g["tipos"] else "18")
        nome_oc = NOMES_TIPO.get(tipo_id, "")

        c = base.buscar(nome=g["nome"], matricula=g["matricula"], codi_emp=g["codi"])
        info = BaseColaboradores.info(c) if c is not None else {}
        obs = []
        if c is None:
            obs.append("NÃO ENCONTRADO NA BASE DO SISTEMA")
        else:
            if info.get("ativo") and info["ativo"] != "Sim":
                obs.append(f"NÃO ATIVO NO SISTEMA (Ativo: {info['ativo']})")
            if info.get("demissao"):
                obs.append(f"DEMISSÃO no sistema: {info['demissao']}")

        # aviso (não corretivo): "superior a 15 dias" sem atestado curto
        # imediatamente antes no mesmo relatório -> pode faltar lançar os
        # primeiros 15 dias como atestado separado (só sinaliza, não mexe na data)
        descricoes_up = {d.upper() for d in g["descricoes"]}
        if any("SUPERIOR A 15" in d or "SUPERIOR A15" in d for d in descricoes_up) \
                and not any("NOVO AFAST" in d for d in descricoes_up):
            curtos = atestados_curtos.get(chave, [])
            tem_atestado_antes = any(cf >= ini_esp - timedelta(days=1) and ci <= ini_esp
                                     for ci, cf in curtos)
            if not tem_atestado_antes:
                obs.append("CONFERIR: doença > 15 dias sem atestado dos 15 dias iniciais "
                          "no mesmo relatório (pode já estar lançado em outro mês)")

        nome_sistema = info.get("nome") or g["nome"]
        empresa = info.get("empresa") or nome_empresa(g["codi"])

        # ---- conferência: junta também os lançamentos já existentes no sistema ----
        status = "NÃO CADASTRADO"
        if lancadas is not None and info:
            recs = lancadas.periodos_colaborador(info.get("parceiro_id", ""), nome_sistema,
                                                 tipos=("18", "35", "41"))
            trechos_sistema = [(rc["ini"], rc["fim"]) for rc in recs]
            merged_sist = merge_intervalos(trechos_sistema, gap_dias=1)
            faixa = melhor_faixa(merged_sist, ini_esp, fim_esp, tolerancia_dias=25)
            if faixa is not None:
                ini_sist, fim_sist = faixa
                erros = []
                if ini_sist != ini_esp:
                    erros.append(f"ERRO DATA INICIO:{ini_sist.strftime('%d/%m/%Y')}")
                if fim_sist != fim_esp:
                    erros.append(f"ERRO DATA FINAL:{fim_sist.strftime('%d/%m/%Y')}")
                status = ("CADASTRADO (" + " OU/E ".join(erros) + ")") if erros else "CADASTRADO"
                tipos_lancados = {rc["tipo"] for rc in recs
                                  if rc["ini"] <= fim_esp and rc["fim"] >= ini_esp}
                if tipos_lancados and tipo_id not in tipos_lancados:
                    nomes_lanc = ", ".join(NOMES_TIPO.get(t, t) for t in tipos_lancados)
                    obs.append(f"TIPO LANÇADO DIVERGENTE: sistema tem '{nomes_lanc}' "
                               f"(esperado '{nome_oc}')")

        if status == "NÃO CADASTRADO":
            imp_rows.append(linha_layout(
                parceiro_id=info.get("parceiro_id", ""),
                nomefuncionario=nome_sistema,
                nomeempresa=empresa,
                nomeescala=info.get("escala", ""),
                nomeposto=info.get("posto", ""),
                datainicio=ini_esp, datafim=fim_esp, hoje=hoje,
                fixos={**FIXOS_AFASTAMENTO, "tipoocorrencia_id": tipo_id, "nomeocorrencia": nome_oc},
            ))

        res_rows.append({
            "STATUS": status,
            "NOME (DOMÍNIO)": g["nome"],
            "EMPRESA": empresa,
            "CODI_EMP": g["codi"],
            "ID SISTEMA (parceiro_id)": info.get("parceiro_id", ""),
            "NOME NO SISTEMA": info.get("nome", ""),
            "TIPO": nome_oc,
            "MOTIVO(S) DOMÍNIO": " / ".join(sorted(g["descricoes"])),
            "INÍCIO (UNIFICADO)": ini_esp.strftime("%d/%m/%Y"),
            "FIM (UNIFICADO)": fim_esp.strftime("%d/%m/%Y"),
            "QTD DIAS": dias_totais,
            "TRECHOS NA DOMÍNIO": len(g["trechos"]),
            "OBS": " | ".join(obs),
        })

    df_imp = pd.DataFrame(imp_rows, columns=COLUNAS_LAYOUT)
    df_res = pd.DataFrame(res_rows)
    return df_imp, df_res
