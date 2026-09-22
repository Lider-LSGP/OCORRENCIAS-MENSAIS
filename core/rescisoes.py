"""Módulo RESCISÕES/AVISOS — processamento completo.

Fonte: DUAS planilhas da Domínio que SE COMPLEMENTAM (uma tem o que falta na
outra) e precisam ser unidas por colaborador antes de decidir a ocorrência:
  * RELAÇÃO DE AVISO      -> codi_emp, i_empregados, nome, data_aviso,
                              data_demissao, concessor (1=Empresa, 2=Colaborador)
  * RELAÇÃO DE RESCISÃO   -> codi_emp, i_empregados, nome, demissao,
                              data_aviso, motivo (código Domínio),
                              aviso_indenizado (S/N)

Regras confirmadas cruzando com o que já está cadastrado no sistema (dados
reais de 17/09/2026):

  1) Quem concede o aviso:
       concessor==1 -> EMPRESA   |   concessor==2 -> COLABORADOR
     Quando falta o `concessor` (só apareceu na Rescisão), infere-se pelo
     `motivo` (ver MOTIVO_EASYAPP / grupos abaixo).

  2) Aviso cumprido ou não:
       dias = (data_fim - data_aviso).days
       dias == 0  -> NÃO CUMPRIU (indenizado / sem trabalhar o aviso)
       dias  > 0  -> CUMPRIU (trabalhou o período de aviso)
     Cruzado com a coluna `aviso_indenizado` (S=indenizado -> não cumpriu)
     quando disponível; se divergir da regra de datas, prevalece a REGRA DE
     DATAS (mais confiável nos testes) e a divergência vai para a OBS.

  3) Tipo de ocorrência (id) e cálculo de datas:
       Términos NATURAIS de contrato (motivo 12 e 22 — fim de prazo
       determinado / fim natural de experiência, sem antecipação por
       ninguém) -> 39 ENCERRAMENTO DE CONTRATO, 1 dia na data da demissão.

       EMPRESA + CUMPRIU    -> 38 AVISO PREVIO EMPRESA
                                datainicio = data_aviso
                                datafim    = data_fim - 7 dias  (regra do
                                             desconto de aviso cumprido)
       EMPRESA + NÃO CUMPRIU -> 40 AVISO PREVIO EMPRESA 1 DIA
                                datainicio = datafim = data_fim (1 dia)

       COLABORADOR + CUMPRIU     -> 33 AVISO PREVIO COLABORADOR
                                     datainicio = data_aviso
                                     datafim    = data_fim  (SEM desconto)
       COLABORADOR + NÃO CUMPRIU -> 36 PEDIDO DE DEMISSAO
                                     datainicio = datafim = data_fim (1 dia)

     Rescisão de CONTRATO DE EXPERIÊNCIA ANTECIPADA (motivo 10/11) segue a
     mesma regra de concessor+cumprimento acima, mas como os dados reais
     mostraram casos divergentes (às vezes 39, às vezes 36/38), a OBS SEMPRE
     recebe um aviso pedindo conferência manual nesse caso específico.

     Transferência (5/6/27) e Morte (8/13/14/40/41/42) NÃO geram ocorrência
     automática — aparecem no resultado só com aviso para tratamento manual.

  4) A DATA DE DEMISSÃO real (a que vai na ficha do colaborador) é SEMPRE a
     data cheia informada pela Domínio — o desconto de 7 dias afeta só o
     `datafim` da OCORRÊNCIA (tipo 38), nunca a data de demissão em si.

  5) `dataprevisaoretorno` fica sempre vazio (não faz sentido para quem está
     saindo da empresa).

  6) Demissões FUTURAS (data_fim > hoje): entram numa lista separada de
     alerta ("DEMISSÃO FUTURA — não demitir ainda"), mas o controle de
     aviso já cadastrado/faltante continua normal.

  7) Fixos (conforme solicitado): descontabeneficio=Sim, geramapabeneficio=
     Sim, tipooccontrolaferias=Não, tipoferiasoc=Não.
"""
from __future__ import annotations

from datetime import date, timedelta
from typing import Optional

import pandas as pd

from .colaboradores import BaseColaboradores, norm_nome
from .common import clean, to_date, lancamento_residual_proximo
from .empresas import nome_empresa
from .layout import linha_layout, COLUNAS_LAYOUT, BASE_DEFAULTS, NOMES_TIPO

DESCONTO_AVISO_CUMPRIDO_EMPRESA = 7
TOLERANCIA_PROXIMIDADE_DIAS = 20

FIXOS_RESCISAO = {
    **BASE_DEFAULTS,
    "descontabeneficio": "Sim",
    "geramapabeneficio": "Sim",
    "tipooccontrolaferias": "Não",
    "tipoferiasoc": "Não",
}

# ---------------------------------------------------------------------------
# Motivo Domínio -> categoria "situação de desligamento" do EasyApp (só
# informativo, aparece no relatório de resultado para anexar na ficha).
# ---------------------------------------------------------------------------
MOTIVO_EASYAPP = {
    "2": "Dispensa S/ Justa Causa", "3": "Dispensa S/ Justa Causa",
    "28": "Dispensa S/ Justa Causa", "29": "Dispensa S/ Justa Causa",
    "30": "Dispensa S/ Justa Causa", "44": "Dispensa S/ Justa Causa",
    "46": "Dispensa S/ Justa Causa", "47": "Dispensa S/ Justa Causa",
    "1": "Dispensa C/ Justa Causa",
    "4": "Dispensa a Pedido (Espontâneo)",
    "22": "Fim do Contrato por Prazo Determinado",
    "23": "Fim do Contrato por Prazo Determinado",
    "24": "Fim do Contrato por Prazo Determinado",
    "10": "Termino de Contrato", "11": "Termino de Contrato", "12": "Termino de Contrato",
    "8": "Morte", "13": "Morte", "14": "Morte", "40": "Morte", "41": "Morte", "42": "Morte",
    "6": "Transferência", "5": "Transferência", "27": "Transferência",
}

GRUPO_TERMINO_NATURAL = {"12", "22"}          # fim de prazo/experiência SEM antecipação
GRUPO_EMPRESA_MOTIVO = {"1", "2", "3", "28", "29", "30", "44", "46", "47", "23", "10"}
GRUPO_COLABORADOR_MOTIVO = {"4", "24", "11"}
GRUPO_ESPECIAL_SEM_OCORRENCIA = {"5", "6", "27", "8", "13", "14", "40", "41", "42"}
GRUPO_EXPERIENCIA_ANTECIPADA = {"10", "11"}


def _tipo_e_datas(concedido_por: str, cumpriu: bool, motivo: str,
                  data_aviso: Optional[date], data_fim: date):
    """Decide (tipo_id, nome_oc, datainicio, datafim_ocorrencia) da ocorrência."""
    if motivo in GRUPO_TERMINO_NATURAL:
        return "39", NOMES_TIPO["39"], data_fim, data_fim

    if concedido_por == "EMPRESA":
        if cumpriu and data_aviso:
            return "38", NOMES_TIPO["38"], data_aviso, data_fim - timedelta(days=DESCONTO_AVISO_CUMPRIDO_EMPRESA)
        return "40", NOMES_TIPO["40"], data_fim, data_fim

    # COLABORADOR
    if cumpriu and data_aviso:
        return "33", NOMES_TIPO["33"], data_aviso, data_fim
    return "36", NOMES_TIPO["36"], data_fim, data_fim


def _concedido_por(concessor: str, motivo: str) -> str:
    if concessor == "1":
        return "EMPRESA"
    if concessor == "2":
        return "COLABORADOR"
    if motivo in GRUPO_COLABORADOR_MOTIVO:
        return "COLABORADOR"
    return "EMPRESA"  # padrão mais comum quando não há info


def _unir_aviso_rescisao(df_aviso: Optional[pd.DataFrame], df_rescisao: Optional[pd.DataFrame]) -> list:
    """Une as duas planilhas por (nome normalizado, codi_emp) — RESCISÃO tem
    prioridade nos campos (motivo/aviso_indenizado/demissao), AVISO cobre o
    que falta (colaboradores só com aviso lançado, ainda sem cálculo)."""
    registros: dict = {}

    if df_rescisao is not None:
        for _, r in df_rescisao.iterrows():
            nome = clean(r.get("nome"))
            codi = clean(r.get("codi_emp"))
            chave = (norm_nome(nome), codi)
            registros[chave] = {
                "nome": nome, "codi": codi,
                "matricula": clean(r.get("i_empregados")),
                "data_aviso": to_date(r.get("data_aviso")),
                "data_fim": to_date(r.get("demissao")),
                "motivo": re_digits(r.get("motivo")),
                "aviso_indenizado": clean(r.get("aviso_indenizado")).upper(),
                "concessor": "",
                "fonte": {"RESCISAO"},
            }

    if df_aviso is not None:
        for _, r in df_aviso.iterrows():
            nome = clean(r.get("nome"))
            codi = clean(r.get("codi_emp"))
            chave = (norm_nome(nome), codi)
            concessor_raw = clean(r.get("concessor"))
            existente = registros.get(chave)
            if existente:
                existente["fonte"].add("AVISO")
                if not existente.get("data_aviso"):
                    existente["data_aviso"] = to_date(r.get("data_aviso"))
                if not existente.get("data_fim"):
                    existente["data_fim"] = to_date(r.get("data_demissao"))
                existente["concessor"] = concessor_raw
            else:
                registros[chave] = {
                    "nome": nome, "codi": codi,
                    "matricula": clean(r.get("i_empregados")),
                    "data_aviso": to_date(r.get("data_aviso")),
                    "data_fim": to_date(r.get("data_demissao")),
                    "motivo": "",
                    "aviso_indenizado": "",
                    "concessor": concessor_raw,
                    "fonte": {"AVISO"},
                }
    return list(registros.values())


def re_digits(v) -> str:
    import re
    return re.sub(r"\D", "", clean(v))


def processar_rescisoes_avisos(df_aviso: Optional[pd.DataFrame] = None,
                               df_rescisao: Optional[pd.DataFrame] = None,
                               base: BaseColaboradores = None, lancadas=None,
                               hoje: date = None, **kw):
    """Processa RELAÇÃO DE AVISO + RELAÇÃO DE RESCISÃO já unidas.

    Retorna (df_imp, df_res, df_futuras):
      * df_imp     -> pronto para o EasyApp (layout completo)
      * df_res     -> conferência (STATUS/OBS/MOTIVO)
      * df_futuras -> apenas quem tem demissão FUTURA (data > hoje), para
                      lembrete — NÃO significa que já deva ser demitido.
    """
    hoje = hoje or date.today()
    regs = _unir_aviso_rescisao(df_aviso, df_rescisao)

    imp_rows, res_rows, futuras_rows = [], [], []
    for g in regs:
        nome_dom, codi = g["nome"], g["codi"]
        data_fim = g["data_fim"]
        if not nome_dom or data_fim is None:
            continue
        data_aviso = g["data_aviso"] or data_fim
        motivo = g["motivo"]
        obs = []

        # ---- quem concede + cumpriu (com cruzamento aviso_indenizado) ----
        concedido_por = _concedido_por(g["concessor"], motivo)
        dias_aviso = (data_fim - data_aviso).days
        cumpriu_por_data = dias_aviso > 0
        cumpriu = cumpriu_por_data
        if g["aviso_indenizado"] in ("S", "N"):
            cumpriu_por_coluna = (g["aviso_indenizado"] == "N")
            if cumpriu_por_coluna != cumpriu_por_data:
                obs.append(f"DIVERGÊNCIA: coluna aviso_indenizado sugere "
                          f"{'CUMPRIU' if cumpriu_por_coluna else 'NÃO CUMPRIU'}, "
                          f"mas as datas sugerem {'CUMPRIU' if cumpriu_por_data else 'NÃO CUMPRIU'} "
                          f"(usada a regra das datas)")

        motivo_categoria = MOTIVO_EASYAPP.get(motivo, "")

        # ---- transferência / morte: fora do fluxo automático ----
        if motivo in GRUPO_ESPECIAL_SEM_OCORRENCIA:
            res_rows.append({
                "STATUS": "CONFERÊNCIA MANUAL",
                "NOME (DOMÍNIO)": nome_dom, "EMPRESA": nome_empresa(codi), "CODI_EMP": codi,
                "TIPO OCORRÊNCIA": "—", "MOTIVO/CATEGORIA": motivo_categoria or "Transferência/Morte",
                "DATA AVISO": data_aviso.strftime("%d/%m/%Y") if data_aviso else "",
                "DATA DEMISSÃO": data_fim.strftime("%d/%m/%Y"),
                "OBS": f"Motivo '{motivo_categoria}' não gera ocorrência automática de aviso — "
                      f"trate manualmente conforme o caso.",
            })
            continue

        if motivo in GRUPO_EXPERIENCIA_ANTECIPADA:
            obs.append("CONFERIR: rescisão de contrato de experiência antecipada — "
                      "confirme se o tipo de ocorrência está correto (casos reais "
                      "já divergiram entre Encerramento de Contrato e Aviso/Pedido)")

        tipo_id, nome_oc, ini_oc, fim_oc = _tipo_e_datas(concedido_por, cumpriu, motivo, data_aviso, data_fim)

        c = base.buscar(nome=nome_dom, matricula=g["matricula"], codi_emp=codi) if base else None
        info = BaseColaboradores.info(c) if c is not None else {}
        if c is None:
            obs.append("NÃO ENCONTRADO NA BASE — pode ser admissão recente; "
                      "cadastrar o colaborador antes de lançar a demissão")
        else:
            if info.get("ativo") and info["ativo"] != "Sim":
                obs.append(f"JÁ NÃO ATIVO NO SISTEMA (Ativo: {info['ativo']})")
            if info.get("demissao"):
                obs.append(f"JÁ TEM DEMISSÃO REGISTRADA NA FICHA: {info['demissao']}")

        nome_sistema = info.get("nome") or nome_dom
        empresa = info.get("empresa") or nome_empresa(codi)
        futura = data_fim > hoje
        if futura:
            obs.append("DEMISSÃO FUTURA — colaborador ainda ATIVO, NÃO demitir agora")

        # ---- conferência contra o que já está lançado ----
        status = "NÃO CADASTRADO"
        if lancadas is not None and info:
            recs = lancadas.periodos_colaborador(info.get("parceiro_id", ""), nome_sistema,
                                                 tipos=("33", "36", "38", "39", "40"))
            # CORREÇÃO: só considera "já cadastrado" se houver um lançamento do
            # MESMO tipo com a MESMA data de início. Se houver do mesmo tipo mas
            # com datas diferentes, marca como CADASTRADO com ERRO. Se houver de
            # outro tipo com data igual, avisa mas mantém a linha para lançar.
            mesmo_tipo = [r for r in recs if r["tipo"] == tipo_id]
            exato = next((r for r in mesmo_tipo
                          if r["ini"] == ini_oc and r["fim"] == fim_oc), None)
            aprox = None
            if not exato and mesmo_tipo:
                aprox = min(mesmo_tipo,
                            key=lambda r: abs((r["ini"] - ini_oc).days)
                            + abs((r["fim"] - fim_oc).days))
                if abs((aprox["ini"] - ini_oc).days) > TOLERANCIA_PROXIMIDADE_DIAS \
                        and abs((aprox["fim"] - fim_oc).days) > TOLERANCIA_PROXIMIDADE_DIAS:
                    aprox = None
            if exato:
                status = "CADASTRADO"
            elif aprox is not None:
                erros = []
                if aprox["ini"] != ini_oc:
                    erros.append(f"ERRO DATA INICIO:{aprox['ini'].strftime('%d/%m/%Y')}")
                if aprox["fim"] != fim_oc:
                    erros.append(f"ERRO DATA FINAL:{aprox['fim'].strftime('%d/%m/%Y')}")
                status = "CADASTRADO (" + " OU/E ".join(erros) + ")" if erros else "CADASTRADO"
            else:
                outro_tipo_msmo_dia = next(
                    (r for r in recs if r["tipo"] != tipo_id and r["ini"] == ini_oc), None)
                if outro_tipo_msmo_dia is not None:
                    nome_lancado = NOMES_TIPO.get(outro_tipo_msmo_dia["tipo"],
                                                  outro_tipo_msmo_dia["tipo"])
                    obs.append(f"Já existe ocorrência tipo(s) {outro_tipo_msmo_dia['tipo']} "
                               f"({nome_lancado}) com início em "
                               f"{outro_tipo_msmo_dia['ini'].strftime('%d/%m/%Y')} — "
                               f"conferir antes de lançar o {nome_oc}")
                elif recs:
                    # nenhum lançamento do mesmo tipo no dia esperado, mas pode
                    # haver um lançamento "residual" de desligamento (a Domínio
                    # cancelou/substituiu o aviso e o antigo continua ATIVO no
                    # sistema interno). Só avisa, não altera o status.
                    residual = lancamento_residual_proximo(recs, ini_oc, fim_oc)
                    if residual is not None:
                        nome_lancado = NOMES_TIPO.get(residual["tipo"], residual["tipo"])
                        obs.append(
                            f"ATENÇÃO: há um lançamento de desligamento ({nome_lancado}) no "
                            f"sistema em {residual['ini'].strftime('%d/%m/%Y')} a "
                            f"{residual['fim'].strftime('%d/%m/%Y')} que pode ter sido "
                            f"CANCELADO/SUBSTITUÍDO na Domínio mas continua ATIVO no sistema — "
                            f"confira se precisa cancelar antes de lançar o novo aviso")
                    else:
                        obs.append("Havia lançamento(s) de desligamento no sistema, mas com "
                                  "data(s) muito diferentes — possível aviso cancelado/"
                                  "substituído; confira manualmente")

        if status == "NÃO CADASTRADO" and c is not None:
            imp_rows.append(linha_layout(
                parceiro_id=info.get("parceiro_id", ""),
                nomefuncionario=nome_sistema, nomeempresa=empresa,
                nomeescala=info.get("escala", ""), nomeposto=info.get("posto", ""),
                datainicio=ini_oc, datafim=fim_oc, hoje=hoje,
                fixos={**FIXOS_RESCISAO, "tipoocorrencia_id": tipo_id, "nomeocorrencia": nome_oc},
                preencher_retorno=False,
            ))
        elif status == "NÃO CADASTRADO" and c is None:
            # colaborador não está na base — ainda assim gera a linha (com o
            # que temos) para facilitar cadastro+lançamento manual, mas fica
            # bem sinalizado no STATUS.
            status = "NÃO CADASTRADO (COLABORADOR NÃO ENCONTRADO)"

        # PARA DESATIVAR = já tem o aviso/rescisão CONFERIDO (verde) e ainda
        # está ATIVO no sistema; a operação faz o desligamento à mão. Demissão
        # futura sai numa aba separada e nunca entra aqui.
        ativo_no_sistema = (info.get("ativo") or "").strip().lower() in ("sim", "s")
        precisa_desativar = (
            str(status).startswith("CADASTRADO") and not futura
            and ativo_no_sistema and not info.get("demissao"))
        status_final = status
        if futura:
            status_final = status + " | DEMISSÃO FUTURA"
        elif precisa_desativar:
            status_final = status + " | PARA DESATIVAR"

        linha_res = {
            "STATUS": status_final,
            "NOME (DOMÍNIO)": nome_dom,
            "EMPRESA": empresa,
            "CODI_EMP": codi,
            "CPF": info.get("cpf", ""),
            "ADMISSÃO": info.get("admissao", ""),
            "ID SISTEMA (parceiro_id)": info.get("parceiro_id", ""),
            "NOME NO SISTEMA": info.get("nome", ""),
            "TIPO OCORRÊNCIA": nome_oc,
            "MOTIVO/CATEGORIA": motivo_categoria,
            "CONCEDIDO POR": concedido_por,
            "AVISO CUMPRIDO?": "Sim" if cumpriu else "Não",
            "DATA AVISO": data_aviso.strftime("%d/%m/%Y") if data_aviso else "",
            "DATA DEMISSÃO (FICHA)": data_fim.strftime("%d/%m/%Y"),
            "DATAFIM OCORRÊNCIA": fim_oc.strftime("%d/%m/%Y"),
            "PARA DESATIVAR": "SIM" if precisa_desativar else "",
            "OBS": " | ".join(obs),
        }
        res_rows.append(linha_res)
        if futura:
            futuras_rows.append(linha_res)

    df_imp = pd.DataFrame(imp_rows, columns=COLUNAS_LAYOUT)
    df_res = pd.DataFrame(res_rows)
    df_futuras = pd.DataFrame(futuras_rows)
    return df_imp, df_res, df_futuras
