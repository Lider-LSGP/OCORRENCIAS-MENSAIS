"""Ocorrências 2.0 — conferência e geração de importações EasyApp.

O app CLASSIFICA CADA ARQUIVO PELO CONTEÚDO: não importa em qual campo o
arquivo foi enviado, ele vai para o papel certo (base / lançadas / Domínio).
O layout de importação enviado por engano vira apenas referência e gera aviso.

Módulos: FALTAS MÊS TODO, FÉRIAS, AFASTAMENTOS e RESCISÕES/AVISOS —
todos completos e testados com dados reais.

Rodar:
    pip install -r requirements.txt
    streamlit run app.py
"""
from __future__ import annotations

import tempfile
from datetime import date, datetime

import pandas as pd
import plotly.express as px
import streamlit as st

from core.leitura import ler_arquivo_upload
from core.classificador import classificar
from core.colaboradores import BaseColaboradores
from core.ocorrencias import OcorrenciasLancadas
from core.faltas import (detectar_tipo_relatorio, processar_faltas,
                         processar_ferias, processar_afastamentos)
from core.rescisoes import processar_rescisoes_avisos
from core.saida import montar_arquivos, zip_bytes, resultado_xlsx_bytes, _nome_bonitinho
from core.empresas import EMPRESAS, APELIDOS, CORES_EMPRESA, apelido_empresa

st.set_page_config(page_title="Ocorrências 2.0", page_icon="📋", layout="wide")

st.markdown("""
<style>
.stApp { background: linear-gradient(180deg, #f5f8fc 0%, #ffffff 60%); }
.block-container { max-width: 1250px; }
h1, h2, h3 { color: #0e2c70; }
div[data-testid="stMetric"] {
    background: #ffffff; border: 1px solid #e6ecf5; border-radius: 12px;
    padding: 10px 14px; box-shadow: 0 1px 3px rgba(14,44,112,.06);
}
.sugestao-card {
    background: #ffffff; border-left: 4px solid #2563EB; border-radius: 10px;
    padding: 12px 16px; margin-bottom: 10px; box-shadow: 0 1px 3px rgba(14,44,112,.06);
}
.sugestao-card.alerta { border-left-color: #D97706; }
.sugestao-card.critico { border-left-color: #DC2626; }
.sugestao-card.ok { border-left-color: #059669; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Estado persistente entre reruns (corrige o bug em que clicar num botão de
# download "sumia" com os resultados — o Streamlit reexecuta o script inteiro
# a cada clique, e sem session_state os DataFrames processados se perdiam).
# ---------------------------------------------------------------------------
st.session_state.setdefault("reset_ctr", 0)
st.session_state.setdefault("settings_ctr", 0)
st.session_state.setdefault("resultados", None)
st.session_state.setdefault("df_futuras_total", None)
st.session_state.setdefault("erros", [])
st.session_state.setdefault("resumo_papeis", None)
st.session_state.setdefault("processed_at", None)

PROCESSADORES = {
    "FALTAS": processar_faltas,
    "FERIAS": processar_ferias,
    "AFASTAMENTOS": processar_afastamentos,
}

LABEL_PAPEL = {
    "BASE": "🗃️ Base de colaboradores",
    "LANCADAS": "✅ Ocorrências já lançadas (conferência)",
    "LAYOUT": "📐 Layout de importação (modelo)",
    "DOMINIO": "📄 Relatório da Domínio",
}

st.title("📋 Ocorrências 2.0 — Conferência & Importação")
st.caption("Domínio → EasyApp • Faltas, Férias, Afastamentos, Rescisões/Avisos • "
           "os arquivos são identificados automaticamente pelo conteúdo")

tab_proc, tab_dash = st.tabs(["🚀 Processamento", "📊 Dashboard"])

# ---------------------------------------------------------------------------
# Sidebar — entradas e ações
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ Entradas")
    st.caption("Envie tudo aqui — o app identifica cada arquivo sozinho.")
    up_files = st.file_uploader(
        "Arquivos (base de colaboradores, relatório de lançadas, relatórios da Domínio)",
        type=["xls", "xlsx"], accept_multiple_files=True,
        key=f"up_{st.session_state['reset_ctr']}")
    st.divider()
    limiar = st.number_input("Faltas: mínimo de dias p/ mês todo", 1, 31, 20,
                             key=f"limiar_{st.session_state['settings_ctr']}")
    dividir = st.radio("Saída", ["Juntar empresas (1 arquivo por tipo)",
                                 "Dividir por tipo + empresa"],
                       key=f"dividir_{st.session_state['settings_ctr']}"
                       ) == "Dividir por tipo + empresa"
    processar = st.button("🚀 Processar", type="primary", use_container_width=True)
    col_a, col_b = st.columns(2)
    nova_busca = col_a.button("🔄 Nova busca", use_container_width=True,
                              help="Limpa os resultados e os arquivos enviados para você "
                                   "começar uma nova conferência do zero (mantém limiar/saída).")
    limpar_tudo = col_b.button("🗑️ Limpar tudo", use_container_width=True,
                               help="Limpa TUDO — arquivos, resultados e configurações — "
                                    "e volta o app ao estado inicial.")
    if st.session_state["processed_at"]:
        st.caption(f"🕒 Última busca: {st.session_state['processed_at'].strftime('%d/%m/%Y %H:%M:%S')}")

# ---------------------------------------------------------------------------
# Botões "Nova busca" / "Limpar tudo" — apenas ajustam o session_state e
# disparam um rerun; isso NÃO interrompe downloads já iniciados no navegador
# (o download_button entrega os bytes numa requisição própria, independente
# do ciclo de execução do script).
# ---------------------------------------------------------------------------
if nova_busca or limpar_tudo:
    st.session_state["resultados"] = None
    st.session_state["df_futuras_total"] = None
    st.session_state["erros"] = []
    st.session_state["resumo_papeis"] = None
    st.session_state["processed_at"] = None
    st.session_state["reset_ctr"] += 1
    if limpar_tudo:
        st.session_state["settings_ctr"] += 1
    st.rerun()

# ---------------------------------------------------------------------------
# Processamento — SEMPRE descarta qualquer resultado anterior antes de
# começar, garantindo que o botão realiza uma busca 100% nova a cada clique
# (nunca reaproveita um resultado desatualizado de uma execução anterior).
# ---------------------------------------------------------------------------
if processar:
    st.session_state["resultados"] = None
    st.session_state["df_futuras_total"] = None
    st.session_state["erros"] = []
    st.session_state["resumo_papeis"] = None

    if not up_files:
        st.error("Envie ao menos a **base de colaboradores** e um **relatório da Domínio**.")
        st.stop()

    tmp = tempfile.mkdtemp()

    # ---------- 1) Classificar cada arquivo pelo conteúdo ----------
    papeis = {"BASE": [], "LANCADAS": [], "LAYOUT": [], "DOMINIO": [], None: []}
    dfs = {}
    with st.spinner("Identificando arquivos..."):
        for up in up_files:
            try:
                df = ler_arquivo_upload(up, tmp)
            except Exception as e:
                st.warning(f"**{up.name}**: não consegui ler — {e}")
                papeis[None].append(up.name)
                continue
            papel = classificar(df)
            papeis[papel].append(up.name)
            dfs[up.name] = (papel, df)

    st.subheader("🔎 Arquivos identificados")
    for papel, nomes in papeis.items():
        if papel is None:
            for n in nomes:
                st.warning(f"❓ **{n}** — tipo não reconhecido (ignorado).")
            continue
        for n in nomes:
            st.write(f"{LABEL_PAPEL[papel]}: `{n}`")

    if papeis["LAYOUT"]:
        st.info("O layout de importação foi reconhecido como **modelo de referência** — "
                "ele não precisa ser enviado; as colunas já estão embutidas no app.")

    if not papeis["BASE"]:
        st.error("**Base de colaboradores não encontrada.** Envie a exportação do sistema "
                 "(a planilha com colunas Id:, Nome, Matricula, Empresa:, ...).")
        st.stop()
    if not papeis["DOMINIO"]:
        st.error("**Nenhum relatório da Domínio encontrado.** Envie ao menos um "
                 "(Faltas, Férias, Afastamentos, Rescisões ou Avisos).")
        st.stop()

    # Rescisões/Avisos precisam das DUAS planilhas (Aviso + Rescisão) juntas,
    # pois uma completa a outra (motivo só existe na Rescisão, datas mais
    # confiáveis às vezes só existem na Aviso). Reunimos antes de processar.
    dfs_aviso, dfs_rescisao = [], []

    # ---------- 2) Base de colaboradores ----------
    with st.spinner("Carregando base de colaboradores..."):
        nome_base = papeis["BASE"][0]
        base = BaseColaboradores(dfs[nome_base][1])
        st.success(f"Base carregada: {len(dfs[nome_base][1])} colaboradores.")

    # ---------- 3) Conferência (lançadas) ----------
    lancadas = None
    if papeis["LANCADAS"]:
        with st.spinner("Indexando ocorrências já lançadas..."):
            nome_l = papeis["LANCADAS"][0]
            try:
                lancadas = OcorrenciasLancadas(dfs[nome_l][1])
                st.success(f"Conferência ativa: {len(lancadas.df)} ocorrências no sistema.")
            except Exception as e:
                st.warning(f"Não consegui indexar as lançadas ({e}). Conferência DESATIVADA.")
    else:
        st.warning("Sem relatório de lançadas — conferência de duplicidade **DESATIVADA**.")

    # ---------- 4) Processar relatórios da Domínio ----------
    resultados, erros = {}, []
    for nome in papeis["DOMINIO"]:
        df = dfs[nome][1]
        try:
            tipo = detectar_tipo_relatorio(df)
            if not tipo:
                erros.append(f"**{nome}**: tipo de relatório Domínio não reconhecido.")
                continue
            if tipo == "AVISOS":
                dfs_aviso.append(df)
                continue
            if tipo == "RESCISOES":
                dfs_rescisao.append(df)
                continue
            df_imp, df_res = PROCESSADORES[tipo](df, base, lancadas,
                                                 limiar=limiar, hoje=date.today())
            if tipo in resultados:  # consolida vários arquivos do mesmo tipo
                ai, ar = resultados[tipo]
                resultados[tipo] = (pd.concat([ai, df_imp], ignore_index=True),
                                    pd.concat([ar, df_res], ignore_index=True))
            else:
                resultados[tipo] = (df_imp, df_res)
        except NotImplementedError as e:
            erros.append(f"**{nome}**: {e}")
        except Exception as e:
            erros.append(f"**{nome}**: erro — {e}")

    # ---------- 4b) Rescisões/Avisos (junta as duas planilhas, se houver) ----------
    df_futuras_total = None
    if dfs_aviso or dfs_rescisao:
        try:
            df_aviso_all = pd.concat(dfs_aviso, ignore_index=True) if dfs_aviso else None
            df_rescisao_all = pd.concat(dfs_rescisao, ignore_index=True) if dfs_rescisao else None
            df_imp, df_res, df_futuras = processar_rescisoes_avisos(
                df_aviso=df_aviso_all, df_rescisao=df_rescisao_all,
                base=base, lancadas=lancadas, hoje=date.today())
            resultados["RESCISOES"] = (df_imp, df_res)
            df_futuras_total = df_futuras
        except Exception as e:
            erros.append(f"**Rescisões/Avisos**: erro — {e}")

    st.session_state["resultados"] = resultados
    st.session_state["df_futuras_total"] = df_futuras_total
    st.session_state["erros"] = erros
    st.session_state["resumo_papeis"] = {k: v for k, v in papeis.items() if k is not None}
    st.session_state["processed_at"] = datetime.now()
    st.rerun()

# ---------------------------------------------------------------------------
# Renderização — SEMPRE a partir do session_state, nunca de variáveis locais
# do bloco "if processar". Assim, clicar em qualquer download_button (que
# provoca um novo rerun do script) continua mostrando os mesmos resultados.
# ---------------------------------------------------------------------------
resultados = st.session_state["resultados"]
df_futuras_total = st.session_state["df_futuras_total"]
erros = st.session_state["erros"] or []

ICONES_STATUS = {
    "CADASTRADO": "✅", "VÁLIDO": "✅",
    "ERRO": "❌",
    "NÃO CADASTRADO": "🕒", "EM ABERTO": "🕒",
    "JÁ LANÇADO": "⏭️",
    "CONFERÊNCIA MANUAL": "🛈",
}


def _icone(status: str) -> str:
    s = str(status).upper()
    if "ERRO" in s:
        return "❌"
    if s.startswith("CADASTRADO") or s == "VÁLIDO":
        return "✅"
    if s.startswith("NÃO CADASTRADO") or s == "EM ABERTO":
        return "🕒"
    if s == "JÁ LANÇADO":
        return "⏭️"
    if s == "CONFERÊNCIA MANUAL":
        return "🛈"
    return "ℹ️"


with tab_proc:
    if resultados:
        for e in erros:
            st.warning(e)

        for tipo, (df_imp, df_res) in resultados.items():
            st.subheader(f"🗂️ {_nome_bonitinho(tipo)}")
            contagem = df_res["STATUS"].value_counts()
            cols = st.columns(min(len(contagem) + 1, 5))
            cols[0].metric("Linhas do relatório", len(df_res))
            for i, (status, qtd) in enumerate(contagem.items(), start=1):
                if i >= len(cols):
                    break
                cols[i].metric(f"{_icone(status)} {status[:22]}", int(qtd))
            st.caption(f"➡️ **{len(df_imp)}** linha(s) vão para o CSV de importação.")
            st.dataframe(df_res, use_container_width=True, height=320)

            if tipo == "RESCISOES" and df_futuras_total is not None and not df_futuras_total.empty:
                st.warning(f"⏳ **{len(df_futuras_total)}** colaborador(es) com demissão FUTURA — "
                          "ainda ATIVOS, não demitir agora. Veja a aba/arquivo de alerta.")
                st.dataframe(df_futuras_total, use_container_width=True, height=200)

        # arquivos são remontados a cada renderização (barato — só formata os
        # DataFrames já calculados) para que os botões de download continuem
        # funcionando após qualquer rerun, sem depender de estado adicional.
        pasta_mes = f"MES {date.today().month:02d}"
        arquivos = montar_arquivos(resultados, dividir, pasta_mes=pasta_mes)
        if df_futuras_total is not None and not df_futuras_total.empty:
            arquivos[f"{pasta_mes}/ALERTA - DEMISSOES FUTURAS.xlsx"] = resultado_xlsx_bytes(
                df_futuras_total, titulo_aba="Demissoes Futuras")

        st.divider()
        st.subheader("📦 Downloads")
        st.download_button("⬇️ Baixar tudo (ZIP)", zip_bytes(arquivos),
                           file_name="ocorrencias_resultado.zip",
                           mime="application/zip", type="primary",
                           key="dl_zip")
        for i, (caminho, data) in enumerate(arquivos.items()):
            mime = "text/csv" if caminho.endswith(".csv") else \
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            st.download_button(f"⬇️ {caminho}", data, file_name=caminho.split("/")[-1],
                               mime=mime, key=f"dl_{i}_{caminho}")
    else:
        st.info("Envie os arquivos na barra lateral e clique em **Processar**. "
                "O app reconhece sozinho: base de colaboradores, lançadas e relatórios da Domínio.")
        st.markdown("""
**Como funciona (visão geral):**
1. Lê os relatórios da Domínio e casa cada colaborador com a base do sistema
   (matrícula eSocial → nome+empresa → nome).
2. Confere contra o que já está lançado no sistema (EasyApp).
3. Gera **CSV de importação** (pronto para o EasyApp) + **XLSX de resultado** com
   STATUS colorido e OBS (não encontrado, nome divergente, demitido/não ativo,
   PCD, demissão futura, aviso possivelmente cancelado etc.).
4. Use **🔄 Nova busca** para conferir um novo lote de arquivos, ou
   **🗑️ Limpar tudo** para reiniciar o app do zero.
""")

# ---------------------------------------------------------------------------
# Dashboard — visão consolidada + sugestões inteligentes
# ---------------------------------------------------------------------------
with tab_dash:
    if not resultados:
        st.info("Processe algum lote de arquivos na aba **🚀 Processamento** para ver o "
                "Dashboard com gráficos e sugestões.")
    else:
        total_linhas = sum(len(r) for _, r in resultados.values())
        total_importacao = sum(len(i) for i, _ in resultados.values())
        total_erro_data = sum(int(r["STATUS"].astype(str).str.contains("ERRO").sum())
                              for _, r in resultados.values())
        total_nao_encontrado = sum(
            int(r["OBS"].astype(str).str.contains("NÃO ENCONTRADO", na=False).sum())
            if "OBS" in r.columns else 0 for _, r in resultados.values())
        total_futuras = len(df_futuras_total) if df_futuras_total is not None else 0

        st.subheader("📊 Visão geral")
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("📄 Linhas processadas", total_linhas)
        k2.metric("📥 P/ importar", total_importacao)
        k3.metric("❌ Erros de data", total_erro_data)
        k4.metric("👤 Não encontrados", total_nao_encontrado)
        k5.metric("⏳ Demissões futuras", total_futuras)

        col_g1, col_g2 = st.columns(2)

        # ---- gráfico 1: status por tipo de ocorrência (barras empilhadas) ----
        linhas_status = []
        for tipo, (_, df_res) in resultados.items():
            for status, qtd in df_res["STATUS"].value_counts().items():
                linhas_status.append({
                    "Tipo": _nome_bonitinho(tipo),
                    "Status": status,
                    "Qtd": int(qtd),
                })
        if linhas_status:
            df_plot = pd.DataFrame(linhas_status)
            fig1 = px.bar(df_plot, x="Tipo", y="Qtd", color="Status",
                         title="Status por tipo de ocorrência", text="Qtd")
            fig1.update_layout(legend_title="", height=420,
                               margin=dict(t=50, b=10, l=10, r=10))
            col_g1.plotly_chart(fig1, use_container_width=True)

        # ---- gráfico 2: distribuição por empresa (linhas para importação) ----
        linhas_emp = []
        for tipo, (_, df_res) in resultados.items():
            if "CODI_EMP" not in df_res.columns:
                continue
            for codi, qtd in df_res["CODI_EMP"].astype(str).value_counts().items():
                linhas_emp.append({"Empresa": APELIDOS.get(codi, codi), "Qtd": int(qtd)})
        if linhas_emp:
            df_emp = pd.DataFrame(linhas_emp).groupby("Empresa", as_index=False)["Qtd"].sum()
            cores = [CORES_EMPRESA.get(cod, "#94A3B8")
                    for cod in [k for k, v in APELIDOS.items() if v in df_emp["Empresa"].values]]
            fig2 = px.pie(df_emp, names="Empresa", values="Qtd", hole=0.45,
                         title="Distribuição de registros por empresa",
                         color="Empresa",
                         color_discrete_map={v: CORES_EMPRESA.get(k, "#94A3B8")
                                             for k, v in APELIDOS.items()})
            fig2.update_layout(height=420, margin=dict(t=50, b=10, l=10, r=10))
            col_g2.plotly_chart(fig2, use_container_width=True)

        # ---- sugestões inteligentes (baseadas em regras sobre os resultados) ----
        st.subheader("💡 Sugestões inteligentes")
        sugestoes = []  # (nivel, texto)  nivel: ok | alerta | critico

        for tipo, (df_imp, df_res) in resultados.items():
            nome_tipo = _nome_bonitinho(tipo)
            n_erro = int(df_res["STATUS"].astype(str).str.contains("ERRO").sum())
            if n_erro:
                sugestoes.append(("critico",
                    f"**{nome_tipo}**: {n_erro} linha(s) com **erro de data** entre o que já "
                    f"está lançado e o que a Domínio informa — vale conferir manualmente antes "
                    f"de considerar encerrado."))
            n_nao_enc = int(df_res["OBS"].astype(str).str.contains("NÃO ENCONTRADO", na=False).sum()) \
                if "OBS" in df_res.columns else 0
            if n_nao_enc:
                sugestoes.append(("alerta",
                    f"**{nome_tipo}**: {n_nao_enc} colaborador(es) da Domínio não foram "
                    f"encontrados na base do sistema — pode ser admissão recente ainda não "
                    f"cadastrada, ou divergência de nome/matrícula."))
            n_atencao = int(df_res["OBS"].astype(str).str.contains("ATENÇÃO", na=False).sum()) \
                if "OBS" in df_res.columns else 0
            if n_atencao:
                sugestoes.append(("alerta",
                    f"**{nome_tipo}**: {n_atencao} caso(s) com um lançamento antigo parecido "
                    f"ainda ATIVO no sistema, mas fora da data atual — pode ter sido "
                    f"**cancelado/substituído na Domínio** sem cancelar no sistema. Confira se "
                    f"precisa cancelar o lançamento antigo antes de registrar o novo."))
            if tipo == "RESCISOES":
                n_manual = int((df_res["STATUS"] == "CONFERÊNCIA MANUAL").sum())
                if n_manual:
                    sugestoes.append(("alerta",
                        f"**Rescisões/Avisos**: {n_manual} caso(s) de Transferência/Morte não "
                        f"geram ocorrência automática — trate manualmente na ficha do "
                        f"colaborador."))
                n_exp = int(df_res["OBS"].astype(str).str.contains("experiência antecipada",
                                                                   na=False).sum())
                if n_exp:
                    sugestoes.append(("alerta",
                        f"**Rescisões/Avisos**: {n_exp} rescisão(ões) de contrato de experiência "
                        f"antecipada — casos reais já divergiram entre tipos de ocorrência; "
                        f"confirme manualmente qual tipo aplicar."))

        if total_futuras:
            top_emp_futuras = ""
            if df_futuras_total is not None and "EMPRESA" in df_futuras_total.columns:
                contagem_emp = df_futuras_total["EMPRESA"].value_counts()
                if not contagem_emp.empty:
                    top_emp_futuras = f" A maioria está em **{contagem_emp.index[0]}**."
            sugestoes.append(("alerta",
                f"Existem **{total_futuras}** demissão(ões) com data futura — esses "
                f"colaboradores ainda estão ATIVOS e não devem ser desligados agora.{top_emp_futuras}"))

        if linhas_emp:
            df_emp_rank = pd.DataFrame(linhas_emp).groupby("Empresa", as_index=False)["Qtd"].sum() \
                .sort_values("Qtd", ascending=False)
            if not df_emp_rank.empty:
                top = df_emp_rank.iloc[0]
                sugestoes.append(("ok" if len(df_emp_rank) <= 1 else "alerta",
                    f"**{top['Empresa']}** concentra o maior volume de registros "
                    f"({int(top['Qtd'])}) neste lote — priorize a conferência dessa empresa "
                    f"se o tempo for curto."))

        if total_erro_data == 0 and total_nao_encontrado == 0 and total_futuras == 0:
            sugestoes.append(("ok", "Nenhuma divergência crítica encontrada neste lote — os "
                                   "dados parecem consistentes entre Domínio e o sistema."))

        if not sugestoes:
            st.caption("Sem observações relevantes para este lote.")
        for nivel, texto in sugestoes:
            st.markdown(f'<div class="sugestao-card {nivel}">{texto}</div>',
                       unsafe_allow_html=True)
