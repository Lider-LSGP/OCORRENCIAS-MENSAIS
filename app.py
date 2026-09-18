"""Ocorrências 2.0 — conferência e geração de importações EasyApp.

O app CLASSIFICA CADA ARQUIVO PELO CONTEÚDO: não importa em qual campo o
arquivo foi enviado, ele vai para o papel certo (base / lançadas / Domínio).
O layout de importação enviado por engano vira apenas referência e gera aviso.

Módulos: FALTAS MÊS TODO (completo) | FÉRIAS, AFASTAMENTOS e
RESCISÕES/AVISOS (esqueleto — mesma estrutura de saída).

Rodar:
    pip install -r requirements.txt
    streamlit run app.py
"""
from __future__ import annotations

import tempfile
from datetime import date

import pandas as pd
import streamlit as st

from core.leitura import ler_arquivo_upload
from core.classificador import classificar
from core.colaboradores import BaseColaboradores
from core.ocorrencias import OcorrenciasLancadas
from core.faltas import (detectar_tipo_relatorio, processar_faltas,
                         processar_ferias, processar_afastamentos,
                         processar_rescisoes_avisos)
from core.saida import montar_arquivos, zip_bytes

st.set_page_config(page_title="Ocorrências 2.0", page_icon="📋", layout="wide")

st.markdown("""
<style>
.stApp { background: linear-gradient(180deg, #f5f8fc 0%, #ffffff 60%); }
.block-container { max-width: 1150px; }
h1, h2, h3 { color: #0e2c70; }
</style>
""", unsafe_allow_html=True)

st.title("📋 Ocorrências 2.0 — Conferência & Importação")
st.caption("Domínio → EasyApp • Faltas, Férias, Afastamentos, Rescisões/Avisos • "
           "os arquivos são identificados automaticamente pelo conteúdo")

PROCESSADORES = {
    "FALTAS": processar_faltas,
    "FERIAS": processar_ferias,
    "AFASTAMENTOS": processar_afastamentos,
    "AVISOS": processar_rescisoes_avisos,
    "RESCISOES": processar_rescisoes_avisos,
}

LABEL_PAPEL = {
    "BASE": "🗃️ Base de colaboradores",
    "LANCADAS": "✅ Ocorrências já lançadas (conferência)",
    "LAYOUT": "📐 Layout de importação (modelo)",
    "DOMINIO": "📄 Relatório da Domínio",
}

with st.sidebar:
    st.header("⚙️ Entradas")
    st.caption("Envie tudo aqui — o app identifica cada arquivo sozinho.")
    up_files = st.file_uploader(
        "Arquivos (base de colaboradores, relatório de lançadas, relatórios da Domínio)",
        type=["xls", "xlsx"], accept_multiple_files=True)
    st.divider()
    limiar = st.number_input("Faltas: mínimo de dias p/ mês todo", 1, 31, 20)
    dividir = st.radio("Saída", ["Juntar empresas (1 arquivo por tipo)",
                                 "Dividir por tipo + empresa"]) == "Dividir por tipo + empresa"
    processar = st.button("🚀 Processar", type="primary", use_container_width=True)

if processar:
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
            if tipo in ("AVISOS", "RESCISOES"):
                erros.append(f"**{nome}** ({tipo}): módulo em desenvolvimento — envie as regras para ativarmos.")
                continue
            df_imp, df_res = PROCESSADORES[tipo](df, base, lancadas,
                                                 limiar=limiar, hoje=date.today())
            if tipo in resultados:  # consolida vários arquivos do mesmo tipo
                ai, ar = resultados[tipo]
                resultados[tipo] = (pd.concat([ai, df_imp], ignore_index=True),
                                    pd.concat([ar, df_res], ignore_index=True))
            else:
                resultados[tipo] = (df_imp, df_res)

            st.subheader(f"🗂️ {nome} → {tipo}")
            contagem = df_res["STATUS"].value_counts()
            n_ir_importacao = len(df_imp)
            cols = st.columns(min(len(contagem) + 1, 5))
            cols[0].metric("Linhas do relatório", len(df_res))
            for i, (status, qtd) in enumerate(contagem.items(), start=1):
                if i >= len(cols):
                    break
                icone = ("✅" if status.startswith("CADASTRADO") and "ERRO" not in status else
                        "❌" if "ERRO" in status else
                        "🕒" if status in ("NÃO CADASTRADO", "EM ABERTO") else
                        "⏭️" if status == "JÁ LANÇADO" else
                        "✅" if status == "VÁLIDO" else "ℹ️")
                cols[i].metric(f"{icone} {status[:22]}", int(qtd))
            st.caption(f"➡️ **{n_ir_importacao}** linha(s) vão para o CSV de importação.")
            st.dataframe(df_res, use_container_width=True, height=320)
        except NotImplementedError as e:
            erros.append(f"**{nome}**: {e}")
        except Exception as e:
            erros.append(f"**{nome}**: erro — {e}")

    for e in erros:
        st.warning(e)

    if resultados:
        arquivos = montar_arquivos(resultados, dividir)
        st.divider()
        st.subheader("📦 Downloads")
        st.download_button("⬇️ Baixar tudo (ZIP)", zip_bytes(arquivos),
                           file_name="ocorrencias_resultado.zip",
                           mime="application/zip", type="primary")
        for caminho, data in arquivos.items():
            mime = "text/csv" if caminho.endswith(".csv") else \
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            st.download_button(f"⬇️ {caminho}", data, file_name=caminho.split("/")[-1], mime=mime)
else:
    st.info("Envie os arquivos na barra lateral e clique em **Processar**. "
            "O app reconhece sozinho: base de colaboradores, lançadas e relatórios da Domínio.")
    st.markdown("""
**Como funciona (Faltas Mês Todo):**
1. Lê o relatório Domínio (`valor_inf` ≥ limiar ⇒ mês todo).
2. Casa cada colaborador com a base do sistema (matrícula eSocial → nome+empresa → nome).
3. Confere contra o que já está lançado (mesma data inicial ⇒ fora da importação).
4. Gera **CSV de importação** (layout EasyApp, tipo 13) + **XLSX de resultado** com
   STATUS (VÁLIDO / EM ABERTO / JÁ LANÇADO) e OBS (não encontrado, nome divergente,
   demitido/não ativo, PCD...).
""")
