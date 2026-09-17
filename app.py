"""Ocorrências 2.0 — conferência e geração de importações EasyApp.

Módulos: FALTAS MÊS TODO (completo) | FÉRIAS, AFASTAMENTOS e
RESCISÕES/AVISOS (esqueleto — mesma estrutura de saída).

Rodar localmente:
    pip install -r requirements.txt
    streamlit run app.py
"""
from __future__ import annotations

import tempfile
from datetime import date

import pandas as pd
import streamlit as st

from core.leitura import ler_arquivo_upload
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
st.caption("Domínio → EasyApp • Faltas, Férias, Afastamentos, Rescisões/Avisos")

PROCESSADORES = {
    "FALTAS": processar_faltas,
    "FERIAS": processar_ferias,
    "AFASTAMENTOS": processar_afastamentos,
    "AVISOS": processar_rescisoes_avisos,
    "RESCISOES": processar_rescisoes_avisos,
}

with st.sidebar:
    st.header("⚙️ Entradas")
    up_base = st.file_uploader("Base de colaboradores do sistema (SISTEMA.xls/xlsx)",
                               type=["xls", "xlsx"])
    up_lanc = st.file_uploader("Relatório de ocorrências JÁ LANÇADAS (EasyApp)",
                               type=["xls", "xlsx"])
    up_rels = st.file_uploader("Relatórios da Domínio (um ou vários)",
                               type=["xls", "xlsx"], accept_multiple_files=True)
    st.divider()
    limiar = st.number_input("Faltas: mínimo de dias p/ mês todo", 1, 31, 20)
    dividir = st.radio("Saída", ["Juntar empresas (1 arquivo por tipo)",
                                 "Dividir por tipo + empresa"]) == "Dividir por tipo + empresa"
    processar = st.button("🚀 Processar", type="primary", use_container_width=True)

if processar:
    if not up_base or not up_rels:
        st.error("Envie ao menos a **base de colaboradores** e um **relatório da Domínio**.")
        st.stop()

    tmp = tempfile.mkdtemp()
    with st.spinner("Carregando base de colaboradores..."):
        df_base = ler_arquivo_upload(up_base, tmp)
        base = BaseColaboradores(df_base)
        st.success(f"Base carregada: {len(df_base)} colaboradores.")

    lancadas = None
    if up_lanc:
        with st.spinner("Carregando ocorrências lançadas (conferência)..."):
            df_lanc = ler_arquivo_upload(up_lanc, tmp)
            lancadas = OcorrenciasLancadas(df_lanc)
            st.success(f"Conferência ativa: {len(lancadas.df)} ocorrências no sistema.")
    else:
        st.warning("Sem relatório de lançadas — a conferência de duplicidade ficou DESATIVADA.")

    resultados, erros = {}, []
    for up in up_rels:
        try:
            df = ler_arquivo_upload(up, tmp)
            tipo = detectar_tipo_relatorio(df)
            if not tipo:
                erros.append(f"**{up.name}**: tipo de relatório não reconhecido.")
                continue
            if tipo != "FALTAS":
                erros.append(f"**{up.name}** ({tipo}): módulo em desenvolvimento — envie as regras para ativarmos.")
                continue
            df_imp, df_res = PROCESSADORES[tipo](df, base, lancadas,
                                                 limiar=limiar, hoje=date.today())
            resultados.setdefault(tipo, (df_imp, df_res))
            # se já existia outro arquivo do mesmo tipo, consolida
            st.subheader(f"🗂️ {up.name} → {tipo}")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Linhas do relatório", len(df_res))
            c2.metric("✅ Válidos", int((df_res["STATUS"] == "VÁLIDO").sum()))
            c3.metric("⚠️ Em aberto", int((df_res["STATUS"] == "EM ABERTO").sum()))
            c4.metric("⏭️ Já lançados", int((df_res["STATUS"] == "JÁ LANÇADO").sum()))
            st.dataframe(df_res, use_container_width=True, height=320)
        except NotImplementedError as e:
            erros.append(f"**{up.name}**: {e}")
        except Exception as e:
            erros.append(f"**{up.name}**: erro — {e}")

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
    st.info("Envie os arquivos na barra lateral e clique em **Processar**.")
    st.markdown("""
**Como funciona (Faltas Mês Todo):**
1. Lê o relatório Domínio (`valor_inf` ≥ limiar ⇒ mês todo).
2. Casa cada colaborador com a base do sistema (matrícula eSocial → nome+empresa → nome).
3. Confere contra o que já está lançado (mesma data inicial ⇒ fora da importação).
4. Gera **CSV de importação** (layout EasyApp, tipo 13) + **XLSX de resultado** com
   STATUS (VÁLIDO / EM ABERTO / JÁ LANÇADO) e OBS (não encontrado, nome divergente,
   demitido/não ativo, PCD...).
""")
