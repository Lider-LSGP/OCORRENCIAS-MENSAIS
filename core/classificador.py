"""Classificação automática de arquivos enviados — o app identifica pelo
CONTEÚDO (colunas), não pelo campo em que o usuário fez o upload.

Papéis:
  BASE      -> base de colaboradores do sistema interno (Id:, Nome, Matricula, Empresa:...)
  LANCADAS  -> Relatório Sintético de Ocorrências do EasyApp (PARCEIROID...)
  LAYOUT    -> Layout de importação (nomefuncionario, tipoocorrencia_id...) — modelo de referência
  DOMINIO   -> relatório da Domínio (codi_emp + i_empregados + colunas do tipo)
"""
from __future__ import annotations

from typing import Optional
import pandas as pd


def _cols(df: pd.DataFrame) -> set:
    return {str(c).strip().lower() for c in df.columns}


def _contem_cabecalho_em_linhas(df: pd.DataFrame, marcador: str, linhas: int = 20) -> bool:
    marcador = marcador.upper()
    cols = [str(c).strip().upper() for c in df.columns]
    if marcador in cols:
        return True
    for i in range(min(linhas, len(df))):
        if marcador in [str(v).strip().upper() for v in df.iloc[i]]:
            return True
    return False


def classificar(df: pd.DataFrame) -> Optional[str]:
    c = _cols(df)

    # Relatório de ocorrências lançadas (cabeçalho pode estar nas primeiras linhas)
    if _contem_cabecalho_em_linhas(df, "PARCEIROID"):
        return "LANCADAS"

    # Base de colaboradores do sistema interno
    if {"id:", "nome", "matricula"} <= c or ("matricula" in c and "empresa:" in c):
        return "BASE"

    # Layout de importação do EasyApp (modelo)
    if "tipoocorrencia_id" in c and "nomefuncionario" in c:
        return "LAYOUT"

    # Relatórios da Domínio
    if "codi_emp" in c and "i_empregados" in c:
        return "DOMINIO"
    if "foempregados_nome" in c or "inicio_gozo" in c or "i_afastamentos" in c \
            or "data_aviso" in c:
        return "DOMINIO"

    return None
