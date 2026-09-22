"""Leitura do Relatório Sintético de Ocorrências do sistema interno (EasyApp).

O relatório tem ~3 linhas de título e o cabeçalho real na linha 4:
PARCEIROID | NOME FUNCIONÁRIO | TIPOOCORRENCIA ID | DATA OCORRÊNCIA | ... |
DATA INÍCIO | DIAS AFASTAMENTO | DATA FIM | ...

A classe aceita caminho de arquivo OU DataFrame (lido com header=0 ou
header=None) — a linha de cabeçalho é sempre localizada automaticamente.

Fornece consultas O(1) por parceiro_id/nome para:
  * ja_lancado(data_inicio, ...)              -> conferência simples (Faltas)
  * periodos_colaborador(parceiro_id, nome)   -> lista de (ini, fim, tipo)
    usada por Férias/Afastamentos para achar o lançamento mais próximo do
    período esperado (mesmo que fragmentado em vários registros).
"""
from __future__ import annotations

import re
import pandas as pd
from datetime import datetime, date
from typing import Optional, Set, Tuple, List, Dict

from .colaboradores import norm_nome
from .leitura import ler_planilha


def _to_date(v) -> Optional[date]:
    if v is None or (isinstance(v, float) and v != v):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s[:19] if len(s) > 10 else s, fmt).date()
        except ValueError:
            continue
    return None


def _localizar_cabecalho(df: pd.DataFrame) -> pd.DataFrame:
    """Garante que as colunas do DataFrame sejam o cabeçalho real (PARCEIROID...)."""
    cols_up = [str(c).strip().upper() for c in df.columns]
    if "PARCEIROID" in cols_up:
        out = df.copy()
        out.columns = [str(c).strip() for c in df.columns]
        return out
    for i in range(min(20, len(df))):
        vals = [str(v).strip().upper() for v in df.iloc[i]]
        if "PARCEIROID" in vals:
            out = df.copy()
            out.columns = [str(v).strip() for v in df.iloc[i]]
            return out.iloc[i + 1:].reset_index(drop=True)
    raise RuntimeError(
        "Cabeçalho PARCEIROID não encontrado no relatório de ocorrências. "
        "Confirme se enviou o Relatório Sintético de Ocorrências do EasyApp.")


class OcorrenciasLancadas:
    """Índice do que já está lançado no sistema interno."""

    def __init__(self, path_ou_df):
        if isinstance(path_ou_df, pd.DataFrame):
            df = path_ou_df
        else:
            df = ler_planilha(path_ou_df, dtype=str, header=None)
        df = _localizar_cabecalho(df)

        df["_parceiro"] = df["PARCEIROID"].apply(lambda v: re.sub(r"\D", "", str(v)))
        df["_nome"] = df["NOME FUNCIONÁRIO"].apply(norm_nome)
        df["_tipo"] = df["TIPOOCORRENCIA ID"].apply(lambda v: re.sub(r"\D", "", str(v)))
        df["_ini"] = df["DATA INÍCIO"].apply(_to_date)
        df["_fim"] = df["DATA FIM"].apply(_to_date) if "DATA FIM" in df.columns else None
        self.df = df

        # índices O(1) por parceiro/nome -> lista de registros {ini, fim, tipo}
        self._por_parceiro: Dict[str, List[dict]] = {}
        self._por_nome: Dict[str, List[dict]] = {}
        self._chave_datas_parceiro: Set[Tuple[str, date]] = set()
        self._chave_datas_nome: Set[Tuple[str, date]] = set()

        for _, r in df.iterrows():
            ini = r["_ini"]
            if ini is None:
                continue
            fim = r["_fim"] if r["_fim"] is not None else ini
            rec = {"ini": ini, "fim": fim, "tipo": r["_tipo"]}
            if r["_parceiro"]:
                self._por_parceiro.setdefault(r["_parceiro"], []).append(rec)
                self._chave_datas_parceiro.add((r["_parceiro"], ini))
            if r["_nome"]:
                self._por_nome.setdefault(r["_nome"], []).append(rec)
                self._chave_datas_nome.add((r["_nome"], ini))

    def ja_lancado(self, data_inicio: date, parceiro_id: str = "", nome: str = "") -> bool:
        """True se existe QUALQUER ocorrência lançada com a mesma data inicial."""
        pid = re.sub(r"\D", "", str(parceiro_id or ""))
        nn = norm_nome(nome) if nome else ""
        if pid and (pid, data_inicio) in self._chave_datas_parceiro:
            return True
        if nn and (nn, data_inicio) in self._chave_datas_nome:
            return True
        return False

    def tipos_do_colaborador(self, data_inicio: date, parceiro_id: str = "", nome: str = "") -> list:
        recs = self.periodos_colaborador(parceiro_id, nome)
        return [r["tipo"] for r in recs if r["ini"] == data_inicio]

    def periodos_colaborador(self, parceiro_id: str = "", nome: str = "",
                             tipos: Optional[tuple] = None) -> List[dict]:
        """Todos os períodos já lançados para o colaborador (opcionalmente filtrado
        por tipoocorrencia_id). Usado para achar o lançamento mais próximo do
        período esperado, mesmo que fragmentado em vários registros."""
        pid = re.sub(r"\D", "", str(parceiro_id or ""))
        nn = norm_nome(nome) if nome else ""
        recs = self._por_parceiro.get(pid, []) if pid else []
        if not recs and nn:
            recs = self._por_nome.get(nn, [])
        if tipos:
            tset = set(tipos)
            recs = [r for r in recs if r["tipo"] in tset]
        return recs

    def tipos_na_data(self, data: date, parceiro_id: str = "", nome: str = "") -> list:
        """Tipos (tipoocorrencia_id) lançados para o colaborador começando em
        uma data específica — ex.: faltas dia a dia."""
        pid = re.sub(r"\D", "", str(parceiro_id or ""))
        nn = norm_nome(nome) if nome else ""
        recs = self._por_parceiro.get(pid, []) if pid else []
        if not recs and nn:
            recs = self._por_nome.get(nn, [])
        return [r["tipo"] for r in recs if r["ini"] == data]

    def tem_tipo_na_data(self, data: date, tipos, parceiro_id: str = "",
                         nome: str = "") -> bool:
        tset = set(tipos)
        return any(t in tset for t in self.tipos_na_data(data, parceiro_id, nome))

    def tem_tipo_no_intervalo(self, ini: date, fim: date, tipos,
                              parceiro_id: str = "", nome: str = "") -> bool:
        """True se existe algum lançamento desses tipos cobrindo parte do intervalo."""
        pid = re.sub(r"\D", "", str(parceiro_id or ""))
        nn = norm_nome(nome) if nome else ""
        recs = self._por_parceiro.get(pid, []) if pid else []
        if not recs and nn:
            recs = self._por_nome.get(nn, [])
        tset = set(tipos)
        for r in recs:
            if r["tipo"] in tset and r["ini"] <= fim and r["fim"] >= ini:
                return True
        return False

