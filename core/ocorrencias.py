"""Leitura do Relatório Sintético de Ocorrências do sistema interno.

O relatório tem ~3 linhas de título e o cabeçalho real na linha 4:
PARCEIROID | NOME FUNCIONÁRIO | TIPOOCORRENCIA ID | DATA OCORRÊNCIA | ... |
DATA INÍCIO | DIAS AFASTAMENTO | DATA FIM | ...
"""
from __future__ import annotations

import re
import pandas as pd
from datetime import datetime, date
from typing import Optional, Set, Tuple

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


class OcorrenciasLancadas:
    """Índice do que já está lançado no sistema interno."""

    def __init__(self, path_ou_df):
        if isinstance(path_ou_df, pd.DataFrame):
            df = path_ou_df
        else:
            raw = ler_planilha(path_ou_df, dtype=str, header=None)
            # localiza linha de cabeçalho (contém PARCEIROID)
            hdr = None
            for i in range(min(10, len(raw))):
                if any(str(v).strip().upper() == "PARCEIROID" for v in raw.iloc[i]):
                    hdr = i
                    break
            if hdr is None:
                raise RuntimeError("Cabeçalho PARCEIROID não encontrado no relatório de ocorrências.")
            raw.columns = [str(v).strip() for v in raw.iloc[hdr]]
            df = raw.iloc[hdr + 1:].reset_index(drop=True)

        self.df = df
        self.df["_parceiro"] = df["PARCEIROID"].apply(lambda v: re.sub(r"\D", "", str(v)))
        self.df["_nome"] = df["NOME FUNCIONÁRIO"].apply(norm_nome)
        self.df["_tipo"] = df["TIPOOCORRENCIA ID"].apply(lambda v: re.sub(r"\D", "", str(v)))
        self.df["_ini"] = df["DATA INÍCIO"].apply(_to_date)

        # chave: (parceiro_id, data_inicio) e (nome, data_inicio) -> tipos lançados
        self._por_parceiro: Set[Tuple[str, date]] = set()
        self._por_nome: Set[Tuple[str, date]] = set()
        for _, r in self.df.iterrows():
            if r["_ini"] is None:
                continue
            if r["_parceiro"]:
                self._por_parceiro.add((r["_parceiro"], r["_ini"]))
            if r["_nome"]:
                self._por_nome.add((r["_nome"], r["_ini"]))

    def ja_lancado(self, data_inicio: date, parceiro_id: str = "", nome: str = "") -> bool:
        """True se existe QUALQUER ocorrência lançada com a mesma data inicial."""
        pid = re.sub(r"\D", "", str(parceiro_id or ""))
        nn = norm_nome(nome) if nome else ""
        if pid and (pid, data_inicio) in self._por_parceiro:
            return True
        if nn and (nn, data_inicio) in self._por_nome:
            return True
        return False

    def tipos_do_colaborador(self, data_inicio: date, parceiro_id: str = "", nome: str = "") -> list:
        pid = re.sub(r"\D", "", str(parceiro_id or ""))
        nn = norm_nome(nome) if nome else ""
        tipos = []
        for _, r in self.df.iterrows():
            if r["_ini"] != data_inicio:
                continue
            if (pid and r["_parceiro"] == pid) or (nn and r["_nome"] == nn):
                tipos.append(r["_tipo"])
        return tipos
