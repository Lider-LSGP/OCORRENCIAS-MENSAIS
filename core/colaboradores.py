"""Base de colaboradores do sistema interno (EasyApp).

Casamento em cascata:
  1) matrícula (i_empregados / matricula_esocial da Domínio  ==  Matricula)
  2) nome normalizado + empresa
  3) nome normalizado (global)
"""
from __future__ import annotations

import re
import unicodedata
import pandas as pd
from typing import Optional

from .empresas import codi_por_nome


def clean(s) -> str:
    if s is None or (isinstance(s, float) and s != s):
        return ""
    return re.sub(r"\s+", " ", str(s).replace("\\n", " ").replace("\n", " ")).strip()


def norm_nome(s) -> str:
    s = clean(s).upper()
    s = "".join(c for c in unicodedata.normalize("NFD", s)
                if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip()


class BaseColaboradores:
    def __init__(self, df: pd.DataFrame):
        df = df.copy()
        df["_nome"] = df["Nome"].apply(norm_nome)
        df["_mat"] = df["Matricula"].apply(lambda v: re.sub(r"\D", "", clean(v)))
        df["_modal"] = (df["Matrícula/Modal:"].apply(lambda v: re.sub(r"\D", "", clean(v)))
                        if "Matrícula/Modal:" in df.columns else "")
        df["_empresa"] = df["Empresa:"].apply(clean)
        df["_codi"] = df["_empresa"].apply(codi_por_nome)
        self.df = df
        self._by_mat = {}
        self._by_nome_emp = {}
        self._by_nome = {}
        for _, r in df.iterrows():
            if r["_mat"]:
                self._by_mat.setdefault((r["_codi"], r["_mat"]), r)
                self._by_mat.setdefault(("", r["_mat"]), r)  # matrícula sem empresa
            if r["_nome"]:
                self._by_nome_emp.setdefault((r["_codi"], r["_nome"]), r)
                self._by_nome.setdefault(r["_nome"], r)

    def buscar(self, nome=None, matricula=None, codi_emp=None) -> Optional[pd.Series]:
        mat = re.sub(r"\D", "", clean(matricula)) if matricula else ""
        nn = norm_nome(nome) if nome else ""
        ce = str(codi_emp).strip() if codi_emp else ""
        if mat:
            r = self._by_mat.get((ce, mat))
            if r is not None:
                return r
        if nn and ce:
            r = self._by_nome_emp.get((ce, nn))
            if r is not None:
                return r
        if nn:
            return self._by_nome.get(nn)
        if mat:
            return self._by_mat.get(("", mat))
        return None

    @staticmethod
    def info(r):
        """Extrai campos úteis de uma linha casada."""
        if r is None:
            return {}
        return {
            "parceiro_id": clean(r.get("Id:")),
            "nome": clean(r.get("Nome")),
            "matricula": clean(r.get("Matricula")),
            "empresa": clean(r.get("Empresa:")),
            "codi_emp": clean(r.get("_codi")),
            "escala": clean(r.get("Tipo Escala")),
            "posto": clean(r.get("Posto Trabalho:")),
            "funcao": clean(r.get("Função:")),
            "ativo": clean(r.get("Ativo")),
            "demissao": clean(r.get("DT/Demissão")),
            "pcd": clean(r.get("Tipo/PCD:")),
        }
