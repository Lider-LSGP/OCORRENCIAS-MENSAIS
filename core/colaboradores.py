"""Base de colaboradores do sistema interno (EasyApp).

Busca por nome e empresa; a matrícula desempata vínculos na mesma empresa.
"""
from __future__ import annotations

import re
import unicodedata
import pandas as pd
from typing import Optional

from .empresas import codi_por_nome, codigo_empresa


def clean(s) -> str:
    if s is None or (isinstance(s, float) and s != s):
        return ""
    return re.sub(r"\s+", " ", str(s).replace("\\n", " ").replace("\n", " ")).strip()


def norm_nome(s) -> str:
    s = clean(s).upper()
    s = "".join(c for c in unicodedata.normalize("NFD", s)
                if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s).strip()


def _cpf(valor):
    n = re.sub(r"\D", "", re.sub(r"\.0+$", "", clean(valor)))
    if not 9 <= len(n) <= 11:
        return ""
    n = n.zfill(11)
    if len(set(n)) == 1:
        return ""
    for i in (9, 10):
        resto = (sum(int(d) * (i+1-j) for j, d in enumerate(n[:i])) * 10) % 11
        if int(n[i]) != (0 if resto == 10 else resto):
            return ""
    return n


def _tipo_pcd(valor):
    v = clean(valor)
    return v if norm_nome(v) not in {"", "NAO", "N", "0", "0.0", "-", "NAN", "N/A", "NAO SE APLICA", "NAO PCD", "SEM DEFICIENCIA"} else ""


class BaseColaboradores:
    def __init__(self, df: pd.DataFrame):
        df = df.copy()
        # o cabeçalho da base às vezes vem colado ("Ativo:Sim" no lugar de "Ativo")
        for c in list(df.columns):
            if str(c).startswith("Ativo") and c != "Ativo":
                df = df.rename(columns={c: "Ativo"})
                break
        df["_nome"] = df["Nome"].apply(norm_nome)
        df["_mat"] = df["Matricula"].apply(lambda v: re.sub(r"\D", "", clean(v)))
        df["_modal"] = (df["Matrícula/Modal:"].apply(lambda v: re.sub(r"\D", "", clean(v)))
                        if "Matrícula/Modal:" in df.columns else "")
        df["_empresa"] = df["Empresa:"].apply(clean)
        df["_codi"] = df["_empresa"].apply(codi_por_nome)
        self.df = df
        self._por_empresa_nome = {}
        self._pcd_por_cpf = {}
        for _, r in df.iterrows():
            if r["_codi"] and r["_nome"]:
                self._por_empresa_nome.setdefault((r["_codi"], r["_nome"]), []).append(r)
            cpf = _cpf(r.get("CPF:"))
            tipo = _tipo_pcd(r.get("Tipo/PCD:"))
            if cpf and tipo:
                self._pcd_por_cpf.setdefault(cpf, tipo)

    def buscar(self, nome=None, matricula=None, codi_emp=None) -> Optional[pd.Series]:
        ce, nn = codigo_empresa(codi_emp), norm_nome(nome)
        if not ce or not nn:
            return None
        candidatos = self._por_empresa_nome.get((ce, nn), [])
        mat = re.sub(r"\D", "", clean(matricula))
        if mat:
            exatos = [r for r in candidatos if r["_mat"] == mat]
            if exatos:
                candidatos = exatos
        return candidatos[0] if len(candidatos) == 1 and clean(candidatos[0].get("Id:")) else None

    def tipo_pcd(self, r):
        if r is None:
            return ""
        return _tipo_pcd(r.get("Tipo/PCD:")) or self._pcd_por_cpf.get(_cpf(r.get("CPF:")), "")

    @staticmethod
    def info(r):
        """Extrai campos úteis de uma linha casada."""
        if r is None:
            return {}
        return {
            "parceiro_id": clean(r.get("Id:")),
            "nome": clean(r.get("Nome")),
            "cpf": clean(r.get("CPF:")),
            "admissao": clean(r.get("DT/Admissão:")),
            "matricula": clean(r.get("Matricula")),
            "empresa": clean(r.get("Empresa:")),
            "codi_emp": clean(r.get("_codi")),
            "escala": clean(r.get("Horario:")) or clean(r.get("Tipo Escala")),
            "escala_tipo": clean(r.get("Tipo Escala")),
            "posto": clean(r.get("Posto Trabalho:")),
            "funcao": clean(r.get("Função:")),
            "ativo": clean(r.get("Ativo")),
            "demissao": clean(r.get("DT/Demissão")),
            "pcd": clean(r.get("Tipo/PCD:")),
        }
