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
        # Nunca reutilizar matrícula ou nome de outra empresa.
        ce = clean(codi_emp)
        if re.fullmatch(r"[1-4](?:\.0+)?", ce):
            ce = ce.split(".")[0]
        elif ce:
            ce = codi_por_nome(ce)
            if not ce:
                return None
        mat = clean(matricula)
        mat = re.sub(r"\.0+$", "", mat)
        mat = re.sub(r"\D", "", mat)
        nn = norm_nome(nome)
        candidatos = self.df
        if ce:
            candidatos = candidatos[candidatos["_codi"] == ce]
        if nn:
            por_nome = candidatos[candidatos["_nome"] == nn]
            if len(por_nome) == 1:
                return por_nome.iloc[0]
            if len(por_nome) > 1 and mat:
                por_mat = por_nome[por_nome["_mat"] == mat]
                if len(por_mat) == 1:
                    return por_mat.iloc[0]
            # Matrícula com nome diferente ou múltiplos vínculos não é confirmação.
            return None
        if mat:
            por_mat = candidatos[candidatos["_mat"] == mat]
            if len(por_mat) == 1:
                return por_mat.iloc[0]
        return None

    def descrever_pendencia(self, nome, codi_emp):
        ce = re.sub(r"\.0+$", "", clean(codi_emp))
        encontrados = self.df[self.df["_nome"] == norm_nome(nome)]
        outros = encontrados[encontrados["_codi"] != ce]
        if not outros.empty:
            empresas = " / ".join(sorted(set(outros["_empresa"])))
            return ("CADASTRO ENCONTRADO SOMENTE/ TAMBÉM EM OUTRA EMPRESA: "
                    + empresas + ". Vínculo da empresa de origem não confirmado; "
                    "fora do CSV. Conferir cadastro/transferência manualmente.")
        if len(encontrados) > 1:
            return "MÚLTIPLOS VÍNCULOS: identificação ambígua; fora do CSV. Conferir manualmente."
        return "VÍNCULO NÃO CONFIRMADO NA EMPRESA DE ORIGEM: fora do CSV; conferir cadastro/matrícula."

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
