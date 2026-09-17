"""Configuração das 4 empresas do grupo."""
from __future__ import annotations
import re, unicodedata

EMPRESAS = {
    "1": "VSP VIGILANCIA E SEGURANCA PATRIMONIAL L",
    "2": "ATIVA TERCEIRIZACAO DE MAO DE OBRA LTDA",
    "3": "LIDER MULTISSERVICOS LTDA",
    "4": "LIDER LIMPE LIMPEZA COMERCIAL LTDA",
}

# aliases normalizados -> codi_emp (cobre abreviações que aparecem na base)
_ALIASES = {
    "vsp": "1",
    "vsp vigilancia": "1",
    "ativa": "2",
    "ativa terceirizacao": "2",
    "lider multisservicos": "3",
    "lider multi": "3",
    "lider limpe": "4",
    "lider limpe limpeza comercial": "4",
}


def _norm(s) -> str:
    if s is None:
        return ""
    s = unicodedata.normalize("NFD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().lower()


def nome_empresa(codi_emp) -> str:
    return EMPRESAS.get(str(codi_emp).strip(), "")


def codi_por_nome(nome: str) -> str:
    """Resolve nome de empresa (como vem nas planilhas) -> codi_emp."""
    n = _norm(nome)
    if not n:
        return ""
    for c, oficial in EMPRESAS.items():
        if _norm(oficial) == n:
            return c
    for alias, c in _ALIASES.items():
        if n.startswith(alias):
            return c
    return ""
