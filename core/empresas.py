"""Configuração das 4 empresas do grupo."""
from __future__ import annotations
import re, unicodedata

EMPRESAS = {
    "1": "VSP VIGILANCIA E SEGURANCA PATRIMONIAL L",
    "2": "ATIVA TERCEIRIZACAO DE MAO DE OBRA LTDA",
    "3": "LIDER MULTISSERVICOS LTDA",
    "4": "LIDER LIMPE LIMPEZA COMERCIAL LTDA",
}

# apelidos curtos usados na tela e nos nomes de arquivo (mais limpo e legível)
APELIDOS = {
    "1": "VSP",
    "2": "ATIVA",
    "3": "LIDER MULTISSERVIÇOS",
    "4": "LIDER LIMPE",
}

# cores de destaque por empresa (usadas nos cards/gráficos do app)
CORES_EMPRESA = {
    "1": "#2563EB",  # azul
    "2": "#059669",  # verde
    "3": "#D97706",  # laranja
    "4": "#7C3AED",  # roxo
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


def apelido_empresa(codi_emp) -> str:
    """Apelido curto (VSP / ATIVA / LIDER MULTISSERVIÇOS / LIDER LIMPE)."""
    c = str(codi_emp).strip()
    if c in APELIDOS:
        return APELIDOS[c]
    # fallback: tenta resolver pelo nome completo, se vier nome em vez de código
    ce = codi_por_nome(codi_emp) if not c.isdigit() else ""
    return APELIDOS.get(ce, nome_empresa(codi_emp) or str(codi_emp))


def apelido_por_nome(nome_qualquer: str) -> str:
    """Resolve um nome de empresa (como vem nas planilhas) direto para o apelido curto."""
    ce = codi_por_nome(nome_qualquer)
    if ce:
        return APELIDOS.get(ce, nome_qualquer)
    # já pode ser um apelido/nome parcial — tenta achar por substring
    n = _norm(nome_qualquer)
    for ce2, apelido in APELIDOS.items():
        if _norm(apelido) in n or _norm(EMPRESAS[ce2]) == n:
            return apelido
    return clean_fallback(nome_qualquer)


def clean_fallback(s):
    return (s or "").strip() or "—"


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
