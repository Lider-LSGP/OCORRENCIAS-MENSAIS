"""Utilitários compartilhados entre os módulos de processamento."""
from __future__ import annotations

import re
from datetime import datetime, date, timedelta
from typing import Optional, List, Tuple
import pandas as pd


def to_date(v) -> Optional[date]:
    """Converte qualquer representação comum de data (Domínio/EasyApp) em date."""
    if v is None or (isinstance(v, float) and v != v):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    if not s or s.lower() == "nan":
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S"):
        try:
            return datetime.strptime(s[:19] if len(s) > 10 else s, fmt).date()
        except ValueError:
            continue
    return None


def merge_intervalos(intervalos: List[Tuple[date, date]], gap_dias: int = 1) -> List[Tuple[date, date]]:
    """Une intervalos sobrepostos/adjacentes (gap <= gap_dias) em faixas contínuas."""
    ivs = sorted([iv for iv in intervalos if iv[0] and iv[1]], key=lambda x: x[0])
    merged: List[Tuple[date, date]] = []
    for ini, fim in ivs:
        if merged and ini <= merged[-1][1] + timedelta(days=gap_dias):
            if fim > merged[-1][1]:
                merged[-1] = (merged[-1][0], fim)
        else:
            merged.append((ini, fim))
    return merged


def melhor_faixa(merged: List[Tuple[date, date]], ini_esp: date, fim_esp: date,
                 tolerancia_dias: int = 25) -> Optional[Tuple[date, date]]:
    """Escolhe a faixa mesclada mais relevante para comparar com o período esperado:
    prioriza sobreposição; sem sobreposição, aceita a mais próxima dentro da tolerância."""
    melhor, melhor_score = None, None
    for ini, fim in merged:
        overlap = min(fim, fim_esp) - max(ini, ini_esp)
        score = overlap.days if overlap.days > 0 else -min(abs((ini - ini_esp).days), abs((fim - fim_esp).days))
        if melhor_score is None or score > melhor_score:
            melhor, melhor_score = (ini, fim), score
    if melhor is None:
        return None
    if melhor_score < 0 and -melhor_score > tolerancia_dias:
        return None
    return melhor


def detectar_tipo_relatorio(df: pd.DataFrame) -> Optional[str]:
    cols = {str(c).strip().lower() for c in df.columns}
    if "foempregados_nome" in cols or ("valor_inf" in cols and "i_eventos" in cols):
        return "FALTAS"
    if "inicio_gozo" in cols:
        return "FERIAS"
    if "data_real" in cols and "i_afastamentos" in cols:
        return "AFASTAMENTOS"
    if "data_aviso" in cols and "salario" not in cols:
        return "AVISOS"
    if "demissao" in cols or "salario" in cols:
        return "RESCISOES"
    return None


def lancamento_residual_proximo(recs, ini_esp, fim_esp,
                                tolerancia_normal: int = 25,
                                janela_ampla: int = 120):
    """Procura, entre os períodos já lançados (`recs`, lista de dicts com
    ini/fim/tipo), um lançamento que esteja PRÓXIMO do período esperado mas
    FORA da tolerância normal de casamento (não é "o mesmo" lançamento) e
    dentro de uma janela mais ampla (mesmo assunto, época diferente).

    Isso é o sinal de um período ANTERIOR que provavelmente foi CANCELADO na
    Domínio (por isso não aparece mais como o período atual) mas continua
    ATIVO no sistema interno, porque cancelamento na Domínio não cancela
    automaticamente o lançamento já feito no sistema. Não altera nada — só
    sinaliza para conferência manual (cancelar o lançamento antigo ou
    confirmar que ele é válido)."""
    melhor, melhor_dist = None, None
    for r in recs:
        ini, fim = r["ini"], r["fim"]
        overlap = min(fim, fim_esp) - max(ini, ini_esp)
        dist = -overlap.days if overlap.days > 0 else \
            min(abs((ini - ini_esp).days), abs((fim - fim_esp).days))
        if dist <= tolerancia_normal:
            continue  # está dentro da tolerância normal — não é "residual"
        if dist > janela_ampla:
            continue  # longe demais, provavelmente assunto diferente
        if melhor_dist is None or dist < melhor_dist:
            melhor, melhor_dist = r, dist
    return melhor


def clean(s) -> str:
    if s is None or (isinstance(s, float) and s != s):
        return ""
    return re.sub(r"\s+", " ", str(s).replace("\\n", " ").replace("\n", " ")).strip()
