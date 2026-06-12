"""
Leitor do PDF "Consulta Operações / Emite Recibos" do Banrisul.
Estrutura por registro (2 linhas):
  1. [DD/MM/YYYY] NSU EFETUADA R$ VALOR Operação Conta
  2. BENEFICIÁRIO - CODIGOBARRAS  (linha "Complemento")
"""
from __future__ import annotations
import re
import pandas as pd
from pathlib import Path

_RE_NSU = re.compile(
    r"(?:(\d{2}/\d{2}/\d{4})\s+)?"   # data opcional na mesma linha
    r"\d{8,}\s+"                        # NSU
    r"(EFETUADA|PENDENTE|CANCELADA)\s+"
    r"R\$\s*([\d.]+,\d{2})\s+"
    r"(T[ií]tulo|Transfer[eê]ncia|Arrecada[cç][aã]o)",
    re.IGNORECASE,
)
_RE_DATA_ONLY = re.compile(r"^\s*(\d{2}/\d{2}/\d{4})\s*$")
_IGNORAR = {
    "data nsu", "complemento", "banco do estado", "consulta opera",
    "situacao", "efetuados", "total", "sac:", "toda transacao",
    "ouvidoria", "agencia/conta",
}


def _eh_cabecalho(linha: str) -> bool:
    l = linha.strip().lower()
    return not l or any(l.startswith(p) for p in _IGNORAR) or "conta:" in l


def ler_consulta_banrisul(caminho) -> pd.DataFrame:
    """
    Retorna DataFrame com: data, valor, beneficiario, operacao
    Uma linha por operação efetuada.
    """
    import pdfplumber
    linhas: list[str] = []
    with pdfplumber.open(Path(caminho)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            linhas.extend(txt.split("\n"))

    registros: list[dict] = []
    data_atual: str | None = None
    aguardando: dict | None = None

    for linha in linhas:
        l = linha.strip()
        if _eh_cabecalho(l):
            continue

        m_nsu = _RE_NSU.search(l)
        if m_nsu:
            if m_nsu.group(1):
                data_atual = m_nsu.group(1)
            valor_str = m_nsu.group(3).replace(".", "").replace(",", ".")
            aguardando = {
                "data":     data_atual,
                "valor":    float(valor_str),
                "operacao": m_nsu.group(4),
            }
            continue

        m_data = _RE_DATA_ONLY.match(l)
        if m_data:
            data_atual = m_data.group(1)
            continue

        if aguardando and aguardando.get("data"):
            partes = l.split(" - ")
            operacao = aguardando["operacao"].lower()

            if "transfer" in operacao:
                beneficiario = partes[-1].strip() if len(partes) >= 2 else partes[0]
            elif "arrecad" in operacao:
                beneficiario = partes[0].strip()
            else:
                beneficiario = partes[0].strip()

            if beneficiario:
                registros.append({
                    "data":        aguardando["data"],
                    "valor":       aguardando["valor"],
                    "beneficiario": beneficiario,
                    "operacao":    aguardando["operacao"],
                })
            aguardando = None

    if not registros:
        return pd.DataFrame(columns=["data", "valor", "beneficiario", "operacao"])

    df = pd.DataFrame(registros)
    df["data"] = pd.to_datetime(df["data"], dayfirst=True, errors="coerce")
    df = df.dropna(subset=["data"])
    return df
