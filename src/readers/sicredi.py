"""
Leitor de extrato Sicredi.
Formato: Data | Descrição | Documento | Valor (R$) | Saldo (R$)
"""

from __future__ import annotations
import re
import pandas as pd
from .base import LeitorBase, _limpar_valor

_COLUNAS_SICREDI = ["Data", "Descrição", "Documento", "Valor (R$)", "Saldo (R$)"]
_RE_DATA = re.compile(r"\d{2}/\d{2}/\d{4}")


class LeitorSicredi(LeitorBase):
    BANCO = "Sicredi"

    def _parse_pdf(self) -> pd.DataFrame:
        import pdfplumber
        registros = []
        with pdfplumber.open(self.caminho) as pdf:
            for page in pdf.pages:
                tbl = page.extract_table()
                if not tbl:
                    continue
                for row in tbl:
                    if not row or len(row) < 5:
                        continue
                    data = str(row[0] or "").strip()
                    desc = str(row[1] or "").strip()
                    doc  = str(row[2] or "").strip()
                    val  = str(row[3] or "").strip()
                    sld  = str(row[4] or "").strip()
                    if not _RE_DATA.match(data):
                        continue
                    registros.append({
                        "data": data, "descricao": desc, "documento": doc,
                        "valor": val, "saldo": sld,
                    })
        if not registros:
            raise ValueError(f"Nenhuma transação no PDF Sicredi: {self.caminho}")
        df = pd.DataFrame(registros)
        return self._separar_credito_debito(df)

    def _parse_excel(self) -> pd.DataFrame:
        for hr in range(0, 10):
            try:
                df = pd.read_excel(self.caminho, header=hr)
                df = self._renomear(df)
                if "valor" in df.columns:
                    return self._separar_credito_debito(df)
            except Exception:
                continue
        raise ValueError(f"Excel Sicredi não reconhecido: {self.caminho}")

    def _parse_csv(self) -> pd.DataFrame:
        for sep in (";", ",", "\t"):
            for enc in ("utf-8", "latin-1", "cp1252"):
                for hr in range(0, 10):
                    try:
                        df = pd.read_csv(self.caminho, sep=sep, encoding=enc,
                                         header=hr, dtype=str, on_bad_lines="skip")
                        df = self._renomear(df)
                        if "valor" in df.columns:
                            return self._separar_credito_debito(df)
                    except Exception:
                        continue
        raise ValueError(f"CSV Sicredi não reconhecido: {self.caminho}")

    _MAP = {
        r"^data$":                              "data",
        r"descri[çc][aã]o|hist[oó]rico":        "descricao",
        r"^documento$|^doc\.?$":                "documento",
        r"^valor":                              "valor",
        r"^saldo":                              "saldo",
    }

    def _renomear(self, df: pd.DataFrame) -> pd.DataFrame:
        mapa = {}
        for col in df.columns:
            cl = str(col).strip().lower()
            for p, n in self._MAP.items():
                if re.search(p, cl):
                    mapa[col] = n
                    break
        return df.rename(columns=mapa)

    def _separar_credito_debito(self, df: pd.DataFrame) -> pd.DataFrame:
        df["valor_num"] = df["valor"].apply(_limpar_valor)
        df["credito"] = df["valor_num"].apply(lambda v: v if v > 0 else 0.0)
        df["debito"]  = df["valor_num"].apply(lambda v: abs(v) if v < 0 else 0.0)
        return df.drop(columns=["valor", "valor_num"], errors="ignore")
