"""
Classe base para leitura de extratos bancários.
"""

from __future__ import annotations
import re
from pathlib import Path
from typing import Union

import pandas as pd


COLUNAS_PADRAO = ["data", "banco", "descricao", "credito", "debito", "saldo", "horario", "documento"]


def _limpar_valor(v) -> float:
    """Converte '1.234,56' ou '-1.234,56' para float."""
    if v is None or (isinstance(v, float) and v != v):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("\xa0", "").replace(" ", "")
    s = re.sub(r"[R$]", "", s)
    if re.search(r"\d\.\d{3},", s):
        s = s.replace(".", "").replace(",", ".")
    elif "," in s and "." not in s:
        s = s.replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "")
    s = re.sub(r"[^\d.\-]", "", s)
    try:
        return float(s)
    except ValueError:
        return 0.0


class LeitorBase:
    BANCO: str = ""

    def __init__(self, caminho: Union[str, Path]):
        self.caminho = Path(caminho)
        self.sufixo = self.caminho.suffix.lower()

    def ler(self) -> pd.DataFrame:
        if self.sufixo == ".pdf":
            df = self._parse_pdf()
        elif self.sufixo in (".xlsx", ".xls"):
            df = self._parse_excel()
        elif self.sufixo == ".csv":
            df = self._parse_csv()
        else:
            raise ValueError(f"Formato não suportado: {self.sufixo}")
        return self._normalizar(df)

    def _parse_pdf(self) -> pd.DataFrame:
        raise NotImplementedError

    def _parse_excel(self) -> pd.DataFrame:
        raise NotImplementedError

    def _parse_csv(self) -> pd.DataFrame:
        raise NotImplementedError

    def _normalizar(self, df: pd.DataFrame) -> pd.DataFrame:
        for col in COLUNAS_PADRAO:
            if col not in df.columns:
                df[col] = None if col in ("horario", "documento") else (0.0 if col in ("credito", "debito", "saldo") else "")

        df["banco"]   = self.BANCO
        df["credito"] = df["credito"].apply(_limpar_valor)
        df["debito"]  = df["debito"].apply(_limpar_valor)
        df["saldo"] = df["saldo"].apply(lambda v: None if v is None else _limpar_valor(v))
        df["saldo"] = pd.to_numeric(df["saldo"], errors="coerce").ffill().fillna(0.0)
        df["data"]    = pd.to_datetime(df["data"], dayfirst=True, errors="coerce")
        df["descricao"] = df["descricao"].astype(str).str.strip()

        df = df.dropna(subset=["data"])
        df = df[~((df["credito"] == 0) & (df["debito"] == 0) & (df["saldo"] == 0))]
        return df[COLUNAS_PADRAO].reset_index(drop=True)
