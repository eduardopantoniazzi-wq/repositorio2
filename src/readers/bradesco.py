"""
Leitor de extrato Bradesco.
Formato PDF Net Empresa: Data | Lançamento | Dcto | Crédito | Débito | Saldo
Formato CSV/Excel exportado manualmente.
"""

from __future__ import annotations
import re
from pathlib import Path

import pandas as pd

from .base import LeitorBase, _ler_pdf_tabelas, _limpar_valor


class LeitorBradesco(LeitorBase):
    BANCO = "Bradesco"

    # Colunas esperadas no PDF (variações de nome aceitas)
    _MAP_COLUNAS = {
        r"data":        "data",
        r"lan[çc]amento|hist[oó]rico|descri[çc][aã]o": "descricao",
        r"dcto|documento|doc\.?": "documento",
        r"cr[eé]dito":  "credito",
        r"d[eé]bito":   "debito",
        r"saldo":       "saldo",
    }

    def _parse_pdf(self) -> pd.DataFrame:
        tabelas = _ler_pdf_tabelas(self.caminho)
        frames = []
        for tbl in tabelas:
            tbl = self._renomear_colunas(tbl)
            if "data" in tbl.columns and "descricao" in tbl.columns:
                frames.append(tbl)
        if not frames:
            raise ValueError(f"Nenhuma tabela reconhecida no PDF Bradesco: {self.caminho}")
        df = pd.concat(frames, ignore_index=True)
        return df

    def _parse_excel(self) -> pd.DataFrame:
        # Tenta várias linhas de cabeçalho
        for header_row in range(0, 10):
            try:
                df = pd.read_excel(self.caminho, header=header_row)
                df = self._renomear_colunas(df)
                if "data" in df.columns and "descricao" in df.columns:
                    return df
            except Exception:
                continue
        raise ValueError(f"Não foi possível ler o Excel Bradesco: {self.caminho}")

    def _parse_csv(self) -> pd.DataFrame:
        for sep in (";", ",", "\t"):
            for enc in ("utf-8", "latin-1", "cp1252"):
                for header_row in range(0, 10):
                    try:
                        df = pd.read_csv(
                            self.caminho, sep=sep, encoding=enc,
                            header=header_row, dtype=str, on_bad_lines="skip"
                        )
                        df = self._renomear_colunas(df)
                        if "data" in df.columns and "descricao" in df.columns:
                            return df
                    except Exception:
                        continue
        raise ValueError(f"Não foi possível ler o CSV Bradesco: {self.caminho}")

    def _renomear_colunas(self, df: pd.DataFrame) -> pd.DataFrame:
        mapa = {}
        for col in df.columns:
            col_lower = str(col).strip().lower()
            for padrao, nome_padrao in self._MAP_COLUNAS.items():
                if re.search(padrao, col_lower):
                    mapa[col] = nome_padrao
                    break
        return df.rename(columns=mapa)
