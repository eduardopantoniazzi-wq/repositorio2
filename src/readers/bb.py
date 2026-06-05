"""
Leitor de extrato Banco do Brasil.
Formato PDF: Dt.balancete | Dt.movimento | Histórico | Documento | Valor | Saldo
Valor positivo = crédito, negativo = débito.
"""

from __future__ import annotations
import re

import pandas as pd

from .base import LeitorBase, _ler_pdf_tabelas, _limpar_valor


class LeitorBB(LeitorBase):
    BANCO = "BB"

    _MAP_COLUNAS = {
        r"dt\.?movimento|data mov":       "data",
        r"dt\.?balancete|data balanc":    "data_balancete",
        r"hist[oó]rico|descri[çc][aã]o": "descricao",
        r"documento|doc\.?|n[uú]m":       "documento",
        r"valor":                          "valor",
        r"saldo":                          "saldo",
    }

    def _parse_pdf(self) -> pd.DataFrame:
        tabelas = _ler_pdf_tabelas(self.caminho)
        frames = []
        for tbl in tabelas:
            tbl = self._renomear_colunas(tbl)
            if "data" in tbl.columns and "valor" in tbl.columns:
                frames.append(tbl)
        if not frames:
            raise ValueError(f"Nenhuma tabela reconhecida no PDF BB: {self.caminho}")
        df = pd.concat(frames, ignore_index=True)
        df = df.drop(columns=["data_balancete"], errors="ignore")
        return self._separar_credito_debito(df)

    def _parse_excel(self) -> pd.DataFrame:
        for header_row in range(0, 10):
            try:
                df = pd.read_excel(self.caminho, header=header_row)
                df = self._renomear_colunas(df)
                if "data" in df.columns and "valor" in df.columns:
                    df = df.drop(columns=["data_balancete"], errors="ignore")
                    return self._separar_credito_debito(df)
            except Exception:
                continue
        raise ValueError(f"Não foi possível ler o Excel BB: {self.caminho}")

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
                        if "data" in df.columns and "valor" in df.columns:
                            df = df.drop(columns=["data_balancete"], errors="ignore")
                            return self._separar_credito_debito(df)
                    except Exception:
                        continue
        raise ValueError(f"Não foi possível ler o CSV BB: {self.caminho}")

    def _renomear_colunas(self, df: pd.DataFrame) -> pd.DataFrame:
        mapa = {}
        for col in df.columns:
            col_lower = str(col).strip().lower()
            for padrao, nome_padrao in self._MAP_COLUNAS.items():
                if re.search(padrao, col_lower):
                    mapa[col] = nome_padrao
                    break
        return df.rename(columns=mapa)

    def _separar_credito_debito(self, df: pd.DataFrame) -> pd.DataFrame:
        df["valor_num"] = df["valor"].apply(_limpar_valor)
        df["credito"] = df["valor_num"].apply(lambda v: v if v > 0 else 0.0)
        df["debito"]  = df["valor_num"].apply(lambda v: abs(v) if v < 0 else 0.0)
        df = df.drop(columns=["valor", "valor_num"], errors="ignore")
        return df
