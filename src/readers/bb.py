"""
Leitor de extrato Banco do Brasil.
Formato: Dt.balancete | Dt.movimento | Ag.origem | Lote | Histórico | Documento | Valor R$ | Saldo
"""

from __future__ import annotations
import re
import pandas as pd
from .base import LeitorBase, _limpar_valor

_RE_DATA    = re.compile(r"\d{2}/\d{2}/\d{4}")
_RE_VAL_CD  = re.compile(r"([\d]{1,3}(?:\.\d{3})*,\d{2})\s*\n?\s*([CD])")
_RE_VAL_NUM = re.compile(r"-?[\d]{1,3}(?:\.\d{3})*,\d{2}")


def _extrair_cd(texto: str):
    t = str(texto or "").replace("\n", " ").strip()
    m = _RE_VAL_CD.search(t)
    if m:
        v = _limpar_valor(m.group(1))
        return (v, 0.0) if m.group(2) == "C" else (0.0, v)
    m2 = _RE_VAL_NUM.search(t)
    if m2:
        v = _limpar_valor(m2.group())
        return (v, 0.0) if v >= 0 else (0.0, abs(v))
    return 0.0, 0.0


class LeitorBB(LeitorBase):
    BANCO = "BB"

    def _parse_pdf(self) -> pd.DataFrame:
        import pdfplumber
        all_rows: list[list] = []
        with pdfplumber.open(self.caminho) as pdf:
            for page in pdf.pages:
                tbl = page.extract_table()
                if not tbl:
                    continue
                start = 0
                for i, row in enumerate(tbl):
                    if row and any("Hist" in str(c or "") or "balancete" in str(c or "").lower() for c in row):
                        start = i + 1
                        break
                all_rows.extend(tbl[start:])

        if not all_rows:
            raise ValueError(f"Nenhuma tabela no PDF BB: {self.caminho}")

        registros = self._reconstruir(all_rows)
        if not registros:
            raise ValueError(f"Nenhuma transação extraída do PDF BB: {self.caminho}")
        return pd.DataFrame(registros)

    def _reconstruir(self, rows: list[list]) -> list[dict]:
        registros: list[dict] = []
        i = 0
        while i < len(rows):
            row = list(rows[i])
            while len(row) < 8:
                row.append(None)

            dt_bal   = str(row[0] or "").strip()
            dt_mov   = str(row[1] or "").strip()
            historico = str(row[4] or "").strip()
            doc      = str(row[5] or "").strip()
            val_raw  = str(row[6] or "").strip()
            sld_raw  = str(row[7] or "").strip()

            tem_data = _RE_DATA.search(dt_bal) or _RE_DATA.search(dt_mov)

            if not tem_data and not _RE_VAL_CD.search(val_raw) and not historico:
                i += 1
                continue

            data = ""
            m = _RE_DATA.search(dt_mov) or _RE_DATA.search(dt_bal)
            if m:
                data = m.group()

            hist_limpo = re.sub(r"^\d{3,5}\s+", "", historico).strip()
            hist_limpo = re.sub(r"^\d{3,5}\s+", "", hist_limpo).strip()

            _INVEST_BB = ("rende facil", "rende fácil", "bb rende", "aplic auto",
                          "aplicacao automatica", "aplicação automática",
                          "resgate automatico", "resgate automático")
            if any(p in hist_limpo.lower() for p in _INVEST_BB):
                i += 1
                continue

            beneficiario = ""
            if i + 1 < len(rows):
                prox = list(rows[i + 1])
                while len(prox) < 8:
                    prox.append(None)
                prox_dt  = str(prox[0] or "").strip()
                prox_hist = str(prox[4] or "").strip()
                prox_val  = str(prox[6] or "").strip()
                if (not _RE_DATA.search(prox_dt)
                        and prox_hist
                        and not _RE_VAL_CD.search(prox_val)
                        and not _RE_VAL_NUM.search(prox_val)):
                    beneficiario = prox_hist
                    i += 1

            descricao = f"{hist_limpo} / {beneficiario}".strip(" /") if beneficiario else hist_limpo

            cred, deb = _extrair_cd(val_raw)
            m_sld = _RE_VAL_NUM.search(sld_raw)
            saldo = _limpar_valor(m_sld.group()) if m_sld else None

            if re.search(r"\bS\s*A\s*L\s*D\s*O\b", hist_limpo, re.IGNORECASE):
                saldo_final = cred if cred > 0 else (deb if deb > 0 else (saldo or 0))
                if saldo_final and saldo_final > 0:
                    if registros:
                        registros[-1]["saldo"] = saldo_final
                    registros.append({
                        "data":      data or (registros[-1]["data"] if registros else "01/01/2000"),
                        "descricao": "SALDO FINAL",
                        "documento": "",
                        "credito":   0.0,
                        "debito":    0.0,
                        "saldo":     saldo_final,
                    })
                i += 1
                continue

            if data:
                registros.append({
                    "data": data, "descricao": descricao,
                    "documento": doc, "credito": cred, "debito": deb, "saldo": saldo,
                })
            i += 1
        return registros

    def _parse_excel(self) -> pd.DataFrame:
        for hr in range(0, 10):
            try:
                df = pd.read_excel(self.caminho, header=hr)
                return self._normalizar_tabular(df)
            except Exception:
                continue
        raise ValueError(f"Excel BB não reconhecido: {self.caminho}")

    def _parse_csv(self) -> pd.DataFrame:
        for sep in (";", ",", "\t"):
            for enc in ("utf-8", "latin-1", "cp1252"):
                for hr in range(0, 10):
                    try:
                        df = pd.read_csv(self.caminho, sep=sep, encoding=enc,
                                         header=hr, dtype=str, on_bad_lines="skip")
                        return self._normalizar_tabular(df)
                    except Exception:
                        continue
        raise ValueError(f"CSV BB não reconhecido: {self.caminho}")

    _MAP = {
        r"dt\.?movimento|data mov": "data",
        r"hist[oó]rico":            "descricao",
        r"documento|doc":           "documento",
        r"valor":                   "valor",
        r"saldo":                   "saldo",
    }

    def _normalizar_tabular(self, df: pd.DataFrame) -> pd.DataFrame:
        mapa = {}
        for col in df.columns:
            cl = str(col).strip().lower()
            for p, n in self._MAP.items():
                if re.search(p, cl):
                    mapa[col] = n
                    break
        df = df.rename(columns=mapa)
        if "valor" in df.columns:
            df["valor_num"] = df["valor"].apply(_limpar_valor)
            df["credito"] = df["valor_num"].apply(lambda v: v if v > 0 else 0.0)
            df["debito"]  = df["valor_num"].apply(lambda v: abs(v) if v < 0 else 0.0)
            df = df.drop(columns=["valor", "valor_num"], errors="ignore")
        return df
