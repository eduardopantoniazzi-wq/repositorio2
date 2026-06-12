"""
Leitor da planilha de saídas do sistema interno (formato .xlsm / .xlsx).

Layout detectado:
  Linha 3: cabeçalho — Data Pag. | Vencimento | Atr. | Título | Tp.Doc | C l i e n t e | Valor Nom. | Juros | Desconto | Desp.Cobr. | Vlr.Total
  Linha 4+: dados

Mapeamento:
  data      <- "Data Pag."  (col 0)
  descricao <- "C l i e n t e" (col 5), sem prefixo "123456 - "
  debito    <- "Vlr.Total" (col 10)
  documento <- "Título" (col 3)
"""
from __future__ import annotations
import re
from pathlib import Path
from typing import Union
import openpyxl
import pandas as pd

_RE_PREFIXO = re.compile(r"^\d+\s*-\s*")
_RODAPE = {"total do dia", "t o t a l g e r a l", "total geral"}

def _limpar_nome(v):
    s = str(v).strip() if v else ""
    return _RE_PREFIXO.sub("", s).strip()

def _limpar_valor(v):
    if v is None: return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace("R$","").replace("\xa0","").replace(" ","")
    s = re.sub(r"[^\d,.\-]","",s)
    if "," in s and "." not in s: s = s.replace(",",".")
    elif "." in s and "," in s: s = s.replace(".","").replace(",",".")
    try: return float(s)
    except: return 0.0

def ler_planilha_sistema(caminho: Union[str, Path]) -> pd.DataFrame:
    caminho = Path(caminho)
    if caminho.suffix.lower() in (".xlsm", ".xlsx", ".xls"):
        return _ler_excel(caminho)
    elif caminho.suffix.lower() == ".csv":
        return _ler_csv(caminho)
    raise ValueError(f"Formato não suportado: {caminho.suffix}")

def _ler_excel(caminho: Path) -> pd.DataFrame:
    wb = openpyxl.load_workbook(caminho, data_only=True)
    ws = wb.active
    header_row = None
    col_data = col_cliente = col_valor = col_doc = None
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        rs = [str(v).strip().lower() if v else "" for v in row]
        if any("data pag" in c or c=="data" for c in rs) and any("cliente" in c for c in rs):
            header_row = i
            for j, c in enumerate(rs):
                if "data pag" in c: col_data = j
                elif "cliente" in c: col_cliente = j
                elif "vlr.total" in c or "vlr total" in c or "valor total" in c: col_valor = j
                elif "valor nom" in c and col_valor is None: col_valor = j
                elif "título" in c or "titulo" in c: col_doc = j
            break
    if header_row is None:
        header_row = 2; col_data, col_cliente, col_valor, col_doc = 0, 5, 10, 3
    registros = []
    for row in ws.iter_rows(min_row=header_row+2, values_only=True):
        cr = row[col_cliente] if col_cliente < len(row) else None
        if cr is None: continue
        if any(r in str(cr).strip().lower() for r in _RODAPE): continue
        dv = row[col_data] if col_data < len(row) else None
        if dv is None: continue
        val = _limpar_valor(row[col_valor] if col_valor < len(row) else None)
        if val == 0.0: continue
        doc = str(row[col_doc]).strip() if col_doc and col_doc < len(row) and row[col_doc] else ""
        registros.append({"data":dv, "descricao":_limpar_nome(cr), "debito":val,
                           "credito":0.0, "saldo":0.0, "banco":"Sistema", "horario":None, "documento":doc})
    if not registros:
        raise ValueError("Nenhum lançamento encontrado na planilha do sistema.")
    df = pd.DataFrame(registros)
    df["data"] = pd.to_datetime(df["data"], errors="coerce")
    df = df.dropna(subset=["data"])
    df["descricao"] = df["descricao"].astype(str).str.strip()
    return df[["data","banco","descricao","credito","debito","saldo","horario","documento"]].reset_index(drop=True)

def _ler_csv(caminho: Path) -> pd.DataFrame:
    for sep in (";",",","\t"):
        for enc in ("utf-8","latin-1","cp1252"):
            try:
                df = pd.read_csv(caminho, sep=sep, encoding=enc, dtype=str, on_bad_lines="skip")
                mapa = {}
                for col in df.columns:
                    cl = col.strip().lower()
                    if "data pag" in cl: mapa[col]="data"
                    elif "cliente" in cl: mapa[col]="descricao"
                    elif "vlr.total" in cl or "valor total" in cl: mapa[col]="valor"
                    elif "valor nom" in cl and "valor" not in mapa.values(): mapa[col]="valor"
                    elif "titulo" in cl: mapa[col]="documento"
                df = df.rename(columns=mapa)
                if "data" in df.columns and "descricao" in df.columns:
                    df["debito"] = df.get("valor", pd.Series([0.0]*len(df))).apply(_limpar_valor)
                    df["credito"]=0.0; df["saldo"]=0.0; df["banco"]="Sistema"; df["horario"]=None
                    if "documento" not in df.columns: df["documento"]=""
                    df["descricao"] = df["descricao"].apply(_limpar_nome)
                    df["data"] = pd.to_datetime(df["data"], dayfirst=True, errors="coerce")
                    df = df.dropna(subset=["data"])[df["debito"]>0]
                    cols=["data","banco","descricao","credito","debito","saldo","horario","documento"]
                    for c in cols:
                        if c not in df.columns: df[c]=""
                    return df[cols].reset_index(drop=True)
            except: continue
    raise ValueError(f"Não foi possível ler CSV: {caminho}")
