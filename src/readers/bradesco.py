"""
Leitor de extrato Bradesco Net Empresa.
Formato: Data | Lançamento (2 linhas) | Dcto. | Crédito (R$) | Débito (R$) | Saldo (R$)

O PDF tem layout em colunas onde cada transação ocupa 2 linhas visuais:
  linha A (sem data): tipo do lançamento
  linha B (com data): dcto + crédito/débito + saldo
  linha C (sem data): beneficiário
"""

from __future__ import annotations
import re
from pathlib import Path

import pandas as pd

from .base import LeitorBase, _limpar_valor, COLUNAS_PADRAO

# Regex para valor monetário: 1.234,56 ou -1.234,56
_RE_VALOR = re.compile(r"-?[\d]{1,3}(?:\.\d{3})*,\d{2}")
_RE_DATA  = re.compile(r"\b(\d{2}/\d{2}/\d{4})\b")


class LeitorBradesco(LeitorBase):
    BANCO = "Bradesco"

    def _parse_pdf(self) -> pd.DataFrame:
        import pdfplumber
        linhas = []
        with pdfplumber.open(self.caminho) as pdf:
            for page in pdf.pages:
                txt = page.extract_text(layout=True)
                if txt:
                    linhas.extend(txt.splitlines())
        return self._parsear_linhas(linhas)

    def _parse_excel(self) -> pd.DataFrame:
        for hr in range(0, 10):
            try:
                df = pd.read_excel(self.caminho, header=hr)
                df = self._renomear_colunas_tabular(df)
                if {"data", "descricao", "saldo"}.issubset(df.columns):
                    return df
            except Exception:
                continue
        raise ValueError(f"Excel Bradesco não reconhecido: {self.caminho}")

    def _parse_csv(self) -> pd.DataFrame:
        for sep in (";", ",", "\t"):
            for enc in ("utf-8", "latin-1", "cp1252"):
                for hr in range(0, 10):
                    try:
                        df = pd.read_csv(self.caminho, sep=sep, encoding=enc,
                                         header=hr, dtype=str, on_bad_lines="skip")
                        df = self._renomear_colunas_tabular(df)
                        if {"data", "descricao", "saldo"}.issubset(df.columns):
                            return df
                    except Exception:
                        continue
        raise ValueError(f"CSV Bradesco não reconhecido: {self.caminho}")

    # ── Parser principal para o layout de texto do PDF ──────────────────────
    def _parsear_linhas(self, linhas: list[str]) -> pd.DataFrame:
        """
        Reconstrói as transações a partir do texto com layout espacial.
        Cada transação tem saldo no final da linha — usamos isso como âncora.
        """
        registros: list[dict] = []
        data_atual = ""
        desc_acumulada: list[str] = []

        # Detecta se linha tem valores financeiros (pelo menos saldo)
        def tem_valores(linha: str) -> bool:
            return len(_RE_VALOR.findall(linha)) >= 1

        def extrair_valores(linha: str):
            """Extrai crédito, débito e saldo de uma linha."""
            vals = _RE_VALOR.findall(linha)
            if not vals:
                return None, None, None
            saldo = vals[-1]
            # Credito/debito: se tiver 3 valores → cred, deb, saldo
            # Se tiver 2 → (cred ou deb), saldo
            # Se tiver 1 → saldo anterior
            if len(vals) >= 3:
                return vals[-3], vals[-2], saldo
            elif len(vals) == 2:
                # Distingue crédito de débito pelo sinal
                v = vals[-2]
                if v.startswith("-"):
                    return None, v, saldo
                else:
                    return v, None, saldo
            return None, None, saldo

        def extrair_dcto(linha: str) -> str:
            """Extrai o número de documento (série de dígitos no meio da linha)."""
            # Remove valores monetários e datas para encontrar o doc
            sem_vals = _RE_VALOR.sub("", linha)
            sem_data = _RE_DATA.sub("", sem_vals)
            # Procura sequência de dígitos isolada (número de documento)
            docs = re.findall(r"\b\d{5,}\b", sem_data)
            return docs[0] if docs else ""

        # Linhas de controle que não são transações
        _IGNORAR = {
            "data", "lançamento", "dcto", "crédito", "débito", "saldo",
            "total", "saldos invest", "histórico", "folha", "os dados",
            "extrato", "agência", "conta", "nome do usuário", "bradesco",
            "net empresa", "data da operação", "últimos lançamentos",
        }

        # Palavras que indicam seção de investimentos — ignorar tudo após detectar
        _SECAO_INVEST = {"saldos invest", "invest fácil", "invest facil",
                         "saldo invest", "aplicação", "cdb", "lci", "lca"}

        def eh_controle(linha: str) -> bool:
            l = linha.strip().lower()
            if not l:
                return True
            for ign in _IGNORAR:
                if l.startswith(ign):
                    return True
            return False

        def eh_secao_invest(linha: str) -> bool:
            l = linha.strip().lower()
            return any(p in l for p in _SECAO_INVEST)

        i = 0
        secao_invest = False   # flag: entrou na seção de investimentos
        while i < len(linhas):
            linha = linhas[i]
            linha_strip = linha.strip()

            # Detecta início da seção "Saldos Invest Fácil / Plus" e para de processar transações
            if eh_secao_invest(linha_strip):
                secao_invest = True

            if secao_invest:
                i += 1
                continue

            if eh_controle(linha) and not _RE_DATA.search(linha):
                i += 1
                continue

            m_data = _RE_DATA.search(linha)

            if tem_valores(linha):
                if m_data:
                    data_atual = m_data.group(1)

                cred, deb, saldo = extrair_valores(linha)
                dcto = extrair_dcto(linha)

                # Descrição: acumula desc_acumulada + texto da linha sem nums/data
                sem_nums = _RE_VALOR.sub("", linha)
                sem_data = _RE_DATA.sub("", sem_nums)
                desc_linha = re.sub(r"\b\d{5,}\b", "", sem_data).strip()

                desc_partes = desc_acumulada.copy()
                if desc_linha:
                    desc_partes.append(desc_linha)
                descricao = " / ".join(p.strip() for p in desc_partes if p.strip())

                # Pega linha seguinte como 2ª parte da descrição (beneficiário)
                if i + 1 < len(linhas):
                    prox = linhas[i + 1].strip()
                    if prox and not tem_valores(prox) and not eh_controle(prox) and not _RE_DATA.match(prox):
                        if descricao:
                            descricao = descricao + " / " + prox
                        else:
                            descricao = prox
                        i += 1

                if data_atual and saldo:
                    registros.append({
                        "data":      data_atual,
                        "descricao": descricao or "",
                        "documento": dcto,
                        "credito":   cred or "",
                        "debito":    deb or "",
                        "saldo":     saldo,
                    })
                desc_acumulada = []
            else:
                # Linha só com texto → acumula descrição
                if linha_strip and not eh_controle(linha):
                    if m_data:
                        data_atual = m_data.group(1)
                        # descrição da mesma linha (ex: "SALDO ANTERIOR")
                        desc = _RE_DATA.sub("", linha_strip).strip()
                        if desc:
                            desc_acumulada = [desc]
                    else:
                        desc_acumulada.append(linha_strip)
                else:
                    if not linha_strip:
                        desc_acumulada = []
            i += 1

        if not registros:
            raise ValueError(f"Nenhuma transação extraída do PDF Bradesco: {self.caminho}")

        df = pd.DataFrame(registros)
        return df

    # ── Renomear colunas para CSV/Excel tabulares ───────────────────────────
    _MAP = {
        r"data":        "data",
        r"lan[çc]amento|hist|descri": "descricao",
        r"dcto|doc":    "documento",
        r"cr[eé]dito":  "credito",
        r"d[eé]bito":   "debito",
        r"saldo":       "saldo",
    }

    def _renomear_colunas_tabular(self, df: pd.DataFrame) -> pd.DataFrame:
        mapa = {}
        for col in df.columns:
            cl = str(col).strip().lower()
            for padrao, nome in self._MAP.items():
                if re.search(padrao, cl):
                    mapa[col] = nome
                    break
        return df.rename(columns=mapa)
