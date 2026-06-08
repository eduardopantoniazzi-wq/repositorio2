"""
Leitor de extrato Banrisul.
Formato texto fixo (layout=True):
  DD HISTORICO                DOCUMENTO    VALOR       (valor com - = débito)
     NOME: beneficiário                               (linha opcional)
  SALDO NA DATA                            SALDO
"""

from __future__ import annotations
import re
import pandas as pd
from .base import LeitorBase, _limpar_valor

_RE_DATA_CABEC = re.compile(r"SALDO ANT EM (\d{2}/\d{2}/\d{4})")
_RE_MES_ANO    = re.compile(r"MOVIMENTOS\s+(\w{3})/(\d{4})", re.IGNORECASE)
_RE_LANCAMENTO = re.compile(
    r"^(\d{2})?\s{1,5}(.{1,45?}?)\s{2,}(\d{6,})\s+([\d.]+,\d{2})(-)?\s*$"
)
_RE_SALDO_DIA  = re.compile(r"SALDO NA DATA\s+([\d.]+,\d{2})", re.IGNORECASE)
_RE_NOME       = re.compile(r"^\s+NOME:\s*(.+)$", re.IGNORECASE)
_RE_VALOR_ONLY = re.compile(r"([\d]{1,3}(?:\.\d{3})*,\d{2})(-)?$")
_RE_DIA_INICIO = re.compile(r"^\s{3,4}(\d{2})\s+(\S.*)")


class LeitorBanrisul(LeitorBase):
    BANCO = "Banrisul"

    def _parse_pdf(self) -> pd.DataFrame:
        import pdfplumber
        linhas_todas: list[str] = []
        with pdfplumber.open(self.caminho) as pdf:
            for page in pdf.pages:
                txt = page.extract_text(layout=True) or ""
                linhas_todas.extend(txt.split("\n"))

        registros = self._parsear(linhas_todas)
        if not registros:
            raise ValueError(f"Nenhuma transação extraída do Banrisul: {self.caminho}")
        return pd.DataFrame(registros)

    def _parsear(self, linhas: list[str]) -> list[dict]:
        registros: list[dict] = []
        dia_atual = ""
        mes_atual = ""
        ano_atual = ""
        saldo_corrente = 0.0
        pendente: dict | None = None   # lançamento aguardando possível linha NOME:

        def _salvar(reg):
            if reg:
                registros.append(reg)

        for linha in linhas:
            # ── Detecta mês/ano (ex: "++ MOVIMENTOS MAI/2026") ──────────────
            m_mes = _RE_MES_ANO.search(linha)
            if m_mes:
                meses_pt = {"JAN":"01","FEV":"02","MAR":"03","ABR":"04",
                            "MAI":"05","JUN":"06","JUL":"07","AGO":"08",
                            "SET":"09","OUT":"10","NOV":"11","DEZ":"12"}
                mes_atual = meses_pt.get(m_mes.group(1).upper(), "01")
                ano_atual = m_mes.group(2)
                continue

            # ── Saldo anterior (pega mês/ano da data) ────────────────────────
            m_ant = _RE_DATA_CABEC.search(linha)
            if m_ant:
                partes = m_ant.group(1).split("/")
                if len(partes) == 3:
                    dia_atual = partes[0]
                    if not mes_atual:
                        mes_atual = partes[1]
                    if not ano_atual:
                        ano_atual = partes[2]
                m_v = _RE_VALOR_ONLY.search(linha)
                if m_v:
                    saldo_corrente = _limpar_valor(m_v.group(1))
                continue

            # ── Saldo na data ─────────────────────────────────────────────────
            m_saldo = _RE_SALDO_DIA.search(linha)
            if m_saldo:
                saldo_corrente = _limpar_valor(m_saldo.group(1))
                # Corrige o saldo do último registro já salvo com o valor exato do extrato
                if pendente:
                    pendente["saldo"] = saldo_corrente
                    _salvar(pendente)
                    pendente = None
                elif registros:
                    registros[-1]["saldo"] = saldo_corrente
                continue

            # ── Linha NOME: do beneficiário ───────────────────────────────────
            m_nome = _RE_NOME.match(linha)
            if m_nome and pendente:
                nome = m_nome.group(1).strip()
                pendente["descricao"] = f"{pendente['descricao']} / {nome}".strip(" /")
                continue

            # ── Detecta início de novo dia (ex: "   29 TED ...") ─────────────
            m_dia = _RE_DIA_INICIO.match(linha)
            if m_dia:
                dia_atual = m_dia.group(1)

            # ── Detecta lançamento com valor no final ─────────────────────────
            m_v = _RE_VALOR_ONLY.search(linha.rstrip())
            if not m_v:
                continue

            # Extrai partes da linha
            valor_str = m_v.group(1)
            negativo  = bool(m_v.group(2))  # termina com '-'
            valor     = _limpar_valor(valor_str)

            # Texto antes do valor = histórico + documento
            antes = linha[:m_v.start()].strip()
            # Remove o dia do início se presente
            antes = re.sub(r"^\d{2}\s+", "", antes).strip()

            # Separa documento (últimos 6+ dígitos colados no final do texto)
            doc = ""
            m_doc = re.search(r"\s+(\d{6,})\s*$", antes)
            if m_doc:
                doc = m_doc.group(1)
                antes = antes[:m_doc.start()].strip()

            historico = antes.strip()

            # Ignora linhas de cabeçalho/rodapé
            if any(p in historico.upper() for p in [
                "DIA HISTORICO", "TARIFA EC", "TEB PJ", "BENEFICIO",
                "SALDO ANT", "MOVIMENTOS", "BLOQUEADO", "SALDO DISP",
                "PARA SIMPLES", "PREZADO", "JUROS DE",
            ]):
                continue

            if not historico or not dia_atual or not mes_atual or not ano_atual:
                continue

            data_str = f"{dia_atual}/{mes_atual}/{ano_atual}"

            cred = valor if not negativo else 0.0
            deb  = valor if negativo else 0.0

            # Atualiza saldo (aproximado — o saldo exato vem do "SALDO NA DATA")
            if not negativo:
                saldo_corrente += valor
            else:
                saldo_corrente -= valor

            _salvar(pendente)
            pendente = {
                "data":      data_str,
                "descricao": historico,
                "documento": doc,
                "credito":   cred,
                "debito":    deb,
                "saldo":     round(saldo_corrente, 2),
            }

        _salvar(pendente)
        return registros

    def _parse_excel(self) -> pd.DataFrame:
        raise NotImplementedError("Banrisul Excel não implementado")

    def _parse_csv(self) -> pd.DataFrame:
        raise NotImplementedError("Banrisul CSV não implementado")
