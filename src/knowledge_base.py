"""
Base de conhecimento: fornecedores e clientes conhecidos.
Cresce automaticamente a cada execução.
"""

from __future__ import annotations
import json
from datetime import datetime
from pathlib import Path

import pandas as pd

from config import BASE_CONHECIMENTO_DIR


_ARQ_FORNECEDORES = BASE_CONHECIMENTO_DIR / "fornecedores_conhecidos.json"
_ARQ_CLIENTES     = BASE_CONHECIMENTO_DIR / "clientes_conhecidos.json"


def _carregar(caminho: Path) -> dict:
    if caminho.exists():
        return json.loads(caminho.read_text(encoding="utf-8"))
    return {}


def _salvar(caminho: Path, dados: dict) -> None:
    caminho.parent.mkdir(parents=True, exist_ok=True)
    caminho.write_text(json.dumps(dados, ensure_ascii=False, indent=2), encoding="utf-8")


def carregar_fornecedores() -> dict:
    return _carregar(_ARQ_FORNECEDORES)


def carregar_clientes() -> dict:
    return _carregar(_ARQ_CLIENTES)


def atualizar_base(df: pd.DataFrame) -> None:
    """
    Analisa o DataFrame consolidado e atualiza fornecedores/clientes conhecidos.
    Fornecedores = quem recebeu débitos.
    Clientes     = quem gerou créditos.
    """
    mes_ano = datetime.now().strftime("%Y-%m")

    fornecedores = carregar_fornecedores()
    clientes     = carregar_clientes()

    # Débitos → fornecedores
    debitos = df[df["debito"] > 0][["descricao", "debito", "data"]].copy()
    for _, row in debitos.iterrows():
        nome = str(row["descricao"]).strip()
        if not nome:
            continue
        if nome not in fornecedores:
            fornecedores[nome] = {"primeira_vez": str(row["data"])[:10], "meses": [], "valores": []}
        if mes_ano not in fornecedores[nome]["meses"]:
            fornecedores[nome]["meses"].append(mes_ano)
        fornecedores[nome]["valores"].append(round(row["debito"], 2))

    # Créditos → clientes
    creditos = df[df["credito"] > 0][["descricao", "credito", "data"]].copy()
    for _, row in creditos.iterrows():
        nome = str(row["descricao"]).strip()
        if not nome:
            continue
        if nome not in clientes:
            clientes[nome] = {"primeira_vez": str(row["data"])[:10], "meses": [], "valores": []}
        if mes_ano not in clientes[nome]["meses"]:
            clientes[nome]["meses"].append(mes_ano)
        clientes[nome]["valores"].append(round(row["credito"], 2))

    _salvar(_ARQ_FORNECEDORES, fornecedores)
    _salvar(_ARQ_CLIENTES, clientes)
    print(f"  ✔ Base atualizada: {len(fornecedores)} fornecedores, {len(clientes)} clientes")
