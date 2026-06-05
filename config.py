"""
Configurações centrais do sistema anti-desvio.
Edite este arquivo antes de rodar pela primeira vez.
"""

from pathlib import Path

# ── Diretórios ─────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
DATA_DIR = BASE_DIR / "data"
EXTRATOS_DIR = DATA_DIR / "extratos"
BASE_CONHECIMENTO_DIR = DATA_DIR / "base_conhecimento"
CREDENTIALS_DIR = BASE_DIR / "credentials"

# ── Google Sheets ──────────────────────────────────────────────────────────
# Coloque o arquivo de credenciais da Service Account aqui:
GOOGLE_CREDENTIALS_FILE = CREDENTIALS_DIR / "google_credentials.json"

# ID da planilha (a parte entre /d/ e /edit na URL do Google Sheets).
# Deixe vazio para criar uma nova planilha automaticamente.
GOOGLE_SHEET_ID = "11-x3mDnzb7f_rGI0FgYI_mCdI3JHsYDHJAriLzSS7M4"

# Nome da planilha (usado ao criar nova)
GOOGLE_SHEET_NAME = "Controle Financeiro - Moinho"

# ── Nomes das abas ─────────────────────────────────────────────────────────
ABA_DASHBOARD = "Dashboard CEO"
ABA_EXTRATO   = "Extrato Consolidado"
ABA_HISTORICO = "Histórico de Alertas"

# ── Regras de desvio ───────────────────────────────────────────────────────
LIMITE_PAGAMENTO_ALTO      = 50_000.00   # R$ — alerta se pagamento acima deste valor sem histórico
LIMITE_VARIACAO_FORNECEDOR = 0.30        # 30 % de variação sobre a média histórica
HORARIO_INICIO_PERMITIDO   = 7           # hora (0-23)
HORARIO_FIM_PERMITIDO      = 19          # hora (0-23)
JANELA_HISTORICO_MESES     = 3           # meses para considerar fornecedor "conhecido"

# ── Unidades contábeis ─────────────────────────────────────────────────────
UNIDADES = ["Santa Maria", "Canoas"]

# ── Mapeamento banco → unidade (ajuste conforme suas contas) ───────────────
# Se uma conta pertence a uma unidade específica, mapeie aqui.
# Deixe em branco para tratar tudo como empresa única.
BANCO_UNIDADE = {
    "Bradesco": "Santa Maria",
    "Sicredi":  "Santa Maria",
    "BB":       "Canoas",
}
