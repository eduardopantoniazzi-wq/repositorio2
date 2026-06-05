# Controle Financeiro Anti-Desvio — Moinho de Trigo

Sistema Python para consolidação de extratos bancários (Bradesco, Sicredi, BB) com detecção automática de desvios financeiros e atualização de Google Sheets.

---

## Instalação rápida

```bash
pip install -r requirements.txt
```

---

## Uso

```bash
# Gerar arquivos de exemplo para teste
python data/exemplos/gerar_exemplos.py

# Rodar sem Google Sheets (modo teste)
python atualizar.py --sem-sheets

# Rodar completo (com Google Sheets)
python atualizar.py
```

---

## Estrutura de pastas

```
atualizar.py                  ← ponto de entrada
config.py                     ← configurações (edite antes de usar)
requirements.txt
credentials/
  google_credentials.json     ← coloque aqui as credenciais da Service Account
data/
  extratos/                   ← coloque aqui os arquivos de extrato (CSV/Excel/PDF)
    bradesco.csv / bradesco.xlsx / bradesco.pdf
    sicredi.csv  / sicredi.xlsx  / sicredi.pdf
    bb.csv       / bb.xlsx       / bb.pdf
  base_conhecimento/
    fornecedores_conhecidos.json   ← gerado automaticamente
    clientes_conhecidos.json       ← gerado automaticamente
  exemplos/
    gerar_exemplos.py
src/
  readers/
    bradesco.py / sicredi.py / bb.py
  consolidator.py
  detector.py
  knowledge_base.py
  sheets.py
  report.py
```

---

## Formatos aceitos por banco

| Banco    | PDF | CSV | Excel |
|----------|-----|-----|-------|
| Bradesco | ✔   | ✔   | ✔     |
| Sicredi  | ✔   | ✔   | ✔     |
| BB       | ✔   | ✔   | ✔     |

O sistema detecta o banco automaticamente pelo nome do arquivo.  
Ex: `bradesco_junho.csv`, `sicredi_2025.pdf`, `bb_canoas.xlsx`

---

## Verificações de desvio

| # | Regra | Nível |
|---|-------|-------|
| 1 | Pagamento > R$ 50.000 sem histórico do beneficiário no mês anterior | CRÍTICO |
| 2 | Mesmo beneficiário com mais de um pagamento no mesmo dia | ATENÇÃO |
| 3 | Beneficiário novo (nunca apareceu nos últimos 3 meses) | ATENÇÃO |
| 4 | Pagamento em horário fora do padrão (antes das 7h ou após 19h) | CRÍTICO |
| 5 | Variação acima de 30% no valor pago a fornecedor recorrente | ATENÇÃO |
| 6 | Crédito recebido de CNPJ não reconhecido na base de clientes | ATENÇÃO |

---

## Configurar Google Sheets

### 1. Criar Service Account no Google Cloud

1. Acesse [console.cloud.google.com](https://console.cloud.google.com)
2. Crie ou selecione um projeto
3. Ative as APIs:
   - Google Sheets API
   - Google Drive API
4. Vá em **IAM e administrador → Contas de serviço → Criar conta de serviço**
5. Baixe o arquivo JSON de credenciais
6. Salve como `credentials/google_credentials.json`

### 2. Compartilhar a planilha com a Service Account

- Copie o e-mail da Service Account (algo como `nome@projeto.iam.gserviceaccount.com`)
- Compartilhe a planilha Google Sheets com esse e-mail (permissão de Editor)
- Copie o ID da planilha da URL e cole em `config.py` → `GOOGLE_SHEET_ID`

### 3. Testar

```bash
python atualizar.py
```

---

## Configurações (config.py)

```python
LIMITE_PAGAMENTO_ALTO      = 50_000.00  # Alerta acima deste valor
LIMITE_VARIACAO_FORNECEDOR = 0.30       # 30% de variação
HORARIO_INICIO_PERMITIDO   = 7          # Horário permitido: das 7h
HORARIO_FIM_PERMITIDO      = 19         # às 19h
JANELA_HISTORICO_MESES     = 3          # Janela de histórico
```

---

## Nomeação dos arquivos de extrato

O sistema detecta o banco pelo nome do arquivo. Certifique-se que o nome contenha:
- `bradesco` para Bradesco
- `sicredi` para Sicredi  
- `bb` ou `bancodobrasil` para Banco do Brasil

Exemplos válidos:
```
bradesco_06_2025.csv
extrato_sicredi_junho.pdf
bb_canoas_20250605.xlsx
```
