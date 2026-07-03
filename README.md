# Financial Statement Analyzer

CLI tool that computes a full set of financial ratios — profitability, liquidity, leverage, efficiency and cash flow — from a company's income statement, balance sheet and cash flow data.

Built from my day-to-day work in **Valuation & Transaction Advisory**: these are the first-pass indicators I look at when assessing a target's financial health in a transaction context.

## Features

- **Excel template generator** — creates a formatted input sheet (`template_financeiro.xlsx`) with section headers and highlighted input cells, so anyone can fill it without touching code
- **Three input modes** — Excel template, CSV (`campo;valor`), or interactive manual entry
- **Robust to real-world data** — protected division handles zero revenue, zero interest expense, negative equity and missing fields, printing `n/d` (not available) instead of crashing
- **20+ ratios across five dimensions** (see below)

## Ratios computed

| Category | Ratios |
|---|---|
| Profitability | Gross margin, EBIT margin, EBITDA margin, net margin, ROA, ROE |
| Liquidity | Current ratio, quick ratio, cash ratio |
| Leverage | Debt/Equity, interest coverage (EBIT/interest), Net Debt/EBITDA |
| Efficiency | DSO, DIO, DPO (360-day basis), cash conversion cycle |
| Cash flow | Free cash flow (OCF − CapEx), earnings quality (OCF/Net income) |

## How to run

```bash
pip install pandas openpyxl
python analisador.py
```

```
Analisador de Demonstrações Financeiras
=========================================

Como deseja inserir os dados?
  0 - Gerar template Excel
  1 - Arquivo CSV
  2 - Arquivo Excel
  3 - Digitação manual
```

Typical flow: run option `0` to generate the template → fill column B → run option `2` and point to the file.

## Sample output

```
=== RENTABILIDADE ===
  Margem Bruta:    40.0%
  Margem EBIT:     20.0%
  Margem EBITDA:   25.0%
  Margem Líquida:  12.0%
  ROA:             13.3%
  ROE:             30.0%

=== ENDIVIDAMENTO ===
  Dívida / PL:           0.75x
  Cobertura de Juros:    6.67x
  Dívida Líq. / EBITDA:  0.88x

=== EFICIÊNCIA ===
  PMR (recebimento):  54.0 dias
  PME (estoque):      60.0 dias
  PMP (pagamento):    54.0 dias
  Ciclo Financeiro:   60.0 dias
```

## Design notes

- `safe_div()` wraps every ratio: financial statements from real engagements routinely contain zeros, negative equity and missing lines — a screener that crashes on them is useless
- EBITDA is derived as `EBIT + D&A` rather than requested as an input, avoiding a common source of inconsistency in filled templates
- Field labels are defined once (`CAMPOS`) and reused by the template generator and manual-entry mode, so the input schema can't drift

## Roadmap

- [ ] Multi-period input for trend analysis (YoY margins, working capital evolution)
- [ ] Export results to Excel report
- [ ] Sector benchmark comparison

## About

Built by [Luca Rivitti](https://www.linkedin.com/) — Valuation & Transaction Advisory @ Grant Thornton Brasil. Part of a series of projects translating transaction advisory workflows into Python. Next up: a **financial model integrity validator** (automated consistency checks for projection models).
