<div align="center">

# Projeto 3 — Automação de Relatório Contábil/Financeiro

**Rede Máxima Supermercados Ltda. · Pacote de Fechamento Mensal**

[![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white)](https://python.org)
[![openpyxl](https://img.shields.io/badge/openpyxl-217346?style=flat-square&logo=microsoftexcel&logoColor=white)](https://openpyxl.readthedocs.io)
[![ReportLab](https://img.shields.io/badge/ReportLab-E74C3C?style=flat-square)](https://reportlab.com)
[![schedule](https://img.shields.io/badge/schedule-4A90D9?style=flat-square)](https://schedule.readthedocs.io)

</div>

---

## Sobre o projeto

Automação completa do **pacote de fechamento mensal** para uma rede fictícia de supermercados (5 lojas regionais), cobrindo 24 meses de dados simulados (Jan/2023 – Dez/2024).

O script recebe o mês de competência como argumento e, em segundos, entrega dois artefatos prontos para distribuição executiva:

- **Excel (.xlsx)** — workbook com 5 abas: Resumo Executivo, DRE por Loja, Ranking, Aging de Recebíveis e Dashboard Visual com gráficos embutidos
- **PDF** — relatório de 4 páginas com capa, KPIs semaforizados, gráficos e **comentário executivo gerado automaticamente** com base nos dados do mês

Adicionalmente, um **agendador** (`agendador.py`) dispara o fechamento todo dia 1º às 06:00 sem intervenção humana — replicando o funcionamento real de uma área de Controladoria.

---

## Empresa fictícia

| Campo | Dado |
|---|---|
| Razão social | Rede Máxima Supermercados Ltda. |
| Setor | Varejo alimentar |
| Lojas | L01-Centro · L02-Norte · L03-Sul · L04-Leste · L05-Aeroporto |
| Período dos dados | Janeiro 2023 – Dezembro 2024 (24 meses) |
| Seed de geração | 42 |

---

## Stack técnica

| Camada | Tecnologia |
|---|---|
| Geração de dados | Python · Pandas · NumPy (seed=42) |
| Armazenamento | CSV (5 arquivos em `data/raw/`) |
| ETL e KPIs | Pandas |
| Geração Excel | openpyxl |
| Geração PDF | ReportLab |
| Gráficos | Matplotlib |
| CLI | argparse |
| Agendamento | schedule |
| Versionamento | Git · GitHub |

---

## Estrutura do projeto

```
relatorio-contabil/
│
├── README.md
├── requirements.txt
├── agendador.py                ← Agendador automático (todo dia 1º)
│
├── scripts/
│   ├── gerar_dados.py          ← Gera os 5 CSVs sintéticos
│   └── gerar_relatorio.py      ← Motor principal (ETL + Excel + PDF + CLI)
│
├── data/
│   └── raw/
│       ├── dre.csv             ← 120 linhas (24 meses × 5 lojas)
│       ├── fluxo_caixa.csv     ← 24 linhas (consolidado mensal)
│       ├── contas_receber.csv  ← 600 linhas (25 clientes × 24 meses)
│       ├── centro_custos.csv   ← 192 linhas (8 depto × 24 meses)
│       └── orcamento.csv       ← 10 linhas (5 lojas × 2 anos)
│
└── output/                     ← Relatórios gerados (Excel + PDF)
```

---

## Como usar

### 1. Instalar dependências

```bash
pip install -r requirements.txt
```

### 2. Gerar os dados sintéticos

```bash
python scripts/gerar_dados.py
```

### 3. Gerar o relatório (CLI)

```bash
# Relatório consolidado (todas as lojas) — gera Excel + PDF
python scripts/gerar_relatorio.py --mes 2024-06

# Relatório de uma loja específica
python scripts/gerar_relatorio.py --mes 2024-06 --loja L01-Centro

# Apenas Excel
python scripts/gerar_relatorio.py --mes 2024-06 --formato excel

# Apenas PDF
python scripts/gerar_relatorio.py --mes 2024-06 --formato pdf

# Diretório de saída personalizado
python scripts/gerar_relatorio.py --mes 2024-06 --output /caminho/destino
```

### 4. Iniciar o agendador

```bash
# Inicia o loop (roda todo dia 1º às 06:00)
python agendador.py

# Forçar execução imediata (para testes)
python agendador.py --executar-agora
```

---

## Entregáveis gerados

### Excel — 5 abas

| Aba | Conteúdo |
|---|---|
| Resumo Executivo | KPI cards semaforizados, variações MoM e vs. orçado, comentário executivo automático |
| DRE por Loja | Receita bruta/líquida, CMV, Lucro Bruto, OPEX, EBITDA, Margem, variação vs. orçado |
| Ranking de Lojas | Ordenado por Margem EBITDA com semaforização (verde/amarelo/vermelho) |
| Inadimplência Aging | 25 clientes com faixas: A Vencer · 1-30d · 31-60d · 61-90d · 91-120d · +120d |
| Gráficos | 4 gráficos embutidos com tema dark: Receita vs EBITDA, Ranking, Real vs Orçado, Inadimplência |

### PDF — 4 páginas

| Página | Conteúdo |
|---|---|
| Capa | Logo, competência, emissão, escopo |
| KPIs + Comentário | 12 KPIs semaforizados + comentário executivo gerado automaticamente |
| Gráficos | 4 gráficos analíticos |
| DRE Consolidada | Tabela por loja com totais |

---

## KPIs e metas

| KPI | Fórmula | Meta verde | Meta amarela | Meta vermelha |
|---|---|---|---|---|
| Margem EBITDA | EBITDA / Receita Líquida | ≥ 20% | 15–20% | < 15% |
| Inadimplência | Vencido / Carteira Total | ≤ 10% | 10–18% | > 18% |
| Score Saúde | Margem(40) + FCO(30) + Inadim(30) | ≥ 70 | 50–70 | < 50 |

---

## Comentário executivo automático

Um dos diferenciais do projeto é a **geração automática de texto analítico** a partir dos KPIs calculados. O script analisa os números e produz parágrafos como:

> *"O desempenho financeiro consolidado em Junho de 2024 apresentou receita bruta de R$ 3.174.000, com variação de +2,3% em relação ao mês anterior. A receita ficou 1,8% acima do orçado, enquanto o EBITDA apresentou desvio de +3,1% frente à meta."*

---

## Datasets gerados

### `dre.csv` — 120 linhas × 21 colunas
- 5 lojas × 24 meses
- Receita base: R$ 620K–1,8MM/mês por loja
- Sazonalidade de varejo (pico Dez, vale Fev)
- Crescimento ~9% a.a.
- Outlier: L03-Sul em Ago/2023 (pico de reposição de estoque, CMV +18%)

### `fluxo_caixa.csv` — 24 linhas × 12 colunas
- Saldo inicial: R$ 3,2MM → encerramento ~R$ 5,5MM
- 3 eventos CAPEX: Mar/23, Set/23, Mar/24
- Captação Jan/2023: +600K | Amortização Jan/2024: -180K

### `contas_receber.csv` — 600 linhas × 10 colunas
- 25 clientes fictícios com aging em 6 faixas
- Taxa de inadimplência variável por cliente

### `centro_custos.csv` — 192 linhas × 6 colunas
- 8 departamentos com pesos econômicos realistas
- Outlier: TI em Nov/2023 (migração de sistema, +105% orçamento)

### `orcamento.csv` — 10 linhas × 8 colunas
- Metas anuais por loja (2023 e 2024)

---

## Portfólio

Este é o **Projeto 3** de um portfólio de 6 projetos de Análise de Dados / FP&A:

| # | Projeto | Status |
|---|---|---|
| 1 | Financial Performance Dashboard (Streamlit + Plotly) | ✅ Concluído |
| 2 | Análise Exploratória com Storytelling | ✅ Concluído |
| **3** | **Automação de Relatório Contábil/Financeiro** | ✅ **Concluído** |
| 4 | Dashboard Operacional com KPIs | 🔄 Em desenvolvimento |
| 5 | Análise com SQL | 🔄 Em desenvolvimento |
| 6 | Lean Six Sigma + Dados | 🔄 Em desenvolvimento |

---

## Autor

**Gary Rainer Chumacero Vanderlei**
Analista de Dados · FP&A · BI Financeiro · Controladoria

- LinkedIn: [linkedin.com/in/garyrainercv](https://www.linkedin.com/in/garyrainercv)
- GitHub: [github.com/Gary-Rainer-Chumacero-Vanderlei](https://github.com/Gary-Rainer-Chumacero-Vanderlei)
- Email: garyvanderlei@gmail.com
- Localização: João Pessoa – PB (presencial ou remoto)
