<div align="center">

# 🤖 Automação de Relatório Contábil/Financeiro
### Rede Máxima Supermercados Ltda.

![Python](https://img.shields.io/badge/Python-3.11-4F8EF7?style=flat-square&logo=python&logoColor=white)
![openpyxl](https://img.shields.io/badge/openpyxl-3.1+-217346?style=flat-square&logo=microsoftexcel&logoColor=white)
![ReportLab](https://img.shields.io/badge/ReportLab-4.x-E74C3C?style=flat-square)
![Pandas](https://img.shields.io/badge/Pandas-2.1+-F5A623?style=flat-square&logo=pandas&logoColor=white)
![Matplotlib](https://img.shields.io/badge/Matplotlib-3.8+-11557C?style=flat-square&logo=python&logoColor=white)
![schedule](https://img.shields.io/badge/schedule-1.2+-4A90D9?style=flat-square)
![Status](https://img.shields.io/badge/Status-Concluído-2DD4A0?style=flat-square)

**Projeto de portfólio em Automação de Controladoria e FP&A**  
Pipeline completo de fechamento mensal com geração automática de Excel, PDF e comentário executivo — disparado por CLI ou agendador sem intervenção humana.

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Gary%20Rainer-0A66C2?style=flat-square&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/garyrainercv/)
[![GitHub](https://img.shields.io/badge/GitHub-Gary--Rainer--Chumacero--Vanderlei-181717?style=flat-square&logo=github&logoColor=white)](https://github.com/Gary-Rainer-Chumacero-Vanderlei)

</div>

---

## 📋 Índice

- [Contexto e Problema de Pesquisa](#contexto-e-problema-de-pesquisa)
- [Objetivos](#objetivos)
- [Arquitetura do Projeto](#arquitetura-do-projeto)
- [Datasets](#datasets)
- [Metodologia e Pipeline de Dados](#metodologia-e-pipeline-de-dados)
- [KPIs e Métricas](#kpis-e-métricas)
- [Entregáveis — Excel e PDF](#entregáveis--excel-e-pdf)
- [Stack Tecnológica](#stack-tecnológica)
- [Resultados e Insights](#resultados-e-insights)
- [Como Executar](#como-executar)
- [Estrutura de Pastas](#estrutura-de-pastas)
- [Sobre o Autor](#sobre-o-autor)

---

## 🎯 Contexto e Problema de Pesquisa

Áreas de Controladoria em redes de varejo enfrentam um gargalo recorrente: o fechamento mensal exige consolidação de dados de múltiplas lojas, cálculo manual de KPIs, formatação de planilhas e elaboração de relatórios executivos — processo que consome dias de trabalho de analistas qualificados e está sujeito a erros humanos de digitação, versionamento incorreto de arquivos e atrasos na distribuição para a diretoria.

**Problema central:** Como automatizar integralmente o pacote de fechamento mensal de uma rede de varejo — da ingestão dos dados brutos à entrega de Excel estruturado e PDF executivo prontos para distribuição — eliminando o trabalho manual repetitivo, garantindo consistência de formatação e adicionando inteligência analítica na forma de comentário executivo gerado automaticamente com base nos KPIs do mês?

Este projeto simula o ambiente real da **Rede Máxima Supermercados Ltda.**, rede fictícia de 5 lojas regionais, com dados sintéticos gerados por script Python com características econômicas realistas: sazonalidade de varejo alimentar com pico em dezembro e vale em fevereiro, crescimento anual de ~9%, outliers intencionais de reposição de estoque e migração de sistema, e perfil de inadimplência segmentado por cliente com aging em 6 faixas.

---

## 🎯 Objetivos

### Objetivo Geral
Construir um pipeline de automação completo de fechamento mensal — da geração de dados sintéticos à entrega de artefatos executivos — demonstrando competências em engenharia de dados, modelagem de KPIs, geração programática de Excel e PDF, automação via CLI e agendamento de tarefas.

### Objetivos Específicos
- Gerar dados financeiros sintéticos com distribuições, tendências e sazonalidade realistas para 24 meses (2023–2024) em 5 datasets integrados cobrindo DRE, fluxo de caixa, recebíveis, custos e orçamento
- Construir motor ETL em Pandas que consolide os 5 datasets e calcule os KPIs de fechamento a partir do mês de competência informado via CLI
- Gerar automaticamente workbook Excel com 5 abas estruturadas, semaforização de metas e 4 gráficos embutidos com tema dark
- Gerar automaticamente relatório PDF de 4 páginas com capa, KPIs semaforizados, gráficos e comentário executivo produzido a partir dos dados calculados
- Implementar agendador que dispara o fechamento todo dia 1º às 06:00 sem intervenção humana, replicando o funcionamento real de uma área de Controladoria
- Documentar o processo completo para publicação em portfólio GitHub/LinkedIn

---

## 🏗️ Arquitetura do Projeto

```
Geração de Dados (Python/NumPy)
        │
        ▼
   data/raw/ (5 CSVs)
        │
        ▼
   Motor ETL (gerar_relatorio.py)
   argparse CLI: --mes · --loja · --formato · --output
        │
        ├──► Excel (.xlsx)
        │    5 abas: Resumo Executivo · DRE por Loja · Ranking
        │            Aging de Recebíveis · Gráficos
        │
        └──► PDF (4 páginas)
             Capa · KPIs + Comentário · Gráficos · DRE Consolidada
                        │
                        ▼
              output/ (artefatos prontos para distribuição)
                        ▲
              agendador.py (disparo automático todo dia 1º 06:00)
```

O pipeline segue arquitetura **ETL orientada a entrega**: os dados brutos são ingeridos e transformados em memória pelo motor Pandas, os artefatos são gerados via openpyxl e ReportLab, e o agendador fecha o ciclo operacional eliminando qualquer intervenção manual no processo de fechamento.

---

## 📊 Datasets

Todos os datasets foram gerados com `seed=42` para reprodutibilidade, via `scripts/gerar_dados.py`.

| Dataset | Registros | Descrição |
|---|---|---|
| `dre.csv` | 120 × 21 colunas | DRE mensal por loja (5 lojas × 24 meses) |
| `fluxo_caixa.csv` | 24 × 12 colunas | Fluxo de caixa consolidado mensal |
| `contas_receber.csv` | 600 × 10 colunas | Carteira de 25 clientes com aging em 6 faixas |
| `centro_custos.csv` | 192 × 6 colunas | Custos reais vs orçamento por departamento |
| `orcamento.csv` | 10 × 8 colunas | Metas anuais por loja (2023 e 2024) |

### Características de Realismo dos Dados

**DRE:** Receita base entre R$ 620K e R$ 1,8MM/mês por loja com crescimento de ~9% a.a. e sazonalidade de varejo alimentar (pico em dezembro, vale em fevereiro). Outlier intencional na L03-Sul em agosto/2023 simulando pico extraordinário de reposição de estoque com CMV +18%.

**Fluxo de Caixa:** Saldo inicial de R$ 3,2MM evoluindo para ~R$ 5,5MM em dezembro/2024. Três eventos de CAPEX nos períodos março/2023, setembro/2023 e março/2024, captação de R$ 600K em janeiro/2023 e amortização de R$ 180K em janeiro/2024.

**Contas a Receber:** 25 clientes fictícios com aging distribuído em 6 faixas: A Vencer, 1–30d, 31–60d, 61–90d, 91–120d e +120d. Taxa de inadimplência variável por cliente, permitindo análise de concentração de risco.

**Centro de Custos:** 8 departamentos com pesos econômicos realistas. Outlier intencional em TI durante novembro/2023 (+105% sobre orçamento), simulando projeto de migração de sistema.

**Orçamento:** Metas anuais por loja para 2023 e 2024, viabilizando o comparativo real vs. orçado em todos os KPIs gerados pelo motor ETL.

---

## 🔬 Metodologia e Pipeline de Dados

### 1. Geração de Dados (`scripts/gerar_dados.py`)
Utiliza **NumPy** para distribuições estatísticas controladas e **Pandas** para estruturação dos DataFrames. Cada dataset foi modelado com componente de tendência linear (crescimento anual), componente sazonal via funções trigonométricas (sin/cos), ruído gaussiano para variação mensal realista e eventos pontuais determinísticos (CAPEX, outliers operacionais).

### 2. Motor ETL e KPIs (`scripts/gerar_relatorio.py`)
Pipeline de transformação em **Pandas** que, a partir do mês de competência informado via CLI, realiza a ingestão dos 5 CSVs, filtragem temporal, consolidação por loja, cálculo de todos os KPIs e comparativos MoM e vs. orçado, e alimentação dos módulos de geração de Excel e PDF. Aceita os seguintes parâmetros:

| Parâmetro | Descrição | Exemplo |
|---|---|---|
| `--mes` | Competência no formato AAAA-MM | `--mes 2024-06` |
| `--loja` | Filtra por loja específica (opcional) | `--loja L01-Centro` |
| `--formato` | Saída desejada: `excel`, `pdf` ou ambos | `--formato excel` |
| `--output` | Diretório de destino dos artefatos | `--output /destino` |

### 3. Geração do Excel (`openpyxl`)
Workbook com 5 abas estruturadas, semaforização automática de metas por faixas de cor (verde/amarelo/vermelho), gráficos embutidos gerados via **Matplotlib** e exportados como imagens inseridas nas células, e formatação financeira consistente em todas as tabelas.

### 4. Geração do PDF (`ReportLab`)
Relatório de 4 páginas com layout profissional: capa com identidade da empresa e competência, página de KPIs semaforizados com comentário executivo gerado automaticamente a partir dos valores calculados pelo ETL, página de gráficos analíticos e DRE consolidada por loja com totais.

### 5. Agendador (`agendador.py`)
Loop de agendamento via **schedule** que dispara o fechamento todo dia 1º às 06:00 sem intervenção humana. Suporta flag `--executar-agora` para testes imediatos fora do ciclo programado.

---

## 📐 KPIs e Métricas

| KPI | Fórmula | Meta 🟢 | Atenção 🟡 | Crítico 🔴 |
|---|---|---|---|---|
| Margem Bruta % | Lucro Bruto / Receita Líquida | ≥ 30% | 22–30% | < 22% |
| EBITDA | Lucro Bruto − OPEX | — | — | — |
| Margem EBITDA % | EBITDA / Receita Líquida | ≥ 20% | 15–20% | < 15% |
| Desvio vs. Orçado % | (Real − Orçado) / Orçado | ≥ 0% | −5% a 0% | < −5% |
| Taxa Inadimplência % | Vencido / Carteira Total | ≤ 10% | 10–18% | > 18% |
| FCO | Recebimentos − Total Saídas | Positivo | — | Negativo |
| **Score Saúde Financeira** | Margem(40) + FCO(30) + Inadimplência(30) | ≥ 70 | 50–70 | < 50 |

---

## 📄 Entregáveis — Excel e PDF

### Excel — 5 Abas
**Público:** Controladoria, FP&A, Gestores de Loja

| Aba | Conteúdo |
|---|---|
| Resumo Executivo | KPI cards semaforizados, variações MoM e vs. orçado, comentário executivo automático |
| DRE por Loja | Receita bruta/líquida, CMV, Lucro Bruto, OPEX, EBITDA, Margem e desvio vs. orçado por loja |
| Ranking de Lojas | Ordenação por Margem EBITDA com semaforização verde/amarelo/vermelho |
| Inadimplência Aging | 25 clientes com valores distribuídos nas 6 faixas de aging e total inadimplente |
| Gráficos | 4 gráficos embutidos em tema dark: Receita vs EBITDA · Ranking de Lojas · Real vs Orçado · Inadimplência por Aging |

### PDF — 4 Páginas
**Público:** C-Level, Diretoria, Conselho

| Página | Conteúdo |
|---|---|
| Capa | Logotipo, competência de referência, data de emissão e escopo do relatório |
| KPIs + Comentário | 12 KPIs semaforizados + comentário executivo gerado automaticamente a partir dos dados |
| Gráficos | 4 visualizações analíticas do período |
| DRE Consolidada | Tabela completa por loja com totais e variações |

### Comentário Executivo Automático
Um dos diferenciais do projeto é a geração de texto analítico diretamente a partir dos KPIs calculados pelo ETL, sem intervenção manual. O motor analisa os números do mês e produz parágrafos estruturados como:

> *"O desempenho consolidado em Junho/2024 registrou receita bruta de R$ 3.174.000, com crescimento de +2,3% em relação ao mês anterior e desvio de +1,8% frente ao orçado. O EBITDA apresentou desvio positivo de +3,1% vs. meta, sustentado pela performance acima do esperado nas lojas L01-Centro e L04-Leste."*

---

## 🛠️ Stack Tecnológica

| Camada | Tecnologia | Finalidade |
|---|---|---|
| Geração de Dados | Python · Pandas · NumPy | Simulação de dados realistas com seed |
| Armazenamento | CSV | Persistência e portabilidade dos datasets |
| ETL e KPIs | Pandas | Ingestão, transformação e cálculo de indicadores |
| Geração Excel | openpyxl | Workbook com 5 abas, semaforização e gráficos |
| Geração PDF | ReportLab | Relatório executivo de 4 páginas |
| Gráficos | Matplotlib | Visualizações embutidas em Excel e PDF |
| CLI | argparse | Interface de linha de comando parametrizável |
| Agendamento | schedule | Disparo automático todo dia 1º às 06:00 |
| Versionamento | Git · GitHub | Controle de versão e portfólio |

---

## 📈 Resultados e Insights

A partir dos dados simulados para o período 2023–2024, os principais resultados demonstrados pela automação foram:

**Redução drástica do tempo de fechamento**
O processo que em uma operação manual consumiria um a dois dias de trabalho analítico — consolidação de planilhas, formatação, cálculo de KPIs, elaboração do comentário e exportação do PDF — é executado em segundos pelo motor automatizado, liberando a equipe de Controladoria para análise e não para digitação.

**Crescimento e rentabilidade da rede**
Receita consolidada cresceu ~9% de 2023 para 2024, com L01-Centro e L04-Leste como lojas de maior margem EBITDA. A L03-Sul apresentou desvio positivo de CMV em agosto/2023 (pico de reposição de estoque), detectado automaticamente como outlier no ranking mensal.

**Gestão de caixa robusta**
FCO positivo em todos os meses do período, com saldo crescendo de R$ 3,2MM para ~R$ 5,5MM mesmo após 3 eventos de CAPEX. A automação permite visualizar a evolução do caixa e os eventos extraordinários diretamente na aba de Resumo Executivo do Excel gerado.

**Inadimplência sob controle com concentração monitorável**
Taxa consolidada dentro da faixa de atenção (10–18%), com concentração de risco identificável por cliente na aba de Aging. O aging em 6 faixas revela que a maior parte dos valores vencidos se concentra nas faixas de 31–60d e 61–90d — janela ainda recuperável com ação proativa de cobrança.

**Outlier de TI mapeado e comunicado automaticamente**
O pico de custos de TI em novembro/2023 (+105% do orçamento, migração de sistema) aparece automaticamente como desvio crítico no Resumo Executivo e no comentário gerado, sem necessidade de identificação manual pela Controladoria.

---

## 🚀 Como Executar

### Pré-requisitos
- Python 3.10+
- Git

### Instalação

```bash
# 1. Clonar o repositório
git clone https://github.com/Gary-Rainer-Chumacero-Vanderlei/relatorio-contabil.git
cd relatorio-contabil

# 2. Criar e ativar ambiente virtual
python -m venv venv
# Windows:
venv\Scripts\activate
# Linux/macOS:
source venv/bin/activate

# 3. Instalar dependências
pip install -r requirements.txt

# 4. Gerar os datasets sintéticos
python scripts/gerar_dados.py

# 5. Gerar o relatório via CLI
# Consolidado (todas as lojas) — Excel + PDF
python scripts/gerar_relatorio.py --mes 2024-06

# Loja específica
python scripts/gerar_relatorio.py --mes 2024-06 --loja L01-Centro

# Apenas Excel ou apenas PDF
python scripts/gerar_relatorio.py --mes 2024-06 --formato excel
python scripts/gerar_relatorio.py --mes 2024-06 --formato pdf

# Diretório de saída personalizado
python scripts/gerar_relatorio.py --mes 2024-06 --output /caminho/destino

# 6. Iniciar o agendador (disparo automático todo dia 1º às 06:00)
python agendador.py

# Forçar execução imediata para testes
python agendador.py --executar-agora
```

### Dependências (`requirements.txt`)

```txt
pandas>=2.1.0
numpy>=1.26.0
openpyxl>=3.1.0
reportlab>=4.0.0
matplotlib>=3.8.0
schedule>=1.2.0
```

---

## 📁 Estrutura de Pastas

```
relatorio-contabil/
│
├── README.md                      ← Documentação do projeto
├── requirements.txt               ← Dependências Python
├── agendador.py                   ← Agendador automático (todo dia 1º às 06:00)
│
├── scripts/
│   ├── gerar_dados.py             ← Geração dos 5 datasets sintéticos (seed=42)
│   └── gerar_relatorio.py        ← Motor principal: ETL + Excel + PDF + CLI
│
├── data/
│   └── raw/
│       ├── dre.csv                ← DRE mensal: 120 linhas (5 lojas × 24 meses)
│       ├── fluxo_caixa.csv        ← Fluxo de caixa consolidado: 24 linhas
│       ├── contas_receber.csv     ← Aging de recebíveis: 600 linhas (25 clientes)
│       ├── centro_custos.csv      ← Custos vs orçamento: 192 linhas (8 departamentos)
│       └── orcamento.csv          ← Metas anuais por loja: 10 linhas
│
└── output/                        ← Artefatos gerados (Excel + PDF por competência)
```

---

## 👤 Sobre o Autor

**Gary Rainer Chumacero Vanderlei**

Contador (CRC ativo) e Analista de Dados com experiência em BI Financeiro nos setores energético e de saúde. Master Black Belt em Lean Six Sigma. Formação em Ciências da Computação.

Stack principal: Python · SQL · Java · Streamlit · Plotly

Baseado em João Pessoa — PB · Disponível para posições presenciais ou remotas em Análise de Dados / BI.

<div align="left">

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Conectar-0A66C2?style=for-the-badge&logo=linkedin&logoColor=white)](https://www.linkedin.com/in/garyrainercv/)
[![GitHub](https://img.shields.io/badge/GitHub-Seguir-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/Gary-Rainer-Chumacero-Vanderlei)
[![Email](https://img.shields.io/badge/Email-Contato-D14836?style=for-the-badge&logo=gmail&logoColor=white)](mailto:garyvanderlei@gmail.com)

</div>
