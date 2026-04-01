# Python Portfolio Analysis

> Análise quantitativa de portfólios de investimento com métricas de risco, coleta de dados via APIs públicas e regressão múltipla com Rolling Beta.

![Python](https://img.shields.io/badge/Python-3.10+-blue?logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-2.x-150458?logo=pandas&logoColor=white)
![statsmodels](https://img.shields.io/badge/statsmodels-0.14-4B8BBE)
![matplotlib](https://img.shields.io/badge/matplotlib-3.x-orange)
![License](https://img.shields.io/badge/license-MIT-green)

---

## Sobre o Projeto

Este projeto foi desenvolvido como um coding test para uma family office e cobre quatro áreas principais de análise quantitativa aplicada a fundos de investimento e portfólios:

- **Q1** — Análise de performance de portfólios com ativos locais e offshore
- **Q2** — Módulo de risco com VaR, Expected Shortfall, Semidesvio, Sortino e Calmar
- **Q3** — Coleta automática de dados via API do Banco Central (CDI) e CVM (cotas de fundos)
- **Q4** — Regressão múltipla com tratamento de multicolinearidade (VIF) e Rolling Beta

---

## Funcionalidades

### Q1 — Performance
- Leitura e merge de bases de retornos locais e offshore
- Construção de 3 portfólios com pesos configuráveis
- Cálculo de: Retorno Acumulado, Retorno Anualizado, Volatilidade, Índice de Sharpe e Max Drawdown
- Gráfico de valor acumulado comparando os portfólios com o CDI

### Q2 — Risco
- VaR Histórico e Paramétrico (95%)
- Expected Shortfall / CVaR (95%)
- Semidesvio (Downside Deviation) anualizado
- Índice de Sortino e Índice de Calmar
- Gráfico de drawdown e histograma de retornos com VaR e ES

### Q3 — APIs Públicas
- CDI via API do Banco Central (SGS série 12) usando `pd.read_json`
- Cotas diárias de fundos via dados abertos da CVM (arquivos `.zip` mensais)
- Suporte a período customizável via input interativo
- Gráfico de valor acumulado: fundo vs CDI

### Q4 — Regressão Múltipla
- Regressão OLS completa do fundo vs classes de ativos locais
- Cálculo de VIF com eliminação iterativa de multicolinearidade
- Regressão corrigida com variáveis selecionadas
- Rolling Beta com janela de 252 dias úteis
- Gráficos de betas e rolling beta por classe de ativo

---

## Tecnologias Utilizadas

| Biblioteca | Uso |
|---|---|
| `pandas` | Manipulação de dados e séries temporais |
| `numpy` | Cálculos vetorizados |
| `matplotlib` | Visualizações e gráficos |
| `scipy` | Distribuição Normal (VaR Paramétrico) |
| `statsmodels` | Regressão OLS, Rolling OLS e VIF |
| `requests` | Coleta de dados via API |
| `zipfile` | Extração dos arquivos da CVM |

---

## Como Rodar

### 1. Clone o repositório
```bash
git clone https://github.com/gusjmo/python-portfolio-analysis.git
cd python-portfolio-analysis
```

### 2. Instale as dependências
```bash
pip install -r requirements.txt
```

### 3. Adicione os arquivos de dados
Coloque na mesma pasta do script:
```
retornosdadoslocais.xlsx
retornosdadosoffshore.xlsx
```

### 4. Execute
```bash
python coding_test.py
```
O script irá solicitar o período de análise (início e fim) e coletará os dados automaticamente.

---

## Arquivos Gerados

### Gráficos (PNG)
| Arquivo | Descrição |
|---|---|
| `q1_valor_acumulado.png` | Valor acumulado dos portfólios vs CDI |
| `q2_drawdown.png` | Drawdown histórico das carteiras |
| `q2_histograma_var_es.png` | Histograma de retornos com VaR e ES |
| `q3_fundo_vs_cdi.png` | Fundo ALLOCATION vs CDI |
| `q4_betas.png` | Betas da regressão completa e corrigida |
| `q4_rolling_beta.png` | Rolling Beta por classe de ativo |

### Dados (CSV)
| Arquivo | Descrição |
|---|---|
| `q1_metricas_performance.csv` | Tabela de performance das carteiras |
| `q2_metricas_risco.csv` | Tabela de risco das carteiras |
| `q3_cdi_bcb.csv` | Série histórica do CDI |
| `q3_cotas_fundo.csv` | Cotas diárias do fundo (CVM) |
| `q4_betas_completo.csv` | Betas da regressão completa |
| `q4_betas_corrigido.csv` | Betas após eliminação por VIF |
| `q4_rolling_betas.csv` | Rolling Beta por data e classe |

---

## Conceitos Aplicados

- **Índice de Sharpe** — retorno ajustado ao risco total
- **Índice de Sortino** — retorno ajustado ao risco de queda
- **Índice de Calmar** — retorno anualizado sobre o máximo drawdown
- **VaR Histórico** — percentil empírico dos piores retornos
- **VaR Paramétrico** — baseado na distribuição Normal
- **Expected Shortfall (CVaR)** — média das perdas além do VaR
- **Semidesvio** — volatilidade apenas dos retornos negativos
- **VIF (Variance Inflation Factor)** — diagnóstico de multicolinearidade
- **Rolling Beta** — exposição dinâmica do fundo às classes de ativos

---

## Autor

**Gustavo Juvencio**
Estudante de Análise e Desenvolvimento de Sistemas

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Gustavo%20Juvencio-0077B5?logo=linkedin)](https://linkedin.com/in/gustavo-juvencio)
[![GitHub](https://img.shields.io/badge/GitHub-gusjmo-181717?logo=github)](https://github.com/gusjmo)
