# ============================================================
# CODING TEST — MMZR Family Office
# Arquivo único: Q1 + Q2 + Q3 + Q4
#
# PREMISSAS ADOTADAS:
# - Pesos baseados na carteira MODERADA da Carta Mensal MMZR
# - Alternativos Líquidos (Dólar) alocados na classe Alternativos
# - CDI usado como taxa livre de risco (Sharpe e Sortino)
# - Frequência de 252 dias úteis para anualização
# - VaR Paramétrico assume distribuição Normal dos retornos
# - Multicolinearidade tratada por VIF iterativo (threshold = 10)
# - Rolling beta: janela de 252 dias úteis (1 ano)
# - Dados CVM disponíveis a partir de jan/2021
# ============================================================

import io
import warnings
import zipfile
from datetime import datetime
from pathlib import Path

import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import numpy as np
import pandas as pd
import requests
import statsmodels.api as sm
from scipy import stats
from statsmodels.regression.rolling import RollingOLS
from statsmodels.stats.outliers_influence import variance_inflation_factor

# Suprimir apenas warnings esperados (não TODOS os warnings do processo)
warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=UserWarning, module="statsmodels")

# Diretório base do script (evita dependência do CWD)
BASE_DIR = Path(__file__).parent

# ============================================================
# PARÂMETROS GLOBAIS
# (edite aqui para testar outras configurações)
# ============================================================

CNPJ_FUNDO  = "36671748000163"           # ALLOCATION RETORNO ABSOLUTO FIC FIM
NOME_FUNDO  = "ALLOCATION RETORNO ABSOLUTO FIC FIM"
FREQ        = 252                          # dias úteis por ano
JANELA_ROLL = 252                          # janela do rolling beta (1 ano)
VIF_THRESH  = 10                           # threshold de multicolinearidade

# ── Pesos da Carteira Moderada — EDITE AQUI ─────────────────
PESO_INFLACAO = 0.20   # peso total da classe inflação

PESOS_BASE = {
    "CDI"           : 0.20,   # Pós-Fixado
    "IDkA Pré 3 Anos": 0.10,  # Pré-Fixado
    "IHFA"          : 0.15,   # Retorno Absoluto
    "IFIX"          : 0.10,   # Imobiliário (Fundos Listados)
    "Dólar"         : 0.05,   # Alternativos Líquidos → Alternativos
    "Ibovespa"      : 0.10,   # Renda Variável Local
    "RV Global"     : 0.07,   # Renda Variável Global (offshore)
    "RF Global"     : 0.03,   # Renda Fixa Global (offshore)
}
# ── Fim da área editável ─────────────────────────────────────

CARTEIRAS = {
    "P1 - IMA-B"  : {**PESOS_BASE, "IMA-B"  : PESO_INFLACAO},
    "P2 - IMA-B 5": {**PESOS_BASE, "IMA-B 5": PESO_INFLACAO},
    "P3 - IMA-B 5+": {**PESOS_BASE, "IMA-B 5+": PESO_INFLACAO},
}

CORES = {
    "P1 - IMA-B"  : "#1f4e79",
    "P2 - IMA-B 5": "#2e75b6",
    "P3 - IMA-B 5+": "#00b0f0",
}

# ============================================================
# CONFIGURAÇÃO INTERATIVA DO PERÍODO (Q3 / Desafio)
# ============================================================

DATA_MINIMA_CVM = datetime(2021, 1, 1)

print("=" * 60)
print("CONFIGURAÇÃO DO PERÍODO DE ANÁLISE")
print("=" * 60)
print("Informe o período de análise do fundo.")
print("Deixe em branco para usar os valores padrão.")
print("⚠️  A CVM disponibiliza dados apenas a partir de 01/01/2021.\n")

inicio_input = input("Data de INÍCIO (DD/MM/AAAA) [padrão: 01/01/2021]: ").strip()
if inicio_input == "":
    DATA_INICIO = "2021-01-01"
    print("   Usando início padrão: 01/01/2021")
else:
    try:
        dt = datetime.strptime(inicio_input, "%d/%m/%Y")
        if dt < DATA_MINIMA_CVM:
            print("   ⚠️  Data anterior a jan/2021! Ajustando para 01/01/2021.")
            DATA_INICIO = "2021-01-01"
        else:
            DATA_INICIO = dt.strftime("%Y-%m-%d")
            print(f"   Início: {inicio_input}")
    except ValueError:
        print("   ⚠️  Formato inválido! Usando padrão: 01/01/2021")
        DATA_INICIO = "2021-01-01"

fim_input = input("Data de FIM    (DD/MM/AAAA) [padrão: hoje]:   ").strip()
if fim_input == "":
    DATA_FIM = datetime.today()
    print(f"   Usando fim: hoje ({DATA_FIM.strftime('%d/%m/%Y')})")
else:
    try:
        DATA_FIM = datetime.strptime(fim_input, "%d/%m/%Y")
        if DATA_FIM < DATA_MINIMA_CVM:
            print("   ⚠️  Data anterior a jan/2021! Usando hoje.")
            DATA_FIM = datetime.today()
        else:
            print(f"   Fim: {fim_input}")
    except ValueError:
        print("   ⚠️  Formato inválido! Usando data de hoje.")
        DATA_FIM = datetime.today()

if datetime.strptime(DATA_INICIO, "%Y-%m-%d") > DATA_FIM:
    print("   ⚠️  Início maior que fim! Invertendo.")
    DATA_INICIO = DATA_FIM.strftime("%Y-%m-%d")

print(f"\n   Período: {DATA_INICIO} → {DATA_FIM.strftime('%Y-%m-%d')}\n")


# ============================================================
# FUNÇÕES UTILITÁRIAS — leitura de Excel
# ============================================================

def ler_e_limpar(caminho: "str | Path") -> pd.DataFrame:
    """Lê e limpa os arquivos Excel do Coding Test.
    A linha 0 do Excel é descartada (cabeçalho em branco);
    os nomes das colunas estão na linha 1.
    """
    raw   = pd.read_excel(caminho, header=None)
    nomes = raw.iloc[1].tolist()
    nomes[0] = "Data"
    df = raw.iloc[2:].copy()
    df.columns = nomes
    df = df.reset_index(drop=True)
    for col in df.columns:
        if col != "Data":
            df[col] = pd.to_numeric(df[col], errors="coerce")
    df["Data"] = pd.to_datetime(
        df["Data"], format="mixed", dayfirst=False, errors="coerce"
    )
    df = df.dropna(subset=["Data"]).reset_index(drop=True)
    return df


# ============================================================
# MÓDULO DE MÉTRICAS DE PERFORMANCE (Q1c)
# ============================================================

def retorno_acumulado(ret: pd.Series) -> float:
    """Retorno total acumulado no período."""
    return (1 + ret).prod() - 1


def retorno_anualizado(ret: pd.Series, dias_ano: int = FREQ) -> float:
    """Retorno total anualizado pela regra de juros compostos."""
    n = len(ret)
    return (1 + retorno_acumulado(ret)) ** (dias_ano / n) - 1


def volatilidade_diaria(ret: pd.Series) -> float:
    """Desvio padrão dos retornos diários."""
    return ret.std()


def volatilidade_anual(ret: pd.Series, dias_ano: int = FREQ) -> float:
    """Volatilidade anualizada pela raiz de 252."""
    return ret.std() * np.sqrt(dias_ano)


def sharpe(ret: pd.Series, rf_anual: float = 0.0, dias_ano: int = FREQ) -> float:
    """Índice de Sharpe anualizado: (Ret_anual - RF_anual) / Vol_anual."""
    vol_a = volatilidade_anual(ret, dias_ano)
    ret_a = retorno_anualizado(ret, dias_ano)
    return (ret_a - rf_anual) / vol_a if vol_a != 0 else np.nan


# ============================================================
# MÓDULO DE MÉTRICAS DE RISCO (Q2a)
# ============================================================

def semidesvio(ret: pd.Series, freq: int = FREQ) -> float:
    """
    Semidesvio (Downside Deviation) anualizado.
    Calcula a volatilidade apenas dos retornos negativos (abaixo de zero).
    Penaliza somente o risco de perda, ignorando a volatilidade positiva.
    Útil quando a distribuição de retornos é assimétrica.
    """
    retornos_neg = ret[ret < 0]
    return retornos_neg.std() * np.sqrt(freq)


def sortino(ret: pd.Series, rf_anual: float = 0.0, dias_ano: int = FREQ) -> float:
    """Índice de Sortino anualizado: (Ret_anual - RF_anual) / Semidesvio_anual."""
    ret_a  = retorno_anualizado(ret, dias_ano)
    semi_a = semidesvio(ret, dias_ano)   # já é anualizado
    return (ret_a - rf_anual) / semi_a if semi_a != 0 else np.nan


def max_drawdown(ret: pd.Series):
    """
    Máximo Drawdown (MDD).
    Retorna a série de drawdown diário e o valor mínimo (pior queda).
    """
    acum  = (1 + ret).cumprod()
    pico  = acum.cummax()
    dd    = (acum - pico) / pico
    return dd, dd.min()


def calmar(ret: pd.Series, dias_ano: int = FREQ) -> float:
    """Índice de Calmar: Retorno Anualizado / |Máximo Drawdown|."""
    _, mdd = max_drawdown(ret)
    ret_a  = retorno_anualizado(ret, dias_ano)
    return ret_a / abs(mdd) if mdd != 0 else np.nan


def var_historico(ret: pd.Series, confianca: float = 0.95) -> float:
    """
    VaR Histórico (95%).
    Perda máxima esperada com (confianca)% de probabilidade,
    usando a distribuição empírica dos retornos.
    Não assume nenhuma distribuição estatística.
    """
    return np.percentile(ret, (1 - confianca) * 100)


def var_parametrico(ret: pd.Series, confianca: float = 0.95) -> float:
    """
    VaR Paramétrico (95%).
    Assume distribuição Normal. Fórmula: VaR = μ + z·σ,
    onde z é o quantil negativo da normal padrão.
    """
    media  = ret.mean()
    desvio = ret.std()
    z      = stats.norm.ppf(1 - confianca)
    return media + z * desvio


def expected_shortfall(ret: pd.Series, confianca: float = 0.95) -> float:
    """
    Expected Shortfall / CVaR (95%).
    Média das perdas nos (1 - confianca)% piores dias.
    Captura a severidade das perdas além do VaR.
    """
    var = var_historico(ret, confianca)
    return ret[ret <= var].mean()


# ============================================================
# FUNÇÕES UTILITÁRIAS — APIs
# ============================================================

# Data fixa para o CDI (enunciado Q3a pede explicitamente 01/01/2010)
CDI_BCB_INICIO = "01/01/2010"


def coletar_cdi_bcb(data_inicio: str = "01/01/2010", data_fim=None) -> pd.DataFrame:
    """
    Q3a — Coleta a série histórica do CDI via API do Banco Central.
    Série SGS nº 12 — CDI acumulado no dia (% a.d.).

    Usa pd.read_json diretamente na URL da API (conforme enunciado).
    Caso a requisição única falhe (timeout em períodos muito longos),
    realiza a coleta ano a ano como fallback, ainda usando pd.read_json.
    """
    if data_fim is None:
        data_fim = datetime.today()

    data_fim_str = pd.Timestamp(data_fim).strftime("%d/%m/%Y")

    # Tentativa 1: pd.read_json diretamente na URL (forma pedida pelo enunciado)
    url_completo = (
        f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.12/dados"
        f"?formato=json&dataInicial={data_inicio}&dataFinal={data_fim_str}"
    )
    try:
        df = pd.read_json(url_completo)   # ← pd.read_json direto na URL
        print(f"   ✔ CDI coletado via pd.read_json: {len(df)} registros")
    except Exception as e:
        # Fallback: coleta ano a ano (ainda com pd.read_json)
        print(f"   ⚠️  Falha no read_json único ({e}). Coletando ano a ano...")
        ano_inicio = int(data_inicio.split("/")[2])
        ano_fim_int = pd.Timestamp(data_fim).year
        frames = []
        for ano in range(ano_inicio, ano_fim_int + 1):
            url_ano = (
                f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.12/dados"
                f"?formato=json&dataInicial=01/01/{ano}&dataFinal=31/12/{ano}"
            )
            try:
                df_ano = pd.read_json(url_ano)   # ← pd.read_json em cada ano
                frames.append(df_ano)
                print(f"   ✔ CDI {ano}: {len(df_ano)} registros")
            except Exception as e2:
                print(f"   ✗ CDI {ano}: {e2}")
        if not frames:
            raise RuntimeError(
                "CDI: nenhum dado retornado. Verifique conectividade com o BCB."
            )
        df = pd.concat(frames, ignore_index=True)

    df.columns = ["Data", "CDI_pct"]
    df["Data"]    = pd.to_datetime(df["Data"], dayfirst=True)
    df["CDI_pct"] = pd.to_numeric(df["CDI_pct"], errors="coerce")
    df["CDI_dec"] = df["CDI_pct"] / 100
    df = df.set_index("Data").sort_index()

    inicio_ts = pd.Timestamp(
        data_inicio.split("/")[2]
        + "-" + data_inicio.split("/")[1]
        + "-" + data_inicio.split("/")[0]
    )
    df = df[(df.index >= inicio_ts) & (df.index <= pd.Timestamp(data_fim))]
    return df


def ultimo_mes_cvm(data_ref) -> tuple:
    """Encontra o último mês/ano com arquivo disponível na CVM (tenta até 7 meses atrás)."""
    data_ref = pd.Timestamp(data_ref)
    for delta in range(7):
        dt   = data_ref - pd.DateOffset(months=delta)
        ano  = dt.year
        mes  = str(dt.month).zfill(2)
        url  = (
            f"https://dados.cvm.gov.br/dados/FI/DOC/INF_DIARIO/"
            f"DADOS/inf_diario_fi_{ano}{mes}.zip"
        )
        try:
            r = requests.get(url, stream=True, timeout=20)
            if r.status_code == 200:
                print(f"   Último arquivo CVM disponível: {ano}/{mes}")
                return int(ano), int(mes)
            else:
                print(f"   ✗ CVM {ano}/{mes}: status {r.status_code}, tentando anterior...")
        except Exception as e:
            print(f"   ✗ CVM {ano}/{mes}: {e}, tentando anterior...")
    raise Exception("Arquivo da CVM não encontrado nos últimos 7 meses.")


def coletar_cotas_fundo(
    cnpj: str, data_inicio_str: str = "2021-01-01", data_fim=None
) -> pd.DataFrame:
    """
    Q3b — Coleta as cotas diárias de um fundo na CVM (dados abertos).
    Download mês a mês, filtrando pelo CNPJ do fundo.
    """
    if data_fim is None:
        data_fim = datetime.today()

    data_inicio  = pd.to_datetime(data_inicio_str)
    ano_fim, mes_fim = ultimo_mes_cvm(data_fim)
    data_limite  = pd.Timestamp(year=ano_fim, month=mes_fim,   day=1)
    data_atual   = pd.Timestamp(year=data_inicio.year, month=data_inicio.month, day=1)

    print(f"   Coletando cotas de {data_atual.date()} até {data_limite.date()}...")
    frames = []

    while data_atual <= data_limite:
        ano = data_atual.year
        mes = str(data_atual.month).zfill(2)
        url = (
            f"https://dados.cvm.gov.br/dados/FI/DOC/INF_DIARIO/"
            f"DADOS/inf_diario_fi_{ano}{mes}.zip"
        )
        try:
            r = requests.get(url, timeout=90)
            if r.status_code == 200:
                z        = zipfile.ZipFile(io.BytesIO(r.content))
                csv_name = [f for f in z.namelist() if f.endswith(".csv")][0]
                df_mes   = pd.read_csv(
                    z.open(csv_name), sep=";", encoding="latin1", low_memory=False
                )
                col_cnpj = (
                    "CNPJ_FUNDO" if "CNPJ_FUNDO" in df_mes.columns
                    else "CNPJ_FUNDO_CLASSE" if "CNPJ_FUNDO_CLASSE" in df_mes.columns
                    else None
                )
                if col_cnpj is None:
                    print(f"   ✗ {ano}/{mes}: coluna CNPJ não encontrada")
                    data_atual += pd.DateOffset(months=1)
                    continue

                cnpj_limpo = df_mes[col_cnpj].str.replace(r"[.\/\-]", "", regex=True)
                df_fundo   = df_mes[cnpj_limpo == cnpj]
                if not df_fundo.empty:
                    frames.append(df_fundo)
                    print(f"   ✔ {ano}/{mes}: {len(df_fundo)} registros")
                else:
                    print(f"   - {ano}/{mes}: fundo não encontrado neste mês")
            else:
                print(f"   ✗ {ano}/{mes}: status {r.status_code}")
        except Exception as e:
            print(f"   ✗ {ano}/{mes}: {e}")
        data_atual += pd.DateOffset(months=1)

    if not frames:
        raise Exception(f"Nenhuma cota encontrada para CNPJ {cnpj}.")

    df       = pd.concat(frames, ignore_index=True)
    col_data = "DT_COMPTC" if "DT_COMPTC" in df.columns else df.columns[0]
    df[col_data] = pd.to_datetime(df[col_data])
    df = df.rename(columns={col_data: "DT_COMPTC"}).sort_values("DT_COMPTC").reset_index(drop=True)
    df = df[
        (df["DT_COMPTC"] >= pd.Timestamp(data_inicio_str))
        & (df["DT_COMPTC"] <= pd.Timestamp(data_fim))
    ]
    return df


# ============================================================
# Q1 — TRATAMENTO DE DADOS E ANÁLISE DE PORTFÓLIOS
# ============================================================

print("\n" + "=" * 60)
print("Q1) Tratamento de dados e análise de portfólios")
print("=" * 60)

# Q1a — Leitura e merge
locais   = ler_e_limpar(BASE_DIR / "retornos_dados_locais.xlsx")
offshore = ler_e_limpar(BASE_DIR / "retornos_dados_offshore.xlsx")

df = (
    pd.merge(locais, offshore, on="Data", how="inner")
    .set_index("Data")
    .sort_index()
)
print(f"   Período dados: {df.index.min().date()} → {df.index.max().date()}")
print(f"   Dias úteis   : {len(df)}")
print(f"   Colunas      : {df.columns.tolist()}")

# Validação dos pesos (usa ValueError em vez de assert, que pode ser desabilitado com -O)
for nome, pesos in CARTEIRAS.items():
    total = sum(pesos.values())
    if abs(total - 1.0) > 1e-9:
        raise ValueError(f"Pesos {nome} somam {total:.4f} ≠ 1.0")
print("   ✅ Pesos validados!")

# Q1b — Retornos diários dos 3 portfólios
def retorno_carteira(df_ret: pd.DataFrame, pesos_dict: dict) -> pd.Series:
    """Calcula o retorno diário ponderado da carteira."""
    cols  = list(pesos_dict.keys())
    pesos = np.array([pesos_dict[c] for c in cols])
    ausentes = [c for c in cols if c not in df_ret.columns]
    if ausentes:
        raise ValueError(f"Colunas ausentes no DataFrame: {ausentes}")
    return df_ret[cols].dot(pesos)

portfolios = {
    nome: retorno_carteira(df, pesos) for nome, pesos in CARTEIRAS.items()
}

# Taxa livre de risco (CDI anualizado) para Sharpe e Sortino
rf_serie    = df["CDI"].astype(float)
rf_total    = (1 + rf_serie).prod() - 1
rf_anual    = (1 + rf_total) ** (FREQ / len(rf_serie)) - 1
print(f"   CDI anualizado (rf): {rf_anual:.4%}")

# Q1c — Métricas de performance em um único DataFrame (Desafio Q1)
# Nomenclatura fiel ao enunciado: "Volatilidade Total" e "Volatilidade Total Anualizada"
rows_perf = []
for nome, ret in portfolios.items():
    rows_perf.append({
        "Portfólio"                  : nome,
        "Retorno Total"              : f"{retorno_acumulado(ret):.4%}",
        "Retorno Total Anualizado"   : f"{retorno_anualizado(ret):.4%}",
        "Volatilidade Total"         : f"{volatilidade_diaria(ret):.4%}",   # std diário
        "Volatilidade Total Anualiz.": f"{volatilidade_anual(ret):.4%}",    # × sqrt(252)
        "Sharpe"                     : f"{sharpe(ret, rf_anual):.4f}",
    })

resumo_perf = pd.DataFrame(rows_perf).set_index("Portfólio")
print("\n📊 Q1 — Performance dos Portfólios:")
print(resumo_perf.to_string())

# Desafio Q1 — Gráfico: valor acumulado dos 3 portfólios + CDI
fig, ax = plt.subplots(figsize=(12, 6))
for nome, ret in portfolios.items():
    acum = (1 + ret).cumprod()
    ax.plot(acum.index, acum, label=nome, linewidth=2, color=CORES[nome])
acum_cdi = (1 + rf_serie).cumprod()
ax.plot(acum_cdi.index, acum_cdi, label="CDI (benchmark)",
        linewidth=1.5, color="#c00000", linestyle="--")
ax.set_title(
    "Carteira Moderada — Valor Acumulado por Portfólio\n(MMZR Family Office | Coding Test)",
    fontsize=14, fontweight="bold", pad=15,
)
ax.set_xlabel("Data", fontsize=11)
ax.set_ylabel("Valor Acumulado (base = 1,00)", fontsize=11)
ax.yaxis.set_major_formatter(mtick.FormatStrFormatter("%.2f"))
ax.legend(fontsize=10)
ax.grid(True, alpha=0.3)
ax.set_facecolor("#f9f9f9")
fig.patch.set_facecolor("white")
plt.tight_layout()
plt.savefig("q1_valor_acumulado.png", dpi=150, bbox_inches="tight")
plt.show()
print("✅ Gráfico salvo: 'q1_valor_acumulado.png'")


# ============================================================
# Q2 — MÓDULO DE RISCO
# ============================================================

print("\n" + "=" * 60)
print("Q2) Módulo de risco — métricas dos portfólios")
print("=" * 60)

rows_risco = []
for nome, ret in portfolios.items():
    _, mdd = max_drawdown(ret)
    rows_risco.append({
        "Portfólio"          : nome,
        "Semidesvio Anual"   : f"{semidesvio(ret):.4%}",
        "Máx. Drawdown"      : f"{mdd:.4%}",
        "VaR 95% Histórico"  : f"{var_historico(ret):.4%}",
        "VaR 95% Paramétrico": f"{var_parametrico(ret):.4%}",
        "ES 95%"             : f"{expected_shortfall(ret):.4%}",
        "Sortino"            : f"{sortino(ret, rf_anual):.4f}",
        "Calmar"             : f"{calmar(ret):.4f}",
    })

resumo_risco = pd.DataFrame(rows_risco).set_index("Portfólio")
print("\n📊 Q2 — Métricas de Risco:")
print(resumo_risco.to_string())

# Gráfico Q2 — Drawdown
fig, ax = plt.subplots(figsize=(12, 5))
for nome, ret in portfolios.items():
    dd_serie, mdd = max_drawdown(ret)
    ax.fill_between(dd_serie.index, dd_serie * 100, 0, alpha=0.20, color=CORES[nome])
    ax.plot(dd_serie.index, dd_serie * 100,
            label=f"{nome} (MDD: {mdd:.2%})", linewidth=1.5, color=CORES[nome])
ax.set_title(
    "Drawdown dos Portfólios — Carteira Moderada\n(MMZR Family Office | Coding Test)",
    fontsize=13, fontweight="bold", pad=12,
)
ax.set_xlabel("Data", fontsize=11)
ax.set_ylabel("Drawdown (%)", fontsize=11)
ax.yaxis.set_major_formatter(mtick.PercentFormatter())
ax.legend(fontsize=10)
ax.grid(True, alpha=0.3)
ax.set_facecolor("#f9f9f9")
fig.patch.set_facecolor("white")
plt.tight_layout()
plt.savefig("q2_drawdown.png", dpi=150, bbox_inches="tight")
plt.show()
print("✅ Gráfico salvo: 'q2_drawdown.png'")

# Desafio Q2 — Histograma com VaR e ES
fig, axes = plt.subplots(1, 3, figsize=(16, 5))
fig.suptitle(
    "Histograma de Retornos Diários com VaR 95% e ES 95%\n(MMZR Family Office | Coding Test)",
    fontsize=13, fontweight="bold",
)
for ax, (nome, ret) in zip(axes, portfolios.items()):
    var_h  = var_historico(ret)
    var_p  = var_parametrico(ret)
    es_val = expected_shortfall(ret)

    counts, bins, patches = ax.hist(
        ret * 100, bins=60, color=CORES[nome], alpha=0.6,
        edgecolor="white", linewidth=0.3,
    )
    for patch, left in zip(patches, bins[:-1]):
        if left <= var_h * 100:
            patch.set_facecolor("#c00000")
            patch.set_alpha(0.8)

    ax.axvline(var_h  * 100, color="#c00000", linewidth=2, linestyle="--",
               label=f"VaR Hist. 95%: {var_h:.3%}")
    ax.axvline(var_p  * 100, color="#ff7f00", linewidth=2, linestyle=":",
               label=f"VaR Param. 95%: {var_p:.3%}")
    ax.axvline(es_val * 100, color="#7f0000", linewidth=2, linestyle="-.",
               label=f"ES 95%: {es_val:.3%}")

    ax.set_title(nome, fontsize=11, fontweight="bold")
    ax.set_xlabel("Retorno Diário (%)", fontsize=9)
    ax.set_ylabel("Frequência", fontsize=9)
    ax.legend(fontsize=7.5)
    ax.grid(True, alpha=0.3)
    ax.set_facecolor("#f9f9f9")

fig.patch.set_facecolor("white")
plt.tight_layout()
plt.savefig("q2_histograma_var_es.png", dpi=150, bbox_inches="tight")
plt.show()
print("✅ Gráfico salvo: 'q2_histograma_var_es.png'")

# Respostas escritas Q2
print("""
╔══════════════════════════════════════════════════════════════╗
║              Q2 — RESPOSTAS ESCRITAS                         ║
╚══════════════════════════════════════════════════════════════╝

─────────────────────────────────────────────────────────────
Q2b) VaR Paramétrico vs VaR Histórico
─────────────────────────────────────────────────────────────

VaR PARAMÉTRICO:
  Vantagens:
    - Simples e rápido de calcular (só precisa de média e desvio)
    - Funciona bem com poucos dados históricos
    - Fácil de interpretar e comunicar
  Desvantagens:
    - Assume distribuição Normal dos retornos, raramente verdade
      em ativos financeiros (caudas mais pesadas que a Normal)
    - Subestima perdas em eventos extremos (crises, crashes)
  Quando usar:
    - Portfólios com retornos aproximadamente simétricos
    - Quando há pouco histórico disponível
    - Relatórios de risco simplificados do dia a dia

VaR HISTÓRICO:
  Vantagens:
    - Não assume nenhuma distribuição estatística
    - Captura eventos extremos reais já ocorridos (ex: 2020, 2008)
    - Mais robusto para distribuições assimétricas e com
      caudas pesadas
  Desvantagens:
    - Depende totalmente do histórico disponível
    - Não captura eventos que ainda não ocorreram
    - Pode ser instável se o período histórico for curto
  Quando usar:
    - Portfólios com distribuição assimétrica de retornos
    - Quando há histórico longo e representativo disponível
    - Análises de risco mais conservadoras e realistas

─────────────────────────────────────────────────────────────
Q2c) O que é o Semidesvio
─────────────────────────────────────────────────────────────

O semidesvio é uma medida de risco que calcula o desvio padrão
considerando APENAS os retornos abaixo de zero (ou abaixo de
um benchmark).

Vantagem: a volatilidade convencional penaliza igualmente
retornos positivos e negativos. Para o investidor, apenas a
volatilidade negativa é um risco real. O semidesvio corrige
isso, medindo somente o risco de perda — tornando a análise
mais alinhada com a percepção real do investidor.

É especialmente útil quando os retornos têm distribuição
assimétrica, onde a volatilidade total seria enganosa.

─────────────────────────────────────────────────────────────
Q2d) O que é o Expected Shortfall (ES / CVaR)
─────────────────────────────────────────────────────────────

O Expected Shortfall (ES), também chamado CVaR (Conditional
Value at Risk), é a média das perdas que ocorrem além do VaR.

Exemplo: se o VaR 95% é -0,30% ao dia, isso significa que em
5% dos dias a perda supera 0,30%. O ES 95% responde:
"Nesse pior cenário de 5%, qual é a perda MÉDIA?"

Vantagem sobre o VaR: o VaR diz apenas ONDE começa a cauda
de perdas, mas não captura a MAGNITUDE dessas perdas extremas.
O ES corrige isso, sendo uma métrica mais conservadora e
informativa. Além disso, o ES é uma medida coerente de risco
(satisfaz subaditividade), adequada para otimização de portfólios.

─────────────────────────────────────────────────────────────
Q2e) Diferença entre Módulo, Biblioteca e Pacote em Python
─────────────────────────────────────────────────────────────

MÓDULO   → Um único arquivo .py com funções, classes e variáveis.
            Exemplo: criar 'metricas_risco.py' com as funções
            semidesvio(), var_historico(), etc.
            Uso: import metricas_risco

PACOTE   → Uma pasta com múltiplos módulos + arquivo __init__.py.
            O __init__.py indica ao Python que a pasta é um pacote.
            Exemplo: o próprio pandas internamente é um pacote,
            contendo pandas.core, pandas.io, etc.

BIBLIOTECA → Um conjunto amplo de pacotes/módulos prontos para
              uso, desenvolvidos para resolver problemas específicos.
              Exemplos: pandas, numpy, matplotlib, scipy.
              O termo é usado de forma mais ampla e informal.

Hierarquia:
  Módulo (.py) → agrupado em → Pacote (pasta + __init__.py)
  → vários pacotes formam → Biblioteca
""")


# ============================================================
# Q3 — COLETA DE DADOS PÚBLICOS VIA API
# ============================================================

print("=" * 60)
print("Q3a) Coletando CDI via API do Banco Central...")
print("=" * 60)

# Q3a — CDI sempre desde 01/01/2010 (conforme enunciado), independente do período do fundo
cdi_raw = coletar_cdi_bcb(CDI_BCB_INICIO, data_fim=DATA_FIM)

print(f"\n✅ CDI coletado: {len(cdi_raw)} dias")
print(f"   Período: {cdi_raw.index.min().date()} → {cdi_raw.index.max().date()}")

print("\n" + "=" * 60)
print("Q3b) Coletando cotas do fundo via CVM...")
print("=" * 60)

df_cotas = coletar_cotas_fundo(
    CNPJ_FUNDO, data_inicio_str=DATA_INICIO, data_fim=DATA_FIM
)
print(f"\n✅ Cotas coletadas: {len(df_cotas)} dias")
print(f"   Período: {df_cotas['DT_COMPTC'].min().date()} → {df_cotas['DT_COMPTC'].max().date()}")

# Retornos do fundo
fundo = df_cotas[["DT_COMPTC", "VL_QUOTA"]].copy()
fundo.columns = ["Data", "Cota"]
fundo = fundo.set_index("Data").sort_index()
fundo["Cota"]    = pd.to_numeric(fundo["Cota"], errors="coerce")
fundo["Retorno"] = fundo["Cota"].pct_change()
fundo = fundo.dropna()

# Alinhamento de datas com CDI
cdi_periodo  = cdi_raw.loc[fundo.index.min():fundo.index.max(), "CDI_dec"]
datas_comuns = fundo.index.intersection(cdi_periodo.index)
if datas_comuns.empty:
    raise RuntimeError(
        "Nenhuma data em comum entre o fundo e o CDI. "
        "Verifique se o CDI foi coletado para o período correto."
    )
ret_fundo  = fundo.loc[datas_comuns, "Retorno"]
ret_cdi_q3 = cdi_periodo.loc[datas_comuns]

# Q3 — retornos totais
ret_total_fundo = (1 + ret_fundo).prod() - 1
ret_total_cdi   = (1 + ret_cdi_q3).prod() - 1

print(f"\n   Retorno Total do Fundo : {ret_total_fundo:.2%}")
print(f"   Retorno Total do CDI   : {ret_total_cdi:.2%}")
print(f"   Fundo vs CDI           : {ret_total_fundo / ret_total_cdi:.2%} do CDI")

# Gráfico Q3 — Fundo vs CDI
acum_fundo = (1 + ret_fundo).cumprod()
acum_cdi_q3 = (1 + ret_cdi_q3).cumprod()

fig, ax = plt.subplots(figsize=(12, 6))
ax.plot(acum_fundo.index, acum_fundo,
        label=f"{NOME_FUNDO} ({ret_total_fundo:.2%})", linewidth=2, color="#1f4e79")
ax.plot(acum_cdi_q3.index, acum_cdi_q3,
        label=f"CDI ({ret_total_cdi:.2%})", linewidth=1.5, color="#c00000", linestyle="--")
ax.set_title(
    f"{NOME_FUNDO}\nValor Acumulado vs CDI",
    fontsize=13, fontweight="bold", pad=12,
)
ax.set_xlabel("Data", fontsize=11)
ax.set_ylabel("Valor Acumulado (base = 1,00)", fontsize=11)
ax.yaxis.set_major_formatter(mtick.FormatStrFormatter("%.2f"))
ax.legend(fontsize=10)
ax.grid(True, alpha=0.3)
ax.set_facecolor("#f9f9f9")
fig.patch.set_facecolor("white")
ax.annotate(
    f"Período: {DATA_INICIO} → {DATA_FIM.strftime('%d/%m/%Y')}  |  "
    f"Executado em: {datetime.today().strftime('%d/%m/%Y')}",
    xy=(0.01, 0.02), xycoords="axes fraction", fontsize=8, color="gray",
)
plt.tight_layout()
plt.savefig("q3_fundo_vs_cdi.png", dpi=150, bbox_inches="tight")
plt.show()
print("✅ Gráfico salvo: 'q3_fundo_vs_cdi.png'")


# ============================================================
# Q4 — ANÁLISE DE REGRESSÃO MÚLTIPLA
# ============================================================

print("\n" + "=" * 60)
print("Q4) Regressão Múltipla — Fundo vs Classes Locais")
print("=" * 60)

# Q4a — preparar dados: y = retorno do fundo, X = classes locais
df_local  = locais.set_index("Data").sort_index()
fundo_ret = fundo["Retorno"]
fundo_ret.name = "Fundo"

df_reg = df_local.join(fundo_ret, how="inner").dropna()
y      = df_reg["Fundo"]
X      = df_reg.drop(columns=["Fundo"])

print(f"   Período regressão: {df_reg.index.min().date()} → {df_reg.index.max().date()}")
print(f"   Observações      : {len(df_reg)}")
print(f"   Regressores (X)  : {X.columns.tolist()}")

# Q4a — Regressão completa
X_const        = sm.add_constant(X)
modelo_completo = sm.OLS(y, X_const).fit()
print("\n📋 Q4a — Regressão Múltipla Completa:")
print(modelo_completo.summary())

betas_completo = pd.DataFrame({
    "Beta"     : modelo_completo.params,
    "Std Error": modelo_completo.bse,
    "t-stat"   : modelo_completo.tvalues,
    "p-valor"  : modelo_completo.pvalues,
}).drop(index="const", errors="ignore")

# Q4b — VIF (Multicolinearidade)
def calcular_vif(X_df: pd.DataFrame) -> pd.DataFrame:
    """Calcula o VIF para cada variável do DataFrame."""
    vif_vals = [variance_inflation_factor(X_df.values, i) for i in range(X_df.shape[1])]
    return (
        pd.DataFrame({"Variável": X_df.columns, "VIF": vif_vals})
        .set_index("Variável")
        .sort_values("VIF", ascending=False)
    )

vif_inicial = calcular_vif(X)
print("\n📋 Q4b — VIF (Variance Inflation Factor):")
print(vif_inicial.round(2).to_string())

tem_multicol = (vif_inicial["VIF"] > VIF_THRESH).any()
if tem_multicol:
    vars_problema = vif_inicial[vif_inicial["VIF"] > VIF_THRESH].index.tolist()
    print(f"\n⚠️  Multicolinearidade detectada! VIF > {VIF_THRESH} em: {vars_problema}")
    print(
        "\n   Métodos para tratar multicolinearidade:"
        "\n   1. Eliminação iterativa por VIF (usado aqui): remove a variável de maior"
        "\n      VIF até todas ficarem abaixo do threshold. Simples e interpretável."
        "\n   2. Ridge Regression (L2): penaliza o tamanho dos coeficientes, estabilizando"
        "\n      os betas sem remover variáveis."
        "\n   3. PCA: transforma os regressores em componentes ortogonais, eliminando a"
        "\n      multicolinearidade por construção (perde interpretabilidade direta)."
        "\n   4. Remoção por correlação: remove um de cada par com correlação > 0,85."
    )
else:
    print(f"\n✅ Sem multicolinearidade severa (todos VIF ≤ {VIF_THRESH}).")

# Q4c — Regressão corrigida por VIF iterativo
def selecionar_por_vif(X_df: pd.DataFrame, threshold: int = 10) -> list:
    """Remove iterativamente a variável com maior VIF até threshold."""
    cols = list(X_df.columns)
    it   = 1
    while True:
        if len(cols) <= 1:
            print("   ⚠️  Apenas 1 variável restante. Encerrando eliminação.")
            break
        vifs    = [variance_inflation_factor(X_df[cols].values, i) for i in range(len(cols))]
        max_vif = max(vifs)
        if max_vif <= threshold:
            break
        removida = cols[vifs.index(max_vif)]
        print(f"   Iteração {it}: removendo '{removida}' (VIF={max_vif:.1f})")
        cols.remove(removida)
        it += 1
    print(f"   → Convergiu com variáveis: {cols}")
    return cols

print(f"\n📋 Q4c — Eliminação iterativa (VIF threshold = {VIF_THRESH}):")
cols_finais      = selecionar_por_vif(X, threshold=VIF_THRESH)
X2               = X[cols_finais]
vif_final        = calcular_vif(X2)
X2_const         = sm.add_constant(X2)
modelo_corrigido = sm.OLS(y, X2_const).fit()

print("\n📋 VIF Final (após eliminação):")
print(vif_final.round(2).to_string())
print("\n📋 Regressão Corrigida:")
print(modelo_corrigido.summary())

betas_corrigido = pd.DataFrame({
    "Beta"     : modelo_corrigido.params,
    "Std Error": modelo_corrigido.bse,
    "t-stat"   : modelo_corrigido.tvalues,
    "p-valor"  : modelo_corrigido.pvalues,
}).drop(index="const", errors="ignore")

print(f"\n   R² ajustado — Completo : {modelo_completo.rsquared_adj:.4f}")
print(f"   R² ajustado — Corrigido: {modelo_corrigido.rsquared_adj:.4f}")

# Gráfico Q4 — Betas (completo vs corrigido)
fig, axes = plt.subplots(1, 2, figsize=(14, 5))
fig.suptitle(
    f"Q4 — Betas: {NOME_FUNDO} vs Classes Locais",
    fontsize=13, fontweight="bold",
)
for ax, (titulo, betas) in zip(axes, [
    ("Modelo Completo", betas_completo),
    (f"Modelo Corrigido\n(VIF ≤ {VIF_THRESH})", betas_corrigido),
]):
    cores_barra = ["#c00000" if v < 0 else "#1f4e79" for v in betas["Beta"]]
    betas["Beta"].plot(kind="barh", ax=ax, color=cores_barra,
                       edgecolor="white", alpha=0.85)
    ax.axvline(0, color="black", linewidth=0.8, linestyle="--")
    ax.set_title(titulo, fontsize=11, fontweight="bold")
    ax.set_xlabel("Beta", fontsize=10)
    ax.grid(True, alpha=0.3, axis="x")
    ax.set_facecolor("#f9f9f9")
fig.patch.set_facecolor("white")
plt.tight_layout()
plt.savefig("q4_betas.png", dpi=150, bbox_inches="tight")
plt.show()
print("✅ Gráfico salvo: 'q4_betas.png'")

# Desafio Q4 — Rolling Beta (1 ano = 252 dias úteis)
print(f"\n📋 Q4 Desafio — Rolling Beta (janela {JANELA_ROLL} dias úteis):")

X2_roll        = sm.add_constant(X2)
rolling_result = RollingOLS(
    endog   = y,
    exog    = X2_roll,
    window  = JANELA_ROLL,
    min_nobs= int(JANELA_ROLL * 0.8),
).fit(params_only=True)

rolling_betas = (
    rolling_result.params
    .drop(columns=["const"], errors="ignore")
    .dropna(how="all")
)
print(f"   Observações com beta: {len(rolling_betas)}")

palette  = ["#1f4e79","#2e75b6","#00b0f0","#c00000","#ff7f00","#4b0082","#006400","#8B4513"]
n_vars   = len(rolling_betas.columns)
fig, axes = plt.subplots(n_vars, 1, figsize=(13, 3 * n_vars), sharex=True)
if n_vars == 1:
    axes = [axes]
fig.suptitle(
    f"Rolling Beta (janela {JANELA_ROLL} dias úteis)\n{NOME_FUNDO} vs Classes Locais",
    fontsize=13, fontweight="bold",
)
for ax, col, cor in zip(axes, rolling_betas.columns, palette):
    s = rolling_betas[col]
    ax.plot(s.index, s, color=cor, linewidth=1.5)
    ax.axhline(0, color="black", linewidth=0.7, linestyle="--", alpha=0.5)
    ax.fill_between(s.index, s, 0, where=(s >= 0), alpha=0.12, color=cor)
    ax.fill_between(s.index, s, 0, where=(s <  0), alpha=0.12, color="#c00000")
    ax.set_ylabel("Beta", fontsize=9)
    ax.set_title(col, fontsize=10, fontweight="bold", loc="left")
    ax.grid(True, alpha=0.3)
    ax.set_facecolor("#f9f9f9")
axes[-1].set_xlabel("Data", fontsize=10)
fig.patch.set_facecolor("white")
plt.tight_layout()
plt.savefig("q4_rolling_beta.png", dpi=150, bbox_inches="tight")
plt.show()
print("✅ Gráfico salvo: 'q4_rolling_beta.png'")


# ============================================================
# EXPORTAÇÃO FINAL
# ============================================================

resumo_perf.to_csv("q1_metricas_performance.csv")
resumo_risco.to_csv("q2_metricas_risco.csv")
cdi_raw.to_csv("q3_cdi_bcb.csv")
df_cotas.to_csv("q3_cotas_fundo.csv", index=False)
betas_completo.to_csv("q4_betas_completo.csv")
betas_corrigido.to_csv("q4_betas_corrigido.csv")
vif_inicial.to_csv("q4_vif_inicial.csv")
vif_final.to_csv("q4_vif_final.csv")
rolling_betas.to_csv("q4_rolling_betas.csv")

print("""
╔══════════════════════════════════════════════════════════════╗
║            CODING TEST — CONCLUÍDO COM SUCESSO!              ║
╠══════════════════════════════════════════════════════════════╣
║  Gráficos gerados:                                           ║
║   → q1_valor_acumulado.png                                   ║
║   → q2_drawdown.png                                          ║
║   → q2_histograma_var_es.png                                 ║
║   → q3_fundo_vs_cdi.png                                      ║
║   → q4_betas.png                                             ║
║   → q4_rolling_beta.png                                      ║
║  CSVs gerados:                                               ║
║   → q1_metricas_performance.csv                              ║
║   → q2_metricas_risco.csv                                    ║
║   → q3_cdi_bcb.csv                                           ║
║   → q3_cotas_fundo.csv                                       ║
║   → q4_betas_completo.csv                                    ║
║   → q4_betas_corrigido.csv                                   ║
║   → q4_vif_inicial.csv / q4_vif_final.csv                   ║
║   → q4_rolling_betas.csv                                     ║
╚══════════════════════════════════════════════════════════════╝
""")
