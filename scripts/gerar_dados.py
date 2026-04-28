"""
gerar_dados.py
Gera os 5 CSVs sintéticos para a Rede Máxima Supermercados Ltda.
Período: Janeiro 2023 – Dezembro 2024 (24 meses × 5 lojas)
Seed: 42
"""

import numpy as np
import pandas as pd
from pathlib import Path

RNG = np.random.default_rng(42)
MESES = pd.date_range("2023-01", periods=24, freq="MS")
LOJAS = {
    "L01-Centro":     {"receita_base": 1_800_000, "peso_custo": 0.64, "peso_desp": 0.12},
    "L02-Norte":      {"receita_base": 1_200_000, "peso_custo": 0.66, "peso_desp": 0.13},
    "L03-Sul":        {"receita_base": 980_000,   "peso_custo": 0.65, "peso_desp": 0.14},
    "L04-Leste":      {"receita_base": 750_000,   "peso_custo": 0.67, "peso_desp": 0.13},
    "L05-Aeroporto":  {"receita_base": 620_000,   "peso_custo": 0.60, "peso_desp": 0.11},
}
OUT = Path(__file__).parent.parent / "data" / "raw"
OUT.mkdir(parents=True, exist_ok=True)

# ── sazonalidade mensal do varejo alimentar ──────────────────────────────────
SAZO = np.array([0.92, 0.88, 0.95, 0.97, 0.98, 1.02,
                  1.05, 1.04, 1.00, 1.01, 1.08, 1.18])


def _crescimento(n_mes: int) -> float:
    return 1 + 0.09 * (n_mes / 23)


# ── 1. DRE por loja ─────────────────────────────────────────────────────────
def gerar_dre():
    rows = []
    for i, mes in enumerate(MESES):
        sazo = SAZO[mes.month - 1]
        cresc = _crescimento(i)
        for loja, p in LOJAS.items():
            receita_bruta = (
                p["receita_base"] * sazo * cresc
                * (1 + RNG.normal(0, 0.025))
            )
            deducoes = receita_bruta * RNG.uniform(0.03, 0.05)
            receita_liq = receita_bruta - deducoes
            cmv = receita_liq * (p["peso_custo"] + RNG.normal(0, 0.012))

            # outlier: L03-Sul em Ago/2023 — pico de reposição de estoque
            if loja == "L03-Sul" and mes == pd.Timestamp("2023-08"):
                cmv *= 1.18

            lucro_bruto = receita_liq - cmv
            desp_pessoal = receita_liq * (p["peso_desp"] * 0.55 + RNG.normal(0, 0.008))
            desp_aluguel = receita_liq * (p["peso_desp"] * 0.20)
            desp_energia = receita_liq * (p["peso_desp"] * 0.12 + RNG.normal(0, 0.005))
            desp_outras  = receita_liq * (p["peso_desp"] * 0.13 + RNG.normal(0, 0.006))
            total_opex   = desp_pessoal + desp_aluguel + desp_energia + desp_outras
            ebitda       = lucro_bruto - total_opex
            deprec       = receita_liq * RNG.uniform(0.018, 0.024)
            ebit         = ebitda - deprec
            desp_fin     = receita_liq * RNG.uniform(0.005, 0.012)
            lair         = ebit - desp_fin
            ir_csll      = max(0, lair * 0.265)
            lucro_liq    = lair - ir_csll

            # orçamento (± 5%)
            receita_orc = receita_bruta * (1 + RNG.normal(0, 0.045))
            ebitda_orc  = ebitda * (1 + RNG.normal(0, 0.05))

            rows.append({
                "mes": mes.strftime("%Y-%m"),
                "loja": loja,
                "receita_bruta": round(receita_bruta, 2),
                "deducoes": round(deducoes, 2),
                "receita_liquida": round(receita_liq, 2),
                "cmv": round(cmv, 2),
                "lucro_bruto": round(lucro_bruto, 2),
                "desp_pessoal": round(desp_pessoal, 2),
                "desp_aluguel": round(desp_aluguel, 2),
                "desp_energia": round(desp_energia, 2),
                "desp_outras": round(desp_outras, 2),
                "total_opex": round(total_opex, 2),
                "ebitda": round(ebitda, 2),
                "depreciacao": round(deprec, 2),
                "ebit": round(ebit, 2),
                "desp_financeiras": round(desp_fin, 2),
                "lair": round(lair, 2),
                "ir_csll": round(ir_csll, 2),
                "lucro_liquido": round(lucro_liq, 2),
                "receita_orcada": round(receita_orc, 2),
                "ebitda_orcado": round(ebitda_orc, 2),
            })
    df = pd.DataFrame(rows)
    df.to_csv(OUT / "dre.csv", index=False)
    print(f"  dre.csv            → {len(df):>4} linhas")
    return df


# ── 2. Fluxo de Caixa consolidado ───────────────────────────────────────────
def gerar_fluxo_caixa(dre: pd.DataFrame):
    rows = []
    saldo = 3_200_000.0
    dre_cons = dre.groupby("mes").agg(
        receita_liquida=("receita_liquida", "sum"),
        lucro_liquido=("lucro_liquido", "sum"),
        ebitda=("ebitda", "sum"),
    ).reset_index()

    for i, row in dre_cons.iterrows():
        rec = row["receita_liquida"] * RNG.uniform(0.88, 0.96)
        pagamentos = row["receita_liquida"] * RNG.uniform(0.75, 0.85)
        impostos_pag = max(0, row["lucro_liquido"] * RNG.uniform(0.20, 0.28))
        fco = rec - pagamentos - impostos_pag

        mes_dt = pd.Timestamp(row["mes"])
        # CAPEX: Mar/23, Set/23, Mar/24
        capex = 0.0
        if mes_dt in [pd.Timestamp("2023-03"), pd.Timestamp("2023-09"),
                      pd.Timestamp("2024-03")]:
            capex = RNG.uniform(280_000, 520_000)

        captacao = 600_000.0 if mes_dt == pd.Timestamp("2023-01") else 0.0
        amortizacao = 180_000.0 if mes_dt == pd.Timestamp("2024-01") else 0.0
        fcf = fco - capex
        fcl = fcf + captacao - amortizacao
        saldo_ini = saldo
        saldo += fcl

        rows.append({
            "mes": row["mes"],
            "saldo_inicial": round(saldo_ini, 2),
            "recebimentos": round(rec, 2),
            "pagamentos_operacionais": round(pagamentos, 2),
            "impostos_pagos": round(impostos_pag, 2),
            "fco": round(fco, 2),
            "capex": round(-capex, 2),
            "fcf": round(fcf, 2),
            "captacao_emprestimos": round(captacao, 2),
            "amortizacao_emprestimos": round(-amortizacao, 2),
            "variacao_liquida": round(fcl, 2),
            "saldo_final": round(saldo, 2),
        })
    df = pd.DataFrame(rows)
    df.to_csv(OUT / "fluxo_caixa.csv", index=False)
    print(f"  fluxo_caixa.csv    → {len(df):>4} linhas")
    return df


# ── 3. Contas a Receber / Aging ──────────────────────────────────────────────
def gerar_contas_receber():
    clientes = [f"Cliente {str(i).zfill(2)}" for i in range(1, 26)]
    faixas = ["A vencer", "1-30d", "31-60d", "61-90d", "91-120d", ">120d"]
    rows = []
    for mes in MESES:
        for cliente in clientes:
            base = RNG.uniform(15_000, 180_000)
            a_vencer = base * RNG.uniform(0.55, 0.70)
            f30  = base * RNG.uniform(0.10, 0.18)
            f60  = base * RNG.uniform(0.05, 0.10)
            f90  = base * RNG.uniform(0.02, 0.07)
            f120 = base * RNG.uniform(0.01, 0.04)
            f120p = base * RNG.uniform(0.005, 0.025)
            total = a_vencer + f30 + f60 + f90 + f120 + f120p
            vencido = f30 + f60 + f90 + f120 + f120p
            inadim = vencido / total

            rows.append({
                "mes": mes.strftime("%Y-%m"),
                "cliente": cliente,
                "a_vencer": round(a_vencer, 2),
                "venc_1_30": round(f30, 2),
                "venc_31_60": round(f60, 2),
                "venc_61_90": round(f90, 2),
                "venc_91_120": round(f120, 2),
                "venc_acima_120": round(f120p, 2),
                "total_carteira": round(total, 2),
                "total_vencido": round(vencido, 2),
                "taxa_inadimplencia": round(inadim, 4),
            })
    df = pd.DataFrame(rows)
    df.to_csv(OUT / "contas_receber.csv", index=False)
    print(f"  contas_receber.csv → {len(df):>4} linhas")
    return df


# ── 4. Centro de Custos ──────────────────────────────────────────────────────
def gerar_centro_custos(dre: pd.DataFrame):
    deptos = {
        "Operações":      0.35,
        "Logística":      0.20,
        "RH":             0.15,
        "Financeiro":     0.08,
        "TI":             0.07,
        "Marketing":      0.08,
        "Administrativo": 0.05,
        "Expansão":       0.02,
    }
    total_opex = dre.groupby("mes")["total_opex"].sum().reset_index()
    rows = []
    for _, row in total_opex.iterrows():
        for depto, peso in deptos.items():
            real = row["total_opex"] * peso * (1 + RNG.normal(0, 0.04))
            orc = row["total_opex"] * peso * (1 + RNG.normal(0, 0.03))
            # outlier: TI em Nov/2023 (migração de sistema)
            if depto == "TI" and row["mes"] == "2023-11":
                real *= 2.05
            var_abs = real - orc
            var_pct = var_abs / orc if orc != 0 else 0
            rows.append({
                "mes": row["mes"],
                "departamento": depto,
                "custo_real": round(real, 2),
                "custo_orcado": round(orc, 2),
                "variacao_absoluta": round(var_abs, 2),
                "variacao_percentual": round(var_pct, 4),
            })
    df = pd.DataFrame(rows)
    df.to_csv(OUT / "centro_custos.csv", index=False)
    print(f"  centro_custos.csv  → {len(df):>4} linhas")
    return df


# ── 5. Orçamento anual por loja ──────────────────────────────────────────────
def gerar_orcamento():
    rows = []
    for ano in [2023, 2024]:
        for loja, p in LOJAS.items():
            cresc_base = 1.09 if ano == 2024 else 1.0
            receita_anual = p["receita_base"] * 12 * cresc_base * (1 + RNG.normal(0, 0.02))
            cmv_anual = receita_anual * (p["peso_custo"] + RNG.normal(0, 0.01))
            opex_anual = receita_anual * (p["peso_desp"] + RNG.normal(0, 0.01))
            ebitda_anual = receita_anual - cmv_anual - opex_anual
            rows.append({
                "ano": ano,
                "loja": loja,
                "receita_orcada": round(receita_anual, 2),
                "cmv_orcado": round(cmv_anual, 2),
                "opex_orcado": round(opex_anual, 2),
                "ebitda_orcado": round(ebitda_anual, 2),
                "meta_margem_bruta": round((receita_anual - cmv_anual) / receita_anual, 4),
                "meta_margem_ebitda": round(ebitda_anual / receita_anual, 4),
            })
    df = pd.DataFrame(rows)
    df.to_csv(OUT / "orcamento.csv", index=False)
    print(f"  orcamento.csv      → {len(df):>4} linhas")
    return df


if __name__ == "__main__":
    print("Gerando dados — Rede Máxima Supermercados Ltda.")
    dre = gerar_dre()
    gerar_fluxo_caixa(dre)
    gerar_contas_receber()
    gerar_centro_custos(dre)
    gerar_orcamento()
    print("Concluído. Arquivos em data/raw/")
