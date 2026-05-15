"""
Gerador de Modelo de Frete
Preencha os 6 parâmetros estruturais e baixe o Excel.
Todos os demais dados (preços, taxas, PL, etc.) são editados direto no Excel.

Lógica de fim de semana inteiramente em fórmulas Excel:
  - WEEKDAY($D,2)>5 → zeros em E/F/G/H/J/K/L/N/O/P
  - WORKDAY() → calcula Dt e Dv nas colunas auxiliares AQ/AR/AS/AT
  - SUMIF(AQ,$D,E) → agrega 2º desembolso e levantamento por data
"""

import io
import datetime
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell import MergedCell
from openpyxl.styles import PatternFill

st.set_page_config(page_title="Gerador — Modelo de Frete", page_icon="🚛")
st.title("🚛 Gerador — Modelo de Frete")
st.divider()

col1, col2 = st.columns(2)

with col1:
    n_dias   = st.number_input("Número de dias", min_value=50, max_value=1800, value=395, step=10)
    transito = st.number_input("Dias de trânsito (corridos)", min_value=1, max_value=60, value=5, step=1)
    pct_ant  = st.number_input("% antecipado no carregamento (D0)", min_value=1, max_value=99, value=80, step=5) / 100

with col2:
    d1 = st.number_input("Prazo importador cart. 1 (dias corridos)", min_value=1, max_value=365, value=10, step=1)
    d2 = st.number_input("Prazo importador cart. 2 (dias corridos)", min_value=1, max_value=365, value=15, step=1)
    d3 = st.number_input("Prazo importador cart. 3 (dias corridos)", min_value=1, max_value=365, value=20, step=1)
    if not (d1 <= d2 <= d3):
        st.warning("Os prazos devem estar em ordem crescente (d1 ≤ d2 ≤ d3).")

st.divider()


# ════════════════════════════════════════════════════════════════════════════
# GERADOR EXCEL
# ════════════════════════════════════════════════════════════════════════════

def gerar_excel(template_bytes, n_dias, d1, d2, d3, transito, pct_ant):
    N     = n_dias
    LR    = N + 4          # última linha de dados
    START = datetime.datetime(2026, 6, 1)  # Segunda-feira

    # ── ORIG: último dia de originação (Python computa, Excel usa via I7) ─
    def is_wd(n):
        return (n - 1) % 7 <= 4  # 0=seg..4=sex, 5=sáb, 6=dom

    def next_wd(n):
        while not is_wd(n):
            n += 1
        return n

    ORIG = 0
    for n0 in range(1, N + 1):
        if not is_wd(n0):
            continue
        dt  = next_wd(n0 + transito)
        if dt > N:
            break
        dv3 = next_wd(dt + d3)
        if dv3 > N:
            break
        ORIG = n0

    wb   = load_workbook(io.BytesIO(template_bytes))
    ws_f = wb["Fundo"]
    ws_c = wb["Custos"]
    ws_d = wb["Dashboard"]

    def sc(ws, row, col, val):
        cell = ws.cell(row, col)
        if not isinstance(cell, MergedCell):
            cell.value = val

    def clear_rows(ws, from_row):
        for r in range(from_row, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                sc(ws, r, col, None)

    # ── Dashboard ─────────────────────────────────────────────────────────
    ws_d["H5"].value  = "% Antecipado D0"
    ws_d["I5"].value  = pct_ant
    ws_d["H6"].value  = "Trânsito (dias corr.)"
    ws_d["I6"].value  = transito
    ws_d["I6"].number_format = "0"       # número inteiro, não contábil
    ws_d["H7"].value  = "Último dia orig."
    ws_d["I7"].value  = ORIG        # referenciado por E/I/M via $C<=I7
    ws_d["H9"].value  = f"% cart. {d1}d"
    ws_d["H10"].value = f"% cart. {d2}d"
    ws_d["H11"].value = f"% cart. {d3}d"
    ws_d["K9"].value  = d1          # referenciado por AR (Dv1)
    ws_d["K10"].value = d2          # referenciado por AS (Dv2)
    ws_d["K11"].value = d3          # referenciado por AT (Dv3)
    ws_d["L9"].value  = "dias prazo"
    ws_d["L10"].value = "dias prazo"
    ws_d["L11"].value = "dias prazo"
    ws_d["E5"].value  = f"=Fundo!AD5+Fundo!AD{LR}"
    ws_d["E6"].value  = f"=Fundo!AJ5+Fundo!AJ{LR}"
    ws_d["E7"].value  = f"=Fundo!AN5+Fundo!AN{LR}"
    ws_d["C23"].value = f"=Fundo!U{LR}/Dashboard!D14"

    # ── Fundo: cabeçalhos ─────────────────────────────────────────────────
    ws_f["E2"].value  = f"Carteira {d1}d"
    ws_f["I2"].value  = f"Carteira {d2}d"
    ws_f["M2"].value  = f"Carteira {d3}d"
    ws_f["F3"].value  = "1º Desemb. (-)"
    ws_f["G3"].value  = "2º Desemb. (-)"  # nova posição
    ws_f["H3"].value  = "Levantamento"     # nova posição
    ws_f["J3"].value  = "1º Desemb. (-)"
    ws_f["K3"].value  = "2º Desemb. (-)"
    ws_f["L3"].value  = "Levantamento"
    ws_f["N3"].value  = "1º Desemb. (-)"
    ws_f["O3"].value  = "2º Desemb. (-)"
    ws_f["P3"].value  = "Levantamento"
    ws_f["R3"].value  = "Total Desemb. (-)"
    ws_f["S3"].value  = "Levantamento"
    ws_f["AQ3"].value = "Dt chegada"
    ws_f["AR3"].value = "Dv1 pgto"
    ws_f["AS3"].value = "Dv2 pgto"
    ws_f["AT3"].value = "Dv3 pgto"

    # ── Fundo: linha 4 e IRRs ─────────────────────────────────────────────
    ws_f["U4"].value  = "=Z5+AF5+AL5"
    ws_f["X2"].value  = f"=IRR(X5:X{LR},0.001)"
    ws_f["AD2"].value = f"=IRR(AD5:AD{LR},0.001)"
    ws_f["AJ2"].value = f"=IRR(AJ5:AJ{LR},0.001)"
    ws_f["AN2"].value = f"=IRR(AN5:AN{LR},0.001)"

    clear_rows(ws_f, 5)
    clear_rows(ws_c, 5)

    yellow    = PatternFill(patternType="solid", fgColor="FFFFFFCC")
    light_red = PatternFill(patternType="solid", fgColor="FFFFCCCC")

    # Atalhos de range para SUMIF (absolutos, usados em todas as linhas)
    rng_aq = f"$AQ$5:$AQ${LR}"
    rng_ar = f"$AR$5:$AR${LR}"
    rng_as = f"$AS$5:$AS${LR}"
    rng_at = f"$AT$5:$AT${LR}"
    rng_e  = f"$E$5:$E${LR}"
    rng_i  = f"$I$5:$I${LR}"
    rng_m  = f"$M$5:$M${LR}"

    for n in range(1, N + 1):
        r  = n + 4
        pr = r - 1
        first = (n == 1)
        last  = (n == N)
        wd    = is_wd(n)   # para coloração — lógica de negócio fica no Excel

        # ── B: Benchmark ──────────────────────────────────────────────────
        sc(ws_f,r,2, "=Dashboard!D14" if first
                     else f"=B{pr}*(((1+Dashboard!$C$20)^(1/30)))")

        # ── C / D: contador e data ────────────────────────────────────────
        sc(ws_f,r,3, n)
        sc(ws_f,r,4, START if first else f"=D{pr}+1")

        # ── E / I / M: originações — Excel checa fim de semana e ORIG ────
        e_f = f"=IF(AND(WEEKDAY($D{r},2)<=5,C{r}<=Dashboard!$I$7),Dashboard!$I$8*Dashboard!$I$9,0)"
        i_f = f"=IF(AND(WEEKDAY($D{r},2)<=5,C{r}<=Dashboard!$I$7),Dashboard!$I$8*Dashboard!$I$10,0)"
        m_f = f"=IF(AND(WEEKDAY($D{r},2)<=5,C{r}<=Dashboard!$I$7),Dashboard!$I$8*Dashboard!$I$11,0)"
        sc(ws_f,r,5,  e_f)
        sc(ws_f,r,9,  i_f)
        sc(ws_f,r,13, m_f)

        # ── F / J / N: 1º desembolso (80%) ───────────────────────────────
        sc(ws_f,r,6,  f"=IF(WEEKDAY($D{r},2)>5,0,E{r}*Dashboard!$I$3*Dashboard!$I$5)")
        sc(ws_f,r,10, f"=IF(WEEKDAY($D{r},2)>5,0,I{r}*Dashboard!$I$3*Dashboard!$I$5)")
        sc(ws_f,r,14, f"=IF(WEEKDAY($D{r},2)>5,0,M{r}*Dashboard!$I$3*Dashboard!$I$5)")

        # ── G / K / O: 2º desembolso (20%) — SUMIF sobre Dt (col AQ) ────
        g2_base = f"*Dashboard!$I$3*(1-Dashboard!$I$5)"
        sc(ws_f,r,7,  f"=IF(WEEKDAY($D{r},2)>5,0,SUMIF({rng_aq},$D{r},{rng_e}){g2_base})")
        sc(ws_f,r,11, f"=IF(WEEKDAY($D{r},2)>5,0,SUMIF({rng_aq},$D{r},{rng_i}){g2_base})")
        sc(ws_f,r,15, f"=IF(WEEKDAY($D{r},2)>5,0,SUMIF({rng_aq},$D{r},{rng_m}){g2_base})")

        # ── H / L / P: levantamento — SUMIF sobre Dv (cols AR/AS/AT) ────
        sc(ws_f,r,8,  f"=IF(WEEKDAY($D{r},2)>5,0,SUMIF({rng_ar},$D{r},{rng_e})*Dashboard!$I$4)")
        sc(ws_f,r,12, f"=IF(WEEKDAY($D{r},2)>5,0,SUMIF({rng_as},$D{r},{rng_i})*Dashboard!$I$4)")
        sc(ws_f,r,16, f"=IF(WEEKDAY($D{r},2)>5,0,SUMIF({rng_at},$D{r},{rng_m})*Dashboard!$I$4)")

        # ── Q / R / S ─────────────────────────────────────────────────────
        sc(ws_f,r,17, f"=E{r}+I{r}+M{r}")
        sc(ws_f,r,18, f"=F{r}+G{r}+J{r}+K{r}+N{r}+O{r}")   # total desembolso
        sc(ws_f,r,19, f"=H{r}+L{r}+P{r}")                   # total levantamento

        # ── T: P&L ativos (levantamentos – custo de aquisição) ────────────
        cost = (f"(SUMIF({rng_ar},$D{r},{rng_e})"
                f"+SUMIF({rng_as},$D{r},{rng_i})"
                f"+SUMIF({rng_at},$D{r},{rng_m}))*Dashboard!$I$3")
        sc(ws_f,r,20, f"=S{r}-{cost}")

        # ── V: rendimento CDI ─────────────────────────────────────────────
        sc(ws_f,r,22, 0 if first
                      else f"=U{pr}*(((1+Dashboard!$C$17)^(1/30))-1)")

        # ── U: caixa ──────────────────────────────────────────────────────
        if first:
            sc(ws_f,r,21, f"=U4-R{r}+S{r}+V{r}-Custos!D{r}")
        elif n == 2:
            sc(ws_f,r,21, f"=U{pr}-R{r}+S{r}+V{r}-Custos!D{r}"
                          f"-AB{pr}-AH{pr}-AM{pr}")
        else:
            sc(ws_f,r,21, f"=U{pr}-R{r}+S{r}+V{r}-Custos!D{r}"
                          f"+Z{pr}+AF{pr}+AL{pr}-AB{pr}-AH{pr}-AM{pr}")

        sc(ws_f,r,23, f"=U{r}+R{r}")
        sc(ws_f,r,24, f"=-R{r}+S{r}-Custos!D{r}")

        # ── Sênior ────────────────────────────────────────────────────────
        sc(ws_f,r,26, "=Z4" if first else 0)
        sc(ws_f,r,27, f"=AC{pr}*(((1+Dashboard!$C$20)^(1/30))-1)")
        sc(ws_f,r,28, f"=AC{r}" if last else 0)
        sc(ws_f,r,29, f"=Z{r}" if first else f"=AC{pr}+AA{r}-AB{pr}+Z{r}")
        sc(ws_f,r,30, f"=-Z{r}+AB{r}")

        # ── Mezanino ──────────────────────────────────────────────────────
        sc(ws_f,r,32, "=AF4" if first else 0)
        sc(ws_f,r,33, f"=AI{pr}*(((1+Dashboard!$C$21)^(1/30))-1)")
        sc(ws_f,r,34, f"=AI{r}" if last else 0)
        sc(ws_f,r,35, f"=AF{r}" if first else f"=AI{pr}+AG{r}-AH{pr}+AF{r}")
        sc(ws_f,r,36, f"=-AF{r}+AH{r}")

        # ── Júnior ────────────────────────────────────────────────────────
        sc(ws_f,r,38, "=AL4" if first else 0)
        sc(ws_f,r,39, f"=U{r}-AB{r}-AH{r}" if last else 0)
        sc(ws_f,r,40, f"=-AL{r}+AM{r}")

        # ── Colunas auxiliares de datas (AQ/AR/AS/AT) ─────────────────────
        # AQ: Dt = próximo dia útil após D0 + trânsito (chegada ao porto)
        sc(ws_f,r,43, f"=IF(E{r}>0,WORKDAY($D{r}+Dashboard!$I$6-1,1),\"\")")
        # AR/AS/AT: Dv1/2/3 = próximo dia útil após Dt + prazo importador
        sc(ws_f,r,44, f"=IF(AQ{r}<>\"\",WORKDAY(AQ{r}+Dashboard!$K$9-1,1),\"\")")
        sc(ws_f,r,45, f"=IF(AQ{r}<>\"\",WORKDAY(AQ{r}+Dashboard!$K$10-1,1),\"\")")
        sc(ws_f,r,46, f"=IF(AQ{r}<>\"\",WORKDAY(AQ{r}+Dashboard!$K$11-1,1),\"\")")

        # ── Cores ─────────────────────────────────────────────────────────
        if not last:
            for col in [26, 28, 32, 34, 38, 39]:   # amarelo: integr./amort.
                cell = ws_f.cell(r, col)
                if not isinstance(cell, MergedCell):
                    cell.fill = yellow

        if not wd:                                   # vermelho: fins de semana
            for col in range(2, 17):
                cell = ws_f.cell(r, col)
                if not isinstance(cell, MergedCell):
                    cell.fill = light_red

        # ── Custos ────────────────────────────────────────────────────────
        sc(ws_c,r,2, n)
        sc(ws_c,r,3, START if first else f"=C{pr}+1")
        sc(ws_c,r,5, "=Dashboard!$J$16/30" if first
                     else f"=MAX(Dashboard!$J$16,(Dashboard!$I$16*Fundo!W{pr})/360)/30")
        sc(ws_c,r,6, "=Dashboard!$J$15/30" if first
                     else f"=MAX(Dashboard!$J$15,(Dashboard!$I$15*Fundo!W{pr})/360)/30")
        sc(ws_c,r,7, 0 if first else f"=Fundo!R{pr}*Dashboard!$I$14")
        sc(ws_c,r,8, 0)
        sc(ws_c,r,9, 10000 if first else None)
        sc(ws_c,r,4, f"=E{r}+F{r}+G{r}+H{r}+I{r}")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ════════════════════════════════════════════════════════════════════════════
# DOWNLOAD
# ════════════════════════════════════════════════════════════════════════════

try:
    with open("Modelagem_Frete_v5.xlsx", "rb") as f:
        template_bytes = f.read()
except FileNotFoundError:
    st.error("❌ Arquivo `Modelagem_Frete_v5.xlsx` não encontrado no repositório.")
    st.stop()

with st.spinner("Gerando Excel..."):
    excel_bytes = gerar_excel(template_bytes, n_dias, d1, d2, d3, transito, pct_ant)

st.download_button(
    label=f"⬇️ Baixar Modelagem_Frete_{n_dias}d.xlsx",
    data=excel_bytes,
    file_name=f"Modelagem_Frete_{n_dias}d.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True,
)
