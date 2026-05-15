"""
Gerador de Modelo de Frete
Preencha os 6 parâmetros estruturais e baixe o Excel.
Todos os demais dados (preços, taxas, PL, etc.) são editados direto no Excel.
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
    START = datetime.datetime(2026, 6, 1)  # Segunda-feira

    # ── Dias úteis ────────────────────────────────────────────────────────
    def is_wd(n):
        return (n - 1) % 7 <= 4  # 0=seg..4=sex, 5=sáb, 6=dom

    def next_wd(n):
        while not is_wd(n):
            n += 1
        return n

    # ── Pré-computar eventos por dia ──────────────────────────────────────
    h_c, g_c, k_c, o_c = {}, {}, {}, {}
    ORIG = 0

    for n0 in range(1, N + 1):
        if not is_wd(n0):
            continue
        dt  = next_wd(n0 + transito)
        if dt > N:
            break
        dv1 = next_wd(dt + d1)
        dv2 = next_wd(dt + d2)
        dv3 = next_wd(dt + d3)
        if dv3 > N:
            break
        ORIG = n0
        h_c.setdefault(dt,  []).append(n0)
        g_c.setdefault(dv1, []).append(n0)
        k_c.setdefault(dv2, []).append(n0)
        o_c.setdefault(dv3, []).append(n0)

    LR = N + 4

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

    def refs(col, n0_list):
        parts = [f"{col}{n0 + 4}" for n0 in n0_list]
        return parts[0] if len(parts) == 1 else "(" + "+".join(parts) + ")"

    # ── Dashboard ─────────────────────────────────────────────────────────
    ws_d["H5"].value = "% Antecipado D0"
    ws_d["I5"].value = pct_ant
    ws_d["H6"].value = "Trânsito (dias corr.)"
    ws_d["I6"].value = transito
    ws_d["H9"].value  = f"% cart. {d1}d"
    ws_d["H10"].value = f"% cart. {d2}d"
    ws_d["H11"].value = f"% cart. {d3}d"
    ws_d["E5"].value  = f"=Fundo!AD5+Fundo!AD{LR}"
    ws_d["E6"].value  = f"=Fundo!AJ5+Fundo!AJ{LR}"
    ws_d["E7"].value  = f"=Fundo!AN5+Fundo!AN{LR}"
    ws_d["C23"].value = f"=Fundo!U{LR}/Dashboard!D14"

    # ── Fundo: cabeçalhos ─────────────────────────────────────────────────
    ws_f["E2"].value = f"Carteira {d1}d"
    ws_f["I2"].value = f"Carteira {d2}d"
    ws_f["M2"].value = f"Carteira {d3}d"
    ws_f["F3"].value = "1º Desemb. (-)"
    ws_f["H3"].value = "2º Desemb. (-)"
    ws_f["J3"].value = "1º Desemb. (-)"
    ws_f["L3"].value = "2º Desemb. (-)"
    ws_f["N3"].value = "1º Desemb. (-)"
    ws_f["P3"].value = "2º Desemb. (-)"
    ws_f["R3"].value = "Total Desemb. (-)"

    # ── Linha 4 e IRRs ────────────────────────────────────────────────────
    ws_f["U4"].value  = "=Z5+AF5+AL5"
    ws_f["X2"].value  = f"=IRR(X5:X{LR},0.001)"
    ws_f["AD2"].value = f"=IRR(AD5:AD{LR},0.001)"
    ws_f["AJ2"].value = f"=IRR(AJ5:AJ{LR},0.001)"
    ws_f["AN2"].value = f"=IRR(AN5:AN{LR},0.001)"

    clear_rows(ws_f, 5)
    clear_rows(ws_c, 5)

    yellow    = PatternFill(patternType="solid", fgColor="FFFFFFCC")
    light_red = PatternFill(patternType="solid", fgColor="FFFFCCCC")

    for n in range(1, N + 1):
        r  = n + 4
        pr = r - 1
        wd    = is_wd(n)
        first = (n == 1)
        last  = (n == N)
        orig  = wd and n <= ORIG

        # ── B: Benchmark ──────────────────────────────────────────────────
        sc(ws_f,r,2, "=Dashboard!D14" if first
                     else f"=B{pr}*(((1+Dashboard!$C$20)^(1/30)))")

        # ── C / D: contador e data ────────────────────────────────────────
        sc(ws_f,r,3, n)
        sc(ws_f,r,4, START if first else f"=D{pr}+1")

        # ── E / I / M: originações (0 em fds e após ORIG) ─────────────────
        sc(ws_f,r,5,  "=Dashboard!$I$8*Dashboard!$I$9"  if orig else 0)
        sc(ws_f,r,9,  "=Dashboard!$I$8*Dashboard!$I$10" if orig else 0)
        sc(ws_f,r,13, "=Dashboard!$I$8*Dashboard!$I$11" if orig else 0)

        # ── F / J / N: 1º desembolso (80%) — carregamento, só dias úteis ─
        sc(ws_f,r,6,  f"=E{r}*Dashboard!$I$3*Dashboard!$I$5"  if orig else 0)
        sc(ws_f,r,10, f"=I{r}*Dashboard!$I$3*Dashboard!$I$5"  if orig else 0)
        sc(ws_f,r,14, f"=M{r}*Dashboard!$I$3*Dashboard!$I$5"  if orig else 0)

        # ── H / L / P: 2º desembolso (20%) — chegada ao porto ────────────
        h_list = h_c.get(n, [])
        if wd and h_list:
            sc(ws_f,r,8,  f"={refs('E',h_list)}*Dashboard!$I$3*(1-Dashboard!$I$5)")
            sc(ws_f,r,12, f"={refs('I',h_list)}*Dashboard!$I$3*(1-Dashboard!$I$5)")
            sc(ws_f,r,16, f"={refs('M',h_list)}*Dashboard!$I$3*(1-Dashboard!$I$5)")
        else:
            sc(ws_f,r,8, 0); sc(ws_f,r,12, 0); sc(ws_f,r,16, 0)

        # ── G / K / O: levantamento — pagamento do importador ────────────
        g_list = g_c.get(n, [])
        k_list = k_c.get(n, [])
        o_list = o_c.get(n, [])
        sc(ws_f,r,7,  f"={refs('E',g_list)}*Dashboard!$I$4" if wd and g_list else 0)
        sc(ws_f,r,11, f"={refs('I',k_list)}*Dashboard!$I$4" if wd and k_list else 0)
        sc(ws_f,r,15, f"={refs('M',o_list)}*Dashboard!$I$4" if wd and o_list else 0)

        # ── Q / R / S ─────────────────────────────────────────────────────
        sc(ws_f,r,17, "=Dashboard!$I$8" if orig else 0)
        sc(ws_f,r,18, f"=F{r}+H{r}+J{r}+L{r}+N{r}+P{r}")
        sc(ws_f,r,19, f"=G{r}+K{r}+O{r}")

        # ── T: P&L ativos ─────────────────────────────────────────────────
        all_lev = ([(n0,"E") for n0 in g_list] +
                   [(n0,"I") for n0 in k_list] +
                   [(n0,"M") for n0 in o_list])
        if all_lev:
            cost = "+".join(f"{col}{n0+4}" for n0, col in all_lev)
            sc(ws_f,r,20, f"=S{r}-({cost})*Dashboard!$I$3")
        else:
            sc(ws_f,r,20, 0)

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

        # ── Cores ─────────────────────────────────────────────────────────
        if not last:
            for col in [26, 28, 32, 34, 38, 39]:   # amarelo: integr./amort.
                cell = ws_f.cell(r, col)
                if not isinstance(cell, MergedCell):
                    cell.fill = yellow

        if not wd:
            for col in range(2, 17):                # vermelho: fins de semana
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
