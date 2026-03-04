import io
from datetime import date

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
import requests

# ─────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────
IDX_DATA_ROW  = 5   # Excel row: datas
IDX_GROUP_ROW = 6   # Excel row: grupos (MZ1, EST...)
IDX_POS_START = 7   # Excel row: início das posições

GREEN_RGB = "FF00FF00"   # única cor verde presente na planilha

GITHUB_URL = "https://github.com/Djalmandre/Inventario26/raw/refs/heads/main/CRONOGRAMA%202026%20RECAP.xlsm"


def fetch_file_bytes(url: str) -> bytes:
    headers = {"User-Agent": "streamlit-app"}
    resp = requests.get(url, headers=headers, timeout=90)
    resp.raise_for_status()
    return resp.content


@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    """
    Lê a planilha em modo read_only (baixo consumo de memória).
    Conta por coluna: total de posições, verdes (inventariadas) e pendentes.
    """
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    ws = wb[sheet_name]

    col_data  = {}   # col_idx -> datetime
    col_grupo = {}   # col_idx -> str
    col_total = {}   # col_idx -> int
    col_verde = {}   # col_idx -> int

    for row in ws.iter_rows(min_row=1):
        # descobrir número da linha
        row_num = None
        for cell in row:
            if hasattr(cell, "row") and cell.row is not None:
                row_num = cell.row
                break
        if row_num is None:
            continue

        if row_num == IDX_DATA_ROW:
            for cell in row:
                if hasattr(cell, "column") and cell.value is not None:
                    col_data[cell.column] = cell.value

        elif row_num == IDX_GROUP_ROW:
            for cell in row:
                if hasattr(cell, "column") and cell.value is not None:
                    col_grupo[cell.column] = str(cell.value)

        elif row_num >= IDX_POS_START:
            for cell in row:
                if not hasattr(cell, "column"):
                    continue
                c = cell.column
                if c not in col_data:
                    continue
                if cell.value is None or str(cell.value).strip() == "":
                    continue

                col_total[c] = col_total.get(c, 0) + 1

                try:
                    fill = cell.fill
                    if (fill and fill.fill_type == "solid"
                            and fill.fgColor
                            and fill.fgColor.type == "rgb"
                            and str(fill.fgColor.rgb).upper() == GREEN_RGB):
                        col_verde[c] = col_verde.get(c, 0) + 1
                except Exception:
                    pass

    wb.close()

    records = []
    for c in sorted(col_data.keys()):
        total = col_total.get(c, 0)
        if total == 0:
            continue
        verde = col_verde.get(c, 0)
        records.append({
            "col":          c,
            "Data":         pd.to_datetime(col_data[c], errors="coerce"),
            "Grupo":        col_grupo.get(c, ""),
            "Total":        total,
            "Inventariado": verde,
            "Pendente":     total - verde,
        })

    df = pd.DataFrame(records).sort_values("Data").reset_index(drop=True)
    return df


def main():
    st.set_page_config(page_title="Painel Micro Inventário", layout="wide")
    st.title("📦 Painel de Micro Inventário — CRONOGRAMA 2026")

    # ── Sidebar ───────────────────────────────────────────────────────────────
    st.sidebar.header("⚙️ Parâmetros")
    sheet_name = st.sidebar.text_input("Nome da aba", value="CRONOGRAMA")
    ignorar_passado = st.sidebar.checkbox(
        "Considerar apenas dias a partir de hoje nos cálculos de meta",
        value=False,
    )
    st.sidebar.caption("Planilha carregada automaticamente do GitHub.")

    # ── Download ──────────────────────────────────────────────────────────────
    with st.spinner("⬇️ Baixando planilha do GitHub..."):
        try:
            file_bytes = fetch_file_bytes(GITHUB_URL)
        except Exception as e:
            st.error(f"Erro ao baixar arquivo: {e}")
            st.stop()

    # ── Processamento ─────────────────────────────────────────────────────────
    with st.spinner("🔍 Lendo células e cores da planilha..."):
        try:
            df = load_data(file_bytes, sheet_name)
        except Exception as e:
            st.error(f"Erro ao processar planilha: {e}")
            st.stop()

    if df.empty:
        st.warning("Nenhuma posição encontrada.")
        st.stop()

    today = pd.Timestamp(date.today())

    # ── Totais ────────────────────────────────────────────────────────────────
    total_pos  = int(df["Total"].sum())
    total_inv  = int(df["Inventariado"].sum())
    total_pend = int(df["Pendente"].sum())
    pct_inv    = round(total_inv / total_pos * 100, 1) if total_pos > 0 else 0.0

    mask_aberto = df["Pendente"] > 0
    if ignorar_passado:
        mask_aberto = mask_aberto & (df["Data"] >= today)
    dias_abertos_df = df[mask_aberto]
    n_dias = len(dias_abertos_df)
    ideal  = int((total_pend + n_dias - 1) / n_dias) if n_dias > 0 else 0

    # ── Métricas ──────────────────────────────────────────────────────────────
    st.subheader("📊 Resumo Geral")
    c1, c2, c3 = st.columns(3)
    c1.metric("📦 Total de Posições",   f"{total_pos:,}".replace(",", "."))
    c2.metric("✅ Inventariadas",        f"{total_inv:,}".replace(",", "."), f"{pct_inv}%")
    c3.metric("⬜ Pendentes",            f"{total_pend:,}".replace(",", "."))

    st.subheader("🎯 Planejamento de Uniformidade")
    p1, p2, p3 = st.columns(3)
    p1.metric("Posições a inventariar", f"{total_pend:,}".replace(",", "."))
    p2.metric("Dias úteis em aberto",   n_dias)
    p3.metric("Meta ideal por dia",     f"{ideal} pos/dia")

    st.info(
        f"💡 Para concluir o inventário de forma uniforme, processe "
        f"**{ideal} posições/dia** ao longo de **{n_dias} dias úteis**."
    )

    # ── Tabela ────────────────────────────────────────────────────────────────
    st.subheader("📋 Detalhamento por Dia")
    df_disp = df.copy()
    df_disp["Data"] = df_disp["Data"].dt.strftime("%d/%m/%Y")
    df_disp["% Concluído"] = df.apply(
        lambda x: f"{round(x['Inventariado'] / x['Total'] * 100, 1)}%"
        if x["Total"] > 0 else "0%",
        axis=1,
    )
    df_disp["Meta Ideal"] = df.apply(
        lambda x: ideal if x["Pendente"] > 0 else "✅", axis=1
    )
    st.dataframe(
        df_disp[["Data", "Grupo", "Total", "Inventariado", "Pendente",
                 "% Concluído", "Meta Ideal"]],
        use_container_width=True,
        hide_index=True,
    )

    # ── Gráficos ──────────────────────────────────────────────────────────────
    st.subheader("📈 Gráficos")
    tab1, tab2 = st.tabs(["Progresso por Dia", "Meta vs Pendente (Dias em Aberto)"])

    with tab1:
        chart_df = df.set_index("Data")[["Inventariado", "Pendente"]]
        st.bar_chart(chart_df, color=["#2ecc71", "#bdc3c7"])

    with tab2:
        if n_dias > 0:
            meta_df = dias_abertos_df.copy().set_index("Data")[["Pendente"]]
            meta_df["Meta Ideal"] = ideal
            st.bar_chart(meta_df, color=["#bdc3c7", "#3498db"])
        else:
            st.success("🎉 Todos os dias já foram concluídos!")

    st.subheader("ℹ️ Legenda de Cores")
    st.markdown("""
    | Cor da célula | Status |
    |---|---|
    | 🟢 Verde | Inventariado |
    | ⬜ Sem cor | Pendente |
    """)


if __name__ == "__main__":
    main()
