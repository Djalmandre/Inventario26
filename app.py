import io
from datetime import date, datetime

import pandas as pd
import requests
import streamlit as st
from openpyxl import load_workbook

# ── Constantes ────────────────────────────────────────────────────────────────
IDX_DATA_ROW  = 5
IDX_GROUP_ROW = 6
IDX_POS_START = 7
GREEN_RGB     = "FF00FF00"
CACHE_VER     = "v4-20260304"

GITHUB_URL = (
    "https://github.com/Djalmandre/Inventario26/raw/refs/heads/main/"
    "CRONOGRAMA%202026%20RECAP.xlsm"
)


# ── Helpers ───────────────────────────────────────────────────────────────────
def fetch_file_bytes(url: str) -> bytes:
    resp = requests.get(url, headers={"User-Agent": "streamlit-app"}, timeout=90)
    resp.raise_for_status()
    return resp.content


def parse_excel_date(value):
    if value is None:
        return None
    try:
        if isinstance(value, (datetime, date)):
            ts = pd.Timestamp(value)
        elif isinstance(value, (int, float)):
            ts = pd.Timestamp("1899-12-30") + pd.Timedelta(days=int(value))
        elif isinstance(value, str):
            ts = pd.to_datetime(value, dayfirst=True)
        else:
            return None
        return ts if 2000 <= ts.year <= 2100 else None
    except Exception:
        return None


# ── Carga de dados ────────────────────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes, sheet_name: str, _v: str = CACHE_VER) -> pd.DataFrame:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)

    if sheet_name not in wb.sheetnames:
        raise ValueError(
            f"Aba '{sheet_name}' não encontrada. "
            f"Abas disponíveis: {wb.sheetnames}"
        )

    ws = wb[sheet_name]
    col_data, col_grupo, col_total, col_verde = {}, {}, {}, {}

    for row in ws.iter_rows(min_row=1):
        row_num = None
        for cell in row:
            if hasattr(cell, "row") and cell.row is not None:
                row_num = cell.row
                break
        if row_num is None:
            continue

        if row_num == IDX_DATA_ROW:
            for cell in row:
                if not hasattr(cell, "column") or cell.value is None:
                    continue
                ts = parse_excel_date(cell.value)
                if ts is not None:
                    col_data[cell.column] = ts

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
                    if (
                        fill
                        and fill.fill_type == "solid"
                        and fill.fgColor
                        and fill.fgColor.type == "rgb"
                        and str(fill.fgColor.rgb).upper() == GREEN_RGB
                    ):
                        col_verde.setdefault(c, set()).add(
                            str(cell.value).strip().upper()
                        )
                except Exception:
                    pass

    wb.close()

    ja_inv = set()
    records = []
    for c in sorted(col_data.keys()):
        total = col_total.get(c, 0)
        if total == 0:
            continue
        verdes = col_verde.get(c, set())
        novas  = verdes - ja_inv
        ja_inv |= verdes
        records.append(
            {
                "Data":         col_data[c],
                "Grupo":        col_grupo.get(c, ""),
                "Total":        total,
                "Inventariado": len(novas),
                "Pendente":     total - len(novas),
            }
        )

    if not records:
        return pd.DataFrame(
            columns=["Data", "Grupo", "Total", "Inventariado", "Pendente"]
        )

    return pd.DataFrame(records).sort_values("Data").reset_index(drop=True)


# ── App ───────────────────────────────────────────────────────────────────────
def main():
    st.set_page_config(page_title="Painel Micro Inventário", layout="wide")
    st.title("📦 Painel de Micro Inventário — CRONOGRAMA 2026")

    # Sidebar
    st.sidebar.header("⚙️ Parâmetros")
    sheet_name = st.sidebar.text_input("Nome da aba", value="CRONOGRAMA")
    ignorar_passado = st.sidebar.checkbox(
        "Considerar apenas dias a partir de hoje nos cálculos de meta",
        value=False,
    )
    st.sidebar.caption("📂 Planilha carregada automaticamente do GitHub.")

    # Download
    with st.spinner("⬇️ Baixando planilha do GitHub..."):
        try:
            file_bytes = fetch_file_bytes(GITHUB_URL)
        except Exception as e:
            st.error(f"Erro ao baixar arquivo: {e}")
            st.stop()

    # Processamento
    with st.spinner("🔍 Lendo células e cores da planilha..."):
        try:
            df = load_data(file_bytes, sheet_name)
        except Exception as e:
            st.error(f"Erro ao processar planilha: {e}")
            st.stop()

    # Validações
    if df is None or df.empty:
        st.error("Nenhuma posição encontrada. Verifique o nome da aba e a linha de datas (linha 5).")
        st.stop()

    if "Data" not in df.columns:
        st.error(f"Coluna 'Data' ausente. Colunas retornadas: {list(df.columns)}")
        st.stop()

    if df["Data"].isna().all():
        st.error("Todas as datas estão inválidas. Verifique a linha 5 da planilha.")
        st.stop()

    today = pd.Timestamp(date.today())

    # Totais
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

    # Métricas
    st.subheader("📊 Resumo Geral")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("📦 Total de Posições",   f"{total_pos:,}".replace(",", "."))
    c2.metric("✅ Inventariadas",        f"{total_inv:,}".replace(",", "."), f"{pct_inv}%")
    c3.metric("⬜ Pendentes",            f"{total_pend:,}".replace(",", "."))
    c4.metric("📅 Dias úteis em aberto", n_dias)

    # Planejamento
    st.subheader("🎯 Planejamento de Uniformidade")
    p1, p2 = st.columns(2)
    p1.metric("Meta ideal por dia", f"{ideal} pos/dia")
    p2.metric("Progresso geral",    f"{pct_inv}%")
    st.info(
        f"💡 Para concluir o inventário de forma uniforme, processe "
        f"**{ideal} posições/dia** ao longo de **{n_dias} dias úteis**."
    )

    # Tabela
    st.subheader("📋 Detalhamento por Dia")
    df_disp = df.copy()
    df_disp["Data"] = df_disp["Data"].dt.strftime("%d/%m/%Y")
    df_disp["% Concluído"] = df.apply(
        lambda x: f"{round(x['Inventariado'] / x['Total'] * 100, 1)}%"
        if x["Total"] > 0 else "0%",
        axis=1,
    )
    df_disp["Meta Ideal"] = df.apply(
        lambda x: f"{ideal} pos" if x["Pendente"] > 0 else "✅ Concluído",
        axis=1,
    )
    st.dataframe(
        df_disp[["Data", "Grupo", "Total", "Inventariado", "Pendente",
                 "% Concluído", "Meta Ideal"]],
        use_container_width=True,
        hide_index=True,
    )

    # Gráficos
    st.subheader("📈 Gráficos")
    tab1, tab2, tab3 = st.tabs([
        "📊 Progresso por Dia",
        "🎯 Meta vs Pendente",
        "📉 Curva Acumulada",
    ])

    with tab1:
        c_df = df.copy()
        c_df["Data_str"] = c_df["Data"].dt.strftime("%d/%m")
        st.bar_chart(
            c_df.set_index("Data_str")[["Inventariado", "Pendente"]],
            color=["#2ecc71", "#bdc3c7"],
        )

    with tab2:
        if n_dias > 0:
            m_df = dias_abertos_df.copy()
            m_df["Data_str"] = m_df["Data"].dt.strftime("%d/%m")
            m_df = m_df.set_index("Data_str")[["Pendente"]]
            m_df["Meta Ideal"] = ideal
            st.bar_chart(m_df, color=["#bdc3c7", "#3498db"])
        else:
            st.success("🎉 Todos os dias já foram concluídos!")

    with tab3:
        s_df = df.sort_values("Data").copy()
        s_df["Inv. Acumulado"]   = s_df["Inventariado"].cumsum()
        s_df["Total Acumulado"]  = s_df["Total"].cumsum()
        s_df["Data_str"] = s_df["Data"].dt.strftime("%d/%m")
        st.line_chart(
            s_df.set_index("Data_str")[["Inv. Acumulado", "Total Acumulado"]],
            color=["#2ecc71", "#3498db"],
        )
        st.caption("🟢 Inventariado acumulado  |  🔵 Total de posições acumulado")

    # Legenda
    st.subheader("ℹ️ Legenda")
    st.markdown("""
    | Cor | Status |
    |---|---|
    | 🟢 Verde (`FF00FF00`) | Inventariado |
    | ⬜ Sem cor | Pendente |

    > Cada posição é contada **uma única vez**, mesmo que apareça verde em múltiplos dias.
    """)


if __name__ == "__main__":
    main()
