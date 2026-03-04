@st.cache_data(show_spinner=False)
def load_data(file_bytes: bytes, sheet_name: str) -> pd.DataFrame:
    wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    ws = wb[sheet_name]

    GREEN_RGB = "FF00FF00"
    IDX_DATA_ROW  = 5
    IDX_GROUP_ROW = 6
    IDX_POS_START = 7

    col_data  = {}
    col_grupo = {}
    col_total = {}
    col_verde_vals = {}   # col -> set de valores únicos verdes

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
                        val = str(cell.value).strip().upper()
                        if c not in col_verde_vals:
                            col_verde_vals[c] = set()
                        col_verde_vals[c].add(val)
                except Exception:
                    pass

    wb.close()

    # Conjunto global de posições já inventariadas (para não contar duplicatas entre colunas)
    ja_inventariadas = set()

    records = []
    for c in sorted(col_data.keys()):
        total = col_total.get(c, 0)
        if total == 0:
            continue

        verdes_desta_col = col_verde_vals.get(c, set())
        # Conta apenas as que ainda NÃO foram contadas em colunas anteriores
        novas = verdes_desta_col - ja_inventariadas
        ja_inventariadas |= verdes_desta_col

        verde = len(novas)

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
