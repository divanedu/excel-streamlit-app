import io
import pandas as pd
import streamlit as st

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string


# =========================
# CODE 1 -> function
# =========================
def run_code_1(file_bytes: bytes) -> bytes:
    # open workbook from memory
    in_buf = io.BytesIO(file_bytes)
    wb = load_workbook(in_buf)

    # IMPORTANT:
    # pandas needs a file-like object too; simplest is to re-use bytes
    xls_buf = io.BytesIO(file_bytes)
    xls = pd.ExcelFile(xls_buf)

    target_sheets = ["1210", "1710", "3310", "3510"]
    sheet_to_col_index = {"1210": 6, "1710": 6, "3310": 7, "3510": 7}

    available_sheets = xls.sheet_names
    sheet_data = {}

    for sheet in target_sheets:
        if sheet not in available_sheets:
            continue

        df = pd.read_excel(xls, sheet_name=sheet, header=0)
        col_idx = sheet_to_col_index[sheet]
        if df.shape[1] <= col_idx:
            continue

        temp = df.iloc[:, [0, col_idx]].copy()
        temp.columns = ["Контрагент", "value"]
        temp["value"] = pd.to_numeric(temp["value"], errors="coerce").fillna(0)
        temp = temp[temp["value"] != 0]
        temp["Контрагент"] = temp["Контрагент"].astype(str).str.strip()
        temp = temp[temp["Контрагент"] != ""]
        temp = temp[~temp["Контрагент"].isin(target_sheets)]
        temp = temp[~temp["Контрагент"].str.lower().str.startswith("итого")]

        if temp.empty:
            continue

        sheet_data[sheet] = temp.groupby("Контрагент")["value"].sum()

    s1210 = sheet_data.get("1210", pd.Series(dtype=float))
    s3510 = sheet_data.get("3510", pd.Series(dtype=float))
    s1710 = sheet_data.get("1710", pd.Series(dtype=float))
    s3310 = sheet_data.get("3310", pd.Series(dtype=float))

    cust_set = set(s1210.index).union(set(s3510.index))
    supp_set = set(s1710.index).union(set(s3310.index))
    all_set = cust_set.union(supp_set)

    if not cust_set and not supp_set:
        raise ValueError("Нет данных для формирования отчёта (код 1).")

    # ---- Build dfs
    if cust_set:
        df_cust = pd.DataFrame(sorted(cust_set), columns=["Контрагент"])
        df_cust["1210"] = df_cust["Контрагент"].map(s1210).fillna(0) / 1000
        df_cust["3510"] = df_cust["Контрагент"].map(s3510).fillna(0) / 1000
        df_cust["сальдо заказчики"] = df_cust["1210"] - df_cust["3510"]
        df_cust = df_cust.sort_values(by="сальдо заказчики", ascending=False).reset_index(drop=True)
    else:
        df_cust = pd.DataFrame(columns=["Контрагент", "1210", "3510", "сальдо заказчики"])

    if supp_set:
        df_supp = pd.DataFrame(sorted(supp_set), columns=["Контрагент"])
        df_supp["1710"] = df_supp["Контрагент"].map(s1710).fillna(0) / 1000
        df_supp["3310"] = df_supp["Контрагент"].map(s3310).fillna(0) / 1000
        df_supp["сальдо поставщики"] = df_supp["1710"] - df_supp["3310"]
        df_supp = df_supp.sort_values(by="сальдо поставщики", ascending=False).reset_index(drop=True)
    else:
        df_supp = pd.DataFrame(columns=["Контрагент", "1710", "3310", "сальдо поставщики"])

    if all_set:
        df_total = pd.DataFrame(sorted(all_set), columns=["Контрагент"])
        df_total["1210"] = df_total["Контрагент"].map(s1210).fillna(0) / 1000
        df_total["1710"] = df_total["Контрагент"].map(s1710).fillna(0) / 1000
        df_total["3310"] = df_total["Контрагент"].map(s3310).fillna(0) / 1000
        df_total["3510"] = df_total["Контрагент"].map(s3510).fillna(0) / 1000
        df_total["общее сальдо"] = (df_total["1210"] + df_total["1710"] - df_total["3310"] - df_total["3510"])
        df_total = df_total.sort_values(by="общее сальдо", ascending=False).reset_index(drop=True)
    else:
        df_total = pd.DataFrame(columns=["Контрагент", "общее сальдо"])

    # ---- Write to sheet
    if "Сальдо PY" in wb.sheetnames:
        wb.remove(wb["Сальдо PY"])
    ws = wb.create_sheet("Сальдо PY")

    ws["A1"] = "Все значения указаны в тысячах тенге"
    ws["A1"].font = Font(name="Arial", size=10, bold=True)

    start_row = 2
    start_col = 2  # B

    font_header = Font(name="Arial", size=10, bold=True)
    font_body = Font(name="Arial", size=10)
    font_bold_body = Font(name="Arial", size=10, bold=True)
    align_center = Alignment(horizontal="center")
    align_left = Alignment(horizontal="left")
    number_format_acc = '#,##0;[Red](#,##0)'

    col_cust_contr = start_col
    col_cust_1210 = start_col + 1
    col_cust_3510 = start_col + 2
    col_cust_saldo = start_col + 3

    col_supp_contr = start_col + 5
    col_supp_1710 = start_col + 6
    col_supp_3310 = start_col + 7
    col_supp_saldo = start_col + 8

    col_total_contr = start_col + 10
    col_total_saldo = start_col + 11

    # block 1
    if not df_cust.empty:
        headers_cust = {
            col_cust_contr: "Контрагент",
            col_cust_1210: "1210",
            col_cust_3510: "3510",
            col_cust_saldo: "сальдо с заказчиками"
        }
        for col, text in headers_cust.items():
            c = ws.cell(row=start_row, column=col, value=text)
            c.font = font_header
            c.alignment = align_center

        for i, (_, row) in enumerate(df_cust.iterrows(), start=start_row + 1):
            r = i
            c_contr = ws.cell(row=r, column=col_cust_contr, value=row["Контрагент"])
            c_contr.font = font_body
            c_contr.alignment = align_left

            for col, val, style in [
                (col_cust_1210, row["1210"], font_body),
                (col_cust_3510, row["3510"], font_body),
                (col_cust_saldo, row["сальдо заказчики"], font_bold_body),
            ]:
                cell = ws.cell(row=r, column=col, value=val)
                cell.font = style
                cell.alignment = align_center
                cell.number_format = number_format_acc

    # block 2
    if not df_supp.empty:
        headers_supp = {
            col_supp_contr: "Контрагент",
            col_supp_1710: "1710",
            col_supp_3310: "3310",
            col_supp_saldo: "сальдо с поставщиками"
        }
        for col, text in headers_supp.items():
            c = ws.cell(row=start_row, column=col, value=text)
            c.font = font_header
            c.alignment = align_center

        for i, (_, row) in enumerate(df_supp.iterrows(), start=start_row + 1):
            r = i
            c_contr = ws.cell(row=r, column=col_supp_contr, value=row["Контрагент"])
            c_contr.font = font_body
            c_contr.alignment = align_left

            for col, val, style in [
                (col_supp_1710, row["1710"], font_body),
                (col_supp_3310, row["3310"], font_body),
                (col_supp_saldo, row["сальдо поставщики"], font_bold_body),
            ]:
                cell = ws.cell(row=r, column=col, value=val)
                cell.font = style
                cell.alignment = align_center
                cell.number_format = number_format_acc

    # block 3
    if not df_total.empty:
        headers_total = {
            col_total_contr: "Контрагент",
            col_total_saldo: "общее сальдо"
        }
        for col, text in headers_total.items():
            c = ws.cell(row=start_row, column=col, value=text)
            c.font = font_header
            c.alignment = align_center

        for i, (_, row) in enumerate(df_total.iterrows(), start=start_row + 1):
            r = i
            c_contr = ws.cell(row=r, column=col_total_contr, value=row["Контрагент"])
            c_contr.font = font_body
            c_contr.alignment = align_left

            cell = ws.cell(row=r, column=col_total_saldo, value=row["общее сальдо"])
            cell.font = font_bold_body
            cell.alignment = align_center
            cell.number_format = number_format_acc

    # widths
    WIDTH_CONTR = 30
    WIDTH_NUM = 18

    for col in [col_cust_contr, col_supp_contr, col_total_contr]:
        ws.column_dimensions[get_column_letter(col)].width = WIDTH_CONTR

    for col in [
        col_cust_1210, col_cust_3510, col_cust_saldo,
        col_supp_1710, col_supp_3310, col_supp_saldo,
        col_total_saldo
    ]:
        ws.column_dimensions[get_column_letter(col)].width = WIDTH_NUM

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================
# CODE 2 -> function
# =========================
def run_code_2(file_bytes: bytes) -> bytes:
    wb = load_workbook(io.BytesIO(file_bytes))

    if 'Md' not in wb.sheetnames or 'Wd' not in wb.sheetnames:
        raise ValueError("Нет листов Md и Wd (код 2).")

    md_ws = wb['Md']
    whd_ws = wb['Wd']

    from collections import defaultdict
    payments_year = defaultdict(lambda: [0.0, 0.0, 0.0])
    performance_year = defaultdict(lambda: [0.0, 0.0, 0.0])
    payments_2025_monthly = defaultdict(lambda: [0.0] * 12)
    performance_2025_monthly = defaultdict(lambda: [0.0] * 12)

    def collect_yearly(sheet, target_dict):
        for row in range(2, sheet.max_row + 1):
            n = sheet[f"A{row}"].value
            c = sheet[f"B{row}"].value
            if not n and not c:
                continue
            key = (str(n).strip() if n else "", str(c).strip() if c else "")
            for idx, col in enumerate(['C', 'D', 'E']):
                v = sheet[f"{col}{row}"].value
                if v is None:
                    continue
                try:
                    target_dict[key][idx] += float(v)
                except:
                    pass

    def collect_monthly_2025(sheet, target_dict):
        start_col = column_index_from_string('AE')
        for row in range(2, sheet.max_row + 1):
            n = sheet[f"A{row}"].value
            c = sheet[f"B{row}"].value
            if not n and not c:
                continue
            key = (str(n).strip() if n else "", str(c).strip() if c else "")
            for i in range(12):
                v = sheet.cell(row=row, column=start_col + i).value
                if v is None:
                    continue
                try:
                    target_dict[key][i] += float(v)
                except:
                    pass

    collect_yearly(whd_ws, payments_year)
    collect_yearly(md_ws, performance_year)
    collect_monthly_2025(whd_ws, payments_2025_monthly)
    collect_monthly_2025(md_ws, performance_2025_monthly)

    all_keys = sorted(
        set(payments_year.keys()) |
        set(performance_year.keys()) |
        set(payments_2025_monthly.keys()) |
        set(performance_2025_monthly.keys()),
        key=lambda x: (x[0], x[1])
    )

    if 'Контракты py' in wb.sheetnames:
        del wb['Контракты py']
    ws = wb.create_sheet('Контракты py')

    ws['A1'] = "ИТОГО в тыс тенге"
    ws['A2'] = "Контрагент"
    ws['B2'] = "Договор"

    ws['C1'] = "оплата"
    ws['C2'] = 2023
    ws['D2'] = 2024
    ws['E2'] = 2025
    ws['F2'] = "Total"

    ws['G1'] = "выполнения с ндс"
    ws['G2'] = 2023
    ws['H2'] = 2024
    ws['I2'] = 2025
    ws['J2'] = "Total"

    ws['K2'] = "дз/(аванс)"

    ws['M1'] = 'оплата'
    months = [f"2025_{str(i).zfill(2)}" for i in range(1, 13)]
    start_col_pay = column_index_from_string('M')
    for i, label in enumerate(months):
        ws[f"{get_column_letter(start_col_pay + i)}2"] = label

    ws['Y1'] = 'выполнения с ндс'
    start_col_perf = column_index_from_string('Y')
    for i, label in enumerate(months):
        ws[f"{get_column_letter(start_col_perf + i)}2"] = label

    start_row = 3
    for idx, key in enumerate(all_keys):
        row = start_row + idx
        name, contract = key

        ws[f"A{row}"] = name
        ws[f"B{row}"] = contract

        py = payments_year.get(key, [0, 0, 0])
        ws[f"C{row}"] = py[0]
        ws[f"D{row}"] = py[1]
        ws[f"E{row}"] = py[2]

        pf = performance_year.get(key, [0, 0, 0])
        ws[f"G{row}"] = pf[0] * 1.12
        ws[f"H{row}"] = pf[1] * 1.12
        ws[f"I{row}"] = pf[2] * 1.12

        ws[f"F{row}"] = f"=SUM(C{row}:E{row})"
        ws[f"J{row}"] = f"=SUM(G{row}:I{row})"
        ws[f"K{row}"] = f"=J{row}-F{row}"

        mp = payments_2025_monthly.get(key, [0]*12)
        for i in range(12):
            ws[f"{get_column_letter(start_col_pay+i)}{row}"] = mp[i]

        mf = performance_2025_monthly.get(key, [0]*12)
        for i in range(12):
            ws[f"{get_column_letter(start_col_perf+i)}{row}"] = mf[i] * 1.12

    last_row = start_row + len(all_keys) - 1

    regular = Font(name="Arial", size=10)
    bold = Font(name="Arial", size=10, bold=True)

    for row in ws.iter_rows(min_row=1, max_row=last_row, min_col=1, max_col=column_index_from_string("AJ")):
        for c in row:
            c.font = regular

    for col in range(1, column_index_from_string("AJ")+1):
        ws[f"{get_column_letter(col)}2"].font = bold

    for addr in ["A1", "C1", "G1", "M1", "Y1", "K2"]:
        ws[addr].font = bold

    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")

    for col in range(1, column_index_from_string("AJ")+1):
        addr = f"{get_column_letter(col)}2"
        ws[addr].alignment = left if addr in ["A2", "B2"] else center

    num_format = '#,##0;[Red](#,##0)'
    numeric_cols = list("CDEFGHIJK")
    numeric_cols += [get_column_letter(c) for c in range(start_col_pay, start_col_pay+12)]
    numeric_cols += [get_column_letter(c) for c in range(start_col_perf, start_col_perf+12)]

    for col in numeric_cols:
        for row in range(3, last_row + 1):
            cell = ws[f"{col}{row}"]
            cell.alignment = center
            cell.number_format = num_format

    for row in range(1, last_row+1):
        ws[f"A{row}"].alignment = left
        ws[f"B{row}"].alignment = left

    ws["C1"].alignment = left
    ws["G1"].alignment = left
    ws["M1"].alignment = left
    ws["Y1"].alignment = left

    ws.column_dimensions['A'].width = 38
    ws.column_dimensions['B'].width = 38
    for col in numeric_cols:
        if column_index_from_string(col) < start_col_pay:
            ws.column_dimensions[col].width = 12.6
        else:
            ws.column_dimensions[col].width = 12.2

    thin = Side(border_style="thin", color="000000")
    border_cols = ["C", "G", "K", "M", "Y"]
    for row in range(1, last_row + 1):
        for col in border_cols:
            cell = ws[f"{col}{row}"]
            cell.border = Border(
                left=thin,
                right=cell.border.right,
                top=cell.border.top,
                bottom=cell.border.bottom
            )

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="Excel Processor", layout="wide")
st.title("Обработка Excel файла")

uploaded = st.file_uploader("Загрузи Excel (.xlsx / .xlsm)", type=["xlsx", "xlsm"])

col1, col2, col3 = st.columns(3)
with col1:
    do_1 = st.checkbox("Запустить код 1 (Сальдо PY)", value=True)
with col2:
    do_2 = st.checkbox("Запустить код 2 (Контракты py)", value=False)
with col3:
    preview_rows = st.number_input("Строк в превью", min_value=5, max_value=200, value=30, step=5)

run = st.button("Запустить обработку", type="primary", disabled=(uploaded is None) or (not do_1 and not do_2))

if run and uploaded is not None:
    file_bytes = uploaded.getvalue()

    try:
        out_bytes = file_bytes
        if do_1:
            out_bytes = run_code_1(out_bytes)
        if do_2:
            out_bytes = run_code_2(out_bytes)

        st.success("Готово! Ниже превью и кнопка скачивания.")

        # Preview: read workbook and show target sheets if they exist
        wb_prev = load_workbook(io.BytesIO(out_bytes), data_only=True)

        prev_cols = st.columns(2)
        with prev_cols[0]:
            if "Сальдо PY" in wb_prev.sheetnames:
                st.subheader("Превью: Сальдо PY")
                df_prev = pd.read_excel(io.BytesIO(out_bytes), sheet_name="Сальдо PY")
                st.dataframe(df_prev.head(int(preview_rows)), use_container_width=True)
            else:
                st.info("Лист 'Сальдо PY' не найден (возможно код 1 не запускался).")

        with prev_cols[1]:
            if "Контракты py" in wb_prev.sheetnames:
                st.subheader("Превью: Контракты py")
                df_prev2 = pd.read_excel(io.BytesIO(out_bytes), sheet_name="Контракты py")
                st.dataframe(df_prev2.head(int(preview_rows)), use_container_width=True)
            else:
                st.info("Лист 'Контракты py' не найден (возможно код 2 не запускался).")

        st.download_button(
            label="Скачать обработанный файл",
            data=out_bytes,
            file_name=f"processed_{uploaded.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Ошибка: {e}")