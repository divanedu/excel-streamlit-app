import io
import time
from collections import defaultdict
from typing import Dict, Tuple, Optional

import pandas as pd
import streamlit as st

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string


# =========================
# Helpers
# =========================
def safe_sheet_name(name: str) -> str:
    """
    Excel sheet name constraints:
    - max 31 chars
    - cannot contain: : \ / ? * [ ]
    """
    banned = [":", "\\", "/", "?", "*", "[", "]"]
    for ch in banned:
        name = name.replace(ch, "")
    name = name.strip()
    if not name:
        name = "лист"
    return name[:31]


def split_prefix_suffix(name: str) -> Tuple[str, str]:
    """
    Returns (prefix, suffix) where suffix is last 4 characters (lowercased)
    and prefix is everything before that (original case).
    """
    if len(name) < 4:
        return name, ""
    return name[:-4], name[-4:].lower()


def split_prefix_2(name: str) -> Tuple[str, str]:
    """
    Returns (prefix, suffix2) where suffix2 is last 2 characters (lowercased),
    prefix is everything before that (original case).
    """
    if len(name) < 2:
        return name, ""
    return name[:-2], name[-2:].lower()


def normalize_prefix(prefix: str) -> str:
    return prefix.strip()


# =========================
# CODE 1 (Saldo) -> multi-company
# =========================
def run_code_1(file_bytes: bytes) -> bytes:
    wb = load_workbook(io.BytesIO(file_bytes))
    xls = pd.ExcelFile(io.BytesIO(file_bytes))

    # Map suffix -> "account code" and column index
    target_suffixes = {"1210": 6, "1710": 6, "3310": 7, "3510": 7}

    # Collect which sheets belong to which prefix
    # prefix can be "" (no prefix), or "S", or "ABC", etc.
    prefix_to_sheets: Dict[str, Dict[str, str]] = defaultdict(dict)
    for sh in xls.sheet_names:
        prefix, suf = split_prefix_suffix(sh)
        if suf in target_suffixes:
            prefix = normalize_prefix(prefix)
            # store exact sheet name for that suffix (if duplicates exist, keep the first)
            if suf not in prefix_to_sheets[prefix]:
                prefix_to_sheets[prefix][suf] = sh

    if not prefix_to_sheets:
        raise ValueError("Код 1: не найдено ни одного листа, заканчивающегося на 1210/1710/3310/3510.")

    # For each company prefix, build a separate saldo sheet
    for prefix, suf_map in prefix_to_sheets.items():
        sheet_data = {}

        for suf, col_idx in target_suffixes.items():
            if suf not in suf_map:
                continue

            sheet_name = suf_map[suf]
            df = pd.read_excel(xls, sheet_name=sheet_name, header=0)

            if df.shape[1] <= col_idx:
                continue

            temp = df.iloc[:, [0, col_idx]].copy()
            temp.columns = ["Контрагент", "value"]
            temp["value"] = pd.to_numeric(temp["value"], errors="coerce").fillna(0)
            temp = temp[temp["value"] != 0]
            temp["Контрагент"] = temp["Контрагент"].astype(str).str.strip()
            temp = temp[temp["Контрагент"] != ""]
            temp = temp[~temp["Контрагент"].isin(["1210", "1710", "3310", "3510"])]
            temp = temp[~temp["Контрагент"].str.lower().str.startswith("итого")]

            if temp.empty:
                continue

            sheet_data[suf] = temp.groupby("Контрагент")["value"].sum()

        s1210 = sheet_data.get("1210", pd.Series(dtype=float))
        s3510 = sheet_data.get("3510", pd.Series(dtype=float))
        s1710 = sheet_data.get("1710", pd.Series(dtype=float))
        s3310 = sheet_data.get("3310", pd.Series(dtype=float))

        cust_set = set(s1210.index).union(set(s3510.index))
        supp_set = set(s1710.index).union(set(s3310.index))
        all_set = cust_set.union(supp_set)

        # If a specific prefix has no usable data, just skip it (do not fail whole run)
        if not cust_set and not supp_set:
            continue

        # Block 1: customers
        if cust_set:
            df_cust = pd.DataFrame(sorted(cust_set), columns=["Контрагент"])
            df_cust["1210"] = df_cust["Контрагент"].map(s1210).fillna(0) / 1000
            df_cust["3510"] = df_cust["Контрагент"].map(s3510).fillna(0) / 1000
            df_cust["сальдо заказчики"] = df_cust["1210"] - df_cust["3510"]
            df_cust = df_cust.sort_values(by="сальдо заказчики", ascending=False).reset_index(drop=True)
        else:
            df_cust = pd.DataFrame(columns=["Контрагент", "1210", "3510", "сальдо заказчики"])

        # Block 2: suppliers
        if supp_set:
            df_supp = pd.DataFrame(sorted(supp_set), columns=["Контрагент"])
            df_supp["1710"] = df_supp["Контрагент"].map(s1710).fillna(0) / 1000
            df_supp["3310"] = df_supp["Контрагент"].map(s3310).fillna(0) / 1000
            df_supp["сальдо поставщики"] = df_supp["1710"] - df_supp["3310"]
            df_supp = df_supp.sort_values(by="сальдо поставщики", ascending=False).reset_index(drop=True)
        else:
            df_supp = pd.DataFrame(columns=["Контрагент", "1710", "3310", "сальдо поставщики"])

        # Block 3: total
        if all_set:
            df_total = pd.DataFrame(sorted(all_set), columns=["Контрагент"])
            df_total["1210"] = df_total["Контрагент"].map(s1210).fillna(0) / 1000
            df_total["1710"] = df_total["Контрагент"].map(s1710).fillna(0) / 1000
            df_total["3310"] = df_total["Контрагент"].map(s3310).fillna(0) / 1000
            df_total["3510"] = df_total["Контрагент"].map(s3510).fillna(0) / 1000
            df_total["общее сальдо"] = df_total["1210"] + df_total["1710"] - df_total["3310"] - df_total["3510"]
            df_total = df_total.sort_values(by="общее сальдо", ascending=False).reset_index(drop=True)
        else:
            df_total = pd.DataFrame(columns=["Контрагент", "общее сальдо"])

        out_sheet_name = safe_sheet_name(f"{prefix}сальд" if prefix else "сальд")
        if out_sheet_name in wb.sheetnames:
            wb.remove(wb[out_sheet_name])
        ws = wb.create_sheet(out_sheet_name)

        ws["A1"] = "Все значения указаны в тысячах тенге"
        ws["A1"].font = Font(name="Arial", size=10, bold=True)

        start_row = 2
        start_col = 2  # B

        font_header = Font(name="Arial", size=10, bold=True)
        font_body = Font(name="Arial", size=10)
        font_bold_body = Font(name="Arial", size=10, bold=True)
        align_center = Alignment(horizontal="center")
        align_left = Alignment(horizontal="left")
        number_format_acc = "#,##0;[Red](#,##0)"

        # Block 1 coords
        col_cust_contr = start_col
        col_cust_1210 = start_col + 1
        col_cust_3510 = start_col + 2
        col_cust_saldo = start_col + 3

        # Block 2 coords (+1 empty col)
        col_supp_contr = start_col + 5
        col_supp_1710 = start_col + 6
        col_supp_3310 = start_col + 7
        col_supp_saldo = start_col + 8

        # Block 3 coords (+1 empty col)
        col_total_contr = start_col + 10
        col_total_saldo = start_col + 11

        # Block 1
        if not df_cust.empty:
            headers_cust = {
                col_cust_contr: "Контрагент",
                col_cust_1210: "1210",
                col_cust_3510: "3510",
                col_cust_saldo: "сальдо с заказчиками",
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

        # Block 2
        if not df_supp.empty:
            headers_supp = {
                col_supp_contr: "Контрагент",
                col_supp_1710: "1710",
                col_supp_3310: "3310",
                col_supp_saldo: "сальдо с поставщиками",
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

        # Block 3
        if not df_total.empty:
            headers_total = {col_total_contr: "Контрагент", col_total_saldo: "общее сальдо"}
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
            col_cust_1210,
            col_cust_3510,
            col_cust_saldo,
            col_supp_1710,
            col_supp_3310,
            col_supp_saldo,
            col_total_saldo,
        ]:
            ws.column_dimensions[get_column_letter(col)].width = WIDTH_NUM

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================
# CODE 2 (Contracts) -> multi-company
# =========================
def run_code_2(file_bytes: bytes) -> bytes:
    wb = load_workbook(io.BytesIO(file_bytes))

    # Find sheets ending with "wd"/"md" (case-insensitive) and group by prefix
    prefix_to_pair: Dict[str, Dict[str, str]] = defaultdict(dict)
    for sh in wb.sheetnames:
        prefix, suf2 = split_prefix_2(sh)
        if suf2 in ("wd", "md"):
            prefix = normalize_prefix(prefix)
            # keep first if duplicates
            if suf2 not in prefix_to_pair[prefix]:
                prefix_to_pair[prefix][suf2] = sh

    # keep only prefixes having both md and wd
    valid_prefixes = [p for p, m in prefix_to_pair.items() if "wd" in m and "md" in m]
    if not valid_prefixes:
        raise ValueError("Код 2: не найдено пар листов (Wd/Md) с одинаковым префиксом (регистр не важен).")

    for prefix in valid_prefixes:
        md_ws = wb[prefix_to_pair[prefix]["md"]]
        wd_ws = wb[prefix_to_pair[prefix]["wd"]]

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
                for idx, col in enumerate(["C", "D", "E"]):
                    v = sheet[f"{col}{row}"].value
                    if v is None:
                        continue
                    try:
                        target_dict[key][idx] += float(v)
                    except:
                        pass

        def collect_monthly_2025(sheet, target_dict):
            start_col = column_index_from_string("AE")
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

        collect_yearly(wd_ws, payments_year)
        collect_yearly(md_ws, performance_year)
        collect_monthly_2025(wd_ws, payments_2025_monthly)
        collect_monthly_2025(md_ws, performance_2025_monthly)

        all_keys = sorted(
            set(payments_year.keys())
            | set(performance_year.keys())
            | set(payments_2025_monthly.keys())
            | set(performance_2025_monthly.keys()),
            key=lambda x: (x[0], x[1]),
        )

        out_sheet_name = safe_sheet_name(f"{prefix}контр" if prefix else "контр")
        if out_sheet_name in wb.sheetnames:
            del wb[out_sheet_name]
        ws = wb.create_sheet(out_sheet_name)

        ws["A1"] = "ИТОГО в тыс тенге"

        ws["A2"] = "Контрагент"
        ws["B2"] = "Договор"

        ws["C1"] = "оплата"
        ws["C2"] = 2023
        ws["D2"] = 2024
        ws["E2"] = 2025
        ws["F2"] = "Total"

        ws["G1"] = "выполнения с ндс"
        ws["G2"] = 2023
        ws["H2"] = 2024
        ws["I2"] = 2025
        ws["J2"] = "Total"

        ws["K2"] = "дз/(аванс)"

        ws["M1"] = "оплата"
        months = [f"2025_{str(i).zfill(2)}" for i in range(1, 13)]

        start_col_pay = column_index_from_string("M")
        for i, label in enumerate(months):
            ws[f"{get_column_letter(start_col_pay + i)}2"] = label

        ws["Y1"] = "выполнения с ндс"
        start_col_perf = column_index_from_string("Y")
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

            mp = payments_2025_monthly.get(key, [0] * 12)
            for i in range(12):
                ws[f"{get_column_letter(start_col_pay + i)}{row}"] = mp[i]

            mf = performance_2025_monthly.get(key, [0] * 12)
            for i in range(12):
                ws[f"{get_column_letter(start_col_perf + i)}{row}"] = mf[i] * 1.12

        last_row = start_row + len(all_keys) - 1 if all_keys else 2

        regular = Font(name="Arial", size=10)
        bold = Font(name="Arial", size=10, bold=True)

        for row in ws.iter_rows(min_row=1, max_row=last_row, min_col=1, max_col=column_index_from_string("AJ")):
            for c in row:
                c.font = regular

        for col in range(1, column_index_from_string("AJ") + 1):
            ws[f"{get_column_letter(col)}2"].font = bold

        for addr in ["A1", "C1", "G1", "M1", "Y1", "K2"]:
            ws[addr].font = bold

        center = Alignment(horizontal="center", vertical="center")
        left = Alignment(horizontal="left", vertical="center")

        for col in range(1, column_index_from_string("AJ") + 1):
            addr = f"{get_column_letter(col)}2"
            ws[addr].alignment = left if addr in ["A2", "B2"] else center

        num_format = "#,##0;[Red](#,##0)"
        numeric_cols = list("CDEFGHIJK")
        numeric_cols += [get_column_letter(c) for c in range(start_col_pay, start_col_pay + 12)]
        numeric_cols += [get_column_letter(c) for c in range(start_col_perf, start_col_perf + 12)]

        for col in numeric_cols:
            for r in range(3, last_row + 1):
                cell = ws[f"{col}{r}"]
                cell.alignment = center
                cell.number_format = num_format

        for r in range(1, last_row + 1):
            ws[f"A{r}"].alignment = left
            ws[f"B{r}"].alignment = left

        for addr in ["C1", "G1", "M1", "Y1"]:
            ws[addr].alignment = left

        ws.column_dimensions["A"].width = 38
        ws.column_dimensions["B"].width = 38
        for col in numeric_cols:
            ws.column_dimensions[col].width = 12.2 if column_index_from_string(col) >= start_col_pay else 12.6

        thin = Side(border_style="thin", color="000000")
        border_cols = ["C", "G", "K", "M", "Y"]
        for r in range(1, last_row + 1):
            for col in border_cols:
                cell = ws[f"{col}{r}"]
                cell.border = Border(
                    left=thin,
                    right=cell.border.right,
                    top=cell.border.top,
                    bottom=cell.border.bottom,
                )

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================
# Old money UI (Streamlit)
# =========================
st.set_page_config(page_title=" ", page_icon=" ", layout="wide", initial_sidebar_state="collapsed")

st.markdown(
    """
    <style>
      #MainMenu {visibility: hidden;}
      footer {visibility: hidden;}
      header {visibility: hidden;}

      /* Old money palette */
      :root{
        --bg1: #F6F1E6;         /* warm ivory */
        --bg2: #EFE6D6;         /* parchment */
        --card: rgba(255,255,255,0.68);
        --border: rgba(47, 63, 51, 0.18); /* muted green border */
        --text: #1F2A22;        /* deep forest text */
        --muted: rgba(31, 42, 34, 0.68);
        --accent: #2F3F33;      /* forest */
        --gold: #8A6F3A;        /* antique gold */
      }

      html, body, [class*="css"], .stApp, .stMarkdown, .stText, .stButton button, .stDownloadButton button {
        font-family: "Times New Roman", Times, serif !important;
        color: var(--text) !important;
      }

      .stApp {
        background:
          radial-gradient(900px circle at 10% 0%, rgba(138, 111, 58, 0.12), transparent 55%),
          radial-gradient(900px circle at 90% 10%, rgba(47, 63, 51, 0.10), transparent 55%),
          linear-gradient(180deg, var(--bg1) 0%, var(--bg2) 55%, var(--bg1) 100%);
      }

      .block-container {
        max-width: 980px;
        padding-top: 2.2rem;
        padding-bottom: 2.4rem;
      }

      .card {
        border-radius: 16px;
        padding: 18px;
        background: var(--card);
        border: 1px solid var(--border);
        box-shadow: 0 18px 42px rgba(0,0,0,0.10);
        backdrop-filter: blur(6px);
      }

      /* Clean labels */
      label, .stCaption, .stMarkdown p, .stMarkdown span {
        color: var(--text) !important;
      }

      /* Inputs look */
      [data-testid="stFileUploader"] section {
        border-radius: 14px;
        padding: 10px;
        border: 1px solid var(--border);
        background: rgba(255,255,255,0.55);
      }

      [role="radiogroup"] {
        border-radius: 14px;
        padding: 12px 12px 6px 12px;
        border: 1px solid var(--border);
        background: rgba(255,255,255,0.55);
      }

      /* Buttons: forest + gold */
      div.stButton > button {
        background: var(--accent) !important;
        color: #FFFFFF !important;
        border: 1px solid rgba(0,0,0,0.08) !important;
        border-radius: 12px !important;
        padding: 0.70rem 1rem !important;
        font-weight: 700 !important;
      }
      div.stButton > button:hover {
        background: #253328 !important;
      }

      div.stDownloadButton > button {
        background: #FFFFFF !important;
        color: var(--accent) !important;
        border: 1px solid rgba(47,63,51,0.30) !important;
        border-radius: 12px !important;
        padding: 0.70rem 1rem !important;
        font-weight: 700 !important;
      }
      div.stDownloadButton > button:hover {
        border-color: rgba(138,111,58,0.55) !important;
        box-shadow: 0 10px 24px rgba(138,111,58,0.12);
      }

      /* Progress bar */
      div[data-testid="stProgress"] > div > div {
        background-color: var(--gold) !important;
      }
    </style>
    """,
    unsafe_allow_html=True,
)

# No title/hero block as requested
left, right = st.columns([1.05, 0.95], gap="large")

with left:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    uploaded = st.file_uploader("Загрузите Excel файл", type=["xlsx", "xlsm"])
    st.markdown("</div>", unsafe_allow_html=True)

    st.write("")

    st.markdown('<div class="card">', unsafe_allow_html=True)
    mode = st.radio(
        "Выберите обработку",
        options=["Код 1 — Сальдо", "Код 2 — Контракты", "Оба (Код 1 → Код 2)"],
        index=0,
    )
    st.markdown("</div>", unsafe_allow_html=True)

with right:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.markdown("**Запуск и скачивание**")
    st.caption("Нажмите «Обработать». После завершения появится кнопка скачивания готового файла.")

    run_btn = st.button("Обработать", type="primary", disabled=(uploaded is None))

    status_box = st.empty()
    progress = st.progress(0)

    if uploaded is not None:
        size_mb = len(uploaded.getvalue()) / (1024 * 1024)
        st.markdown(f"Файл: `{uploaded.name}`  \nРазмер: `{size_mb:.2f} MB`")

    st.markdown("</div>", unsafe_allow_html=True)

st.write("")

if run_btn and uploaded is not None:
    file_bytes = uploaded.getvalue()

    try:
        status_box.info("Подготовка...")
        progress.progress(10)
        time.sleep(0.12)

        out_bytes = file_bytes

        if mode in ["Код 1 — Сальдо", "Оба (Код 1 → Код 2)"]:
            status_box.info("Выполняется обработка (Сальдо)...")
            progress.progress(35)
            out_bytes = run_code_1(out_bytes)
            progress.progress(60)

        if mode in ["Код 2 — Контракты", "Оба (Код 1 → Код 2)"]:
            status_box.info("Выполняется обработка (Контракты)...")
            progress.progress(75)
            out_bytes = run_code_2(out_bytes)
            progress.progress(92)

        status_box.success("Готово. Можно скачать результат.")
        progress.progress(100)

        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown("**Результат**")

        st.download_button(
            label="Скачать обработанный Excel",
            data=out_bytes,
            file_name=f"processed_{uploaded.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)

    except Exception as e:
        progress.progress(0)
        status_box.error(f"Ошибка: {e}")
