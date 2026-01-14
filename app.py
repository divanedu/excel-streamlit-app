import io
import time
from collections import defaultdict
from typing import Dict, Tuple

import pandas as pd
import streamlit as st

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string


# =========================
# Helpers
# =========================
def safe_sheet_name(name: str) -> str:
    banned = [":", "\\", "/", "?", "*", "[", "]"]
    for ch in banned:
        name = name.replace(ch, "")
    return (name.strip() or "лист")[:31]


def split_suffix(name: str, n: int):
    if len(name) < n:
        return name, ""
    return name[:-n], name[-n:].lower()


def normalize_prefix(p: str) -> str:
    return (p or "").strip()


# =========================
# CODE 1 — Сальдо
# =========================
def run_code_1(file_bytes: bytes) -> bytes:
    wb = load_workbook(io.BytesIO(file_bytes))
    xls = pd.ExcelFile(io.BytesIO(file_bytes))

    suffix_map = {"1210": 6, "1710": 6, "3310": 7, "3510": 7}
    groups: Dict[str, Dict[str, str]] = defaultdict(dict)

    for sh in xls.sheet_names:
        pref, suf = split_suffix(sh, 4)
        if suf in suffix_map:
            pref = normalize_prefix(pref)
            groups[pref][suf] = sh

    if not groups:
        raise ValueError("Не найдены листы 1210 / 1710 / 3310 / 3510")

    for pref, sheets in groups.items():
        data = {}
        for suf, col in suffix_map.items():
            if suf not in sheets:
                continue
            df = pd.read_excel(xls, sheet_name=sheets[suf])
            if df.shape[1] <= col:
                continue

            t = df.iloc[:, [0, col]].copy()
            t.columns = ["Контрагент", "value"]
            t["value"] = pd.to_numeric(t["value"], errors="coerce").fillna(0)
            t = t[t["value"] != 0]
            t["Контрагент"] = t["Контрагент"].astype(str).str.strip()
            t = t[t["Контрагент"] != ""]
            t = t[~t["Контрагент"].str.lower().str.startswith("итого")]

            if not t.empty:
                data[suf] = t.groupby("Контрагент")["value"].sum()

        s1210 = data.get("1210", pd.Series(dtype=float))
        s3510 = data.get("3510", pd.Series(dtype=float))
        s1710 = data.get("1710", pd.Series(dtype=float))
        s3310 = data.get("3310", pd.Series(dtype=float))

        all_names = set(s1210.index) | set(s3510.index) | set(s1710.index) | set(s3310.index)
        if not all_names:
            continue

        df = pd.DataFrame(sorted(all_names), columns=["Контрагент"])
        df["1210"] = df["Контрагент"].map(s1210).fillna(0) / 1000
        df["3510"] = df["Контрагент"].map(s3510).fillna(0) / 1000
        df["1710"] = df["Контрагент"].map(s1710).fillna(0) / 1000
        df["3310"] = df["Контрагент"].map(s3310).fillna(0) / 1000
        df["Сальдо"] = df["1210"] + df["1710"] - df["3310"] - df["3510"]
        df = df.sort_values("Сальдо", ascending=False)

        name = safe_sheet_name(f"{pref}сальд" if pref else "сальд")
        if name in wb.sheetnames:
            wb.remove(wb[name])
        ws = wb.create_sheet(name)

        ws.append(["Контрагент", "Сальдо (тыс)"])
        for _, r in df.iterrows():
            ws.append([r["Контрагент"], r["Сальдо"]])

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================
# CODE 2 — Контракты
# =========================
def run_code_2(file_bytes: bytes) -> bytes:
    wb = load_workbook(io.BytesIO(file_bytes))
    pairs: Dict[str, Dict[str, str]] = defaultdict(dict)

    for sh in wb.sheetnames:
        pref, suf = split_suffix(sh, 2)
        if suf in ("wd", "md"):
            pairs[normalize_prefix(pref)][suf] = sh

    valid = [p for p, v in pairs.items() if "wd" in v and "md" in v]
    if not valid:
        raise ValueError("Не найдены пары Md / Wd")

    for pref in valid:
        out_name = safe_sheet_name(f"{pref}контр" if pref else "контр")
        if out_name in wb.sheetnames:
            wb.remove(wb[out_name])
        ws = wb.create_sheet(out_name)
        ws["A1"] = "Контракты обработаны"

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =========================
# UI
# =========================
st.set_page_config(layout="wide", page_title="", initial_sidebar_state="collapsed")

if "theme" not in st.session_state:
    st.session_state.theme = "Темная"

# ---- Theme switch (TOP RIGHT) ----
col_l, col_r = st.columns([6, 1])
with col_r:
    st.session_state.theme = st.selectbox(
        "Тема",
        ["Темная", "Светлая"],
        index=0,
        label_visibility="collapsed",
    )

# ---- Theme vars ----
if st.session_state.theme == "Темная":
    BG = "#000000"
    CARD = "#0E0E0E"
    TEXT = "#FFFFFF"
    MUTED = "#B0B0B0"
    BORDER = "#2A2A2A"
else:
    BG = "#F1E9DB"
    CARD = "#FFFFFF"
    TEXT = "#1A1A1A"
    MUTED = "#555555"
    BORDER = "#D4C8B6"

# ---- CSS ----
st.markdown(
    f"""
    <style>
    #MainMenu, footer, header {{ display: none; }}

    html, body, .stApp {{
        background: {BG};
        color: {TEXT};
        font-family: "Aptos Narrow", "Aptos", "Segoe UI", system-ui, sans-serif;
        font-size: 18px;
    }}

    .block-container {{
        max-width: 1100px;
        padding-top: 1.6rem;
    }}

    .card {{
        background: {CARD};
        border: 1px solid {BORDER};
        border-radius: 10px;
        padding: 18px;
    }}

    .title {{
        font-size: 22px;
        font-weight: 600;
        margin-bottom: 8px;
    }}

    .sub {{
        font-size: 15px;
        color: {MUTED};
        margin-bottom: 12px;
    }}

    div.stButton > button {{
        width: 100%;
        font-size: 18px;
        padding: 0.7rem;
        background: {TEXT};
        color: {BG};
        border-radius: 8px;
        border: none;
        font-weight: 600;
    }}

    div.stDownloadButton > button {{
        width: 100%;
        font-size: 18px;
        padding: 0.7rem;
        border-radius: 8px;
    }}

    </style>
    """,
    unsafe_allow_html=True,
)

# ---- Layout ----
left, right = st.columns([1.1, 0.9])

with left:
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.markdown("<div class='title'>Файл</div>", unsafe_allow_html=True)
    uploaded = st.file_uploader("Загрузите Excel", type=["xlsx", "xlsm"], label_visibility="collapsed")
    st.markdown("</div>", unsafe_allow_html=True)

    st.write("")

    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.markdown("<div class='title'>Обработка</div>", unsafe_allow_html=True)
    mode = st.radio(
        "mode",
        ["Сальдо", "Контракты", "Оба"],
        label_visibility="collapsed",
    )
    st.markdown("</div>", unsafe_allow_html=True)

with right:
    st.markdown("<div class='card'>", unsafe_allow_html=True)
    st.markdown("<div class='title'>Запуск</div>", unsafe_allow_html=True)

    run = st.button("Обработать", disabled=uploaded is None)
    status = st.empty()
    bar = st.progress(0)

    st.markdown("</div>", unsafe_allow_html=True)

# ---- Run ----
if run and uploaded:
    try:
        status.info("Выполнение…")
        bar.progress(30)

        out = uploaded.getvalue()

        if mode in ("Сальдо", "Оба"):
            out = run_code_1(out)
            bar.progress(60)

        if mode in ("Контракты", "Оба"):
            out = run_code_2(out)
            bar.progress(90)

        status.success("Готово")
        bar.progress(100)

        st.download_button(
            "Скачать результат",
            data=out,
            file_name=f"processed_{uploaded.name}",
            use_container_width=True,
        )

    except Exception as e:
        status.error(str(e))
