"""
Load & chuẩn hóa raw data MBNT / TTQT từ Excel hoặc CSV
Ánh xạ cột theo README Excel template
"""

from __future__ import annotations
import pandas as pd
from typing import Optional, Tuple, Callable
from pathlib import Path

try:
    import streamlit as st
except ImportError:
    st = None


# Mapping cột chuẩn (index 0-based theo header Excel)
MBNT_COLS = {
    "Ngay": ["Ngày khởi tạo", "Ngay_khoi_tao", "Ngày giao dịch", "Ngay_GD", "Ngày", "Ngay", "A"],
    "Ma_CIF": ["Ma_CIF", "Mã CIF", "CIF", "MaCIF", "D"],
    "Ten_Khach_Hang": ["Ten_Khach_Hang", "Tên Khách Hàng", "TenKH", "E"],
    "Phong_Ban": ["Phong_Ban", "Phòng Ban", "Phòng", "Phong", "F"],
    "Chieu_NH": ["Chieu_NH", "Chiều NH", "Chieu", "M"],
    "DS_Quy_USD": ["DS_Quy_USD", "DS", "Doanh_So", "T"],
    "Loi_Nhuan_VND": ["Loi_Nhuan_VND", "LN", "Loi_Nhuan", "U"],
    "Nguon_Mua_KH": ["Nguon_Mua_KH", "Nguon", "W"],
    "Muc_Dich_KH": ["Muc_Dich_KH", "Muc_Dich", "X"],
}

TTQT_COLS = {
    "Ngay_GD": ["Ngay_GD", "Ngày GD", "Ngay", "A"],
    "Ma_CIF": ["Ma_CIF", "Mã CIF", "CIF", "B"],
    "Ten_Khach_Hang": ["Ten_Khach_Hang", "Tên Khách Hàng", "C"],
    "Phong_Ban": ["Phong_Ban", "Phòng Ban", "Phòng", "D"],
    "DS_Quy_USD": ["DS_Quy_USD", "DS", "AMOUNT_USD", "E"],
}


def _find_col(df: pd.DataFrame, candidates: list) -> Optional[str]:
    cols_lower = {str(c).strip().lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.lower().strip()
        if key in cols_lower:
            return cols_lower[key]
        # partial
        for k, orig in cols_lower.items():
            if key in k or k in key:
                return orig
    return None


def normalize_mbnt(df: pd.DataFrame) -> pd.DataFrame:
    """Chuẩn hóa DataFrame MBNT về schema thống nhất"""
    if df is None or df.empty:
        return pd.DataFrame()

    out = pd.DataFrame()
    mapping = {}
    for std, cands in MBNT_COLS.items():
        found = _find_col(df, cands)
        if found:
            mapping[std] = found

    if "Ngay" not in mapping:
        # fallback cột A (index 0)
        if len(df.columns) > 0:
            mapping["Ngay"] = df.columns[0]
    if "Ma_CIF" not in mapping and len(df.columns) > 3:
        mapping["Ma_CIF"] = df.columns[3]
    if "Phong_Ban" not in mapping and len(df.columns) > 5:
        mapping["Phong_Ban"] = df.columns[5]
    if "Chieu_NH" not in mapping and len(df.columns) > 12:
        mapping["Chieu_NH"] = df.columns[12]
    if "DS_Quy_USD" not in mapping and len(df.columns) > 19:
        mapping["DS_Quy_USD"] = df.columns[19]
    if "Loi_Nhuan_VND" not in mapping and len(df.columns) > 20:
        mapping["Loi_Nhuan_VND"] = df.columns[20]

    for std, src in mapping.items():
        out[std] = df[src]

    # Bổ sung cột thiếu
    for col in ["Ngay", "Ma_CIF", "Ten_Khach_Hang", "Phong_Ban", "Chieu_NH", "DS_Quy_USD", "Loi_Nhuan_VND", "Nguon_Mua_KH", "Muc_Dich_KH"]:
        if col not in out.columns:
            out[col] = None

    # Numeric
    out["DS_Quy_USD"] = pd.to_numeric(out["DS_Quy_USD"], errors="coerce").fillna(0)
    out["Loi_Nhuan_VND"] = pd.to_numeric(out["Loi_Nhuan_VND"], errors="coerce").fillna(0)
    out["Ngay"] = pd.to_datetime(out["Ngay"], errors="coerce")

    # Drop hàng trống hoàn toàn
    out = out.dropna(how="all", subset=["Ma_CIF", "DS_Quy_USD"])
    out = out[out["Ma_CIF"].astype(str).str.strip() != ""]
    return out.reset_index(drop=True)


def normalize_ttqt(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    out = pd.DataFrame()
    mapping = {}
    for std, cands in TTQT_COLS.items():
        found = _find_col(df, cands)
        if found:
            mapping[std] = found

    if "Ngay_GD" not in mapping and len(df.columns) > 0:
        mapping["Ngay_GD"] = df.columns[0]
    if "Ma_CIF" not in mapping and len(df.columns) > 1:
        mapping["Ma_CIF"] = df.columns[1]
    if "Phong_Ban" not in mapping and len(df.columns) > 3:
        mapping["Phong_Ban"] = df.columns[3]
    if "DS_Quy_USD" not in mapping and len(df.columns) > 4:
        mapping["DS_Quy_USD"] = df.columns[4]

    for std, src in mapping.items():
        out[std] = df[src]

    for col in ["Ngay_GD", "Ma_CIF", "Ten_Khach_Hang", "Phong_Ban", "DS_Quy_USD"]:
        if col not in out.columns:
            out[col] = None

    out["DS_Quy_USD"] = pd.to_numeric(out["DS_Quy_USD"], errors="coerce").fillna(0)
    out["Ngay_GD"] = pd.to_datetime(out["Ngay_GD"], errors="coerce")
    out = out.dropna(how="all", subset=["Ma_CIF"])
    return out.reset_index(drop=True)


def _cache(fn):
    if st is not None:
        return st.cache_data(show_spinner="Đang đọc file Excel...")(fn)
    return fn

@_cache
def load_excel_sheets(file_bytes: bytes) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """Đọc 2 sheet raw từ file Excel template"""
    try:
        xl = pd.ExcelFile(file_bytes, engine="openpyxl")
        mbnt = pd.DataFrame()
        ttqt = pd.DataFrame()

        # Tìm sheet MBNT
        for name in xl.sheet_names:
            if "mbnt" in name.lower() or "03_data" in name.lower():
                raw = pd.read_excel(xl, sheet_name=name, header=0)
                mbnt = normalize_mbnt(raw)
                break
        if mbnt.empty and "03_Data_Raw_MBNT" in xl.sheet_names:
            mbnt = normalize_mbnt(pd.read_excel(xl, sheet_name="03_Data_Raw_MBNT", header=0))

        # TTQT
        for name in xl.sheet_names:
            if "ttqt" in name.lower() or "tttm" in name.lower():
                raw = pd.read_excel(xl, sheet_name=name, header=0)
                # Bỏ 2 dòng header đặc biệt nếu có
                if len(raw) > 2:
                    raw = raw.iloc[1:].reset_index(drop=True)
                ttqt = normalize_ttqt(raw)
                break

        return mbnt, ttqt
    except Exception as e:
        if st is not None:
            st.error(f"Lỗi đọc Excel: {e}")
        else:
            print(f"Lỗi đọc Excel: {e}")
        return pd.DataFrame(), pd.DataFrame()


def load_csv_mbnt(file) -> pd.DataFrame:
    df = pd.read_csv(file)
    return normalize_mbnt(df)


def load_csv_ttqt(file) -> pd.DataFrame:
    df = pd.read_csv(file)
    return normalize_ttqt(df)
