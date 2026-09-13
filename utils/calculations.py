"""
MBNT Dashboard VCB - Calculation Engine
Tái hiện logic Excel: SUMIFS, RANK, Pareto 20%/10%, IsFirst (unique CIF)
Tương thích nguyên tắc: không circular, chỉ dùng aggregation + rank
"""

from __future__ import annotations
import pandas as pd
import numpy as np
from datetime import datetime, date
from typing import Optional, Tuple, List, Dict


# Danh sách 15 phòng chuẩn (khớp Danh_Sach_Phong)
PHONG_LIST = [
    "Khách hàng doanh nghiệp 1",
    "Khách hàng doanh nghiệp 2",
    "Khách hàng doanh nghiệp 3",
    "Khách hàng FDI",
    "Khách hàng bán lẻ 1",
    "Khách hàng bán lẻ 2",
    "Khách hàng bán lẻ 3",
    "Dịch vụ khách hàng tổ chức 1",
    "Dịch vụ khách hàng tổ chức 2",
    "PGD Tôn Đức Thắng",
    "PGD Mê Linh",
    "PGD Đồng Khởi",
    "PGD Cộng Hòa",
    "PGD Võ Văn Kiệt",
    "PGD Lê Thánh Tôn",
]


def ensure_datetime(series: pd.Series) -> pd.Series:
    """Chuyển cột ngày sang datetime, lỗi → NaT"""
    return pd.to_datetime(series, errors="coerce")


def filter_by_date(
    df: pd.DataFrame,
    date_col: str,
    start: Optional[date | datetime] = None,
    end: Optional[date | datetime] = None,
) -> pd.DataFrame:
    """Lọc theo khoảng ngày (tương đương SUMIFS date range)"""
    if df is None or df.empty:
        return df
    out = df.copy()
    out[date_col] = ensure_datetime(out[date_col])
    if start is not None:
        out = out[out[date_col] >= pd.Timestamp(start)]
    if end is not None:
        out = out[out[date_col] <= pd.Timestamp(end)]
    return out


def filter_by_phong(df: pd.DataFrame, phong_col: str, phong: Optional[str] = None) -> pd.DataFrame:
    """Lọc theo phòng (dropdown)"""
    if df is None or df.empty or not phong or phong == "Tất cả":
        return df
    return df[df[phong_col].astype(str).str.strip() == phong.strip()].copy()


def aggregate_cif(
    df: pd.DataFrame,
    cif_col: str = "Ma_CIF",
    name_col: str = "Ten_Khach_Hang",
    phong_col: str = "Phong_Ban",
    ds_col: str = "DS_Quy_USD",
    ln_col: str = "Loi_Nhuan_VND",
) -> pd.DataFrame:
    """
    Gom nhóm theo CIF (IsFirst logic):
    - Tổng DS_USD, LN_VND
    - Giữ tên + phòng (lấy first)
    """
    if df is None or df.empty:
        return pd.DataFrame(columns=[cif_col, name_col, phong_col, "DS_USD", "LN_VND"])

    # Chuẩn hóa
    work = df.copy()
    work[cif_col] = work[cif_col].astype(str).str.strip()
    work = work[work[cif_col].notna() & (work[cif_col] != "") & (work[cif_col] != "0")]

    if work.empty:
        return pd.DataFrame(columns=[cif_col, name_col, phong_col, "DS_USD", "LN_VND"])

    # Numeric
    work[ds_col] = pd.to_numeric(work[ds_col], errors="coerce").fillna(0)
    work[ln_col] = pd.to_numeric(work[ln_col], errors="coerce").fillna(0)

    agg = (
        work.groupby(cif_col, as_index=False)
        .agg(
            {
                name_col: "first",
                phong_col: "first",
                ds_col: "sum",
                ln_col: "sum",
            }
        )
        .rename(columns={ds_col: "DS_USD", ln_col: "LN_VND"})
    )
    return agg


def add_rank_and_pareto(
    cif_df: pd.DataFrame,
    value_col: str,
    ascending: bool = False,
) -> pd.DataFrame:
    """
    Thêm Rank, % tỷ trọng, % cộng dồn (Pareto)
    Rank 1 = cao nhất (descending)
    """
    if cif_df is None or cif_df.empty:
        return cif_df

    out = cif_df.copy()
    total = out[value_col].sum()
    if total == 0:
        out["Rank"] = 0
        out["Pct"] = 0.0
        out["CumPct"] = 0.0
        return out

    out = out.sort_values(value_col, ascending=ascending).reset_index(drop=True)
    out["Rank"] = range(1, len(out) + 1)
    out["Pct"] = out[value_col] / total
    out["CumPct"] = out["Pct"].cumsum()
    return out


def get_top_pareto(
    cif_df: pd.DataFrame,
    value_col: str,
    pct_threshold: float = 0.20,
    n_max: int = 50,
) -> pd.DataFrame:
    """Top CIF đóng góp đến pct_threshold (20% hoặc 10%)"""
    ranked = add_rank_and_pareto(cif_df, value_col, ascending=False)
    if ranked.empty:
        return ranked
    # Lấy đến khi CumPct >= threshold, hoặc tối đa n_max
    mask = ranked["CumPct"] <= pct_threshold + 1e-9
    top = ranked[mask]
    if top.empty:
        top = ranked.head(1)
    else:
        # Thêm 1 dòng vượt ngưỡng nếu cần
        last_idx = top.index[-1]
        if last_idx + 1 < len(ranked) and ranked.loc[last_idx, "CumPct"] < pct_threshold:
            top = ranked.loc[: last_idx + 1]
    return top.head(n_max)


def get_bottom_pareto(
    cif_df: pd.DataFrame,
    value_col: str,
    pct_threshold: float = 0.10,
    n_max: int = 30,
) -> pd.DataFrame:
    """Bottom 10% (thấp nhất) – rank ascending"""
    ranked = add_rank_and_pareto(cif_df, value_col, ascending=True)
    if ranked.empty:
        return ranked
    mask = ranked["CumPct"] <= pct_threshold + 1e-9
    bottom = ranked[mask]
    if bottom.empty:
        bottom = ranked.head(1)
    return bottom.head(n_max)


def filter_chieu(
    df: pd.DataFrame,
    chieu_col: str = "Chieu_NH",
    chieu: str = "BÁN",
) -> pd.DataFrame:
    """Lọc BÁN / MUA (khớp đúng giá trị cột M)"""
    if df is None or df.empty:
        return df
    return df[df[chieu_col].astype(str).str.strip().str.upper() == chieu.upper()].copy()


def filter_muc_dich(
    df: pd.DataFrame,
    col: str,
    pattern: str = "*trả*",
) -> pd.DataFrame:
    """Wildcard filter như Excel *trả* / *vay*"""
    if df is None or df.empty:
        return df
    # Chuyển *xxx* → contains
    clean = pattern.strip("*").lower()
    return df[df[col].astype(str).str.lower().str.contains(clean, na=False)].copy()


def compute_kpis(
    mbnt: pd.DataFrame,
    ttqt: Optional[pd.DataFrame] = None,
    start: Optional[date] = None,
    end: Optional[date] = None,
    phong: Optional[str] = None,
) -> Dict[str, float]:
    """10 KPI chính như Dashboard"""
    mbnt_f = filter_by_date(mbnt, "Ngay", start, end)
    mbnt_f = filter_by_phong(mbnt_f, "Phong_Ban", phong)

    ds_mbnt = pd.to_numeric(mbnt_f.get("DS_Quy_USD", 0), errors="coerce").fillna(0).sum() / 1_000_000  # tr.USD
    ln_mbnt = pd.to_numeric(mbnt_f.get("Loi_Nhuan_VND", 0), errors="coerce").fillna(0).sum() / 1_000_000_000  # tỷ VND
    so_gd = len(mbnt_f)
    so_cif = mbnt_f["Ma_CIF"].nunique() if "Ma_CIF" in mbnt_f.columns else 0

    # Bán-Trả nợ / Mua-Vay (ước lượng từ Chieu + Muc_Dich)
    ban = filter_chieu(mbnt_f, "Chieu_NH", "BÁN")
    mua = filter_chieu(mbnt_f, "Chieu_NH", "MUA")
    ds_ban = pd.to_numeric(ban.get("DS_Quy_USD", 0), errors="coerce").fillna(0).sum() / 1_000_000
    ds_mua = pd.to_numeric(mua.get("DS_Quy_USD", 0), errors="coerce").fillna(0).sum() / 1_000_000

    # TTQT
    ds_ttqt = 0.0
    so_gd_ttqt = 0
    if ttqt is not None and not ttqt.empty:
        ttqt_f = filter_by_date(ttqt, "Ngay_GD", start, end)
        ttqt_f = filter_by_phong(ttqt_f, "Phong_Ban", phong)
        ds_ttqt = pd.to_numeric(ttqt_f.get("DS_Quy_USD", 0), errors="coerce").fillna(0).sum() / 1_000_000
        so_gd_ttqt = len(ttqt_f)

    return {
        "ds_mbnt": round(ds_mbnt, 2),
        "ln_mbnt": round(ln_mbnt, 2),
        "ds_ttqt": round(ds_ttqt, 2),
        "so_gd": so_gd,
        "so_cif": so_cif,
        "ds_ban": round(ds_ban, 2),
        "ds_mua": round(ds_mua, 2),
        "so_gd_ttqt": so_gd_ttqt,
    }


def prepare_pareto_tables(
    mbnt: pd.DataFrame,
    ttqt: Optional[pd.DataFrame] = None,
    start: Optional[date] = None,
    end: Optional[date] = None,
    phong: Optional[str] = None,
) -> Dict[str, pd.DataFrame]:
    """Tạo 6 bảng Pareto + 4 bảng Bán/Mua như Dashboard"""
    mbnt_f = filter_by_date(mbnt, "Ngay", start, end)
    mbnt_f = filter_by_phong(mbnt_f, "Phong_Ban", phong)

    cif_mbnt = aggregate_cif(mbnt_f)

    # Top 20% DS / LN MBNT
    top20_ds = get_top_pareto(cif_mbnt, "DS_USD", 0.20)
    top20_ln = get_top_pareto(cif_mbnt, "LN_VND", 0.20)

    # Bottom 10% DS / LN MBNT
    bot10_ds = get_bottom_pareto(cif_mbnt, "DS_USD", 0.10)
    bot10_ln = get_bottom_pareto(cif_mbnt, "LN_VND", 0.10)

    # TTQT top/bottom
    top20_ttqt = pd.DataFrame()
    bot10_ttqt = pd.DataFrame()
    if ttqt is not None and not ttqt.empty:
        ttqt_f = filter_by_date(ttqt, "Ngay_GD", start, end)
        ttqt_f = filter_by_phong(ttqt_f, "Phong_Ban", phong)
        cif_ttqt = aggregate_cif(
            ttqt_f,
            cif_col="Ma_CIF",
            name_col="Ten_Khach_Hang",
            phong_col="Phong_Ban",
            ds_col="DS_Quy_USD",
            ln_col="DS_Quy_USD",  # TTQT không có LN riêng
        )
        cif_ttqt = cif_ttqt.rename(columns={"DS_USD": "DS_TTQT"})
        top20_ttqt = get_top_pareto(cif_ttqt, "DS_TTQT", 0.20)
        bot10_ttqt = get_bottom_pareto(cif_ttqt, "DS_TTQT", 0.10)

    # Bán-Trả nợ / Mua-Vay
    ban = filter_chieu(mbnt_f, "Chieu_NH", "BÁN")
    mua = filter_chieu(mbnt_f, "Chieu_NH", "MUA")
    # Có thể tinh chỉnh thêm Muc_Dich nếu cần

    cif_ban = aggregate_cif(ban)
    cif_mua = aggregate_cif(mua)
    top_ban_ds = get_top_pareto(cif_ban, "DS_USD", 0.20)
    top_ban_ln = get_top_pareto(cif_ban, "LN_VND", 0.20)
    top_mua_ds = get_top_pareto(cif_mua, "DS_USD", 0.20)
    top_mua_ln = get_top_pareto(cif_mua, "LN_VND", 0.20)

    def fmt(df: pd.DataFrame, value_cols: List[str]) -> pd.DataFrame:
        if df.empty:
            return df
        out = df.copy()
        for c in value_cols:
            if c in out.columns:
                if "LN" in c or "VND" in c:
                    out[c] = (out[c] / 1_000_000_000).round(3)  # tỷ
                else:
                    out[c] = (out[c] / 1_000_000).round(3)  # tr.USD
        if "Pct" in out.columns:
            out["Pct"] = (out["Pct"] * 100).round(2)
        if "CumPct" in out.columns:
            out["CumPct"] = (out["CumPct"] * 100).round(2)
        return out

    return {
        "top20_ds": fmt(top20_ds, ["DS_USD"]),
        "top20_ln": fmt(top20_ln, ["LN_VND"]),
        "bot10_ds": fmt(bot10_ds, ["DS_USD"]),
        "bot10_ln": fmt(bot10_ln, ["LN_VND"]),
        "top20_ttqt": fmt(top20_ttqt, ["DS_TTQT"]) if not top20_ttqt.empty else top20_ttqt,
        "bot10_ttqt": fmt(bot10_ttqt, ["DS_TTQT"]) if not bot10_ttqt.empty else bot10_ttqt,
        "top_ban_ds": fmt(top_ban_ds, ["DS_USD"]),
        "top_ban_ln": fmt(top_ban_ln, ["LN_VND"]),
        "top_mua_ds": fmt(top_mua_ds, ["DS_USD"]),
        "top_mua_ln": fmt(top_mua_ln, ["LN_VND"]),
    }


def monthly_th_from_raw(
    mbnt: pd.DataFrame,
    year: int = 2026,
    value_col: str = "DS_Quy_USD",
    phong_col: str = "Phong_Ban",
) -> pd.DataFrame:
    """Tính thực hiện (TH) theo tháng & phòng từ raw – cho sheet Kế hoạch"""
    if mbnt is None or mbnt.empty:
        return pd.DataFrame()

    work = mbnt.copy()
    work["Ngay"] = ensure_datetime(work["Ngay"])
    work = work[work["Ngay"].dt.year == year]
    work[value_col] = pd.to_numeric(work[value_col], errors="coerce").fillna(0)
    work["Thang"] = work["Ngay"].dt.month

    pivot = (
        work.groupby([phong_col, "Thang"])[value_col]
        .sum()
        .unstack(fill_value=0)
        .reindex(columns=range(1, 13), fill_value=0)
    )
    pivot.columns = [f"T{m}" for m in range(1, 13)]
    pivot["Cả Năm"] = pivot.sum(axis=1)
    return pivot.reset_index()
