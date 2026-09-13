"""
MBNT Dashboard VCB – Web Version
Tái hiện cấu trúc Excel: Dashboard / Báo cáo phòng / Kế hoạch / Báo cáo lời
Deploy: Render / GitHub
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import date, datetime, timedelta
from pathlib import Path
import sys
import io

sys.path.insert(0, str(Path(__file__).parent))

from utils.calculations import (
    PHONG_LIST,
    compute_kpis,
    prepare_pareto_tables,
    monthly_th_from_raw,
    filter_by_date,
    filter_by_phong,
)
from utils.data_loader import load_excel_sheets, load_csv_mbnt, load_csv_ttqt, normalize_mbnt, normalize_ttqt

# ──────────────────────────────────────────────
# Page config
# ──────────────────────────────────────────────
st.set_page_config(
    page_title="MBNT Dashboard | Vietcombank",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom CSS – professional banking look
st.markdown(
    """
<style>
    .main-header {
        font-size: 1.8rem;
        font-weight: 700;
        color: #003366;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        color: #555;
        font-size: 0.95rem;
        margin-bottom: 1.5rem;
    }
    .kpi-card {
        background: linear-gradient(135deg, #f0f7ff 0%, #e6f0fa 100%);
        border-radius: 12px;
        padding: 1rem 1.2rem;
        border-left: 5px solid #003366;
        box-shadow: 0 2px 6px rgba(0,0,0,0.06);
    }
    .kpi-label { font-size: 0.8rem; color: #555; margin-bottom: 0.3rem; }
    .kpi-value { font-size: 1.6rem; font-weight: 700; color: #003366; }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px 8px 0 0;
        padding: 8px 16px;
        background: #f5f5f5;
    }
    div[data-testid="stMetricValue"] { font-size: 1.4rem; }
</style>
""",
    unsafe_allow_html=True,
)

# ──────────────────────────────────────────────
# Sidebar – Data & Filters
# ──────────────────────────────────────────────
# Navigation state
if "nav_page" not in st.session_state:
    st.session_state.nav_page = "🏠 Home / Dashboard"

NAV_OPTIONS = [
    "🏠 Home / Dashboard",
    "🏢 Báo cáo phòng",
    "📋 Kế hoạch",
    "📝 Báo cáo lời",
    "📁 Data",
    "📖 Hướng dẫn",
]

with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/1b/Vietcombank_logo.svg/200px-Vietcombank_logo.svg.png", width=160)

    # ── Điều hướng ──
    st.markdown("### 🧭 Điều hướng")
    if st.button("🏠 Về Home", use_container_width=True, type="primary"):
        st.session_state.nav_page = "🏠 Home / Dashboard"
        st.rerun()

    st.session_state.nav_page = st.radio(
        "Chọn trang",
        NAV_OPTIONS,
        index=NAV_OPTIONS.index(st.session_state.nav_page) if st.session_state.nav_page in NAV_OPTIONS else 0,
        label_visibility="collapsed",
        key="nav_radio",
    )

    st.divider()
    st.markdown("### 📁 Nạp dữ liệu")
    st.caption("Tải template → điền data → upload lại")

    _tpl_dir = Path(__file__).parent / "data"
    _tpl_xlsx = _tpl_dir / "Template_MBNT_TTQT.xlsx"
    _tpl_mbnt_csv = _tpl_dir / "Template_MBNT.csv"
    _tpl_ttqt_csv = _tpl_dir / "Template_TTQT.csv"
    _tpl_phong = _tpl_dir / "Danh_Sach_Phong.csv"

    st.markdown("**1. Tải template**")
    t1, t2, t3 = st.columns(3)
    with t1:
        if _tpl_xlsx.exists():
            st.download_button(
                "Excel",
                data=_tpl_xlsx.read_bytes(),
                file_name="Template_MBNT_TTQT.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="dl_tpl_xlsx",
                help="Sheet: MBNT + TTQT + Danh_Sach_Phong",
            )
    with t2:
        if _tpl_mbnt_csv.exists():
            st.download_button(
                "MBNT",
                data=_tpl_mbnt_csv.read_bytes(),
                file_name="Template_MBNT.csv",
                mime="text/csv",
                use_container_width=True,
                key="dl_tpl_mbnt",
            )
    with t3:
        if _tpl_ttqt_csv.exists():
            st.download_button(
                "TTQT",
                data=_tpl_ttqt_csv.read_bytes(),
                file_name="Template_TTQT.csv",
                mime="text/csv",
                use_container_width=True,
                key="dl_tpl_ttqt",
            )
    if _tpl_phong.exists():
        st.download_button(
            "📋 Danh sách 15 phòng",
            data=_tpl_phong.read_bytes(),
            file_name="Danh_Sach_Phong.csv",
            mime="text/csv",
            use_container_width=True,
            key="dl_tpl_phong",
        )

    st.markdown("**2. Upload file đã điền**")
    uploaded = st.file_uploader(
        "Excel (.xlsx) – sheet MBNT / TTQT",
        type=["xlsx", "xls"],
        help="Dùng Template_MBNT_TTQT.xlsx đã điền, hoặc file Excel gốc",
    )

    col_a, col_b = st.columns(2)
    with col_a:
        csv_mbnt = st.file_uploader("CSV MBNT", type=["csv"], key="csv_mbnt")
    with col_b:
        csv_ttqt = st.file_uploader("CSV TTQT", type=["csv"], key="csv_ttqt")

    st.divider()
    st.markdown("### 📅 Kỳ lọc")
    today = date.today()
    default_start = date(today.year, today.month, 1) - timedelta(days=30)
    start_date = st.date_input("Từ ngày", value=default_start, key="start")
    end_date = st.date_input("Đến ngày", value=today, key="end")

    st.markdown("### 🏢 Phòng")
    phong_options = ["Tất cả"] + PHONG_LIST
    selected_phong = st.selectbox("Chọn phòng", phong_options, index=0)

    st.divider()
    st.caption("Tên phòng phải khớp 100% cột Phong_Ban")
    st.caption("Chieu_NH = BÁN hoặc MUA · Ngày = DATE")

# ──────────────────────────────────────────────
# Load data
# ──────────────────────────────────────────────
@st.cache_data
def get_sample_data():
    """Sample demo cả năm 2026 – mỗi tháng khác nhau để thấy rõ khi đổi kỳ lọc"""
    import numpy as np
    rng = np.random.default_rng(42)
    rows = []
    for month in range(1, 13):
        n_m = 40 + month * 3  # tháng sau nhiều GD hơn
        for i in range(n_m):
            day = int(rng.integers(1, 28))
            ds = float(rng.uniform(20_000, 500_000) * (1 + month * 0.15))
            rows.append({
                "Ngay": pd.Timestamp(year=2026, month=month, day=day),
                "Ma_CIF": f"CIF{1000 + (month * 10 + i) % 50}",
                "Ten_Khach_Hang": f"Khách hàng {(month * 10 + i) % 50 + 1}",
                "Phong_Ban": PHONG_LIST[(month + i) % len(PHONG_LIST)],
                "Chieu_NH": "BÁN" if i % 2 == 0 else "MUA",
                "DS_Quy_USD": ds,
                "Loi_Nhuan_VND": ds * float(rng.uniform(80, 250)),
                "Nguon_Mua_KH": "Nội địa",
                "Muc_Dich_KH": ["Trả nợ vay", "Thanh toán nhập khẩu", "Vay ngoại tệ"][i % 3],
            })
    mbnt = pd.DataFrame(rows)
    ttqt = mbnt.sample(min(200, len(mbnt)), random_state=42).copy()
    ttqt = ttqt.rename(columns={"Ngay": "Ngay_GD"})
    ttqt["DS_Quy_USD"] = ttqt["DS_Quy_USD"] * 0.6
    return mbnt, ttqt


mbnt_df = pd.DataFrame()
ttqt_df = pd.DataFrame()

if uploaded is not None:
    mbnt_df, ttqt_df = load_excel_sheets(uploaded.getvalue())
elif csv_mbnt is not None:
    mbnt_df = load_csv_mbnt(csv_mbnt)
    if csv_ttqt is not None:
        ttqt_df = load_csv_ttqt(csv_ttqt)
else:
    mbnt_df, ttqt_df = get_sample_data()
    st.sidebar.info("⚡ Đang dùng dữ liệu mẫu demo. Upload file thật để phân tích thực tế.")

# Session state cho kế hoạch
if "ke_hoach" not in st.session_state:
    st.session_state.ke_hoach = {
        "DS_MBNT": {p: {f"T{m}": 0.0 for m in range(1, 13)} for p in PHONG_LIST},
        "LN_MBNT": {p: {f"T{m}": 0.0 for m in range(1, 13)} for p in PHONG_LIST},
        "DS_TTQT": {p: {f"T{m}": 0.0 for m in range(1, 13)} for p in PHONG_LIST},
    }


def _df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8-sig")


def _df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Data") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return buf.getvalue()


# Sidebar: View / Download data đang load
with st.sidebar:
    st.divider()
    st.markdown("### 👁 Data đang load")
    st.caption(f"MBNT: **{len(mbnt_df):,}** dòng | TTQT: **{len(ttqt_df):,}** dòng")

    show_mbnt = st.checkbox("Xem MBNT", key="view_mbnt")
    show_ttqt = st.checkbox("Xem TTQT", key="view_ttqt")

    c_dl1, c_dl2 = st.columns(2)
    with c_dl1:
        if not mbnt_df.empty:
            st.download_button(
                "⬇ MBNT CSV",
                data=_df_to_csv_bytes(mbnt_df),
                file_name=f"MBNT_{start_date}_{end_date}.csv",
                mime="text/csv",
                use_container_width=True,
            )
            st.download_button(
                "⬇ MBNT Excel",
                data=_df_to_excel_bytes(mbnt_df, "MBNT"),
                file_name=f"MBNT_{start_date}_{end_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    with c_dl2:
        if not ttqt_df.empty:
            st.download_button(
                "⬇ TTQT CSV",
                data=_df_to_csv_bytes(ttqt_df),
                file_name=f"TTQT_{start_date}_{end_date}.csv",
                mime="text/csv",
                use_container_width=True,
            )
            st.download_button(
                "⬇ TTQT Excel",
                data=_df_to_excel_bytes(ttqt_df, "TTQT"),
                file_name=f"TTQT_{start_date}_{end_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

# ──────────────────────────────────────────────
# Header
# ──────────────────────────────────────────────
st.markdown('<div class="main-header">⚡ DASHBOARD LÃNH ĐẠO – MBNT & TTQT-TTTM | VIETCOMBANK</div>', unsafe_allow_html=True)
st.markdown(
    f'<div class="sub-header">Kỳ: {start_date.strftime("%d/%m/%Y")} → {end_date.strftime("%d/%m/%Y")} '
    f'| Phòng: <b>{selected_phong}</b> | {len(mbnt_df):,} GD MBNT | {len(ttqt_df):,} GD TTQT</div>',
    unsafe_allow_html=True,
)

# Hiển thị raw data khi tick checkbox sidebar
if show_mbnt and not mbnt_df.empty:
    with st.expander(f"📋 Raw MBNT đang load ({len(mbnt_df):,} dòng)", expanded=True):
        st.dataframe(mbnt_df, use_container_width=True, height=360)
if show_ttqt and not ttqt_df.empty:
    with st.expander(f"📋 Raw TTQT đang load ({len(ttqt_df):,} dòng)", expanded=True):
        st.dataframe(ttqt_df, use_container_width=True, height=360)

# ──────────────────────────────────────────────
# Nội dung theo điều hướng sidebar
# ──────────────────────────────────────────────

def show_table(df: pd.DataFrame, title: str, value_label: str):
    """Hiển thị bảng Pareto – dùng chung Dashboard & Báo cáo phòng"""
    st.markdown(f"**{title}**")
    if df is None or df.empty:
        st.info("Không có dữ liệu")
        return
    display = df.copy()

    if "DS (tr.USD)" in value_label or value_label.startswith("DS"):
        drop_cols = [c for c in ["LN_VND"] if c in display.columns]
        display = display.drop(columns=drop_cols, errors="ignore")
        value_src = "DS_TTQT" if "DS_TTQT" in display.columns else "DS_USD"
    else:
        drop_cols = [c for c in ["DS_USD", "DS_TTQT"] if c in display.columns]
        display = display.drop(columns=drop_cols, errors="ignore")
        value_src = "LN_VND"

    rename = {
        "Ma_CIF": "Mã CIF",
        "Ten_Khach_Hang": "Tên Khách Hàng",
        "Phong_Ban": "Phòng QL",
        value_src: value_label,
        "Pct": "Tỷ trọng %",
        "CumPct": "Cộng dồn %",
        "Rank": "Rank",
    }
    display = display.rename(columns={k: v for k, v in rename.items() if k in display.columns})
    display = display.loc[:, ~display.columns.duplicated()]
    cols_show = [c for c in ["Rank", "Mã CIF", "Tên Khách Hàng", "Phòng QL", value_label, "Tỷ trọng %", "Cộng dồn %"] if c in display.columns]
    st.dataframe(display[cols_show], use_container_width=True, hide_index=True, height=280)


_nav = st.session_state.get("nav_page", "🏠 Home / Dashboard")

def _back_home_btn():
    if st.button("← Về Home", key=f"back_{_nav}"):
        st.session_state.nav_page = "🏠 Home / Dashboard"
        st.rerun()

# ══════════════════════════════════════════════
# HOME / DASHBOARD
# ══════════════════════════════════════════════
if _nav == "🏠 Home / Dashboard":
    phong_filter = None if selected_phong == "Tất cả" else selected_phong
    kpis = compute_kpis(mbnt_df, ttqt_df, start_date, end_date, phong_filter)
    _mbnt_f = filter_by_date(mbnt_df, "Ngay", start_date, end_date)
    if phong_filter:
        _mbnt_f = filter_by_phong(_mbnt_f, "Phong_Ban", phong_filter)
    st.caption(
        f"📌 Sau lọc kỳ **{start_date.strftime('%d/%m/%Y')} → {end_date.strftime('%d/%m/%Y')}**"
        f" · Phòng: **{selected_phong}** · "
        f"GD MBNT: **{len(_mbnt_f):,}** / {len(mbnt_df):,} · "
        f"CIF: **{_mbnt_f['Ma_CIF'].nunique() if not _mbnt_f.empty else 0}**"
    )

    # KPI row
    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.metric("💰 DS MBNT (tr.USD)", f"{kpis['ds_mbnt']:,.2f}")
    with c2:
        st.metric("📈 LN MBNT (tỷ VND)", f"{kpis['ln_mbnt']:,.2f}")
    with c3:
        st.metric("🌐 DS TTQT (tr.USD)", f"{kpis['ds_ttqt']:,.2f}")
    with c4:
        st.metric("🔢 Số GD MBNT", f"{kpis['so_gd']:,}")
    with c5:
        st.metric("👥 Số CIF", f"{kpis['so_cif']:,}")

    c6, c7, c8, c9 = st.columns(4)
    with c6:
        st.metric("📊 DS Bán (tr.USD)", f"{kpis['ds_ban']:,.2f}")
    with c7:
        st.metric("📈 DS Mua (tr.USD)", f"{kpis['ds_mua']:,.2f}")
    with c8:
        st.metric("🔢 GD TTQT", f"{kpis['so_gd_ttqt']:,}")
    with c9:
        ratio = (kpis["ds_ban"] / kpis["ds_mua"] * 100) if kpis["ds_mua"] else 0
        st.metric("Bán/Mua %", f"{ratio:.1f}%")

    st.divider()

    # Pareto tables
    tables = prepare_pareto_tables(mbnt_df, ttqt_df, start_date, end_date, phong_filter)

    r1c1, r1c2 = st.columns(2)
    with r1c1:
        show_table(tables["top20_ds"], "🏆 20% DS MBNT CAO NHẤT", "DS (tr.USD)")
    with r1c2:
        show_table(tables["top20_ln"], "🏆 20% LN MBNT CAO NHẤT", "LN (tỷ VND)")

    r2c1, r2c2 = st.columns(2)
    with r2c1:
        show_table(tables["bot10_ds"], "⚠ 10% DS MBNT THẤP NHẤT", "DS (tr.USD)")
    with r2c2:
        show_table(tables["bot10_ln"], "⚠ 10% LN MBNT THẤP NHẤT", "LN (tỷ VND)")

    r3c1, r3c2 = st.columns(2)
    with r3c1:
        show_table(tables["top20_ttqt"], "🌐 20% DS TTQT CAO NHẤT", "DS TTQT (tr.USD)")
    with r3c2:
        show_table(tables["bot10_ttqt"], "⚠ 10% DS TTQT THẤP NHẤT", "DS TTQT (tr.USD)")

    st.markdown("---")
    st.markdown("#### 📊 BÁN-TRẢ NỢ / MUA-VAY")
    r4c1, r4c2, r4c3, r4c4 = st.columns(4)
    with r4c1:
        show_table(tables["top_ban_ds"], "Bán – Top DS", "DS (tr.USD)")
    with r4c2:
        show_table(tables["top_ban_ln"], "Bán – Top LN", "LN (tỷ VND)")
    with r4c3:
        show_table(tables["top_mua_ds"], "Mua – Top DS", "DS (tr.USD)")
    with r4c4:
        show_table(tables["top_mua_ln"], "Mua – Top LN", "LN (tỷ VND)")

    # Donut charts
    st.markdown("---")
    st.markdown("#### 🍩 Phân bổ theo Phòng")
    mbnt_f = filter_by_date(mbnt_df, "Ngay", start_date, end_date)
    if phong_filter:
        mbnt_f = filter_by_phong(mbnt_f, "Phong_Ban", phong_filter)

    if not mbnt_f.empty:
        by_phong = (
            mbnt_f.groupby("Phong_Ban")
            .agg(DS=("DS_Quy_USD", "sum"), LN=("Loi_Nhuan_VND", "sum"), GD=("Ma_CIF", "count"))
            .reset_index()
        )
        by_phong["DS"] = by_phong["DS"] / 1e6
        by_phong["LN"] = by_phong["LN"] / 1e9

        dc1, dc2, dc3 = st.columns(3)
        with dc1:
            fig = px.pie(by_phong, values="DS", names="Phong_Ban", title="DS MBNT theo Phòng", hole=0.45)
            fig.update_layout(margin=dict(t=40, b=0, l=0, r=0), height=320, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
        with dc2:
            fig = px.pie(by_phong, values="LN", names="Phong_Ban", title="LN MBNT theo Phòng", hole=0.45)
            fig.update_layout(margin=dict(t=40, b=0, l=0, r=0), height=320, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)
        with dc3:
            fig = px.pie(by_phong, values="GD", names="Phong_Ban", title="Số GD theo Phòng", hole=0.45)
            fig.update_layout(margin=dict(t=40, b=0, l=0, r=0), height=320, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

# ══════════════════════════════════════════════
# TAB 2: BÁO CÁO PHÒNG (luôn filter 1 phòng)
# ══════════════════════════════════════════════
elif _nav == "🏢 Báo cáo phòng":
    _back_home_btn()
    st.markdown("### 🏢 Báo cáo theo Phòng")
    phong_p = st.selectbox("Chọn phòng để xem chi tiết", PHONG_LIST, key="phong_detail")
    kpis_p = compute_kpis(mbnt_df, ttqt_df, start_date, end_date, phong_p)
    _pf = filter_by_date(mbnt_df, "Ngay", start_date, end_date)
    _pf = filter_by_phong(_pf, "Phong_Ban", phong_p)
    st.caption(
        f"📌 Kỳ **{start_date.strftime('%d/%m/%Y')} → {end_date.strftime('%d/%m/%Y')}**"
        f" · Phòng **{phong_p}** · GD sau lọc: **{len(_pf):,}**"
    )

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("DS Phòng (tr.USD)", f"{kpis_p['ds_mbnt']:,.2f}")
    m2.metric("LN Phòng (tỷ VND)", f"{kpis_p['ln_mbnt']:,.2f}")
    m3.metric("DS TTQT Phòng", f"{kpis_p['ds_ttqt']:,.2f}")
    m4.metric("Số CIF Phòng", f"{kpis_p['so_cif']:,}")

    tables_p = prepare_pareto_tables(mbnt_df, ttqt_df, start_date, end_date, phong_p)

    pc1, pc2 = st.columns(2)
    with pc1:
        show_table(tables_p["top20_ds"], f"🏆 Top DS – {phong_p}", "DS (tr.USD)")
    with pc2:
        show_table(tables_p["top20_ln"], f"🏆 Top LN – {phong_p}", "LN (tỷ VND)")

# ══════════════════════════════════════════════
# TAB 3: KẾ HOẠCH
# ══════════════════════════════════════════════
elif _nav == "📋 Kế hoạch":
    _back_home_btn()
    st.markdown("### 📋 Kế hoạch & % Hoàn thành")
    st.caption("Nhập KH (ô vàng tương đương). TH tự tính từ raw theo tháng. %HT = TH/KH")

    metric_choice = st.radio("Chỉ tiêu", ["DS MBNT (tr.USD)", "LN MBNT (tỷ VND)", "DS TTQT (tr.USD)"], horizontal=True)
    key_map = {
        "DS MBNT (tr.USD)": ("DS_MBNT", "DS_Quy_USD", 1e6),
        "LN MBNT (tỷ VND)": ("LN_MBNT", "Loi_Nhuan_VND", 1e9),
        "DS TTQT (tr.USD)": ("DS_TTQT", "DS_Quy_USD", 1e6),
    }
    kh_key, val_col, divisor = key_map[metric_choice]

    # Tính TH thực tế
    if "TTQT" in metric_choice:
        th_pivot = monthly_th_from_raw(ttqt_df.rename(columns={"Ngay_GD": "Ngay"}) if not ttqt_df.empty else pd.DataFrame(), 2026, "DS_Quy_USD")
    else:
        th_pivot = monthly_th_from_raw(mbnt_df, 2026, val_col)

    # Editor cho KH
    st.markdown("#### Nhập Kế hoạch (KH)")
    kh_data = []
    for p in PHONG_LIST:
        row = {"Phòng": p}
        for m in range(1, 13):
            row[f"T{m}"] = st.session_state.ke_hoach[kh_key][p].get(f"T{m}", 0.0)
        kh_data.append(row)
    kh_df = pd.DataFrame(kh_data)

    edited = st.data_editor(
        kh_df,
        use_container_width=True,
        num_rows="fixed",
        key=f"editor_{kh_key}",
        column_config={f"T{m}": st.column_config.NumberColumn(f"T{m}", format="%.2f") for m in range(1, 13)},
    )

    # Lưu lại session
    for _, r in edited.iterrows():
        p = r["Phòng"]
        for m in range(1, 13):
            st.session_state.ke_hoach[kh_key][p][f"T{m}"] = float(r[f"T{m}"] or 0)

    # Bảng %HT
    st.markdown("#### % Hoàn thành (TH / KH)")
    ht_rows = []
    for p in PHONG_LIST:
        row = {"Phòng": p}
        th_row = th_pivot[th_pivot.get("Phong_Ban", th_pivot.columns[0]) == p] if not th_pivot.empty else pd.DataFrame()
        for m in range(1, 13):
            kh_val = st.session_state.ke_hoach[kh_key][p].get(f"T{m}", 0)
            th_val = 0.0
            if not th_row.empty and f"T{m}" in th_row.columns:
                th_val = float(th_row[f"T{m}"].values[0]) / divisor
            pct = (th_val / kh_val * 100) if kh_val else 0
            row[f"T{m}"] = round(pct, 1)
        ht_rows.append(row)
    ht_df = pd.DataFrame(ht_rows)
    st.dataframe(ht_df, use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════
# TAB 4: BÁO CÁO LỜI
# ══════════════════════════════════════════════
elif _nav == "📝 Báo cáo lời":
    _back_home_btn()
    st.markdown("### 📝 Báo cáo lời – Phân tích & Tổng hợp")
    st.caption("Soạn nội dung + dán nguồn ngoại (NHNN, Bloomberg…). Download / In báo cáo bên dưới.")

    sections = [
        ("I. BỐI CẢNH THỊ TRƯỜNG", "Nhập nội dung từ NHNN, TTCK, bộ/ngành..."),
        ("II. DIỄN BIẾN TỶ GIÁ", "Phân tích USD/VND, EUR và các cặp chính trong kỳ"),
        ("III. KẾT QUẢ MBNT", "Tổng hợp DS, LN, số GD, so sánh KH vs TH"),
        ("IV. PHÂN TÍCH PARETO", "Top 20% CIF đóng góp, CIF nổi bật"),
        ("V. BÁN-TRẢ NỢ / MUA-VAY", "Phân tích chiều giao dịch & khách vay ngoại tệ"),
        ("VI. TTQT – TTTM", "Kết quả DS TTQT, tương quan MBNT"),
        ("VII. HUY ĐỘNG & CHO VAY", "Tương quan HĐV/CV với MBNT"),
        ("VIII. ĐÁNH GIÁ & ĐỀ XUẤT", "Nhận xét tổng thể, kiến nghị kỳ tới"),
        ("IX. TÀI LIỆU ĐÍNH KÈM", "Link / tóm tắt nguồn"),
    ]

    report_parts = []
    phong_filter_r = None if selected_phong == "Tất cả" else selected_phong
    kpis_r = compute_kpis(mbnt_df, ttqt_df, start_date, end_date, phong_filter_r)

    kpi_block = f"""Kỳ báo cáo: {start_date.strftime('%d/%m/%Y')} → {end_date.strftime('%d/%m/%Y')}
Phòng: {selected_phong}
DS MBNT: {kpis_r['ds_mbnt']:,.2f} tr.USD
LN MBNT: {kpis_r['ln_mbnt']:,.2f} tỷ VND
DS TTQT: {kpis_r['ds_ttqt']:,.2f} tr.USD
Số GD MBNT: {kpis_r['so_gd']:,} | Số CIF: {kpis_r['so_cif']:,}
DS Bán: {kpis_r['ds_ban']:,.2f} | DS Mua: {kpis_r['ds_mua']:,.2f} tr.USD
"""
    report_parts.append("=== TÓM TẮT KPI ===\n" + kpi_block)

    for title, placeholder in sections:
        with st.expander(title, expanded=(title.startswith("I.") or title.startswith("III."))):
            txt = st.text_area(label=title, placeholder=placeholder, height=120, key=f"bc_{title}", label_visibility="collapsed")
            if txt and txt.strip():
                report_parts.append(f"\n=== {title} ===\n{txt.strip()}")

    full_report = "\n".join(report_parts)
    full_report_html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>Báo cáo MBNT VCB</title>
<style>
body {{ font-family: 'Segoe UI', Arial, sans-serif; max-width: 900px; margin: 40px auto; padding: 20px; color: #222; }}
h1 {{ color: #003366; border-bottom: 2px solid #003366; padding-bottom: 8px; }}
h2 {{ color: #003366; margin-top: 28px; }}
.kpi {{ background: #f0f7ff; padding: 16px; border-radius: 8px; border-left: 4px solid #003366; white-space: pre-wrap; }}
pre {{ white-space: pre-wrap; }}
@media print {{ body {{ margin: 0; }} }}
</style></head><body>
<h1>⚡ BÁO CÁO LỜI – MBNT & TTQT | VIETCOMBANK</h1>
<div class="kpi">{kpi_block}</div>
"""
    for part in report_parts[1:]:
        if part.startswith("\n=== "):
            title_line = part.split("\n")[1].replace("=== ", "").replace(" ===", "")
            body = "\n".join(part.split("\n")[2:])
            full_report_html += f"<h2>{title_line}</h2><pre>{body}</pre>"
    full_report_html += """
<script>/* Nút in */</script>
<p style="margin-top:40px;color:#888;font-size:12px;">In: Ctrl+P hoặc nút Print bên dưới · MBNT Dashboard VCB</p>
</body></html>"""

    st.divider()
    st.markdown("#### 📤 Xuất / In báo cáo")
    rc1, rc2, rc3 = st.columns(3)
    with rc1:
        st.download_button(
            "⬇ Download TXT",
            data=full_report.encode("utf-8-sig"),
            file_name=f"BaoCao_MBNT_{start_date}_{end_date}.txt",
            mime="text/plain",
            use_container_width=True,
        )
    with rc2:
        st.download_button(
            "⬇ Download HTML",
            data=full_report_html.encode("utf-8"),
            file_name=f"BaoCao_MBNT_{start_date}_{end_date}.html",
            mime="text/html",
            use_container_width=True,
            help="Mở file HTML → Ctrl+P để in / Save as PDF",
        )
    with rc3:
        if st.button("🖨 Xem bản in (HTML)", use_container_width=True):
            st.session_state["show_print_preview"] = True

    if st.session_state.get("show_print_preview"):
        st.markdown("---")
        st.markdown("##### Bản xem trước (in bằng Ctrl+P)")
        st.components.v1.html(full_report_html, height=600, scrolling=True)

# ══════════════════════════════════════════════
# TAB 5: DATA – View & Download
# ══════════════════════════════════════════════
elif _nav == "📁 Data":
    _back_home_btn()
    st.markdown("### 📁 Data đang load – View & Download")
    st.caption("Xem toàn bộ raw data hiện tại (mẫu demo hoặc file đã upload) và tải về CSV/Excel.")

    d1, d2 = st.columns(2)
    with d1:
        st.metric("Số dòng MBNT", f"{len(mbnt_df):,}")
    with d2:
        st.metric("Số dòng TTQT", f"{len(ttqt_df):,}")

    st.markdown("#### MBNT")
    if mbnt_df.empty:
        st.info("Chưa có data MBNT")
    else:
        st.dataframe(mbnt_df, use_container_width=True, height=400)
        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                "⬇ Download MBNT CSV",
                data=_df_to_csv_bytes(mbnt_df),
                file_name=f"MBNT_{start_date}_{end_date}.csv",
                mime="text/csv",
                key="dl_mbnt_csv_tab",
            )
        with c2:
            st.download_button(
                "⬇ Download MBNT Excel",
                data=_df_to_excel_bytes(mbnt_df, "MBNT"),
                file_name=f"MBNT_{start_date}_{end_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_mbnt_xlsx_tab",
            )

    st.markdown("#### TTQT")
    if ttqt_df.empty:
        st.info("Chưa có data TTQT")
    else:
        st.dataframe(ttqt_df, use_container_width=True, height=400)
        c3, c4 = st.columns(2)
        with c3:
            st.download_button(
                "⬇ Download TTQT CSV",
                data=_df_to_csv_bytes(ttqt_df),
                file_name=f"TTQT_{start_date}_{end_date}.csv",
                mime="text/csv",
                key="dl_ttqt_csv_tab",
            )
        with c4:
            st.download_button(
                "⬇ Download TTQT Excel",
                data=_df_to_excel_bytes(ttqt_df, "TTQT"),
                file_name=f"TTQT_{start_date}_{end_date}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_ttqt_xlsx_tab",
            )

# ══════════════════════════════════════════════
# TAB 6: HƯỚNG DẪN
# ══════════════════════════════════════════════
elif _nav == "📖 Hướng dẫn":
    _back_home_btn()
    st.markdown("""
## 📖 Hướng dẫn sử dụng – MBNT Dashboard Web (VCB)

### Cấu trúc tương đương Excel
| Sheet Excel | Tab Web | Mục đích |
|-------------|---------|----------|
| 01_DASHBOARD | 📊 Dashboard | 10 KPI + 10 bảng Pareto/Bán-Mua + Donut |
| 02_BAO_CAO_PHONG | 🏢 Báo cáo phòng | Lọc 1 phòng + bảng chi tiết |
| 03_KE_HOACH | 📋 Kế hoạch | Nhập KH, TH tự tính, %HT |
| 05_BAO_CAO_LOI | 📝 Báo cáo lời | Soạn phân tích + đề xuất |
| 03_Data_Raw_MBNT | Upload | Dán / upload raw MBNT |
| Data_Raw_TTQT_TTTM | Upload | Dán / upload raw TTQT |

### Cách nạp dữ liệu
1. **Upload Excel template** (file gốc) → tự đọc 2 sheet raw.
2. Hoặc upload **CSV riêng** cho MBNT và TTQT.
3. Cột bắt buộc (tên có thể gần đúng):
   - MBNT: Ngày, Ma_CIF, Ten_Khach_Hang, Phong_Ban, Chieu_NH, DS_Quy_USD, Loi_Nhuan_VND
   - TTQT: Ngay_GD, Ma_CIF, Phong_Ban, DS_Quy_USD

### Lưu ý quan trọng (giữ nguyên tắc Excel)
- **Tên phòng** phải khớp 100% danh sách 15 phòng trong `Danh_Sach_Phong`.
- Ngày phải là định dạng DATE (không text).
- Chieu_NH dùng đúng `"BÁN"` hoặc `"MUA"`.
- Logic Pareto / Rank dùng `groupby` + `rank` (tương đương SUMIFS + RANK + IsFirst).

### Deploy lên Render
```bash
# 1. Push lên GitHub
git init
git add .
git commit -m "MBNT Dashboard VCB web"
git remote add origin https://github.com/<user>/mbnt-dashboard-vcb.git
git push -u origin main

# 2. Trên Render.com
# - New → Web Service
# - Connect repo
# - Runtime: Python
# - Build: pip install -r requirements.txt
# - Start: streamlit run app.py --server.port $PORT --server.address 0.0.0.0
```

### Local
```bash
pip install -r requirements.txt
streamlit run app.py
```
""")

# Footer
st.markdown("---")
st.caption("MBNT Dashboard VCB Web v1.0 | Tái hiện nguyên tắc Excel (SUMIFS + RANK + IsFirst) | Không circular reference")
