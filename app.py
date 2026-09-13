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
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/1b/Vietcombank_logo.svg/200px-Vietcombank_logo.svg.png", width=160)
    st.markdown("### 📁 Nạp dữ liệu")
    st.caption("Dán raw giống Excel: sheet 03_Data_Raw_MBNT & Data_Raw_TTQT_TTTM")

    uploaded = st.file_uploader(
        "Upload Excel template (.xlsx)",
        type=["xlsx", "xls"],
        help="File MBNT Dashboard VCB – sẽ đọc 2 sheet raw",
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
    st.caption("Nguyên tắc: Tên phòng phải khớp 100% cột Phong_Ban trong raw")
    st.caption("Excel 2016/365 compatible logic → pandas rank + groupby")

# ──────────────────────────────────────────────
# Load data
# ──────────────────────────────────────────────
@st.cache_data
def get_sample_data():
    """Tạo sample nhỏ để demo khi chưa upload"""
    import numpy as np
    rng = np.random.default_rng(42)
    n = 200
    dates = pd.date_range("2026-08-01", "2026-09-12", periods=n)
    cifs = [f"CIF{1000+i%40}" for i in range(n)]
    names = [f"Khách hàng {i%40+1}" for i in range(n)]
    phongs = rng.choice(PHONG_LIST, n)
    chieu = rng.choice(["BÁN", "MUA"], n)
    ds = rng.uniform(10_000, 2_000_000, n)
    ln = ds * rng.uniform(50, 300)  # margin approx
    mbnt = pd.DataFrame({
        "Ngay": dates,
        "Ma_CIF": cifs,
        "Ten_Khach_Hang": names,
        "Phong_Ban": phongs,
        "Chieu_NH": chieu,
        "DS_Quy_USD": ds,
        "Loi_Nhuan_VND": ln,
        "Nguon_Mua_KH": "Nội địa",
        "Muc_Dich_KH": rng.choice(["Trả nợ vay", "Thanh toán nhập khẩu", "Vay ngoại tệ"], n),
    })
    ttqt = mbnt.sample(80).copy()
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

# ──────────────────────────────────────────────
# Header
# ──────────────────────────────────────────────
st.markdown('<div class="main-header">⚡ DASHBOARD LÃNH ĐẠO – MBNT & TTQT-TTTM | VIETCOMBANK</div>', unsafe_allow_html=True)
st.markdown(
    f'<div class="sub-header">Kỳ: {start_date.strftime("%d/%m/%Y")} → {end_date.strftime("%d/%m/%Y")} '
    f'| Phòng: <b>{selected_phong}</b> | {len(mbnt_df):,} GD MBNT | {len(ttqt_df):,} GD TTQT</div>',
    unsafe_allow_html=True,
)

# ──────────────────────────────────────────────
# Tabs = Sheets
# ──────────────────────────────────────────────
tab_dash, tab_phong, tab_kh, tab_baocao, tab_help = st.tabs(
    ["📊 01_DASHBOARD", "🏢 02_BÁO CÁO PHÒNG", "📋 03_KẾ HOẠCH", "📝 05_BÁO CÁO LỜI", "📖 Hướng dẫn"]
)

# ══════════════════════════════════════════════
# TAB 1: DASHBOARD (chi nhánh hoặc đã lọc phòng)
# ══════════════════════════════════════════════
with tab_dash:
    phong_filter = None if selected_phong == "Tất cả" else selected_phong
    kpis = compute_kpis(mbnt_df, ttqt_df, start_date, end_date, phong_filter)

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

    def show_table(df: pd.DataFrame, title: str, value_label: str):
        st.markdown(f"**{title}**")
        if df is None or df.empty:
            st.info("Không có dữ liệu")
            return
        display = df.copy()

        # Chỉ giữ 1 cột giá trị phù hợp với value_label, tránh trùng tên
        if "DS (tr.USD)" in value_label or value_label.startswith("DS"):
            # Bảng DS → giữ DS_USD / DS_TTQT, bỏ LN_VND
            drop_cols = [c for c in ["LN_VND"] if c in display.columns]
            display = display.drop(columns=drop_cols, errors="ignore")
            value_src = "DS_TTQT" if "DS_TTQT" in display.columns else "DS_USD"
        else:
            # Bảng LN → giữ LN_VND, bỏ DS
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
        # Loại bỏ cột trùng nếu còn
        display = display.loc[:, ~display.columns.duplicated()]
        cols_show = [c for c in ["Rank", "Mã CIF", "Tên Khách Hàng", "Phòng QL", value_label, "Tỷ trọng %", "Cộng dồn %"] if c in display.columns]
        st.dataframe(display[cols_show], use_container_width=True, hide_index=True, height=280)

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
with tab_phong:
    st.markdown("### 🏢 Báo cáo theo Phòng")
    phong_p = st.selectbox("Chọn phòng để xem chi tiết", PHONG_LIST, key="phong_detail")
    kpis_p = compute_kpis(mbnt_df, ttqt_df, start_date, end_date, phong_p)

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
with tab_kh:
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
with tab_baocao:
    st.markdown("### 📝 Báo cáo lời – Phân tích & Tổng hợp")
    st.caption("Soạn nội dung + dán nguồn ngoại (NHNN, Bloomberg…). Có thể export PDF từ trình duyệt.")

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

    for title, placeholder in sections:
        with st.expander(title, expanded=(title.startswith("I.") or title.startswith("III."))):
            st.text_area(label=title, placeholder=placeholder, height=120, key=f"bc_{title}", label_visibility="collapsed")

    if st.button("📄 Tóm tắt KPI vào báo cáo (copy)"):
        summary = f"""
Kỳ báo cáo: {start_date} → {end_date}
Phòng: {selected_phong}
DS MBNT: {kpis['ds_mbnt']:,.2f} tr.USD
LN MBNT: {kpis['ln_mbnt']:,.2f} tỷ VND
DS TTQT: {kpis['ds_ttqt']:,.2f} tr.USD
Số GD: {kpis['so_gd']:,} | Số CIF: {kpis['so_cif']:,}
"""
        st.code(summary)

# ══════════════════════════════════════════════
# TAB 5: HƯỚNG DẪN
# ══════════════════════════════════════════════
with tab_help:
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
