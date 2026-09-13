# ⚡ MBNT Dashboard VCB – Web Version

Tái hiện đầy đủ **cấu trúc & nguyên tắc** file Excel  
`20260913_MBNT_Dashboard_VCB_v7_FINAL_Claude_Blank.xlsx`  
thành ứng dụng web sẵn sàng deploy **GitHub + Render**.

## 📁 Cấu trúc thư mục (tương đương sheet Excel)

```
mbnt-dashboard-vcb/
├── app.py                  # Entry point Streamlit (Dashboard + tabs)
├── requirements.txt
├── README.md
├── .gitignore
├── render.yaml             # Blueprint deploy Render (optional)
├── utils/
│   ├── calculations.py     # Engine: KPI, Pareto 20%/10%, Rank, IsFirst
│   └── data_loader.py      # Chuẩn hóa raw MBNT / TTQT
├── pages/                  # (dự phòng multi-page sau này)
└── data/                   # Nơi chứa sample / cache (không commit data thật)
```

| Sheet Excel              | Tab / Module Web          | Chức năng |
|--------------------------|---------------------------|-----------|
| README                    | Tab Hướng dẫn             | Hướng dẫn đầy đủ |
| 01_DASHBOARD             | 📊 Dashboard              | 10 KPI + 6 Pareto + 4 Bán/Mua + 3 Donut |
| 02_BAO_CAO_PHONG         | 🏢 Báo cáo phòng          | Dropdown phòng + bảng chi tiết |
| 03_KE_HOACH              | 📋 Kế hoạch               | Nhập KH (vàng) → TH tự lấy raw → %HT |
| 03_Data_Raw_MBNT         | Upload Excel/CSV          | Dán từ hàng 2 |
| Data_Raw_TTQT_TTTM       | Upload Excel/CSV          | Dán từ hàng 3 |
| Danh_Sach_Phong          | PHONG_LIST (hard-code)    | 15 phòng – phải khớp 100% |
| CALC_*                   | `utils/calculations.py`   | Không circular, chỉ groupby + rank |
| 05_BAO_CAO_LOI           | 📝 Báo cáo lời            | 9 mục phân tích + đề xuất |

## 🔧 Nguyên tắc giữ nguyên từ Excel

- **Không circular reference** → dùng `groupby` + `rank` thay vì MATCH expanding.
- Chỉ dùng logic tương đương **SUMIFS + INDEX/MATCH + RANK** (không cần XLOOKUP/FILTER/UNIQUE).
- Tên phòng **phải khớp 100%** cột `Phong_Ban` / `Phong`.
- `Chieu_NH` = `"BÁN"` hoặc `"MUA"`.
- Ngày phải là DATE.
- Pareto Top 20% / Bottom 10% theo cộng dồn %.

## 🚀 Chạy local

```bash
cd mbnt-dashboard-vcb
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Mở http://localhost:8501

## ☁️ Deploy lên Render (khuyến nghị)

### Cách 1 – Blueprint (render.yaml)

1. Push repo lên GitHub.
2. Vào [Render Dashboard](https://dashboard.render.com) → **New** → **Blueprint**.
3. Chọn repo → Render đọc `render.yaml` và tạo Web Service tự động.

### Cách 2 – Manual Web Service

1. **New → Web Service** → Connect GitHub repo.
2. Settings:
   - **Runtime**: Python 3
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**:  
     ```
     streamlit run app.py --server.port $PORT --server.address 0.0.0.0 --server.headless true
     ```
   - **Instance**: Free tier đủ dùng demo / nội bộ nhỏ.

3. Sau khi deploy xong → URL dạng `https://mbnt-dashboard-vcb.onrender.com`

> **Lưu ý Free tier**: service ngủ sau 15 phút không dùng. Lần truy cập đầu sẽ cold-start ~30-60s.

## 📤 Upload dữ liệu

1. Upload file Excel gốc (template) → app tự đọc sheet `03_Data_Raw_MBNT` và `Data_Raw_TTQT_TTTM`.
2. Hoặc upload 2 file CSV riêng.
3. Cột được map linh hoạt (tên gần đúng vẫn nhận).

## 📋 Danh sách 15 phòng chuẩn

```
Khách hàng doanh nghiệp 1
Khách hàng doanh nghiệp 2
Khách hàng doanh nghiệp 3
Khách hàng FDI
Khách hàng bán lẻ 1
Khách hàng bán lẻ 2
Khách hàng bán lẻ 3
Dịch vụ khách hàng tổ chức 1
Dịch vụ khách hàng tổ chức 2
PGD Tôn Đức Thắng
PGD Mê Linh
PGD Đồng Khởi
PGD Cộng Hòa
PGD Võ Văn Kiệt
PGD Lê Thánh Tôn
```

## 🔒 Bảo mật

- Không hard-code credential.
- Data upload chỉ nằm trong session Streamlit (không lưu server trừ khi bạn tự thêm DB).
- Nên bật **authentication** (Streamlit secrets / OAuth) nếu dùng production.

## 📜 License

Nội bộ Vietcombank – sử dụng theo quy định đơn vị.
"""
