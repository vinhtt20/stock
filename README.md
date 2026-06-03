# 📊 Phân loại Doanh nghiệp Phi tài chính HOSE theo ICB

> Lập danh sách doanh nghiệp **phi tài chính** niêm yết trên **Sở Giao dịch Chứng khoán TP.HCM (HOSE)** và gán mỗi mã vào một ngành theo chuẩn **ICB cấp 1**.

![Python](https://img.shields.io/badge/Python-3.10+-blue?logo=python&logoColor=white)
![Data](https://img.shields.io/badge/Data-vnstock-orange)
![Tests](https://img.shields.io/badge/tests-39%20passed-brightgreen)
![Source](https://img.shields.io/badge/source-KBS%20(7%20nh%C3%B3m%20ICB)-informational)
![Status](https://img.shields.io/badge/DoD-ho%C3%A0n%20th%C3%A0nh-success)

Sản phẩm là bước **chuẩn bị mẫu (sample construction)** cho nghiên cứu Quản trị lợi nhuận
(*Earnings Management*) bằng mô hình Jones điều chỉnh — kế thừa **Nguyen, Kim & Ali (2024)**.
Tiêu chí nghiệm thu đầy đủ xem [`DoD.txt`](DoD.txt).

---

## 📌 Mục lục

- [Kết quả](#-kết-quả)
- [Phương pháp](#-phương-pháp)
- [Nguồn dữ liệu](#-nguồn-dữ-liệu)
- [Cài đặt & Chạy](#️-cài-đặt--chạy)
- [Cấu trúc dự án](#-cấu-trúc-dự-án)
- [Luồng code](#-luồng-code)
- [Quyết định phương pháp (DoD §7)](#-quyết-định-phương-pháp-dod-7)

---

## 🎯 Kết quả

> **Ngày chốt dữ liệu:** 03/06/2026 · **Nguồn:** KBS (quy đổi ICB cấp 1) · **Sàn:** HOSE

### Tổng quan mẫu

| Nhóm | Số DN |
|---|---:|
| 🏭 **Phi tài chính** (mẫu nghiên cứu) | **283** |
| 🏢 Bất động sản *(loại)* | 57 |
| 🏦 Tài chính *(loại)* | 47 |
| **Tổng mã HOSE** | **387** |

Khối tài chính bị loại gồm: **21** Ngân hàng · **20** Chứng khoán · **5** Bảo hiểm · **1** Tài chính khác.

### Phân bố 283 DN phi tài chính theo ICB cấp 1

| # | Ngành ICB cấp 1 | Số DN | Tỷ trọng |
|---:|---|---:|---:|
| 1 | Công nghiệp | 124 | 43.8% |
| 2 | Hàng tiêu dùng | 56 | 19.8% |
| 3 | Vật liệu cơ bản | 33 | 11.7% |
| 4 | Tiện ích cộng đồng | 31 | 11.0% |
| 5 | Dịch vụ tiêu dùng | 19 | 6.7% |
| 6 | Y tế | 12 | 4.2% |
| 7 | Công nghệ | 8 | 2.8% |
| | **Tổng** | **283** | **100%** |

```
Công nghiệp          ████████████████████████████████████████████  124
Hàng tiêu dùng       ████████████████████  56
Vật liệu cơ bản      ████████████  33
Tiện ích cộng đồng   ███████████  31
Dịch vụ tiêu dùng    ███████  19
Y tế                 ████  12
Công nghệ            ███  8
```

📂 Chi tiết danh sách từng doanh nghiệp: [`BAO_CAO_PHAN_LOAI_HOSE.md`](BAO_CAO_PHAN_LOAI_HOSE.md) / [`.pdf`](BAO_CAO_PHAN_LOAI_HOSE.pdf) · File Excel: [`HOSE_phi_tai_chinh.xlsx`](HOSE_phi_tai_chinh.xlsx)

---

## 🧪 Phương pháp

| Quyết định | Nội dung | Lý do |
|---|---|---|
| **Sàn** | Chỉ HOSE | Phạm vi nghiên cứu |
| **Loại tài chính** | NH, CK, BH, TC khác | Cấu trúc vốn & chế độ kế toán khác biệt căn bản |
| **Loại bất động sản** | Toàn bộ BĐS | BCTC đặc thù làm méo tham số mô hình Jones |
| **Chuẩn phân ngành** | ICB cấp 1 | Tương thích văn liệu quốc tế, đủ thô để mỗi ngành nhiều quan sát |

---

## 🔌 Nguồn dữ liệu

Dữ liệu lấy qua thư viện **[`vnstock`](https://github.com/thinh-vu/vnstock)** (class `Listing`), không gọi API HTTP trực tiếp:

- `listing.symbols_by_exchange()` → danh sách mã + sàn (lọc `HSX/HOSE`)
- `listing.symbols_by_industries()` → bảng phân ngành

| `--source` | Nhà cung cấp | Trả về | Số nhóm |
|---|---|---|---|
| **`kbs`** *(mặc định)* | KB Securities | Taxonomy riêng 25 ngành → **quy đổi** sang ICB cấp 1 | **7 nhóm** (thiếu Dầu khí, Viễn thông) |
| `vci` | Vietcap | Mã **ICB thật** 4 cấp | 9 nhóm — đôi khi `403 Forbidden` theo IP |
| `auto` | VCI trước → fallback KBS | — | — |

> ℹ️ KBS dùng mã ngành ở cột `industry_code` (1–29), **không phải ICB** — script quy đổi qua bảng `KBS_TO_ICB`.

---

## ⚙️ Cài đặt & Chạy

```bash
# Chạy mặc định (KBS, 7 nhóm ICB — số liệu chốt theo DoD)
uv run --with vnstock --with pandas --with openpyxl python PHAN_LOAI_HOSE_hoan_chinh.py

# Ép nguồn VCI (ICB thật, 9 nhóm — tách Dầu khí + Viễn thông)
python PHAN_LOAI_HOSE_hoan_chinh.py --source vci

# Chạy test offline (KHÔNG cần cài vnstock — nhờ lazy import)
uv run --with pandas --with openpyxl --with pytest pytest test_PHAN_LOAI_HOSE_hoan_chinh.py -q
```

**Tham số CLI:** `--output` · `--reference-output` · `--source {kbs,vci,auto}` · `--kbs-industry-col` · `--include-real-estate`

---

## 📁 Cấu trúc dự án

```
.
├── PHAN_LOAI_HOSE_hoan_chinh.py        # Script chính: fetch → phân loại → xuất Excel
├── test_PHAN_LOAI_HOSE_hoan_chinh.py   # 39 unit test offline (DataFrame tổng hợp)
├── HOSE_phi_tai_chinh.xlsx             # Output 1 — danh sách (Toan_bo, Phi_tai_chinh, mỗi ngành 1 sheet)
├── Co_so_phan_loai_nganh_HOSE.xlsx     # Output 2 — tài liệu phương pháp 6 sheet
├── BAO_CAO_PHAN_LOAI_HOSE.md / .pdf    # Báo cáo kết quả
├── DoD.txt                             # Definition of Done — tiêu chí nghiệm thu
└── README.md
```

---

## 🔄 Luồng code

```
main(argv)
  │
  ├─ argparse ─────────────────── đọc --source, --output, ...
  │
  ├─ resolve_source(prefer) ───── chọn & lấy dữ liệu (lazy import vnstock)
  │     ├─ "kbs"  → fetch_hose_symbols + _fetch_kbs
  │     ├─ "vci"  → fetch_hose_symbols + fetch_icb   (raise nếu lỗi)
  │     └─ "auto" → thử VCI, lỗi thì fallback KBS
  │
  ├─ build_dataset(_kbs) ──────── dựng bảng đã phân loại
  │     ├─ VCI: mã ICB thật → ngành ICB cấp 1
  │     └─ KBS: merge industry_code → ánh xạ KBS_TO_ICB
  │            → áp GAN_TAY (ADG/YEG/CLC) → mã chưa rõ = MISSING_ICB
  │     └─ classify(row): NH / CK / BH / TC khác / Bất động sản / Phi tài chính
  │            (norm() chuẩn hoá tiếng Việt, có fix đ→d)
  │
  ├─ make_reports(df) ─────────── đếm phân bố, cảnh báo "CHƯA RÕ"
  │
  ├─ export_excel(df) ─────────── Output 1: nhiều sheet
  └─ export_reference_workbook ── Output 2: 6 sheet phương pháp (nhúng NGAY_CHOT)
```

---

## ✅ Quyết định phương pháp (DoD §7)

Cả 4 mục mở của DoD đã được chốt:

- [x] **Ngày chốt dữ liệu** — `NGAY_CHOT = 03/06/2026`
- [x] **Số nhóm** — **7 nhóm KBS** (mặc định `--source=kbs`)
- [x] **Chuẩn phân ngành** — **ICB** (không dùng VSIC)
- [x] **3 mã gán tay** — ADG/YEG → *Dịch vụ tiêu dùng*, CLC → *Hàng tiêu dùng*

---

<sub>Ngoài phạm vi task này: dựng panel firm-year, áp ngưỡng ≥10 quan sát/ngành, ước lượng mô hình Jones điều chỉnh để tính dồn tích bất thường |DAC|.</sub>
