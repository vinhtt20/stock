# Phân loại Doanh nghiệp Phi tài chính HOSE theo ICB

Lập danh sách doanh nghiệp **phi tài chính** niêm yết trên **HOSE** và gán mỗi mã vào một
ngành **ICB cấp 1**. Sản phẩm là bước chuẩn bị mẫu cho nghiên cứu Quản trị lợi nhuận
(Earnings Management) — kế thừa Nguyen, Kim & Ali (2024). Chi tiết tiêu chí xem `DoD.txt`.

## Thành phần

| File | Vai trò |
|---|---|
| `PHAN_LOAI_HOSE_hoan_chinh.py` | Script chính: lấy dữ liệu → phân loại → xuất Excel. |
| `test_PHAN_LOAI_HOSE_hoan_chinh.py` | 39 unit test offline (DataFrame tổng hợp, không gọi mạng). |
| `HOSE_phi_tai_chinh.xlsx` | Output 1 — danh sách (sheet `Toan_bo`, `Phi_tai_chinh`, mỗi ngành 1 sheet). |
| `Co_so_phan_loai_nganh_HOSE.xlsx` | Output 2 — tài liệu phương pháp 6 sheet. |
| `BAO_CAO_PHAN_LOAI_HOSE.md` / `.pdf` | Báo cáo tổng hợp kết quả. |
| `DoD.txt` | Definition of Done — tiêu chí nghiệm thu. |

## Chạy

```bash
# Mặc định: KBS (7 nhóm ICB — số liệu chốt theo DoD)
uv run --with vnstock --with pandas --with openpyxl python PHAN_LOAI_HOSE_hoan_chinh.py

# Thử VCI trước, lỗi thì fallback KBS
python PHAN_LOAI_HOSE_hoan_chinh.py --source auto

# Ép nguồn VCI (ICB thật, 9 nhóm — tách Dầu khí + Viễn thông)
python PHAN_LOAI_HOSE_hoan_chinh.py --source vci

# Test offline (không cần vnstock)
uv run --with pandas --with openpyxl --with pytest pytest test_PHAN_LOAI_HOSE_hoan_chinh.py -q
```

**Tham số CLI:** `--output`, `--reference-output`, `--source {auto,vci,kbs}`,
`--kbs-industry-col`, `--include-real-estate`.

## Luồng code

```
main(argv)
  │
  ├─ argparse → đọc --source, --output, ...
  │
  ├─ resolve_source(prefer)            ── chọn & lấy dữ liệu (lazy import vnstock)
  │     ├─ "vci"  → fetch_hose_symbols + fetch_icb        (raise nếu lỗi)
  │     ├─ "kbs"  → fetch_hose_symbols + _fetch_kbs
  │     └─ "auto" → thử VCI, lỗi thì fallback KBS
  │
  ├─ build_dataset / build_dataset_kbs ── dựng bảng đã phân loại
  │     ├─ VCI: dùng mã ICB thật → ngành ICB cấp 1
  │     └─ KBS: merge mã ngành (industry_code) → ánh xạ KBS_TO_ICB
  │            → áp GAN_TAY (ADG/YEG/CLC) → mã chưa rõ = MISSING_ICB
  │     │
  │     └─ classify(row)               ── gán nhãn từng mã:
  │            Ngân hàng / Chứng khoán / Bảo hiểm / TC khác
  │            / Bất động sản (regex "bat dong san")
  │            / Phi tài chính
  │            (norm() chuẩn hoá tiếng Việt, có fix đ→d)
  │
  ├─ make_reports(df)                  ── đếm phân bố, cảnh báo "CHƯA RÕ"
  │
  ├─ export_excel(df, output)          ── Output 1: nhiều sheet
  └─ export_reference_workbook(...)    ── Output 2: 6 sheet phương pháp (nhúng NGAY_CHOT)
```

### Các khối chính trong script

1. **Cấu hình** — hằng số: `EXCLUDE_REAL_ESTATE`, `NGAY_CHOT`, `KBS_INDUSTRY_CODE_COL`,
   bảng ánh xạ `KBS_TO_ICB` (25 ngành KBS → ICB cấp 1), `GAN_TAY` (gán tay 3 mã thiếu ngành),
   các nhãn `LBL_*` và regex `REAL_ESTATE_RE`, `SYMBOL_RE`.
2. **Hàm thuần** — `norm()` (chuẩn hoá dấu tiếng Việt + fix `đ`→`d`), `classify()` phân loại
   một dòng, `all_names()`, `_unique_sheet_name()`.
3. **I/O vnstock** — `fetch_hose_symbols()`, `fetch_icb()`, `_fetch_kbs()`, `resolve_source()`,
   `_assert_schema()` kiểm tra cột bắt buộc.
4. **Dựng dữ liệu** — `build_dataset()` (VCI), `build_dataset_kbs()` (KBS),
   `_non_financial_sample()` lọc mẫu phi tài chính.
5. **Xuất** — `make_reports()`, `export_excel()`, `export_reference_workbook()`.
6. **Orchestrator** — `main()` ghép toàn bộ luồng.

## Lưu ý nguồn dữ liệu

- **KBS** (ổn định): taxonomy riêng 25 ngành (mã 1–29), **không phải** ICB → script quy đổi
  sang ICB cấp 1 ⇒ **7 nhóm** (thiếu Dầu khí, Viễn thông). Cột mã ngành thật là `industry_code`.
- **VCI**: trả mã ICB thật 4 cấp ⇒ **9 nhóm**, nhưng có thể bị `403 Forbidden` theo IP.

## Kết quả (nguồn KBS, chốt 03/06/2026)

387 mã HOSE = **283 phi tài chính** + 57 bất động sản + 47 tài chính (21 NH / 20 CK / 5 BH /
1 TC khác). 283 mã phi tài chính chia 7 nhóm ICB cấp 1: Công nghiệp 124, Hàng tiêu dùng 56,
Vật liệu cơ bản 33, Tiện ích cộng đồng 31, Dịch vụ tiêu dùng 19, Y tế 12, Công nghệ 8.
