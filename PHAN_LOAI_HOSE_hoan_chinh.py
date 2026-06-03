# %% PHÂN NGÀNH HOSE THEO ICB THẬT (nguồn VCI) — Phương án A
# Phân loại dựa trên TÊN ngành ICB + com_type_code (ổn định, không phụ thuộc
# định dạng mã). Kết quả khớp các ngành phi tài chính trong file Excel.
#
# Cấu trúc: hàm thuần (norm/all_names/classify) tách khỏi I/O (fetch_*) và
# khỏi transform/report/export. Không có biến/side-effect ở cấp module — mọi
# thứ chạy qua main(). Nhờ vậy test offline được mà không cần vnstock/mạng.
import argparse
import logging
import re
import sys
import unicodedata

import pandas as pd

logger = logging.getLogger(__name__)

# Hằng số phân loại (không còn là global mutable; chỉ dùng đọc).
FIN_KW = ["ngan hang", "bao hiem", "chung khoan", "dich vu tai chinh", "tai chinh"]
FIN_COM_TYPES = {"NH", "CK", "BH"}

# Nhãn phân loại (giữ tiếng Việt). Bất động sản là nhóm top-level riêng,
# KHÔNG gắn tiền tố "Tài chính".
LBL_NH = "Tài chính - Ngân hàng"
LBL_CK = "Tài chính - Chứng khoán"
LBL_BH = "Tài chính - Bảo hiểm"
LBL_FIN_OTHER = "Tài chính - Khác"
LBL_REAL_ESTATE = "Bất động sản"
LBL_NON_FIN = "Phi tài chính"
LBL_MISSING_ICB = "Chưa phân loại (thiếu ICB)"

# Mã ICB cấp 1..4 -> cột icb_name_lvN. Regex lọc mã 3 chữ cái — neo tường minh
# (^...$) để không phụ thuộc ngầm vào việc pandas fullmatch tự neo.
SYMBOL_RE = re.compile(r"^[A-Za-z]{3}$")
# Regex PHÁT HIỆN bất động sản (không phải loại trừ).
REAL_ESTATE_RE = re.compile(r"bat dong san")

# Sheet dành riêng cho de-dup tên sheet Excel.
RESERVED_SHEETS = ("Toan_bo", "Phi_tai_chinh")

# ---------------------------------------------------------------------------
# THAM SỐ QUAN TRỌNG (operator chỉnh ở đây trước khi chạy)
# ---------------------------------------------------------------------------
EXCLUDE_REAL_ESTATE = True  # True = loại bất động sản khỏi mẫu phi tài chính.
NGAY_CHOT = "03/06/2026"  # Ngày chốt dữ liệu.

# Tên cột mã ngành KBS trong bảng symbols_by_industries(source="kbs").
# CHƯA xác nhận được khi chưa chạy live -> để giá trị đoán hợp lý, có thể
# override qua tham số industry_code_col của build_dataset_kbs.
KBS_INDUSTRY_CODE_COL = "industry_code"  # tên cột thật của KBS (đã xác nhận live)

# ---------------------------------------------------------------------------
# Bảng ánh xạ 25 ngành KBS -> ICB cấp 1 (Phụ lục A của DoD).
# value = (tên ngành KBS, bucket, ICB cấp 1, approx)
#   bucket ∈ {"NH","CK","BH","FIN_OTHER","RE","NONFIN"}
#   icb_cap1 = None với nhóm bị loại (tài chính + BĐS)
#   approx = True chỉ với mã gần đúng (*): 1, 17, 26, 28
# Chép NGUYÊN VĂN từ DoD — KHÔNG suy diễn thêm.
# ---------------------------------------------------------------------------
KBS_TO_ICB = {
    # --- Nhóm bị LOẠI (tài chính + bất động sản) ---
    2: ("Bảo hiểm", "BH", None, False),
    3: ("Bất động sản", "RE", None, False),
    5: ("Chứng khoán", "CK", None, False),
    11: ("Ngân hàng", "NH", None, False),
    29: ("Tài chính khác", "FIN_OTHER", None, False),
    # --- Nhóm PHI TÀI CHÍNH (quy đổi sang ICB cấp 1) ---
    1: ("Bán buôn", "NONFIN", "Công nghiệp", True),
    6: ("Công nghệ và thông tin", "NONFIN", "Công nghệ", False),
    7: ("Bán lẻ", "NONFIN", "Dịch vụ tiêu dùng", False),
    8: ("Chăm sóc sức khỏe", "NONFIN", "Y tế", False),
    10: ("Khai khoáng", "NONFIN", "Vật liệu cơ bản", False),
    12: ("Nông - Lâm - Ngư", "NONFIN", "Hàng tiêu dùng", False),
    15: ("SX Thiết bị, máy móc", "NONFIN", "Công nghiệp", False),
    16: ("SX Hàng gia dụng", "NONFIN", "Hàng tiêu dùng", False),
    17: ("Sản phẩm cao su", "NONFIN", "Hàng tiêu dùng", True),
    18: ("SX Nhựa - Hóa chất", "NONFIN", "Vật liệu cơ bản", False),
    19: ("Thực phẩm - Đồ uống", "NONFIN", "Hàng tiêu dùng", False),
    20: ("Chế biến Thủy sản", "NONFIN", "Hàng tiêu dùng", False),
    21: ("Vật liệu xây dựng", "NONFIN", "Công nghiệp", False),
    22: ("Tiện ích", "NONFIN", "Tiện ích cộng đồng", False),
    23: ("Vận tải - kho bãi", "NONFIN", "Công nghiệp", False),
    24: ("Xây dựng", "NONFIN", "Công nghiệp", False),
    25: ("Dịch vụ lưu trú, ăn uống, giải trí", "NONFIN", "Dịch vụ tiêu dùng", False),
    26: ("SX Phụ trợ", "NONFIN", "Công nghiệp", True),
    27: ("Thiết bị điện", "NONFIN", "Công nghiệp", False),
    28: ("Dịch vụ tư vấn, hỗ trợ", "NONFIN", "Công nghiệp", True),
}

# Bản đồ bucket -> nhãn phan_loai (tái dùng nhãn hiện có).
BUCKET_TO_LABEL = {
    "NH": LBL_NH,
    "CK": LBL_CK,
    "BH": LBL_BH,
    "FIN_OTHER": LBL_FIN_OTHER,
    "RE": LBL_REAL_ESTATE,
    "NONFIN": LBL_NON_FIN,
}

# ---------------------------------------------------------------------------
# GÁN TAY: 3 mã thiếu ngành (DoD §5.5). Operator PHẢI xác nhận ICB cấp 1.
# Đây chỉ là best-effort; chỉnh trực tiếp dict này khi cần.
# ---------------------------------------------------------------------------
GAN_TAY = {
    "ADG": "Dịch vụ tiêu dùng",  # đã xác nhận 03/06/2026
    "YEG": "Dịch vụ tiêu dùng",  # đã xác nhận 03/06/2026
    "CLC": "Hàng tiêu dùng",     # đã xác nhận 03/06/2026
}


# ---------------------------------------------------------------------------
# 1) HÀM THUẦN (không I/O, dễ test)
# ---------------------------------------------------------------------------
def norm(s):
    """Chuẩn hoá: bỏ dấu, lower, strip. NaN/None -> chuỗi rỗng."""
    if s is None or (isinstance(s, float) and pd.isna(s)):
        s = ""
    s = str(s).lower().strip()
    if s == "nan":
        s = ""
    # "đ" là ký tự riêng (không phải base + dấu kết hợp) nên NFD không tách
    # được -> map thủ công, nếu không "bất động sản" sẽ thành "bat dong san".
    s = s.replace("đ", "d")
    return "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )


def all_names(row, name_cols):
    """Ghép tất cả tên ICB (đã norm) thành một chuỗi. Bỏ qua cột không có."""
    return " | ".join(norm(row.get(c, "")) for c in name_cols)


def classify(row, name_cols, real_estate_re=REAL_ESTATE_RE):
    """Phân loại 1 dòng theo TÊN ICB + com_type_code.

    Thứ tự ưu tiên (giữ nguyên hành vi gốc):
      1. Bất động sản thắng tài chính (dòng vừa khớp BĐS vừa khớp KW tài chính
         -> Bất động sản).
      2. Tài chính: Ngân hàng / Chứng khoán / Bảo hiểm / Khác.
      3. Nếu TẤT CẢ cột tên ICB trống/NaN VÀ com_type_code không cho tín hiệu
         tài chính -> "Chưa phân loại (thiếu ICB)".
      4. Còn lại -> "Phi tài chính".
    """
    names = all_names(row, name_cols)
    ctc = norm(row.get("com_type_code", "")).upper()

    is_re = bool(real_estate_re.search(names))
    is_fin = (ctc in FIN_COM_TYPES) or any(k in names for k in FIN_KW)

    if is_re:
        return LBL_REAL_ESTATE
    if is_fin:
        if "ngan hang" in names or ctc == "NH":
            return LBL_NH
        if "chung khoan" in names or ctc == "CK":
            return LBL_CK
        if "bao hiem" in names or ctc == "BH":
            return LBL_BH
        return LBL_FIN_OTHER

    # Thiếu ICB: mọi tên ICB rỗng và không có tín hiệu tài chính.
    names_empty = all(norm(row.get(c, "")) == "" for c in name_cols) if name_cols else True
    if names_empty:
        return LBL_MISSING_ICB

    return LBL_NON_FIN


def _unique_sheet_name(base, used):
    """Tạo tên sheet hợp lệ Excel: <=31 ký tự, không trùng.

    `used` là tập tên đã dùng (gồm cả RESERVED_SHEETS). Khi trùng thì thêm
    hậu tố _2/_3/... và cắt cho vừa 31 ký tự. Trả về tên mới (không tự thêm
    vào `used` — caller chịu trách nhiệm cập nhật).
    """
    base = (base or "sheet").strip() or "sheet"
    candidate = base[:31]
    if candidate not in used:
        return candidate
    i = 2
    while True:
        suffix = f"_{i}"
        candidate = base[: 31 - len(suffix)] + suffix
        if candidate not in used:
            return candidate
        i += 1


# ---------------------------------------------------------------------------
# 2) I/O (chỉ chạm vnstock / mạng — KHÔNG được import ở đường test)
# ---------------------------------------------------------------------------
def _assert_schema(df, required, what):
    if not isinstance(df, pd.DataFrame) or df.empty:
        raise ValueError(f"{what}: dữ liệu rỗng hoặc không phải DataFrame.")
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"{what}: thiếu cột bắt buộc {missing}. Có: {list(df.columns)}")


def fetch_hose_symbols(listing):
    """Lấy danh sách mã sàn HOSE từ vnstock Listing (đã truyền sẵn)."""
    ex = listing.symbols_by_exchange()
    _assert_schema(ex, ["symbol", "organ_name", "exchange"], "symbols_by_exchange")
    return ex


def fetch_icb(listing):
    """Lấy bảng ICB thật (dạng dài) từ vnstock Listing."""
    ind = listing.symbols_by_industries()
    _assert_schema(
        ind,
        ["symbol", "com_type_code", "icb_level", "icb_code", "icb_name"],
        "symbols_by_industries",
    )
    return ind


# ---------------------------------------------------------------------------
# 3) TRANSFORM / REPORT / EXPORT (thuần dữ liệu, test được)
# ---------------------------------------------------------------------------
def build_dataset(hose, icb):
    """Lọc HOSE, trải ICB thành cột, merge và phân loại.

    Trả về (df, name_cols). df có thêm cột phan_loai và nganh_icb_cap1.
    Hành vi thiếu ICB: dòng không có tên ICB và không có tín hiệu tài chính
    -> nhãn "Chưa phân loại (thiếu ICB)".
    """
    ex = hose.copy()
    ex["exchange"] = ex["exchange"].astype(str).str.upper().str.strip()
    h = ex[ex["exchange"].isin(["HSX", "HOSE"])][
        ["symbol", "organ_name"]
    ].drop_duplicates("symbol")

    before = len(h)
    mask3 = h["symbol"].astype(str).str.fullmatch(SYMBOL_RE.pattern)
    h = h[mask3]
    dropped = before - len(h)
    logger.info("Loại %d mã không khớp bộ lọc 3 chữ cái [A-Za-z]{3}.", dropped)

    h = h[
        h["organ_name"].notna()
        & (h["organ_name"].astype(str).str.strip().str.lower() != "none")
    ]

    # Trải ICB dạng dài -> rộng theo icb_level.
    piv = icb.pivot_table(
        index="symbol", columns="icb_level",
        values=["icb_code", "icb_name"], aggfunc="first",
    )
    piv.columns = [f"{a}_lv{int(b)}" for a, b in piv.columns]
    piv = piv.reset_index()
    piv = piv.merge(
        icb[["symbol", "com_type_code"]].drop_duplicates("symbol"),
        on="symbol", how="left",
    )
    df = h.merge(piv, on="symbol", how="left")

    name_cols = [c for c in df.columns if c.startswith("icb_name_lv")]

    df["phan_loai"] = df.apply(lambda r: classify(r, name_cols), axis=1)
    df["nganh_icb_cap1"] = df.apply(
        lambda r: r.get("icb_name_lv1") if r["phan_loai"] == LBL_NON_FIN else "",
        axis=1,
    )
    # Đường VCI dùng ICB thật, không có ngành gốc KBS -> để rỗng cho đồng nhất
    # schema xuất Excel (symbol, organ_name, nganh_goc, nganh_icb_cap1).
    df["nganh_goc"] = ""
    return df, name_cols


def build_dataset_kbs(hose, kbs_df, industry_code_col=KBS_INDUSTRY_CODE_COL,
                      include_real_estate=None):
    """Đường KBS: KBS trả về MỘT mã ngành nguyên/symbol (không pivot 4 cấp).

    Quy đổi mã KBS -> ICB cấp 1 qua KBS_TO_ICB:
      - nganh_goc      : tên ngành KBS (tiếng Việt)
      - bucket -> phan_loai (tái dùng nhãn LBL_*)
      - nganh_icb_cap1 : ICB cấp 1 quy đổi (rỗng với nhóm bị loại)
    GAN_TAY override SAU merge: gán nganh_icb_cap1 + phan_loai=LBL_NON_FIN.
    Mã không có mã KBS và không nằm trong GAN_TAY -> LBL_MISSING_ICB.

    `industry_code_col` là tên cột mã ngành KBS — chưa biết chắc khi chưa chạy
    live nên cho phép override. `include_real_estate` không dùng ở đây (mẫu lọc
    về sau) nhưng giữ tham số cho đồng nhất chữ ký.
    """
    ex = hose.copy()
    ex["exchange"] = ex["exchange"].astype(str).str.upper().str.strip()
    h = ex[ex["exchange"].isin(["HSX", "HOSE"])][
        ["symbol", "organ_name"]
    ].drop_duplicates("symbol")

    before = len(h)
    mask3 = h["symbol"].astype(str).str.fullmatch(SYMBOL_RE.pattern)
    h = h[mask3]
    logger.info("Loại %d mã không khớp bộ lọc 3 chữ cái.", before - len(h))

    h = h[
        h["organ_name"].notna()
        & (h["organ_name"].astype(str).str.strip().str.lower() != "none")
    ]

    if industry_code_col not in kbs_df.columns:
        raise ValueError(
            f"KBS: thiếu cột mã ngành '{industry_code_col}'. "
            f"Có: {list(kbs_df.columns)}. Truyền industry_code_col cho đúng."
        )

    codes = kbs_df[["symbol", industry_code_col]].drop_duplicates("symbol").copy()
    codes = codes.rename(columns={industry_code_col: "kbs_code"})
    df = h.merge(codes, on="symbol", how="left")

    def _to_int(v):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        try:
            return int(float(v))
        except (TypeError, ValueError):
            return None

    nganh_goc, phan_loai, nganh_cap1 = [], [], []
    for _, r in df.iterrows():
        code = _to_int(r.get("kbs_code"))
        entry = KBS_TO_ICB.get(code)
        if entry is None:
            nganh_goc.append("")
            phan_loai.append(LBL_MISSING_ICB)
            nganh_cap1.append("")
        else:
            name_kbs, bucket, icb1, _approx = entry
            nganh_goc.append(name_kbs)
            phan_loai.append(BUCKET_TO_LABEL[bucket])
            nganh_cap1.append(icb1 if (bucket == "NONFIN" and icb1) else "")

    df["nganh_goc"] = nganh_goc
    df["phan_loai"] = phan_loai
    df["nganh_icb_cap1"] = nganh_cap1

    # GAN_TAY override SAU merge / TRƯỚC khi hoàn tất.
    for sym, icb1 in GAN_TAY.items():
        mask = df["symbol"] == sym
        if mask.any():
            df.loc[mask, "nganh_icb_cap1"] = icb1
            df.loc[mask, "phan_loai"] = LBL_NON_FIN
            df.loc[mask & (df["nganh_goc"] == ""), "nganh_goc"] = "Gán tay"

    return df, []


def _non_financial_sample(df, include_real_estate):
    """Chọn mẫu phi tài chính tại thời điểm lấy mẫu (không qua nhãn).

    Phi tài chính luôn vào mẫu; Bất động sản chỉ vào khi include_real_estate.
    Loại "Chưa phân loại (thiếu ICB)" KHÔNG bao giờ vào mẫu.
    """
    labels = [LBL_NON_FIN]
    if include_real_estate:
        labels.append(LBL_REAL_ESTATE)
    return df[df["phan_loai"].isin(labels)].copy()


def make_reports(df, include_real_estate=False):
    """Tạo nội dung báo cáo (dạng text) + cảnh báo thiếu ICB.

    Trả về dict: keys 'lines' (list[str] in ra stdout), 'sample' (DataFrame),
    'missing_symbols' (list), 'warning' (str|None), 'counts' (Series).
    """
    counts = df["phan_loai"].value_counts()
    sample = _non_financial_sample(df, include_real_estate)

    missing_df = df[df["phan_loai"] == LBL_MISSING_ICB]
    missing_symbols = sorted(missing_df["symbol"].astype(str).tolist())
    n_missing = len(missing_symbols)

    warning = None
    if n_missing > 0:
        warning = (
            f"CẢNH BÁO — CHƯA RÕ: {n_missing} mã thiếu ICB (không phân loại "
            f"được, đã loại khỏi mẫu phi tài chính): {', '.join(missing_symbols)}. "
            f"Bổ sung các mã này vào bảng GAN_TAY (đầu script) rồi chạy lại."
        )

    lines = []
    lines.append(
        f"Tổng HOSE: {len(df)} | Mẫu phi tài chính"
        f"{' (gồm BĐS)' if include_real_estate else ''}: {len(sample)}\n"
    )
    lines.append("Phân bố tổng:")
    lines.append(counts.to_string() + "\n")
    lines.append("Phi tài chính theo ngành ICB cấp 1:")
    lines.append(
        sample[sample["phan_loai"] == LBL_NON_FIN]["nganh_icb_cap1"]
        .value_counts().to_string() + "\n"
    )
    lines.append("Mẫu kiểm tra ICB thật (tối đa 10 dòng):")
    cols = [c for c in ["symbol", "icb_code_lv1", "icb_name_lv1",
                        "com_type_code", "phan_loai"] if c in df.columns]
    lines.append(df[cols].head(10).to_string(index=False))

    lines.append("\n" + "=" * 60)
    nonfin = sample[sample["phan_loai"] == LBL_NON_FIN]
    for ng in sorted(nonfin["nganh_icb_cap1"].dropna().unique()):
        if ng == "":
            continue
        nhom = nonfin[nonfin["nganh_icb_cap1"] == ng].sort_values("symbol")
        lines.append(f"\n### {ng}  ({len(nhom)} DN)")
        for _, r in nhom.iterrows():
            lines.append(f"  {str(r['symbol']):<5} {r['organ_name']}")

    return {
        "lines": lines,
        "sample": sample,
        "missing_symbols": missing_symbols,
        "warning": warning,
        "counts": counts,
    }


# Cột bàn giao bắt buộc cho các sheet danh sách (DoD §4).
DELIVERY_COLS = ["symbol", "organ_name", "nganh_goc", "nganh_icb_cap1"]


def _ensure_delivery_cols(frame):
    """Bảo đảm frame có đủ DELIVERY_COLS (rỗng nếu thiếu) và đặt lên đầu."""
    out = frame.copy()
    for c in DELIVERY_COLS:
        if c not in out.columns:
            out[c] = ""
    rest = [c for c in out.columns if c not in DELIVERY_COLS]
    return out[DELIVERY_COLS + rest]


def export_excel(df, path, include_real_estate=False):
    """Xuất Excel nhiều sheet. Không crash khi mẫu phi tài chính rỗng.

    Mọi sheet danh sách mang đủ cột symbol, organ_name, nganh_goc,
    nganh_icb_cap1 (DoD §4) — đặt lên đầu cho dễ đọc.
    """
    df = _ensure_delivery_cols(df)
    sample = _non_financial_sample(df, include_real_estate)
    nonfin = sample[sample["phan_loai"] == LBL_NON_FIN]

    used = set(RESERVED_SHEETS)
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.sort_values("symbol").to_excel(w, sheet_name="Toan_bo", index=False)
        if sample.empty:
            # Vẫn tạo sheet rỗng để file hợp lệ.
            sample.to_excel(w, sheet_name="Phi_tai_chinh", index=False)
        else:
            sample.sort_values(["phan_loai", "nganh_icb_cap1", "symbol"]).to_excel(
                w, sheet_name="Phi_tai_chinh", index=False
            )
            for ng in sorted(nonfin["nganh_icb_cap1"].dropna().unique()):
                if ng == "":
                    continue
                base = norm(ng).replace(" ", "_")
                sheet = _unique_sheet_name(base, used)
                used.add(sheet)
                nonfin[nonfin["nganh_icb_cap1"] == ng].sort_values(
                    "symbol"
                ).to_excel(w, sheet_name=sheet, index=False)
    return path


# ---------------------------------------------------------------------------
# 4) NGUỒN DỮ LIỆU (VCI ưu tiên, fallback KBS) — chạm vnstock, lazy import
# ---------------------------------------------------------------------------
def _build_listing(source):
    """Tạo vnstock Listing(source=...). Tách ra để test inject/monkeypatch."""
    from vnstock import Listing  # import TRỄ -> test offline không chạm

    return Listing(source=source)


def _fetch_kbs(listing_factory):
    """Lấy dữ liệu KBS: (hose, kbs_industries, "KBS")."""
    lst = listing_factory("kbs")
    hose = fetch_hose_symbols(lst)
    kbs = lst.symbols_by_industries()
    return hose, kbs, "KBS"


def resolve_source(listing_factory=_build_listing, prefer="auto"):
    """Chọn nguồn dữ liệu. KHÔNG retry.

    prefer:
      - "auto" (mặc định): thử VCI (ICB thật) trước; lỗi/rỗng -> fallback KBS.
      - "vci": ép dùng VCI (không fallback).
      - "kbs": ép dùng KBS (bỏ qua VCI) — dùng để tái lập số liệu DoD 7 nhóm.

    `listing_factory(source)` trả về một Listing-like object. Trong test có
    thể inject một factory giả để không cần vnstock thật.
    Trả về (hose_df, data_df, source_name) với source_name ∈ {"VCI","KBS"}.
    """
    if prefer == "kbs":
        return _fetch_kbs(listing_factory)

    # --- Thử VCI ---
    try:
        lst = listing_factory("vci")
        hose = fetch_hose_symbols(lst)
        icb = fetch_icb(lst)
        if isinstance(hose, pd.DataFrame) and not hose.empty \
                and isinstance(icb, pd.DataFrame) and not icb.empty:
            return hose, icb, "VCI"
        if prefer == "vci":
            raise ValueError("VCI trả về dữ liệu rỗng/không hợp lệ.")
        logger.warning("VCI trả về dữ liệu rỗng/không hợp lệ -> fallback KBS.")
    except Exception as exc:  # vd 403 Forbidden
        if prefer == "vci":
            raise
        logger.warning("VCI lỗi (%s) -> fallback KBS.", exc)

    # --- Fallback KBS ---
    return _fetch_kbs(listing_factory)


# ---------------------------------------------------------------------------
# 5) TÀI LIỆU PHƯƠNG PHÁP (Output 2) — văn bản soạn tay + 2 sheet tính tự động
# ---------------------------------------------------------------------------
_SOAN_TAY = "(Soạn tay — không tính toán tự động)"

# 4 sheet văn bản tham chiếu (hằng số soạn tay).
REF_PHAN_NGANH_ICB = [
    "Phân ngành ICB (Industry Classification Benchmark) — 4 cấp.",
    "Cấp 1 (Industry): nhóm ngành lớn nhất, dùng để dựng panel theo ngành.",
    "ICB có 10/11 ngành cấp 1; mẫu phi tài chính HOSE thường rơi vào: Công "
    "nghiệp, Hàng tiêu dùng, Dịch vụ tiêu dùng, Vật liệu cơ bản, Y tế, Công "
    "nghệ, Tiện ích cộng đồng, Dầu khí, Viễn thông.",
    "Tài chính (ngân hàng, chứng khoán, bảo hiểm, BĐS) bị LOẠI khỏi mẫu.",
]
REF_VSIC = [
    "VSIC — Hệ thống ngành kinh tế Việt Nam (QĐ 36/2025/QĐ-TTg).",
    "Chuẩn pháp quy VN; KHÔNG dùng làm chuẩn chính trong nghiên cứu này.",
    "Chỉ tham chiếu khi hội đồng yêu cầu chuyển sang chuẩn pháp quy VN.",
]
REF_DOI_CHIEU = [
    "Đối chiếu ICB <-> VSIC (tham khảo, KHÔNG dùng để phân loại tự động).",
    "ICB nhóm theo bản chất tài chính/thị trường; VSIC theo hoạt động kinh tế.",
    "Hai khung không ánh xạ 1-1; bảng đối chiếu chỉ mang tính định hướng.",
]
REF_TRICH_DAN = [
    "Tài liệu trích dẫn:",
    "Nguyen, Kim & Ali (2024) — bài kế thừa về Quản trị lợi nhuận (Earnings "
    "Management) bằng dồn tích bất thường (mô hình Jones điều chỉnh).",
    "FTSE Russell — Industry Classification Benchmark (ICB) methodology.",
    "Ràng buộc mẫu: mỗi ngành cần >=10 quan sát firm-year khi dựng panel.",
]


def export_reference_workbook(path, source_name, ngay_chot, kbs_map):
    """Xuất Output 2: Co_so_phan_loai_nganh_HOSE.xlsx — 6 sheet.

    2 sheet TÍNH TỰ ĐỘNG (data-driven): "Co so phan loai" và "Anh xa KBS-ICB".
    4 sheet VĂN BẢN soạn tay: Phan nganh ICB / VSIC QD36-2025 / Doi chieu
    ICB-VSIC / Tai lieu trich dan (dòng đầu ghi rõ "(Soạn tay...)").
    Tên sheet <=31 ký tự, không dấu để tránh lỗi Excel.
    """
    so_nhom = "9 nhóm ICB (ICB thật)" if source_name == "VCI" \
        else "7 nhóm ICB (KBS quy đổi)"
    co_so_rows = [
        ("Nguồn dữ liệu sử dụng", source_name),
        ("Số nhóm ngành", so_nhom),
        ("Ngày chốt dữ liệu (NGAY_CHOT)", ngay_chot),
        ("Sàn", "HOSE (HSX)"),
        ("Cấp phân ngành", "ICB cấp 1"),
        ("Quy tắc loại trừ — Tài chính",
         "Loại NH/CK/BH/Tài chính khác"),
        ("Quy tắc loại trừ — Bất động sản",
         f"EXCLUDE_REAL_ESTATE = {EXCLUDE_REAL_ESTATE}"),
        ("Gán tay (cần xác nhận)",
         ", ".join(f"{k}->{v}" for k, v in GAN_TAY.items())),
    ]
    df_co_so = pd.DataFrame(co_so_rows, columns=["Mục", "Nội dung"])

    # Ánh xạ KBS->ICB (data-driven), kèm cờ (*) cho mã gần đúng.
    rows = []
    for code in sorted(kbs_map):
        name_kbs, bucket, icb1, approx = kbs_map[code]
        rows.append({
            "Mã KBS": code,
            "Ngành KBS": name_kbs + (" (*)" if approx else ""),
            "Phân loại NC": "LOẠI" if icb1 is None else "Phi tài chính",
            "ICB cấp 1": icb1 if icb1 is not None else "—",
            "Gần đúng (*)": "có" if approx else "",
        })
    df_kbs = pd.DataFrame(rows)

    def _prose_df(lines):
        return pd.DataFrame([_SOAN_TAY] + list(lines), columns=["Nội dung"])

    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df_co_so.to_excel(w, sheet_name="Co so phan loai", index=False)
        _prose_df(REF_PHAN_NGANH_ICB).to_excel(
            w, sheet_name="Phan nganh ICB", index=False)
        _prose_df(REF_VSIC).to_excel(
            w, sheet_name="VSIC QD36-2025", index=False)
        _prose_df(REF_DOI_CHIEU).to_excel(
            w, sheet_name="Doi chieu ICB-VSIC", index=False)
        _prose_df(REF_TRICH_DAN).to_excel(
            w, sheet_name="Tai lieu trich dan", index=False)
        df_kbs.to_excel(w, sheet_name="Anh xa KBS-ICB", index=False)
    return path


# ---------------------------------------------------------------------------
# 6) ORCHESTRATOR
# ---------------------------------------------------------------------------
def main(argv):
    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    parser = argparse.ArgumentParser(
        description="Phân ngành HOSE theo ICB thật (nguồn VCI)."
    )
    parser.add_argument(
        "--output", default="HOSE_phi_tai_chinh.xlsx",
        help="Đường dẫn file Excel danh sách (Output 1).",
    )
    parser.add_argument(
        "--reference-output", default="Co_so_phan_loai_nganh_HOSE.xlsx",
        help="Đường dẫn file tài liệu phương pháp (Output 2).",
    )
    parser.add_argument(
        "--kbs-industry-col", default=KBS_INDUSTRY_CODE_COL,
        help="Tên cột mã ngành KBS (override khi biết tên thật lúc chạy live).",
    )
    parser.add_argument(
        "--include-real-estate", action="store_true", default=False,
        help="Đưa bất động sản vào mẫu phi tài chính (mặc định loại).",
    )
    parser.add_argument(
        "--source", choices=["auto", "vci", "kbs"], default="kbs",
        help="Mặc định kbs (7 nhóm ICB, chốt DoD); auto=VCI rồi fallback KBS; vci=ép VCI.",
    )
    args = parser.parse_args(argv)

    # Mặc định include_real_estate suy từ EXCLUDE_REAL_ESTATE; cờ CLI ép bật.
    include_real_estate = args.include_real_estate or (not EXCLUDE_REAL_ESTATE)

    # Ưu tiên VCI, tự fallback KBS (lazy import vnstock bên trong).
    try:
        hose, data, source_name = resolve_source(prefer=args.source)
    except Exception as exc:  # fail-fast với thông báo dễ đọc
        print(f"LỖI khi lấy dữ liệu vnstock: {exc}", file=sys.stderr)
        return 2

    if source_name == "VCI":
        print("Nguồn dữ liệu: VCI (ICB thật, 9 nhóm)")
        df, _ = build_dataset(hose, data)
    else:
        print("Nguồn dữ liệu: KBS (quy đổi, 7 nhóm)")
        df, _ = build_dataset_kbs(
            hose, data, industry_code_col=args.kbs_industry_col,
            include_real_estate=include_real_estate,
        )

    report = make_reports(df, include_real_estate=include_real_estate)

    print("\n".join(report["lines"]))
    if report["warning"]:
        logger.warning(report["warning"])

    try:
        export_excel(df, args.output, include_real_estate=include_real_estate)
        export_reference_workbook(
            args.reference_output, source_name, NGAY_CHOT, KBS_TO_ICB)
    except Exception as exc:
        print(f"LỖI khi xuất Excel: {exc}", file=sys.stderr)
        return 3

    print(f"\nĐã lưu {args.output}")
    print(f"Đã lưu {args.reference_output}")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
