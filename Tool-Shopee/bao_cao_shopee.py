"""Tạo báo cáo bán hàng Shopee theo mặt hàng từ file xuất từ Shopee Seller Center.

Cách dùng: kéo thả file .xlsx vào Tao-bao-cao.bat
       hoặc: python bao_cao_shopee.py "<đường dẫn file>.xlsx"
Báo cáo được lưu vào thư mục "Báo cáo" nằm cạnh công cụ.
"""
import math
import os
import re
import subprocess
import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

# openpyxl đi kèm trong thư mục "thu-vien" để chạy được trên máy chưa cài thư viện này
sys.path.insert(0, str(Path(__file__).resolve().parent / 'thu-vien'))

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

import licensing  # cấp phép offline 48 giờ (bản sao trong thu-vien)

OUT_DIR = Path(__file__).resolve().parent / 'Báo cáo'

# ---- tên sheet ----
OVERVIEW = 'Tổng quan'
PROD     = 'Theo mặt hàng'
RET      = 'Đơn hoàn & hủy'
DATA     = 'Dữ liệu chi tiết'
HOW      = 'Cách tính'

# ---- màu & định dạng ----
NAVY, BLUE, LIGHT, BAND = '1F3864', '1F4E78', 'DDEBF7', 'F5F8FC'
TOTAL, GREEN, TILE, RED, ORANGE, GRAY, LINE = 'D9E1F2', 'E2EFDA', 'F2F6FC', 'C00000', 'F4B942', '595959', 'BFBFBF'
NUM   = '#,##0;-#,##0;"-"'
PCT   = '0.0%;-0.0%;"-"'
MONEY = '#,##0" đ";-#,##0" đ";"-"'
THIN  = Side(style='thin', color=LINE)
BOX   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def font(size=10, bold=False, color='262626', italic=False, underline=None):
    return Font(name='Arial', size=size, bold=bold, color=color, italic=italic, underline=underline)


def fill(color):
    return PatternFill('solid', fgColor=color)


CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT   = Alignment(horizontal='left',   vertical='center', wrap_text=True)
RIGHT  = Alignment(horizontal='right',  vertical='center')


def num(x):
    if x is None or x == '' or x == '-':
        return 0.0
    try:
        return float(str(x).replace(',', '.'))
    except (ValueError, TypeError):
        return 0.0


def short_name(full):
    """Rút gọn tên: lấy phần trước dấu ' – ' hoặc ' - '."""
    for sep in (' – ', ' - '):
        head = full.split(sep)[0].strip()
        if len(head) >= 10:
            return head
    return full


def parse_date(val):
    """Parse ngày Shopee: '2026-07-10 12:07' hoặc datetime object."""
    if val is None or val == '' or val == '-':
        return None
    if isinstance(val, datetime):
        return val
    s = str(val).strip()
    if not s or s == '-':
        return None
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d'):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def read_orders(src):
    """Đọc file Shopee, trả về list dict theo header dòng 1."""
    wb = load_workbook(src, data_only=True)
    ws = wb.worksheets[0]
    it  = ws.iter_rows(values_only=True)
    hdr = list(next(it))
    data = []
    for r in it:
        d = dict(zip(hdr, r))
        order_id = str(d.get(hdr[0]) or '').strip()
        if order_id:
            data.append(d)
    wb.close()
    return data


def to_rows(data):
    """Chuyển dữ liệu thô → list dict chuẩn hóa.

    Shopee: mỗi dòng = 1 SKU/đơn; 1 đơn có thể nhiều dòng (nhiều sản phẩm).
    Gom theo Mã đơn hàng khi đếm số đơn.
    """
    if not data:
        return []

    hdr = list(data[0].keys())

    def col(idx):   # 1-indexed
        return hdr[idx - 1]

    rows        = []
    seen_orders = set()

    for r in data:
        def v(idx):
            return r.get(col(idx))

        order_id      = str(v(1)  or '').strip()
        order_date    = parse_date(v(3))
        status        = str(v(4)  or '').strip()
        cancel_reason = str(v(6)  or '').strip()
        sent_date     = parse_date(v(13))
        return_status = str(v(15) or '').strip()
        sku           = str(v(16) or '').strip()
        full_name     = re.sub(r'\s+', ' ', str(v(17) or '')).strip()
        raw_size      = str(v(21) or '').strip()
        size          = raw_size if raw_size else '—'   # dùng '—' khi không có phân loại
        price_orig    = num(v(22))
        seller_disc   = num(v(23))
        shopee_disc   = num(v(24))
        price_promo   = num(v(26))
        qty           = int(num(v(27)))
        qty_returned  = int(num(v(28)))
        paid          = num(v(29))   # tiền khách trả theo từng dòng
        payment_method = str(v(49) or '').strip()
        province      = str(v(57) or '').strip()

        prod = short_name(full_name)

        # ---- phân loại đơn ----
        # Shopee: "Đã hủy" → Hủy; hoàn trả được chấp thuận → Hoàn; "Hoàn thành" → Bán;
        # còn lại (chờ xác nhận / chờ lấy hàng / đang giao / đã giao…) → Đang xử lý: chưa tính doanh thu
        is_cancelled = (status == 'Đã hủy')
        is_return    = (return_status == 'Đã Chấp Thuận Yêu Cầu')
        if is_cancelled:
            group = 'Hủy'
        elif is_return:
            group = 'Hoàn'
        elif status == 'Hoàn thành':
            group = 'Bán'
        else:
            group = 'Đang xử lý'

        # Phân loại lý do hủy để hiển thị rõ hơn
        cancel_type = ''
        if is_cancelled:
            lr = cancel_reason.lower()
            if 'tự động hủy' in lr and ('giao hàng thất bại' in lr or 'giao thất bại' in lr):
                cancel_type = 'Giao thất bại'
            elif 'tự động hủy' in lr and 'không gửi hàng đúng hạn' in lr:
                cancel_type = 'Shop giao trễ'
            elif 'tự động hủy' in lr:
                cancel_type = 'Hệ thống tự hủy'
            elif 'hủy bởi người mua' in lr or 'người mua' in lr:
                cancel_type = 'Người mua hủy'
            elif 'hủy bởi shop' in lr or 'người bán' in lr:
                cancel_type = 'Shop hủy'
            else:
                cancel_type = 'Khác'

        # đếm order: 1 lần / mã đơn
        is_first = order_id not in seen_orders
        if is_first:
            seen_orders.add(order_id)
        cnt_o = 1 if is_first else 0

        rows.append(dict(
            order         = order_id,
            created       = order_date,
            status        = status,
            return_status = return_status,
            cancel_reason = cancel_reason,
            cancel_type   = cancel_type,
            prod          = prod,
            full          = full_name,
            sku           = sku,
            size          = size,
            qty           = qty,
            price_orig    = price_orig,
            price_promo   = price_promo,
            seller_disc   = seller_disc,
            shopee_disc   = shopee_disc,
            paid          = paid,
            qty_returned  = qty_returned,
            group         = group,
            sent          = sent_date,
            payment       = payment_method,
            province      = province,
            cnt_o         = cnt_o,
        ))
    return rows


# ---- sheet "Dữ liệu chi tiết" ----
DATA_COLS = [
    ('order',         'Mã đơn',          20, None),
    ('created',       'Ngày đặt',         16, 'dd/mm/yyyy hh:mm'),
    ('status',        'Trạng thái',       14, None),
    ('return_status', 'Trạng thái hoàn',  20, None),
    ('cancel_reason', 'Lý do hủy',        36, None),
    ('cancel_type',   'Loại hủy',         16, None),
    ('prod',          'Mặt hàng',         36, None),
    ('full',          'Tên đầy đủ',       50, None),
    ('sku',           'SKU',              20, None),
    ('size',          'Phân loại',        10, None),
    ('qty',           'Số lượng',         10, NUM),
    ('price_orig',    'Giá gốc',          12, NUM),
    ('price_promo',   'Giá ưu đãi',       12, NUM),
    ('seller_disc',   'Giảm giá shop',    13, NUM),
    ('shopee_disc',   'Shopee trợ giá',   13, NUM),
    ('paid',          'Tiền khách trả',   14, NUM),
    ('qty_returned',  'SL hoàn trả',      11, NUM),
    ('payment',       'Thanh toán',       22, None),
    ('province',      'Tỉnh/Thành phố',   22, None),
    ('group',         'Phân loại đơn',    15, None),
    ('cnt_o',         'Cột phụ: đếm đơn', 12, None),
]
HELPERS = {'cnt_o'}
DCOL = {k: get_column_letter(i) for i, (k, *_) in enumerate(DATA_COLS, 1)}

# ---- sheet "Theo mặt hàng" ----
PROD_COLS = [
    ('stt',         'STT',                     6),
    ('prod',        'Mặt hàng',               36),
    ('size',        'Phân loại',              12),
    ('orders',      'Số đơn',                 10),
    ('qty',         'Tổng SL',                11),
    ('cancelled',   'Đơn hủy',               10),
    ('returned',    'Đã hoàn trả (SL)',        12),
    ('rev',         'Doanh thu (khách trả)',   16),
    ('seller_disc', 'Giảm giá shop',           14),
    ('shopee_disc', 'Shopee trợ giá',          14),
    ('net',         'Doanh thu thuần',         15),
    ('share',       'Tỷ trọng',                9),
]
PIDX = {k: i for i, (k, *_) in enumerate(PROD_COLS, 1)}
PCOL = {k: get_column_letter(i) for k, i in PIDX.items()}


def rng(key, n):
    c = DCOL[key]
    return f"'{DATA}'!${c}$2:${c}${n}"


def row_height(text, width, base=15):
    return base * max(1, math.ceil(len(text) / (width * 0.95))) + 4


def setup_print(ws, title_rows=None, landscape=True):
    ws.page_setup.orientation = 'landscape' if landscape else 'portrait'
    ws.page_setup.paperSize   = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    if title_rows:
        ws.print_title_rows = title_rows


def title_cell(ws, text, sub, col='A'):
    ws[f'{col}1'] = text
    ws[f'{col}1'].font = font(14, True, NAVY)
    ws[f'{col}2'] = sub
    ws[f'{col}2'].font = font(9, color=GRAY, italic=True)
    ws.row_dimensions[1].height = 24


# ─────────────────────────────────────────────────────
# Sheet: Dữ liệu chi tiết
# ─────────────────────────────────────────────────────
def sheet_data(wb, rows):
    ws = wb.create_sheet(DATA)
    ws.sheet_properties.tabColor = '7F7F7F'
    ws.append([h for _, h, *_ in DATA_COLS])
    for r in rows:
        ws.append([r[k] for k, *_ in DATA_COLS])
    n = len(rows) + 1
    for i, (k, _, w, fmt) in enumerate(DATA_COLS, 1):
        c   = ws.cell(1, i)
        c.font      = font(10, True, 'FFFFFF')
        c.fill      = fill('A6A6A6' if k in HELPERS else BLUE)
        c.alignment = CENTER
        c.border    = BOX
        dim = ws.column_dimensions[get_column_letter(i)]
        dim.width = w
        if k in HELPERS:
            dim.outlineLevel, dim.hidden = 1, True
        for (cell,) in ws.iter_rows(min_row=2, max_row=n, min_col=i, max_col=i):
            cell.font = font()
            if fmt:
                cell.number_format = fmt
    ws.row_dimensions[1].height = 32
    ws.freeze_panes = 'B2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(DATA_COLS))}{n}'
    return n


# ─────────────────────────────────────────────────────
# Sheet: Theo mặt hàng
# ─────────────────────────────────────────────────────
def sheet_products(wb, rows, n, period):
    ws = wb.create_sheet(PROD)
    ws.sheet_properties.tabColor = BLUE
    ws.sheet_view.showGridLines  = False
    title_cell(ws, 'DOANH THU THEO MẶT HÀNG',
               f'{period}  •  Chỉ tính đơn Bán (đã trừ tiền hoàn)  •  Sắp xếp theo doanh thu thuần giảm dần')

    # Header dòng 4
    for i, (k, h, w) in enumerate(PROD_COLS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
        c = ws.cell(4, i, h)
        c.font      = font(10, True, 'FFFFFF')
        c.fill      = fill(BLUE)
        c.alignment = CENTER
        c.border    = BOX
    ws.row_dimensions[4].height = 42

    # Xây danh sách combo (prod, size) — dùng size='—' khi trống
    net_by_combo = defaultdict(float)
    combos       = {}  # (prod, size) → True nếu có đơn Bán/Hoàn
    for r in rows:
        key = (r['prod'], r['size'])
        if r['group'] == 'Bán':
            net_by_combo[key] += r['paid']  # doanh thu thuần = tiền đơn Bán
            combos[key] = True
        elif r['group'] == 'Hoàn':
            combos[key] = True
        elif r['group'] == 'Hủy':
            if key not in combos:
                combos[key] = False   # chỉ có đơn hủy

    # Sắp theo doanh thu thuần giảm dần
    sorted_combos = sorted(combos.keys(), key=lambda k: -net_by_combo.get(k, 0))

    first = 5
    last  = first + len(sorted_combos) - 1
    tr    = last + 1
    P     = PCOL

    for k, (prod, size) in enumerate(sorted_combos):
        r = first + k
        b_ban  = f'{rng("prod", n)},$B{r},{rng("size", n)},$C{r},{rng("group", n)},"Bán"'
        b_hoan = f'{rng("prod", n)},$B{r},{rng("size", n)},$C{r},{rng("group", n)},"Hoàn"'
        b_huy  = f'{rng("prod", n)},$B{r},{rng("size", n)},$C{r},{rng("group", n)},"Hủy"'
        f = {
            'stt':         k + 1,
            'prod':        prod,
            'size':        size,
            # Số đơn: chỉ đếm đơn Bán (không tính Hoàn vào cột này để khớp trực giác)
            'orders':      f'=SUMIFS({rng("cnt_o", n)},{b_ban})',
            # Tổng SL xuất: chỉ đơn Bán
            'qty':         f'=SUMIFS({rng("qty", n)},{b_ban})',
            # Đơn hủy
            'cancelled':   f'=SUMIFS({rng("cnt_o", n)},{b_huy})',
            # SL hoàn trả: lấy qty_returned của đơn Bán + qty của đơn Hoàn
            'returned':    f'=SUMIFS({rng("qty_returned", n)},{b_ban})+SUMIFS({rng("qty", n)},{b_hoan})',
            # Doanh thu (khách trả) = tiền khách đã trả cho đơn Bán + đơn Hoàn (đơn Hoàn sau đó được hoàn tiền)
            'rev':         f'=SUMIFS({rng("paid", n)},{b_ban})+SUMIFS({rng("paid", n)},{b_hoan})',
            # Giảm giá shop = tổng seller_disc đơn Bán
            'seller_disc': f'=SUMIFS({rng("seller_disc", n)},{b_ban})',
            # Shopee trợ giá = tổng shopee_disc đơn Bán
            'shopee_disc': f'=SUMIFS({rng("shopee_disc", n)},{b_ban})',
            # DT thuần = Doanh thu (khách trả) − tiền hoàn đơn Hoàn = tiền đơn Bán
            'net':         f'={P["rev"]}{r}-SUMIFS({rng("paid", n)},{b_hoan})',
            'share':       f'=IF(${P["net"]}${tr}=0,0,{P["net"]}{r}/${P["net"]}${tr})',
        }
        for i, (key, *_) in enumerate(PROD_COLS, 1):
            ws.cell(r, i, f[key])
        ws.row_dimensions[r].height = row_height(prod, PROD_COLS[1][2])

    # Dòng tổng cộng
    ws.cell(tr, 2, 'TỔNG CỘNG')
    # Số đơn bán tổng (không trùng)
    ws.cell(tr, 4, f'=SUMIFS({rng("cnt_o", n)},{rng("group", n)},"Bán")')
    for key, *_ in PROD_COLS[4:]:
        c = P[key]
        ws[f'{c}{tr}'] = f'=SUM({c}{first}:{c}{last})'
    ws.row_dimensions[tr].height = 22

    # Định dạng
    for r in range(first, tr + 1):
        is_total = r == tr
        for i, (key, *_) in enumerate(PROD_COLS, 1):
            c = ws.cell(r, i)
            c.border    = BOX
            c.alignment = LEFT if key in ('prod', 'size') else (CENTER if key == 'stt' else RIGHT)
            c.number_format = PCT if key == 'share' else NUM
            color = RED if key == 'returned' else '262626'
            c.font = font(10, is_total or key == 'net', color)
            if is_total:
                c.fill = fill(TOTAL)
            elif key == 'net':
                c.fill = fill(GREEN)
            elif (r - first) % 2:
                c.fill = fill(BAND)
        if is_total:
            for i in range(1, len(PROD_COLS) + 1):
                ws.cell(r, i).border = Border(
                    left=THIN, right=THIN,
                    top=Side(style='medium', color=NAVY),
                    bottom=Side(style='medium', color=NAVY))

    note = tr + 2
    ws.cell(note, 2,
            'Số liệu TẠM TÍNH tại thời điểm xuất file; chưa trừ phí sàn, hoa hồng, thuế. '
            'Giải thích từng cột: xem sheet "Cách tính".').font = font(9, color=GRAY, italic=True)
    ws.freeze_panes = 'D5'
    ws.auto_filter.ref = f'A4:{get_column_letter(len(PROD_COLS))}{last}'
    setup_print(ws, '4:4')
    return tr


# ─────────────────────────────────────────────────────
# Sheet: Đơn hoàn & hủy
# ─────────────────────────────────────────────────────
def sheet_returns(wb, rows, period):
    ws = wb.create_sheet(RET)
    ws.sheet_properties.tabColor = RED
    ws.sheet_view.showGridLines  = False
    title_cell(ws, 'ĐƠN HOÀN & HỦY',
               f'{period}  •  Đơn bị hoàn trả và đơn bị hủy (chia theo loại)')

    cols = [
        ('STT',       5), ('Loại',        14), ('Mã đơn',     20),
        ('Ngày đặt', 16), ('Mặt hàng',    36), ('Phân loại',  10),
        ('Số lượng',  9), ('Tiền hoàn',   14), ('Giá trị đơn', 14),
        ('Lý do',    48),
    ]
    for i, (h, w) in enumerate(cols, 1):
        c = ws.cell(4, i, h)
        c.font, c.fill, c.alignment, c.border = font(10, True, 'FFFFFF'), fill(BLUE), CENTER, BOX
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[4].height = 30

    # ---- dữ liệu ----
    items = []
    for r in rows:
        if r['group'] == 'Hoàn':
            items.append((0, r['created'], 'Hoàn trả', r['order'], r['prod'],
                          r['size'], r['qty'], r['paid'], r['paid'],
                          'Đã Chấp Thuận Yêu Cầu'))
        elif r['group'] == 'Hủy':
            # Phân biệt: giao thất bại vs người mua hủy vs hệ thống hủy
            kind = r['cancel_type'] if r['cancel_type'] else 'Hủy đơn'
            items.append((1, r['created'], kind, r['order'], r['prod'],
                          r['size'], r['qty'], 0, r['paid'],
                          r['cancel_reason'] or ''))
    items.sort(key=lambda x: (x[0], x[1] or datetime.min))

    first = 5
    if not items:
        ws.cell(first, 2,
                'Không có đơn hoàn hay hủy trong file này.').font = font(10, italic=True)
        return

    for k, (_, created, kind, oid, prod, size, q, refund, order_val, reason) in enumerate(items):
        r   = first + k
        vals = [k + 1, kind, oid, created, prod, size, q, refund, order_val, reason]
        for i, val in enumerate(vals, 1):
            c = ws.cell(r, i, val)
            # tiền hoàn đỏ (col 8), giá trị đơn hủy cam (col 9)
            if   i == 8 and val > 0:
                c.font = font(10, color=RED)
            elif i == 9 and refund == 0 and val > 0:
                c.font = font(10, color=ORANGE)
            else:
                c.font = font(10)
            c.border    = BOX
            c.alignment = CENTER if i in (1, 2, 3, 4) else (RIGHT if i in (7, 8, 9) else LEFT)
            if (r - first) % 2:
                c.fill = fill(BAND)
        ws.cell(r, 4).number_format = 'dd/mm/yyyy hh:mm'
        ws.cell(r, 7).number_format = NUM
        ws.cell(r, 8).number_format = NUM
        ws.cell(r, 9).number_format = NUM
        ws.row_dimensions[r].height = max(row_height(prod, 36), row_height(reason, 48))

    tr = first + len(items)
    ws.cell(tr, 5, 'TỔNG CỘNG')
    ws.cell(tr, 7, f'=SUM(G{first}:G{tr - 1})')
    ws.cell(tr, 8, f'=SUM(H{first}:H{tr - 1})')
    ws.cell(tr, 9, f'=SUM(I{first}:I{tr - 1})')
    for i in range(1, len(cols) + 1):
        c = ws.cell(tr, i)
        c.font   = font(10, True, RED if i == 8 else (ORANGE if i == 9 else '262626'))
        c.fill   = fill(TOTAL)
        c.border = Border(left=THIN, right=THIN,
                          top=Side(style='medium', color=NAVY),
                          bottom=Side(style='medium', color=NAVY))
        c.number_format = NUM
        c.alignment = RIGHT if i in (7, 8, 9) else LEFT

    note = tr + 2
    ws.cell(note, 2, (
        'Cột "Tiền hoàn" (đỏ): số tiền thực sự hoàn lại cho khách — chỉ đơn Hoàn trả mới có.  '
        'Cột "Giá trị đơn" (cam): giá trị đơn hủy — khách chưa trả, shop không thu được.'
    )).font = font(9, color=GRAY, italic=True)
    ws.freeze_panes = 'A5'
    ws.auto_filter.ref = f'A4:J{tr - 1}'
    setup_print(ws, '4:4')


# ─────────────────────────────────────────────────────
# Sheet: Tổng quan
# ─────────────────────────────────────────────────────
def sheet_overview(ws, rows, n, tr, period, exported):
    ws.title = OVERVIEW
    ws.sheet_properties.tabColor = NAVY
    ws.sheet_view.showGridLines  = False
    ws.column_dimensions['A'].width = 2
    for c in 'BCDE':
        ws.column_dimensions[c].width = 28
    ws.column_dimensions['F'].width = 66

    ws['B2'] = 'BÁO CÁO BÁN HÀNG SHOPEE'
    ws['B2'].font = font(18, True, NAVY)
    ws.row_dimensions[2].height = 30
    ws['B3'] = f'{period}  •  Dữ liệu xuất lúc {exported:%d/%m/%Y %H:%M}  •  Đơn vị tiền: VNĐ'
    ws['B3'].font = font(10, color=GRAY, italic=True)

    P = lambda key: f"='{PROD}'!{PCOL[key]}{tr}"

    def cnt_grp(val):
        return f'=SUMIFS({rng("cnt_o", n)},{rng("group", n)},"{val}")'

    def sum_col(col, grp):
        return f'=SUMIFS({rng(col, n)},{rng("group", n)},"{grp}")'

    # ---- 3 hàng tile (mỗi hàng 4 tile) ----
    # Hàng 1: tiền
    # Hàng 2: số đơn & số lượng
    # Hàng 3: hoàn trả & hủy
    tiles = [
        [
            ('DOANH THU THUẦN',            P('net'),            MONEY, GREEN, NAVY),
            ('Doanh thu (khách trả)',       P('rev'),            MONEY, TILE, NAVY),
            ('Tiền hoàn cho khách',        f'={P("rev")[1:]}-{P("net")[1:]}', MONEY, TILE, RED),
            ('Giảm giá shop / Shopee bù',  f'={P("seller_disc")[1:]}+{P("shopee_disc")[1:]}', MONEY, TILE, NAVY),
        ],
        [
            ('Số đơn bán',                 cnt_grp('Bán'),      NUM,   TILE, NAVY),
            ('Tổng số lượng xuất (đơn bán)',P('qty'),           NUM,   TILE, NAVY),
            ('SL hoàn trả về kho',         P('returned'),       NUM,   TILE, RED),
            ('Shopee trợ giá',             P('shopee_disc'),    MONEY, TILE, NAVY),
        ],
        [
            ('Số đơn hủy',                 cnt_grp('Hủy'),      NUM,   TILE, GRAY),
            ('Giá trị đơn hủy (chưa thu)', sum_col('paid','Hủy'), MONEY, TILE, ORANGE),
            ('Số đơn hoàn trả',            cnt_grp('Hoàn'),     NUM,   TILE, RED),
            ('Giảm giá shop (đơn bán)',    P('seller_disc'),    MONEY, TILE, NAVY),
        ],
    ]

    r = 5
    for line in tiles:
        for j, (label, formula, fmt, bg, color) in enumerate(line):
            col = 2 + j
            top = ws.cell(r,     col, label)
            val = ws.cell(r + 1, col, formula)
            top.font      = font(9, True, GRAY)
            top.fill      = fill(bg)
            top.alignment = Alignment(horizontal='left', vertical='bottom', indent=1)
            val.font      = font(16, True, color)
            val.fill      = fill(bg)
            val.alignment = Alignment(horizontal='left', vertical='center', indent=1)
            val.number_format = fmt
            top.border = Border(left=THIN, right=THIN, top=THIN)
            val.border = Border(left=THIN, right=THIN, bottom=THIN)
        ws.row_dimensions[r    ].height = 20
        ws.row_dimensions[r + 1].height = 34
        ws.row_dimensions[r + 2].height = 8
        r += 3

    # ---- Bảng doanh thu theo trạng thái ----
    # Chỉ tính DT của đơn BÁN và ĐƠN HOÀN (không tính giá trị đơn hủy vào doanh thu)
    r += 1
    ws.cell(r,     2, 'DOANH THU THEO TRẠNG THÁI ĐƠN').font = font(12, True, NAVY)
    ws.cell(r + 1, 2,
            'Chỉ tính đơn "Bán" (hoàn thành) — không tính đơn Hủy vào doanh thu'
            ).font = font(9, color=GRAY, italic=True)
    h = r + 2
    for i, t in enumerate(['Trạng thái', 'Số đơn', 'Doanh thu (khách trả)', 'Tỷ trọng'], 2):
        c = ws.cell(h, i, t)
        c.font, c.fill, c.alignment, c.border = font(10, True, 'FFFFFF'), fill(BLUE), CENTER, BOX
    ws.row_dimensions[h].height = 24

    # Chỉ 2 trạng thái DT: Bán và Hoàn
    status_rows = [
        ('Đơn Bán (hoàn thành)', 'Bán'),
        ('Đơn Hoàn trả',         'Hoàn'),
    ]
    tot = h + len(status_rows) + 1
    for k, (label, group) in enumerate(status_rows):
        rr          = h + 1 + k
        rev_formula = f'=SUMIFS({rng("paid", n)},{rng("group", n)},"{group}")'
        cnt_formula = f'=SUMIFS({rng("cnt_o", n)},{rng("group", n)},"{group}")'
        # Tỷ trọng chia trên tổng DT của đơn Bán + Hoàn (không bao gồm hủy)
        share_formula = f'=IF($D${tot}=0,0,D{rr}/$D${tot})'
        vals = [label, cnt_formula, rev_formula, share_formula]
        for i, val in enumerate(vals, 2):
            c = ws.cell(rr, i, val)
            c.font      = font(10, i == 2, RED if (group == 'Hoàn' and i in (3,)) else '262626')
            c.border    = BOX
            c.alignment = LEFT if i == 2 else RIGHT
            c.number_format = PCT if i == 5 else NUM
            if k % 2:
                c.fill = fill(BAND)
        ws.row_dimensions[rr].height = 22

    # Dòng tổng: chỉ tổng DT Bán + Hoàn
    ws.cell(tot, 2, 'TỔNG (đơn Bán + Hoàn)')
    ws.cell(tot, 3, f'=SUM(C{h + 1}:C{tot - 1})')
    ws.cell(tot, 4, f'=SUM(D{h + 1}:D{tot - 1})')
    ws.cell(tot, 5, f'=SUM(E{h + 1}:E{tot - 1})')
    for i in range(2, 6):
        c = ws.cell(tot, i)
        c.font      = font(10, True)
        c.fill      = fill(TOTAL)
        c.number_format = PCT if i == 5 else NUM
        c.alignment = LEFT if i == 2 else RIGHT
        c.border    = Border(left=THIN, right=THIN,
                             top=Side(style='medium', color=NAVY),
                             bottom=Side(style='medium', color=NAVY))
    ws.row_dimensions[tot].height = 22

    # Thêm ghi chú về đơn hủy (tách biệt)
    note_r = tot + 1
    ws.cell(note_r, 2,
            f'Đơn Hủy: không tính vào doanh thu. '
            f'Số đơn hủy và giá trị đơn hủy xem ở bảng KPI phía trên và sheet "Đơn hoàn & hủy".'
            ).font = font(9, color=GRAY, italic=True)
    ws.cell(note_r + 1, 2,
            f'="Đơn đang xử lý (chưa hoàn thành — chưa tính doanh thu): "'
            f'&TEXT(SUMIFS({rng("cnt_o", n)},{rng("group", n)},"Đang xử lý"),"#,##0")&" đơn, "'
            f'&TEXT(SUMIFS({rng("paid", n)},{rng("group", n)},"Đang xử lý"),"#,##0")&" đ"'
            ).font = font(9, color=GRAY, italic=True)

    # ---- Link sang các sheet khác ----
    r = tot + 4
    ws.cell(r, 2, 'XEM CHI TIẾT').font = font(12, True, NAVY)
    links = [
        (PROD, 'Doanh thu, số lượng bán theo từng mặt hàng và phân loại'),
        (RET,  'Danh sách đơn hoàn trả và đơn hủy (phân loại lý do)'),
        (DATA, 'Toàn bộ dữ liệu đơn hàng gốc (có thể lọc, tra cứu theo mã đơn)'),
        (HOW,  'Giải thích cách tính từng con số'),
    ]
    for k, (sheet, desc) in enumerate(links, 1):
        c = ws.cell(r + k, 2, f'→ {sheet}')
        c.hyperlink = Hyperlink(ref=c.coordinate, location=f"'{sheet}'!A1")
        c.font      = font(10, True, '0563C1', underline='single')
        ws.cell(r + k, 3, desc).font = font(10, color=GRAY)

    r += len(links) + 2
    ws.cell(r, 2,
            'Lưu ý: Doanh thu là số TẠM TÍNH tại lúc xuất file. '
            'Chưa trừ phí sàn, hoa hồng, thuế. '
            'Tiền thực nhận xem tại Seller Center → Tài chính.').font = font(9, True, RED, italic=True)
    setup_print(ws)


# ─────────────────────────────────────────────────────
# Sheet: Cách tính
# ─────────────────────────────────────────────────────
def sheet_howto(wb, src, exported, period):
    ws = wb.create_sheet(HOW)
    ws.sheet_properties.tabColor = '7F7F7F'
    ws.sheet_view.showGridLines  = False
    ws.column_dimensions['A'].width = 2
    ws.column_dimensions['B'].width = 34
    ws.column_dimensions['C'].width = 100
    ws['B1'] = 'CÁCH TÍNH CÁC CON SỐ TRONG BÁO CÁO'
    ws['B1'].font = font(14, True, NAVY)

    sections = [
        ('NGUỒN DỮ LIỆU', [
            ('File gốc',          f'{src.name} — xuất từ Shopee Seller Center lúc {exported:%d/%m/%Y %H:%M}'),
            ('Khoảng thời gian',  period),
            ('Phạm vi',           'Mỗi dòng trong file gốc là 1 đơn hàng (1 hoặc nhiều sản phẩm trong 1 đơn).'),
        ]),
        ('PHÂN LOẠI ĐƠN', [
            ('Bán',    'Trạng thái "Hoàn thành" và không có yêu cầu hoàn trả — đây là doanh thu chính.'),
            ('Đang xử lý', 'Đơn chưa "Hoàn thành" (chờ xác nhận, chờ lấy hàng, đang giao, đã giao…) — CHƯA tính '
                           'doanh thu vì còn có thể bị hủy / hoàn. Xuất lại file sau khi đơn hoàn thành.'),
            ('Hoàn',   'Trạng thái "Hoàn thành" nhưng "Trạng thái Trả hàng/Hoàn tiền" = "Đã Chấp Thuận Yêu Cầu". '
                       'Tiền hoàn được trừ khỏi doanh thu thuần.'),
            ('Hủy',    'Trạng thái "Đã hủy" — không tính vào doanh thu. '
                       'Cột "Giá trị đơn" trong bảng Đơn hoàn & hủy cho biết đơn đó trị giá bao nhiêu.'),
            ('Loại hủy', 'Phân biệt: Người mua hủy / Giao thất bại (Shopee tự hủy do ship lỗi) / '
                         'Shop giao trễ / Hệ thống tự hủy / Khác.'),
        ]),
        ('CÁC CỘT TRONG BẢNG "THEO MẶT HÀNG"', [
            ('Mặt hàng',               'Tên rút gọn: phần trước dấu " – " trong tên sản phẩm. '
                                       'Tên đầy đủ xem ở "Dữ liệu chi tiết".'),
            ('Phân loại',              'Size/biến thể (500ML, 1L, 2L, 4.8L…). '
                                       'Hiển thị "—" nếu sản phẩm không có phân loại.'),
            ('Số đơn',                 'Số đơn Bán chứa combo mặt hàng + phân loại đó. '
                                       'Không đếm đơn Hoàn và Hủy.'),
            ('Tổng SL',                'Tổng số lượng xuất kho của đơn Bán.'),
            ('Đơn hủy',                'Số đơn bị hủy có chứa combo này.'),
            ('Đã hoàn trả (SL)',       'Số lượng sản phẩm khách trả lại (cột SL hoàn từ file gốc + SL đơn Hoàn).'),
            ('Doanh thu (khách trả)',  'Tổng "Tổng số tiền Người mua thanh toán" của đơn Bán và đơn Hoàn '
                                       '(đơn Hoàn: khách đã trả rồi được hoàn lại).'),
            ('Giảm giá shop',          'Tổng phần người bán trợ giá (đơn Bán).'),
            ('Shopee trợ giá',         'Tổng phần Shopee hỗ trợ giảm giá (đơn Bán).'),
            ('Doanh thu thuần',        'Bằng Doanh thu (khách trả) trừ tổng tiền hoàn của đơn Hoàn (= tiền đơn Bán). '
                                       'Đây là con số sát nhất với tiền thực thu (trước phí sàn).'),
            ('Tỷ trọng',              'Phần trăm doanh thu thuần của combo so với tổng doanh thu thuần.'),
        ]),
        ('CÁC TILE TRONG TỔNG QUAN', [
            ('DOANH THU THUẦN',             'Doanh thu (khách trả) trừ tiền hoàn đơn Hoàn — số quan trọng nhất.'),
            ('Doanh thu (khách trả)',        'Tổng tiền khách đã trả cho đơn Bán và đơn Hoàn (chưa trừ hoàn).'),
            ('Tiền hoàn cho khách',          'Tổng tiền hoàn lại cho đơn Hoàn.'),
            ('Giảm giá shop / Shopee bù',    'Tổng phần shop và Shopee trợ giá cho khách trong đơn Bán.'),
            ('Tổng SL xuất (đơn bán)',       'Tổng số chai/gói xuất kho trong đơn Bán.'),
            ('SL hoàn trả về kho',           'Số lượng sản phẩm đã hoàn về kho (cần cộng vào nhập kho QL).'),
            ('Số đơn hủy',                  'Đơn bị hủy — không thu được tiền.'),
            ('Giá trị đơn hủy (chưa thu)',  'Tổng giá trị đơn hủy — đây là "tiếc" bao nhiêu, không phải DT.'),
        ]),
        ('LƯU Ý', [
            ('Số tạm tính',   'Đây là số tại lúc xuất file. Đơn chưa hoàn tất vẫn có thể bị hoàn sau đó.'),
            ('Chưa trừ phí',  'Chưa trừ phí sàn, hoa hồng, thuế. Tiền thực nhận xem ở Seller Center → Tài chính.'),
            ('Nhiều SKU/đơn', 'Một đơn Shopee có thể có nhiều sản phẩm. Mỗi dòng trong file gốc là 1 dòng sản phẩm. '
                              'Công cụ gom đúng khi đếm số đơn.'),
        ]),
    ]

    r = 3
    for head, items in sections:
        ws.cell(r, 2, head).font = font(11, True, 'FFFFFF')
        for i in (2, 3):
            ws.cell(r, i).fill = fill(BLUE)
        ws.row_dimensions[r].height = 22
        r += 1
        for k, (name, text) in enumerate(items):
            a, b = ws.cell(r, 2, name), ws.cell(r, 3, text)
            a.font, b.font = font(10, True, NAVY), font(10)
            a.alignment, b.alignment = LEFT, LEFT
            a.border = b.border = BOX
            if k % 2:
                a.fill = b.fill = fill(BAND)
            ws.row_dimensions[r].height = row_height(text, 100, 14)
            r += 1
        r += 1
    setup_print(ws)
    ws.page_setup.fitToHeight = 1


# ─────────────────────────────────────────────────────
# Hàm tiện ích
# ─────────────────────────────────────────────────────
def export_time(src):
    """Lấy ngày cuối kỳ từ tên file Order.all.YYYYMMDD_YYYYMMDD.xlsx."""
    m = re.search(r'(\d{8})_(\d{8})', src.stem)
    if m:
        try:
            return datetime.strptime(m.group(2), '%Y%m%d')
        except ValueError:
            pass
    return datetime.fromtimestamp(src.stat().st_mtime)


def build_report(rows, src, exported, out):
    created = [r['created'] for r in rows if r['created']]
    period  = (f'Đơn đặt từ {min(created):%d/%m/%Y} đến {max(created):%d/%m/%Y}'
               if created else 'Khoảng thời gian không xác định')
    wb      = Workbook()
    overview = wb.active
    n  = sheet_data(wb, rows)
    tr = sheet_products(wb, rows, n, period)
    sheet_returns(wb, rows, period)
    sheet_overview(overview, rows, n, tr, period, exported)
    sheet_howto(wb, src, exported, period)
    wb.move_sheet(DATA, offset=len(wb.sheetnames))   # dữ liệu chi tiết nằm cuối
    wb.active = 0
    wb.calculation.fullCalcOnLoad = True
    wb.save(out)


def recalc_with_excel(path):
    """Cho Excel tính sẵn công thức rồi lưu, để file xem được trên điện thoại / Zalo."""
    ps  = ('$x = New-Object -ComObject Excel.Application; $x.DisplayAlerts = $false; '
           'try { $w = $x.Workbooks.Open($env:BAO_CAO_PATH); $x.CalculateFull(); $w.Save(); $w.Close($false) } '
           'finally { $x.Quit(); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($x) }')
    res = subprocess.run(['powershell', '-NoProfile', '-NonInteractive', '-Command', ps],
                         env=dict(os.environ, BAO_CAO_PATH=str(path)), capture_output=True)
    return res.returncode == 0


def pick_file():
    import tkinter as tk
    from tkinter import filedialog
    root = tk.Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title='Chọn file đơn hàng xuất từ Shopee',
        filetypes=[('Excel', '*.xlsx')])


# ─────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────
def main():
    sys.stdout.reconfigure(encoding='utf-8')
    if not licensing.ensure_licensed('Báo cáo Shopee'):
        return
    src = sys.argv[1] if len(sys.argv) > 1 else pick_file()
    if not src:
        print('Chưa chọn file.')
        return
    src  = Path(src)
    print(f'Đang đọc: {src.name}')
    data = read_orders(src)
    print(f'  → {len(data)} dòng dữ liệu')
    rows = to_rows(data)

    ban  = sum(1 for r in rows if r['group'] == 'Bán'  and r['cnt_o'])
    hoan = sum(1 for r in rows if r['group'] == 'Hoàn' and r['cnt_o'])
    huy  = sum(1 for r in rows if r['group'] == 'Hủy'  and r['cnt_o'])
    dang = sum(1 for r in rows if r['group'] == 'Đang xử lý' and r['cnt_o'])
    print(f'  → Bán: {ban} đơn | Hoàn: {hoan} đơn | Hủy: {huy} đơn | Đang xử lý (chưa tính DT): {dang} đơn')

    dt_ban  = sum(r['paid'] for r in rows if r['group'] == 'Bán')
    dt_hoan = sum(r['paid'] for r in rows if r['group'] == 'Hoàn')
    print(f'  → DT (khách trả): {dt_ban + dt_hoan:,.0f}đ | Tiền hoàn: {dt_hoan:,.0f}đ | Thuần: {dt_ban:,.0f}đ')

    exported = export_time(src)
    created  = [r['created'] for r in rows if r['created']]
    OUT_DIR.mkdir(exist_ok=True)

    if created:
        out = OUT_DIR / (f'Báo cáo Shopee {min(created):%d.%m.%Y} - {max(created):%d.%m.%Y} '
                         f'(xuất {exported:%d.%m.%Y %Hh%M}).xlsx')
    else:
        out = OUT_DIR / f'Báo cáo Shopee (xuất {exported:%d.%m.%Y %Hh%M}).xlsx'

    print('Đang tạo báo cáo...')
    try:
        build_report(rows, src, exported, out)
    except PermissionError:
        print(f'Không ghi được — file "{out.name}" đang mở trong Excel. Đóng file rồi chạy lại.')
        return

    print('Đang cho Excel tính sẵn số liệu...')
    if not recalc_with_excel(out):
        print('(!) Không tự tính được bằng Excel. Hãy mở file bằng Excel, bấm Lưu rồi mới gửi đi.')

    print()
    print('XONG! Báo cáo đã lưu tại:')
    print(f'  {out}')
    subprocess.Popen(f'explorer /select,"{out}"')
    os.startfile(out)


if __name__ == '__main__':
    main()
