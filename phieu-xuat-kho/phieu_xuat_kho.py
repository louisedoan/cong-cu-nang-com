# -*- coding: utf-8 -*-
"""Tool phiếu xuất kho: đọc mọi đơn đặt hàng PDF trong thư mục → mỗi đơn 1 phiếu xuất kho (PXK) theo mẫu,
đối chiếu giá với bảng chào giá, tính khối lượng. Nhấp đúp Tao-phieu-xuat-kho.bat (hoặc: python phieu_xuat_kho.py).

Chỉ đọc file trong thư mục này; không gửi gì qua mạng; không sửa file đơn/bảng giá/mẫu/danh mục gốc.
"""
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT / 'thu-vien'))

import licensing  # noqa: E402  cấp phép offline 20 ngày (bản sao trong thu-vien)

from openpyxl import load_workbook  # noqa: E402

OUT_NAME = 'Phiếu xuất kho'
OUT_DIR = ROOT / OUT_NAME
VAT = 1.08          # giá trên đơn ZMART đã gồm VAT 8%; cột đơn giá PXK là chưa VAT
PXK_SHEET = 'PHIEU XUAT KHO'
# Vị trí theo mẫu "PHIEU XUAT KHO": tiêu đề dòng 13, dòng hàng 14.. (mẫu 3 dòng), TỔNG 17, VAT 18–20, chữ ký 23.
_FIRST, _ORIG_N, _ORIG_TONG, _NCOL = 14, 3, 17, 13


def only_digits(s):
    return re.sub(r'\D', '', str(s or ''))


def _norm(s):
    return re.sub(r'\s+', ' ', str(s or '')).strip()


def _key(s):
    return _norm(s).lower()


def _num(v):
    if v is None or v == '':
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(',', ''))
    except ValueError:
        return 0.0


def _money(s):
    """'77,400' -> 77400 (dấu phẩy là phân cách nghìn trên đơn ZMART)."""
    t = str(s or '').replace(',', '').replace('.', '').strip()
    return int(t) if t.isdigit() else 0


# =====================================================================
# file đầu vào
# =====================================================================

def _pxk_sheet(wb):
    """Tên sheet phiếu xuất kho trong workbook (không phân biệt hoa thường) hoặc None."""
    return next((n for n in wb.sheetnames if _key(n) == _key(PXK_SHEET)), None)


def _layout_ok(ws):
    """Đúng bố cục mẫu: tiêu đề bảng ở dòng 13 (STT), dòng TỔNG ở dòng 17 (mẫu 3 dòng hàng).
    PXK đã lập (số dòng khác 3) không đạt -> không bị dùng nhầm làm mẫu."""
    return 'stt' in _key(ws.cell(_FIRST - 1, 1).value) and _key(ws.cell(_ORIG_TONG, 1).value).startswith('tổng')


def _inspect(path):
    """(tên các sheet đã chuẩn hóa, mẫu PXK đúng bố cục?: True/False, None nếu không có sheet PXK)."""
    try:
        wb = load_workbook(path, read_only=True)
    except Exception:
        return [], None
    try:
        sh = _pxk_sheet(wb)
        return [_key(n) for n in wb.sheetnames], (_layout_ok(wb[sh]) if sh else None)
    finally:
        wb.close()


def find_inputs(folder):
    """Tìm trong `folder` (không xét thư mục con): đơn *.pdf, bảng chào giá (sheet 'báo giá'), mẫu PXK (sheet
    'PHIEU XUAT KHO' đúng bố cục mẫu), danh mục tải trọng (tên có 'DANH MỤC').
    Thiếu bảng giá / mẫu, hoặc có hơn 1 bảng giá / mẫu -> RuntimeError (nêu tên file để người dùng chọn)."""
    folder = Path(folder)
    orders = sorted(p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == '.pdf')
    quotes, templates, wrong_layout, weights = [], [], [], None
    for p in sorted(folder.iterdir()):
        if not p.is_file() or p.suffix.lower() not in ('.xlsx', '.xlsm') or p.name.startswith('~$'):
            continue
        if weights is None and 'danh mục' in _key(p.name):
            weights = p
            continue
        names, layout = _inspect(p)
        if any('báo giá' in n or 'bao gia' in n for n in names):
            quotes.append(p)
        elif layout:
            templates.append(p)
        elif layout is False:
            wrong_layout.append(p)
    if not quotes:
        raise RuntimeError('Không thấy BẢNG CHÀO GIÁ (file Excel có sheet "Bảng báo giá") trong thư mục công cụ.')
    if len(quotes) > 1:
        raise RuntimeError('Có nhiều hơn 1 bảng chào giá trong thư mục công cụ, chỉ để lại 1 file: '
                           + ', '.join(p.name for p in quotes))
    if not templates:
        extra = (' (các file sau có sheet "PHIEU XUAT KHO" nhưng không đúng bố cục mẫu 3 dòng hàng: '
                 + ', '.join(p.name for p in wrong_layout) + ')') if wrong_layout else ''
        raise RuntimeError('Không thấy MẪU PHIẾU XUẤT KHO (file Excel có sheet "PHIEU XUAT KHO") trong thư mục công cụ'
                           + extra + '.')
    if len(templates) > 1:
        raise RuntimeError('Có nhiều hơn 1 mẫu phiếu xuất kho trong thư mục công cụ, chỉ để lại 1 file: '
                           + ', '.join(p.name for p in templates))
    return {'orders': orders, 'quote': quotes[0], 'template': templates[0], 'weights': weights}


def _header_cols(ws, needed, max_scan=12):
    """Dòng tiêu đề + cột theo tên (khớp chứa chuỗi). needed: {field: [chuỗi con]}."""
    for r in range(1, max_scan + 1):
        cells = {c: _key(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)}
        pos = {}
        for field, subs in needed.items():
            for sub in subs:
                col = next((c for c, t in cells.items() if t and sub in t and c not in pos.values()), None)
                if col:
                    pos[field] = col
                    break
        if all(f in pos for f in needed):
            return r, pos
    raise RuntimeError(f'Sheet "{ws.title}" thiếu cột tiêu đề cần thiết.')


def load_quote(path):
    """Bảng chào giá -> {barcode: {ten, dvt, quy_cach, gia}}; gia = giá cung cấp (đã gồm VAT 8%)."""
    wb = load_workbook(path, data_only=True)  # cột giá là công thức: đọc giá trị đã tính
    ws = next((s for s in wb.worksheets if 'báo giá' in _key(s.title) or 'bao gia' in _key(s.title)), wb.worksheets[0])
    hr, pos = _header_cols(ws, {'barcode': ['mã vạch'], 'ten': ['tên sản phẩm'], 'dvt': ['đơn vị tính', 'đvt'],
                                'quy_cach': ['quy cách'], 'gia': ['giá cung cấp']})
    out = {}
    for r in range(hr + 1, ws.max_row + 1):
        bc = only_digits(ws.cell(r, pos['barcode']).value)
        if len(bc) < 12:
            continue
        out[bc] = {'ten': _norm(ws.cell(r, pos['ten']).value), 'dvt': _norm(ws.cell(r, pos['dvt']).value).upper(),
                   'quy_cach': int(_num(ws.cell(r, pos['quy_cach']).value)) or 1, 'gia': _num(ws.cell(r, pos['gia']).value)}
    wb.close()
    return out


def load_weights(path):
    """Danh mục tải trọng -> {barcode: {quy_cach, tai_trong}} (tải trọng là kg/thùng của quy cách danh mục)."""
    if not path:
        return {}
    wb = load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    hr, pos = _header_cols(ws, {'barcode': ['barcode', 'mã vạch'], 'quy_cach': ['quy cách'],
                                'tai_trong': ['tải trọng']})
    out = {}
    for r in range(hr + 1, ws.max_row + 1):
        bc = only_digits(ws.cell(r, pos['barcode']).value)
        tt = _num(ws.cell(r, pos['tai_trong']).value)
        if bc and tt > 0:
            out[bc] = {'quy_cach': int(_num(ws.cell(r, pos['quy_cach']).value)) or 1, 'tai_trong': tt}
    wb.close()
    return out


# =====================================================================
# đọc đơn đặt hàng PDF (mẫu ZMART Phú Quốc)
# =====================================================================
# Dòng hàng: "<mã vạch><ĐVT> <SL> <đơn giá> <thuế> <thành tiền>" — khi thuế >= 100.000, pypdf dính "thuế+thành tiền"
# (vd "412,8005,572,800") nên thành tiền tính = SL × đơn giá rồi kiểm phần đuôi.
_LINE = re.compile(r'(893\d{9,10})\s*([A-ZÀ-ỸĐ]{2,5})\s+(\d+)\s+([\d,]+)\s+([\d,\s]+)$')


def pdf_text(path):
    import pypdf
    return '\n'.join((p.extract_text() or '') for p in pypdf.PdfReader(str(path)).pages)


def parse_text(text):
    """Text đơn -> {chain, header, lines, tong, canh_bao}. Không phải mẫu ZMART -> RuntimeError."""
    if 'ZMART' not in text.upper() or 'Mã đơn nhập' not in text:
        raise RuntimeError('Chưa hỗ trợ mẫu đơn này (hiện chỉ đọc được đơn ZMART Phú Quốc).')
    L = [l.strip() for l in text.split('\n') if l.strip()]
    cong_ty = re.search(r'^(CÔNG TY.+?)\s+Mã đơn nhập', text, re.M)
    sieu_thi = next((l for l in L if l.upper().startswith('SIÊU THỊ')), '')
    dia_chi = []
    for i, l in enumerate(L):
        if l.startswith('Địa chỉ kinh doanh'):
            dia_chi.append(l.split(':', 1)[1].strip())
            for nxt in L[i + 1:]:
                if nxt.startswith(('Đơn đặt hàng', 'Nhà cung cấp', 'STT')):
                    break
                dia_chi.append(nxt)
            break
    header = {'cong_ty': _norm(cong_ty.group(1)) if cong_ty else '',
              'mst': (re.search(r'MST\s*:\s*(\d+)', text) or [None, ''])[1],
              'sieu_thi': _norm(sieu_thi),
              'dia_chi': _norm(' '.join(dia_chi)).rstrip('.'),
              'ma_don': (re.search(r'Mã đơn nhập:\s*(\S+)', text) or [None, ''])[1],
              'ngay': (re.search(r'Ngày tạo:\s*([\d\-/]+)', text) or [None, ''])[1]}
    lines, canh_bao = [], []
    for l in L:
        m = _LINE.search(l)
        if not m:
            continue
        bc, dvt, sl, gia, tail = m.group(1), m.group(2), int(m.group(3)), _money(m.group(4)), m.group(5)
        tt = sl * gia
        if not tail.replace(' ', '').endswith(f'{tt:,}'):
            canh_bao.append(f'Dòng {bc}: thành tiền trên đơn không bằng SL × đơn giá ({tt:,}) — kiểm tay.')
        lines.append({'barcode': bc, 'dvt': dvt, 'sl': sl, 'gia': gia, 'thanh_tien': tt})
    # 5 số tổng (SL, tổng tiền hàng, chiết khấu, thuế, giá trị đơn) đứng liền nhau, nhãn nằm cuối trang
    tong = {}
    for i in range(len(L) - 4):
        block = L[i:i + 5]
        if all(re.fullmatch(r'[\d,]+', x) for x in block):
            v = [_money(x) for x in block]
            tong = {'so_luong': v[0], 'tien_hang': v[1], 'chiet_khau': v[2], 'thue': v[3], 'gia_tri': v[4]}
            break
    so_stt = len(re.findall(r'^\d{1,3} \d{6}\b', text, re.M))  # dòng "STT SKU"
    return {'chain': 'zmart', 'header': header, 'lines': lines, 'tong': tong, 'so_stt': so_stt, 'canh_bao': canh_bao}


def parse_order(path):
    return parse_text(pdf_text(path))


def check_order(order):
    """Cảnh báo khiến đơn KHÔNG được tạo PXK tự động (đọc thiếu, lệch tổng, có chiết khấu)."""
    w = list(order.get('canh_bao', []))
    t = order.get('tong') or {}
    s_tt = sum(l['thanh_tien'] for l in order['lines'])
    s_sl = sum(l['sl'] for l in order['lines'])
    if not order['lines']:
        w.append('Không đọc được dòng hàng nào.')
    if not t:
        w.append('Không đọc được phần tổng của đơn.')
    else:
        if s_tt != t.get('tien_hang'):
            w.append(f'Tổng thành tiền các dòng ({s_tt:,}) khác "Tổng tiền hàng" trên đơn ({t.get("tien_hang", 0):,}).')
        if s_sl != t.get('so_luong'):
            w.append(f'Tổng số lượng các dòng ({s_sl:,}) khác "Số lượng" trên đơn ({t.get("so_luong", 0):,}).')
        if t.get('chiet_khau'):
            w.append(f'Đơn có chiết khấu {t["chiet_khau"]:,} — chưa hỗ trợ, lập phiếu tay.')
        if s_tt != t.get('gia_tri'):
            w.append(f'Tổng thành tiền các dòng ({s_tt:,}) khác "Giá trị đơn" trên đơn ({t.get("gia_tri", 0):,}).')
        thue_8 = round(t.get('tien_hang', 0) * 8 / 108)  # giá đã gồm VAT 8%: thuế = tiền hàng × 8/108
        if abs(t.get('thue', 0) - thue_8) > max(2, len(order['lines'])):  # chênh do làm tròn từng dòng
            w.append(f'Thuế trên đơn ({t.get("thue", 0):,}) không khớp VAT 8% ({thue_8:,}) — phiếu tính VAT 8%, '
                     'lập phiếu tay.')
    if order.get('so_stt') and order['so_stt'] != len(order['lines']):
        w.append(f'Đơn có {order["so_stt"]} dòng nhưng chỉ đọc được {len(order["lines"])} dòng.')
    return w


# =====================================================================
# đối chiếu bảng chào giá + tính dòng PXK
# =====================================================================

def build_rows(order, quote, weights):
    """Mỗi dòng đơn -> 1 dòng PXK (không bỏ dòng). Trả (rows, issues)."""
    h = order['header']
    rows, issues = [], []

    def issue(bc, loai, chi_tiet):
        issues.append({'ma_don': h.get('ma_don'), 'sieu_thi': h.get('sieu_thi'), 'barcode': bc, 'loai': loai,
                       'chi_tiet': chi_tiet})

    for ln in order['lines']:
        bc, notes = ln['barcode'], []
        q, w = quote.get(bc), weights.get(bc)
        if q:
            ten, dvt, quy_cach = q['ten'], q['dvt'] or ln['dvt'], q['quy_cach']
            if abs(ln['gia'] - q['gia']) > 0.5:
                txt = f'Giá đơn {ln["gia"]:,.0f} ≠ báo giá {q["gia"]:,.0f}'
                notes.append(txt)
                issue(bc, 'Lệch giá', txt)
        else:
            ten, dvt, quy_cach = f'Mã {bc}', ln['dvt'], (w['quy_cach'] if w else 1)
            notes.append('Không có trong bảng chào giá')
            issue(bc, 'Không có trong bảng chào giá', f'SL {ln["sl"]}, giá đơn {ln["gia"]:,.0f}')
        if w:
            kg = ln['sl'] * w['tai_trong'] / w['quy_cach']  # theo từng chai: kg/thùng ÷ chai/thùng (danh mục)
        else:
            kg = 0
            notes.append('Chưa có tải trọng')
            issue(bc, 'Chưa có tải trọng', 'Bổ sung vào DANH MỤC HH KÊNH MT rồi chạy lại')
        rows.append({'barcode': bc, 'ten': ten, 'dvt': dvt, 'quy_cach': quy_cach, 'sl': ln['sl'],
                     'so_thung': ln['sl'] / quy_cach, 'gia': ln['gia'], 'don_gia_1pcs': ln['gia'] / VAT,
                     'thanh_tien': ln['thanh_tien'], 'kg': kg, 'ghi_chu': '; '.join(notes)})
    return rows, issues


# =====================================================================
# số phiếu: PXK/KHO-PQ<mmyy>/<n>, đếm lại từ 1 mỗi tháng; nhớ ở file "mmyy n" (sửa tay được)
# =====================================================================
COUNTER_FILE = ROOT / 'so-phieu-pxk.txt'
KHO_CODE, KHO_XUAT, XUAT_TAI = 'KHO-PQ', 'KHO-PQ', 'KHO PHÚ QUỐC'


def _max_issued(out_dir, ngay):
    """Số PXK lớn nhất đã lập trong tháng của `ngay` (theo tên file dd.mm-PXK<n>-... và năm sửa file)."""
    from datetime import datetime
    best = 0
    if out_dir and Path(out_dir).is_dir():
        for p in Path(out_dir).glob('*.xlsx'):
            m = re.match(r'\d{2}\.(\d{2})-PXK(\d+)-', p.name)
            if m and m.group(1) == f'{ngay:%m}' and datetime.fromtimestamp(p.stat().st_mtime).year == ngay.year:
                best = max(best, int(m.group(2)))
    return best


def next_number(path, ngay, out_dir=None):
    """Số phiếu kế tiếp trong tháng: theo file đếm, nhưng không thấp hơn số PXK đã có trong thư mục kết quả
    (mất/sửa sai file đếm cũng không cấp trùng số)."""
    n = 0
    try:
        mmyy, last = Path(path).read_text('utf-8').split()[:2]
        if mmyy == f'{ngay:%m%y}':
            n = int(last)
    except (OSError, ValueError):
        pass
    return max(n, _max_issued(out_dir, ngay)) + 1


def save_number(path, ngay, n):
    Path(path).write_text(f'{ngay:%m%y} {n}\n', encoding='utf-8')


# =====================================================================
# điền phiếu xuất kho (bản sao mẫu; dựng lại khối TỔNG/VAT/chữ ký theo số dòng)
# =====================================================================
def _safe(s):
    return re.sub(r'[\\/:*?"<>|]+', '_', str(s or '')).strip() or 'X'


def _short_store(sieu_thi):
    return re.sub(r'(?i)^siêu thị\s*', '', sieu_thi or '').strip() or 'KH'


def pxk_name(order, n, ngay):
    h = order['header']
    return f'{ngay:%d.%m}-PXK{n}-{_safe(_short_store(h.get("sieu_thi")))}-{_safe(h.get("ma_don"))}.xlsx'


def fill_pxk(order, rows, template, out_dir, n, ngay):
    """Điền 1 đơn vào bản sao mẫu PXK. Trả đường dẫn file. Không sửa file mẫu.
    Lỗi giữa chừng -> không để lại file PXK (lần chạy sau lập lại đơn này)."""
    import os
    template = Path(template)
    if not template.exists():
        raise RuntimeError(f'Không tìm thấy mẫu phiếu xuất kho: {template}')
    if not rows:
        raise RuntimeError('Đơn không có dòng hàng.')
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    out = Path(out_dir) / pxk_name(order, n, ngay)
    tmp = out.with_name(out.name + '.tmp')  # đuôi .tmp: không bị nhận là PXK đã lập
    wb = load_workbook(template)            # lưu ra file khác -> mẫu gốc không đổi
    try:
        sh = _pxk_sheet(wb)
        if not sh:
            raise RuntimeError(f'File mẫu thiếu sheet "{PXK_SHEET}".')
        if not _layout_ok(wb[sh]):
            raise RuntimeError('File mẫu không đúng bố cục (tiêu đề dòng 13, dòng TỔNG ở dòng 17).')
        _fill_sheet(wb[sh], order, rows, n, ngay)
        wb.save(tmp)
        os.replace(tmp, out)
    finally:
        wb.close()
        if tmp.exists():
            tmp.unlink()
    return out


def _fill_sheet(ws, order, rows, n, ngay):
    from copy import copy
    from openpyxl.utils import get_column_letter as col
    from openpyxl.utils.cell import range_boundaries
    N, delta = len(rows), len(rows) - _ORIG_N
    max_row = ws.max_row
    data_style = {c: copy(ws.cell(_FIRST, c)._style) for c in range(1, _NCOL + 1)}
    data_height = ws.row_dimensions[_FIRST].height
    tail = [(r, c, ws.cell(r, c).value, copy(ws.cell(r, c)._style), ws.row_dimensions[r].height)
            for r in range(_ORIG_TONG, max_row + 1) for c in range(1, _NCOL + 1)]
    merges = [(m.min_col, m.min_row, m.max_col, m.max_row) for m in ws.merged_cells.ranges]
    for m in list(ws.merged_cells.ranges):          # gỡ ô gộp từ dòng 14 trở xuống (dựng lại theo vị trí mới)
        if m.min_row >= _FIRST:
            ws.unmerge_cells(str(m))
    for r in range(_FIRST, max_row + max(delta, 0) + 1):
        for c in range(1, _NCOL + 1):
            ws.cell(r, c).value = None
    for r, c, val, style, height in tail:            # khối TỔNG/VAT/chữ ký dời theo số dòng (giữ chữ + định dạng)
        cell = ws.cell(r + delta, c)
        cell.value, cell._style = val, style
        if c == 1:
            ws.row_dimensions[r + delta].height = height
    for i, r in enumerate(rows):                     # dòng hàng
        rr = _FIRST + i
        vals = {1: i + 1, 2: KHO_XUAT, 4: r['barcode'], 5: r['ten'], 6: r['dvt'], 7: r['quy_cach'],
                8: r['so_thung'], 9: f'=H{rr}*G{rr}', 10: r['don_gia_1pcs'], 11: f'=J{rr}*G{rr}',
                12: f'=H{rr}*K{rr}', 13: r['ghi_chu'] or None}
        ws.row_dimensions[rr].height = data_height  # dòng hàng cao đều như dòng mẫu
        for c in range(1, _NCOL + 1):
            cell = ws.cell(rr, c)
            cell.value, cell._style = vals.get(c), copy(data_style[c])
        if r['so_thung'] != int(r['so_thung']):
            ws.cell(rr, 8).number_format = 'General'  # số thùng lẻ: hiện đủ (vd 1.5)
    last, t = _FIRST + N - 1, _FIRST + N
    ws.cell(t, 8).value = f'=SUM(H14:H{last})'
    ws.cell(t, 9).value = f'=SUM(I14:I{last})'
    ws.cell(t + 1, 13).value = f'=SUM(L14:L{last})'
    ws.cell(t + 2, 12).value = 0.08
    ws.cell(t + 2, 13).value = f'=M{t + 1}*L{t + 2}'
    ws.cell(t + 3, 13).value = f'=SUM(M{t + 1}:M{t + 2})'
    for c1, r1, c2, r2 in merges:
        if r1 >= _FIRST:
            ws.merge_cells(f'{col(c1)}{r1 + delta}:{col(c2)}{r2 + delta}')
    if ws.auto_filter.ref:  # vùng lọc của mẫu (A13:M20) giãn theo số dòng
        c1, r1, c2, r2 = range_boundaries(ws.auto_filter.ref)
        ws.auto_filter.ref = f'{col(c1)}{r1}:{col(c2)}{r2 + delta if r2 >= _ORIG_TONG else r2}'
    h = order['header']
    kg = sum(r['kg'] for r in rows)
    ws['A5'] = f'Ngày   {ngay:%d}  tháng   {ngay:%m}  năm {ngay:%Y}'
    ws['A7'] = f'Đơn vị nhận hàng: {h.get("cong_ty", "")}' + (f' - MST: {h["mst"]}' if h.get('mst') else '')
    ws['A8'] = f'Người nhận hàng: {h.get("sieu_thi", "")}'
    ws['A9'] = f'Địa chỉ nhận hàng: {h.get("dia_chi", "")}'
    ws['A10'] = 'Lý do xuất: Xuất kho bán hàng'
    ws['A11'] = f'Xuất tại: {XUAT_TAI}'
    ws['I1'] = f'{h.get("ma_don", "")}\nSố: PXK/{KHO_CODE}{ngay:%m%y}/{n}'
    ws['I9'] = f'Đơn hàng(PO) số:{h.get("ma_don", "")}'
    ws['L11'] = f'Khối lượng:  {round(kg, 1):g} Kg'


# =====================================================================
# file tổng hợp đối chiếu
# =====================================================================

def write_summary(results, out_dir, ngay):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    Path(out_dir).mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = 'Tổng hợp'
    head = ['Siêu thị', 'Mã đơn', 'File đơn', 'Trạng thái', 'Số dòng', 'Tổng tiền đơn (đã VAT)', 'Khối lượng (Kg)',
            'Lệch giá', 'Chưa có trong bảng giá', 'Chưa có tải trọng', 'File PXK']
    ws.append(head)
    for r in results:
        iss = r.get('issues', [])
        cnt = lambda k: sum(1 for i in iss if i['loai'] == k)  # noqa: E731
        h = r.get('header') or {}
        ws.append([h.get('sieu_thi'), h.get('ma_don'), r['file'], r['status'], r.get('n_lines'), r.get('tong_don'),
                   round(r.get('kg') or 0, 1), cnt('Lệch giá'), cnt('Không có trong bảng chào giá'),
                   cnt('Chưa có tải trọng'), Path(r['pxk']).name if r.get('pxk') else ''])
    ws2 = wb.create_sheet('Chi tiết cần kiểm')
    ws2.append(['Siêu thị', 'Mã đơn', 'Mã vạch', 'Loại', 'Chi tiết'])
    for r in results:
        for i in r.get('issues', []):
            ws2.append([i.get('sieu_thi'), i.get('ma_don'), i.get('barcode'), i['loai'], i['chi_tiet']])
    for sh in (ws, ws2):
        for c in sh[1]:
            c.font = Font(bold=True, color='FFFFFF')
            c.fill = PatternFill('solid', fgColor='1F4E78')
            c.alignment = Alignment(wrap_text=True, vertical='center')
        for i, w in enumerate([18, 14, 22, 30, 9, 18, 14, 9, 14, 14, 40][:sh.max_column], 1):
            sh.column_dimensions[chr(64 + i)].width = w
    for row in ws.iter_rows(min_row=2, min_col=6, max_col=6):
        for c in row:
            c.number_format = '#,##0'
    out = Path(out_dir) / f'Tong hop doi chieu {ngay:%d.%m.%Y %Hh%M}.xlsx'
    wb.save(out)
    return out


def run(folder, ngay=None, counter=None):
    """Chạy cả loạt cho thư mục: mỗi đơn PDF -> 1 PXK (đơn đã có PXK thì bỏ qua), rồi file tổng hợp.
    Trả (results, summary_path). Thiếu bảng giá / mẫu -> RuntimeError."""
    from datetime import datetime
    folder = Path(folder)
    ngay = ngay or datetime.now()
    counter = Path(counter) if counter else folder / COUNTER_FILE.name
    out_dir = folder / OUT_NAME
    inp = find_inputs(folder)
    quote, weights = load_quote(inp['quote']), load_weights(inp['weights'])
    results = []
    for pdf in inp['orders']:
        res = {'file': pdf.name, 'header': {}, 'issues': [], 'pxk': None, 'n_lines': 0, 'tong_don': 0, 'kg': 0}
        results.append(res)
        try:
            order = parse_order(pdf)
        except Exception as e:  # PDF lỗi / mẫu đơn chưa hỗ trợ
            res['status'] = f'Không tạo PXK: {e}'
            continue
        h = res['header'] = order['header']
        rows, issues = build_rows(order, quote, weights)
        res.update(n_lines=len(rows), issues=issues, kg=sum(r['kg'] for r in rows),
                   tong_don=(order.get('tong') or {}).get('gia_tri') or sum(r['thanh_tien'] for r in rows))
        warn = check_order(order)
        if warn:
            res['status'] = 'Không tạo PXK: ' + ' '.join(warn)
            res['issues'] = issues + [{'ma_don': h.get('ma_don'), 'sieu_thi': h.get('sieu_thi'), 'barcode': '',
                                       'loai': 'Đọc đơn', 'chi_tiet': w} for w in warn]
            continue
        da_co = sorted(out_dir.glob(f'*-{_safe(h.get("ma_don"))}.xlsx')) if out_dir.exists() else []
        if da_co:
            res.update(status='Bỏ qua: đã có phiếu xuất kho của đơn này', pxk=str(da_co[0]))
            continue
        n = next_number(counter, ngay, out_dir)
        try:
            res['pxk'] = str(fill_pxk(order, rows, inp['template'], out_dir, n, ngay))
        except PermissionError:
            res['status'] = 'Không tạo PXK: file đang mở trong Excel — đóng file rồi chạy lại.'
            continue
        except Exception as e:  # lỗi 1 đơn không làm dừng cả loạt; không cấp số
            res['status'] = f'Không tạo PXK: lỗi khi lập phiếu ({e})'
            continue
        save_number(counter, ngay, n)
        res['status'] = 'Đã tạo PXK'
    return results, write_summary(results, out_dir, ngay)


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    if not licensing.ensure_licensed('Phiếu xuất kho'):
        return
    print('Đang đọc các đơn đặt hàng PDF trong thư mục...')
    try:
        results, summary = run(ROOT)
    except RuntimeError as e:
        print('(!)', e)
        return
    if not results:
        print('(!) Không có file đơn PDF nào trong thư mục công cụ.')
    for r in results:
        h = r['header']
        print(f'\n• {r["file"]} — {h.get("sieu_thi", "")} {h.get("ma_don", "")}: {r["status"]}')
        if r.get('pxk'):
            print(f'  {r["n_lines"]} dòng · {r["tong_don"]:,.0f}đ · {r["kg"]:,.1f} Kg → {Path(r["pxk"]).name}')
        for i in r['issues']:
            print(f'  (!) {i["loai"]}: {i["barcode"]} {i["chi_tiet"]}')
    print(f'\nTổng hợp đối chiếu: {summary}')
    print(f'Phiếu xuất kho nằm trong thư mục "{OUT_NAME}". Kiểm các dòng có ghi chú trước khi in.')


if __name__ == '__main__':
    main()
