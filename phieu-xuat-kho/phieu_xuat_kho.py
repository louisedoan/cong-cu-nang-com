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

def _sheet_names(path):
    try:
        wb = load_workbook(path, read_only=True)
        names = list(wb.sheetnames)
        wb.close()
        return names
    except Exception:
        return []


def find_inputs(folder):
    """Tìm trong `folder` (không xét thư mục con): đơn *.pdf, bảng chào giá (sheet 'báo giá'), mẫu PXK (sheet
    'PHIEU XUAT KHO'), danh mục tải trọng (tên có 'DANH MỤC'). Thiếu bảng giá / mẫu -> RuntimeError."""
    folder = Path(folder)
    orders = sorted(p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == '.pdf')
    quote = template = weights = None
    for p in sorted(folder.iterdir()):
        if not p.is_file() or p.suffix.lower() not in ('.xlsx', '.xlsm') or p.name.startswith('~$'):
            continue
        names = [_key(n) for n in _sheet_names(p)]
        if weights is None and 'danh mục' in _key(p.name):
            weights = p
        elif quote is None and any('báo giá' in n or 'bao gia' in n for n in names):
            quote = p
        elif template is None and _key(PXK_SHEET) in names:
            template = p
    if quote is None:
        raise RuntimeError('Không thấy BẢNG CHÀO GIÁ (file Excel có sheet "Bảng báo giá") trong thư mục công cụ.')
    if template is None:
        raise RuntimeError('Không thấy MẪU PHIẾU XUẤT KHO (file Excel có sheet "PHIEU XUAT KHO") trong thư mục công cụ.')
    return {'orders': orders, 'quote': quote, 'template': template, 'weights': weights}


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


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    if not licensing.ensure_licensed('Phiếu xuất kho'):
        return
    find_inputs(ROOT)


if __name__ == '__main__':
    main()
