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

OUT_DIR = ROOT / 'Phiếu xuất kho'


def find_inputs(folder):
    """Tìm đơn PDF, bảng chào giá, mẫu PXK, danh mục tải trọng trong `folder` (Task 2)."""
    raise NotImplementedError


def main():
    sys.stdout.reconfigure(encoding='utf-8')
    if not licensing.ensure_licensed('Phiếu xuất kho'):
        return
    find_inputs(ROOT)


if __name__ == '__main__':
    main()
