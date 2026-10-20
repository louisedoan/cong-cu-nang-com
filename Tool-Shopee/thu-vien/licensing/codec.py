"""Định dạng mã kích hoạt (Activation Code) — dùng chung cho máy người dùng và máy quản trị.

Mã = "NC1-" + Base32( payload 28 byte + chữ ký Ed25519 64 byte ), chia nhóm 4 ký tự.
Payload (big-endian): version(1) product(1) license_id(8) device_id(10) issued_at(4) expires_at(4)
Thời gian là số giây UTC (epoch). Thời hạn và thiết bị nằm trong phần được ký nên không sửa được.
"""
import base64
import hashlib
import struct
from dataclasses import dataclass

from . import ed25519

PREFIX = 'NC1'
VERSION = 1
PRODUCT_NANGCOM = 1
LICENSE_SECONDS = 48 * 3600
DEVICE_BYTES = 10
_PAYLOAD = struct.Struct('>BB8s10sII')
PAYLOAD_SIZE = _PAYLOAD.size  # 28
SIGNATURE_SIZE = 64
CODE_CHARS = 148  # độ dài Base32 của 92 byte (bỏ '=')
_CONFUSABLE = str.maketrans({'0': 'O', '1': 'I', '8': 'B'})


class CodeFormatError(ValueError):
    """Chuỗi nhập vào không đúng định dạng mã kích hoạt."""


def _b32(data):
    return base64.b32encode(data).decode('ascii').rstrip('=')


def _unb32(text):
    return base64.b32decode(text + '=' * (-len(text) % 8))


def group(text, size=4):
    return '-'.join(text[i:i + size] for i in range(0, len(text), size))


# ---------------- Device ID ----------------

def format_device_id(raw):
    if len(raw) != DEVICE_BYTES:
        raise ValueError('Device ID phải dài 10 byte.')
    return group(_b32(raw))


def parse_device_id(text):
    """'abcd-efgh-...' (không phân biệt hoa thường, bỏ khoảng trắng/gạch) -> 10 byte."""
    s = ''.join(str(text or '').split()).replace('-', '').upper().translate(_CONFUSABLE)
    if len(s) != 16:
        raise ValueError('Mã thiết bị phải gồm 16 ký tự (dạng XXXX-XXXX-XXXX-XXXX).')
    try:
        raw = _unb32(s)
    except (ValueError, base64.binascii.Error):
        raise ValueError('Mã thiết bị chứa ký tự không hợp lệ (chỉ gồm A–Z và 2–7).') from None
    if len(raw) != DEVICE_BYTES or _b32(raw) != s:
        raise ValueError('Mã thiết bị không hợp lệ.')
    return raw


# ---------------- Activation Code ----------------

@dataclass(frozen=True)
class LicenseToken:
    version: int
    product: int
    license_id: bytes
    device_id: bytes
    issued_at: int
    expires_at: int
    payload: bytes
    signature: bytes

    @property
    def license_id_hex(self):
        return self.license_id.hex().upper()

    @property
    def device_id_text(self):
        return format_device_id(self.device_id)

    @property
    def code(self):
        return encode_code(self.payload, self.signature)

    @property
    def code_hash(self):
        return code_hash(self.payload + self.signature)

    def verify(self, public_key):
        return ed25519.verify(public_key, self.payload, self.signature)


def pack_payload(license_id, device_id, issued_at, expires_at, product=PRODUCT_NANGCOM, version=VERSION):
    return _PAYLOAD.pack(version, product, license_id, device_id, int(issued_at), int(expires_at))


def encode_code(payload, signature):
    return f'{PREFIX}-' + group(_b32(payload + signature))


def code_hash(raw):
    return hashlib.sha256(raw).hexdigest()


def decode_code(text):
    """Tách mã thành LicenseToken (CHƯA xác minh chữ ký). Sai định dạng -> CodeFormatError."""
    s = ''.join(str(text or '').split()).replace('-', '').upper()
    if s.startswith(PREFIX):
        s = s[len(PREFIX):]
    s = s.translate(_CONFUSABLE)
    if not s:
        raise CodeFormatError('Chưa nhập mã kích hoạt.')
    if len(s) != CODE_CHARS:
        raise CodeFormatError(f'Mã kích hoạt phải có {CODE_CHARS} ký tự (không kể "NC1-" và dấu gạch), '
                              f'mã vừa nhập có {len(s)}. Hãy sao chép đủ cả mã.')
    try:
        raw = _unb32(s)
    except (ValueError, base64.binascii.Error):
        raise CodeFormatError('Mã kích hoạt chứa ký tự không hợp lệ (chỉ gồm A–Z và 2–7).') from None
    if len(raw) != PAYLOAD_SIZE + SIGNATURE_SIZE or _b32(raw) != s:
        raise CodeFormatError('Mã kích hoạt không đúng định dạng.')
    payload, signature = raw[:PAYLOAD_SIZE], raw[PAYLOAD_SIZE:]
    version, product, lid, dev, issued, expires = _PAYLOAD.unpack(payload)
    return LicenseToken(version, product, lid, dev, issued, expires, payload, signature)
