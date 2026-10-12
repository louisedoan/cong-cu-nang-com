# Bộ công cụ Nàng Cơm

| Thư mục | Công cụ | Cách mở |
|---|---|---|
| `DT - DH - CN/Công cụ` | Doanh thu – Đơn hàng – Công nợ NPP (giao diện web cục bộ) | `Mo-cong-cu.bat` |
| `TikTok - Tổng hợp` | Báo cáo doanh thu TikTok + cập nhật tồn kho QL | kéo file vào `Tao-bao-cao.bat` |
| `TikTok - Tổng hợp - V2` | Báo cáo số lượng TikTok (không tiền) + tồn kho QL riêng | `Tao-bao-cao.bat` |
| `TikTok - Tổng hợp - V3` | Số lượng TikTok theo ngày (chốt 14h) | `Tao-bao-cao.bat` |

Hướng dẫn nghiệp vụ từng công cụ nằm trong `HUONG-DAN.txt` của thư mục đó.

## Cài đặt

1. Cài **Python 3.8+** (khi cài nhớ tick *Add python.exe to PATH*). Không cần `pip install` gì: thư viện đi kèm
   trong thư mục `thu-vien` của từng công cụ. Cần Microsoft Excel cho các bước xuất PDF / tính công thức.
2. Lấy mã nguồn:
   ```powershell
   cd D:\
   git clone https://github.com/louisedoan/cong-cu-nang-com.git
   ```
   Cập nhật bản mới sau này: `cd D:\cong-cu-nang-com` rồi `git pull`. Dữ liệu Excel, báo cáo, sao lưu không nằm
   trên git nên `git pull` không đụng tới.

## License (cấp phép) — bắt buộc

Mỗi máy cần **mã kích hoạt** do quản trị viên cấp. Không cần Internet để dùng.

1. Mở công cụ → màn hình **Kích hoạt license** hiện **Mã thiết bị** của máy (dạng `XXXX-XXXX-XXXX-XXXX`).
   Bấm *Sao chép* và gửi cho quản trị viên.
2. Nhận mã bắt đầu bằng `NC1-…` → dán vào ô *Mã kích hoạt* → **Kích hoạt**.
3. Dùng bình thường. Đóng/mở lại không phải nhập lại mã. App web hiện *License còn … giờ … phút* ở góc trái dưới.

Quy tắc:
- Mỗi mã chỉ dùng **một lần**, cho **đúng máy** có Mã thiết bị đó.
- Hiệu lực **48 giờ tính từ lúc quản trị viên tạo mã** (không phải lúc nhập) → nhập ngay khi nhận được.
- Hết hạn: công cụ khóa chức năng cho tới khi nhập mã mới. **Dữ liệu, báo cáo, sao lưu giữ nguyên.**
- Giữ đồng hồ Windows đúng giờ (bật *Đặt giờ tự động*). Lùi đồng hồ sẽ làm license bị khóa.
- Cài lại Windows / đổi máy làm Mã thiết bị thay đổi → gửi Mã thiết bị mới để được cấp lại.
- Quyền riêng tư: Mã thiết bị là giá trị băm một chiều từ định danh cài đặt Windows. Công cụ không thu thập tên,
  tài khoản, nội dung file và không gửi gì qua mạng.

Dòng lệnh (chạy trong thư mục `thu-vien` của công cụ):
```powershell
python -m licensing status          # trạng thái, thời gian còn lại
python -m licensing device-id       # Mã thiết bị
python -m licensing activate NC1-…  # nhập mã
python -m licensing policy          # chính sách license offline
```

| Thông báo | Cách xử lý |
|---|---|
| *Bản cài đặt này chưa được cấu hình khóa xác thực* | Bản công cụ cũ — `git pull` bản mới |
| *Mã kích hoạt không hợp lệ* | Dán thiếu/sai mã — sao chép lại nguyên mã |
| *Mã này được cấp cho thiết bị khác* | Gửi đúng Mã thiết bị (ghi trong thông báo) cho quản trị viên |
| *Mã đã hết hạn* | Mã tạo quá 48 giờ trước — xin mã mới và nhập ngay |
| *Đồng hồ máy đang sớm hơn… / Đồng hồ Windows đang lùi…* | Bật *Đặt giờ tự động* trong Windows rồi mở lại; vẫn lỗi thì liên hệ quản trị viên |
| *License đã lưu không thuộc máy này* | Cài lại Windows / đổi máy — xin cấp lại |

Module license nằm ở `licensing/` và được chép kèm vào `thu-vien/licensing` của mỗi công cụ — không sửa các bản
trong `thu-vien`.
