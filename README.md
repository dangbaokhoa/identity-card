---
title: Vietnamese ID Card QR Reader
emoji: 🪪
colorFrom: blue
colorTo: green
sdk: streamlit
sdk_version: "1.40.0"
app_file: app.py
pinned: false
license: apache-2.0
---

# Vietnamese ID Card QR Reader

Ứng dụng đọc thông tin từ mã QR trên Căn cước công dân Việt Nam.

## Tính năng

- Đọc mã QR từ mặt sau thẻ CCCD (nhanh và chính xác 100%)
- Trích xuất tự động: Số CCCD, Số CMND cũ, Họ tên, Ngày sinh, Giới tính, Địa chỉ thường trú, Ngày cấp
- Xử lý nhiều ảnh cùng lúc (batch processing)
- Tải mẫu Word và tự động điền thông tin
- Tải kết quả dưới dạng file ZIP
- Đối soát số TKHQ từ file Excel: đọc cột B từ dòng `Số TKHQ hàng hóa nhập khẩu đã thông quan` đến `Tổng cộng` và rà soát trùng với file ngày trước

## Cách sử dụng

1. Tải lên file mẫu Word (.docx) với các placeholder: `{{ ho_va_ten }}`, `{{ so }}`, v.v.
2. **Chụp/tải ảnh MẶT SAU CCCD** (mặt có mã QR)
3. Nhấn "Đọc mã QR"
4. Xem và chỉnh sửa kết quả
5. Tạo và tải file kết quả

## Lưu ý quan trọng

⚠️ **Phải chụp mặt SAU** (mặt có mã QR) của thẻ CCCD, không phải mặt trước!

## Công nghệ

- pyzbar: QR code detection and decoding
- OpenCV: Image processing
- Streamlit: Web interface
- python-docx: Word document generation
