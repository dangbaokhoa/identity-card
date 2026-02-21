---
title: Vietnamese ID Card OCR
emoji: 🪪
colorFrom: blue
colorTo: green
sdk: streamlit
sdk_version: "1.40.0"
app_file: app.py
pinned: false
license: apache-2.0
---

# Vietnamese ID Card OCR

Ứng dụng trích xuất thông tin từ Căn cước công dân Việt Nam sử dụng OCR (EasyOCR).

## Tính năng

- Trích xuất tự động các trường: Họ tên, Số CCCD, Ngày sinh, Giới tính, Quốc tịch, Quê quán, Nơi thường trú, Có giá trị đến
- Xử lý nhiều ảnh cùng lúc (batch processing)
- Tải mẫu Word và tự động điền thông tin
- Tải kết quả dưới dạng file ZIP

## Cách sử dụng

1. Tải lên file mẫu Word (.docx) với các placeholder: `{{ ho_va_ten }}`, `{{ so }}`, v.v.
2. Tải lên một hoặc nhiều ảnh CCCD
3. Nhấn "Trích xuất thông tin"
4. Xem và chỉnh sửa kết quả
5. Tạo và tải file kết quả

## Công nghệ

- EasyOCR: Vietnamese + English text recognition
- OpenCV: Image preprocessing
- Streamlit: Web interface
- python-docx: Word document generation
