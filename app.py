import os
import tempfile
import zipfile
from io import BytesIO

os.environ.setdefault("STREAMLIT_SERVER_FILE_WATCHER_TYPE", "none")
os.environ.setdefault("STREAMLIT_SERVER_RUN_ON_SAVE", "false")

import streamlit as st
from docxtpl import DocxTemplate


st.set_page_config(page_title="ID Card QR Reader", page_icon="🪪", layout="centered")
st.title("🪪 Đọc thông tin CCCD từ mã QR")
st.caption("Tải mẫu Word + ảnh mặt SAU (có QR code) của CCCD, xem kết quả, sau đó tải tất cả file kết quả.")


@st.cache_resource
def get_qr_reader():
    print("[APP] Loading QR Reader (cached resource)...")
    from id_card_ocr import IDCardQRReader
    reader = IDCardQRReader()
    print("[APP] ✓ QR Reader initialized")
    return reader


def generate_docx_from_template(data: dict, template_bytes: bytes) -> bytes:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_template:
        temp_template.write(template_bytes)
        temp_template_path = temp_template.name

    temp_output_path = temp_template_path.replace(".docx", "_output.docx")

    try:
        doc = DocxTemplate(temp_template_path)
        doc.render(data)
        doc.save(temp_output_path)

        with open(temp_output_path, "rb") as file:
            output_bytes = file.read()

        return output_bytes
    finally:
        for path in [temp_template_path, temp_output_path]:
            if os.path.exists(path):
                os.remove(path)


def run_qr_on_upload(uploaded_file):
    print(f"[APP] Processing uploaded file: {uploaded_file.name}")
    suffix = os.path.splitext(uploaded_file.name)[1].lower() or ".jpg"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_image:
        temp_image.write(uploaded_file.getbuffer())
        temp_image_path = temp_image.name

    try:
        print("[APP] Getting QR reader...")
        qr_reader = get_qr_reader()
        print("[APP] Running QR reading...")
        data = qr_reader.process_image(temp_image_path)
        print(f"[APP] ✓ QR reading complete for {uploaded_file.name}")
        return data
    except Exception as e:
        print(f"[APP] ✗ Error processing {uploaded_file.name}: {e}")
        raise
    finally:
        if os.path.exists(temp_image_path):
            os.remove(temp_image_path)


def apply_template_aliases(data: dict) -> dict:
    data["full_name"] = data.get("fullname", "")
    data["id_number"] = data.get("no", "")
    data["dob"] = data.get("date_of_birth", "")
    data["gender"] = data.get("sex", "")
    data["ho_va_ten"] = data.get("fullname", "")
    data["so"] = data.get("no", "")
    data["ngay_sinh"] = data.get("date_of_birth", "")
    data["gioi_tinh"] = data.get("sex", "")
    data["quoc_tich"] = data.get("nationality", "")
    data["que_quan"] = data.get("place_of_origin", "")
    data["noi_thuong_tru"] = data.get("residence", "")
    data["co_gia_tri_den"] = data.get("expiry_date", "")
    return data


def safe_output_name(filename: str) -> str:
    base, _ = os.path.splitext(filename)
    cleaned = "".join(ch if ch.isalnum() or ch in ("-", "_") else "_" for ch in base)
    return cleaned or "result"


if "batch_results" not in st.session_state:
    st.session_state["batch_results"] = []

st.subheader("1) Tải mẫu Word")
uploaded_template = st.file_uploader(
    "Tải file mẫu .docx (bắt buộc)",
    type=["docx"],
    key="template_required",
)

st.subheader("2) Tải ảnh mặt SAU CCCD (có mã QR)")
uploaded_images = st.file_uploader(
    "Tải lên một hoặc nhiều ảnh mặt SAU CCCD (có QR code)",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True,
    key="batch_images",
)

can_extract = uploaded_template is not None and uploaded_images
if st.button("Đọc mã QR", type="primary", disabled=not can_extract):
    print(f"[APP] Starting batch QR reading for {len(uploaded_images)} images...")
    with st.spinner("Đang đọc mã QR cho các ảnh..."):
        results = []
        for idx, image_file in enumerate(uploaded_images):
            try:
                print(f"[APP] Processing image {idx+1}/{len(uploaded_images)}: {image_file.name}")
                extracted = run_qr_on_upload(image_file)
                extracted = apply_template_aliases(extracted)
                results.append({
                    "image_name": image_file.name,
                    "data": extracted,
                })
            except Exception as error:
                print(f"[APP] ✗ Failed to process {image_file.name}: {error}")
                st.error(f"Không thể xử lý ảnh {image_file.name}: {error}")
        st.session_state["batch_results"] = results
        print(f"[APP] ✓ Batch QR reading complete: {len(results)} successful")

if uploaded_template is None:
    st.info("Vui lòng tải mẫu Word để tiếp tục.")
elif not uploaded_images:
    st.info("Vui lòng tải lên ít nhất một ảnh mặt SAU CCCD (có QR code) để tiếp tục.")

if st.session_state["batch_results"]:
    st.subheader("3) Xem và chỉnh kết quả")
    st.caption("Bạn có thể chỉnh sửa từng trường trước khi tạo file kết quả.")

    for idx, item in enumerate(st.session_state["batch_results"]):
        image_name = item["image_name"]
        data = item["data"]
        key_prefix = f"card_{idx}"

        with st.expander(f"Ảnh {idx + 1}: {image_name}", expanded=(idx == 0)):
            field_col_1, field_col_2 = st.columns(2)
            with field_col_1:
                data["no"] = st.text_input("Số CCCD", value=data.get("no", ""), key=f"{key_prefix}_no")
                data["old_id"] = st.text_input("Số CMND cũ", value=data.get("old_id", ""), key=f"{key_prefix}_old_id")
                data["fullname"] = st.text_input("Họ và tên", value=data.get("fullname", ""), key=f"{key_prefix}_fullname")
                data["date_of_birth"] = st.text_input("Ngày sinh", value=data.get("date_of_birth", ""), key=f"{key_prefix}_dob")
            with field_col_2:
                data["sex"] = st.text_input("Giới tính", value=data.get("sex", ""), key=f"{key_prefix}_sex")
                data["nationality"] = st.text_input("Quốc tịch", value=data.get("nationality", ""), key=f"{key_prefix}_nationality")
                data["issue_date"] = st.text_input("Ngày cấp", value=data.get("issue_date", ""), key=f"{key_prefix}_issue")
                data["expiry_date"] = st.text_input("Có giá trị đến", value=data.get("expiry_date", ""), key=f"{key_prefix}_expiry")

            data["residence"] = st.text_input("Nơi thường trú", value=data.get("residence", ""), key=f"{key_prefix}_residence")
            item["data"] = apply_template_aliases(data)

    st.subheader("4) Tải file kết quả")
    if st.button("Tạo file kết quả"):
        zip_buffer = BytesIO()
        template_bytes = uploaded_template.getvalue()
        with zipfile.ZipFile(zip_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as archive:
            for item in st.session_state["batch_results"]:
                output_bytes = generate_docx_from_template(item["data"], template_bytes)
                result_name = safe_output_name(item["image_name"])
                archive.writestr(f"{result_name}_result.docx", output_bytes)

        zip_buffer.seek(0)
        st.download_button(
            label="Tải tất cả kết quả (.zip)",
            data=zip_buffer,
            file_name="ocr_results.zip",
            mime="application/zip",
        )

st.markdown("---")
st.caption("Mẹo: Dùng ảnh mặt trước rõ nét. Bạn có thể sửa từng trường trước khi tải file.")
