import os
import tempfile
import zipfile
import importlib
from io import BytesIO

os.environ.setdefault("STREAMLIT_SERVER_FILE_WATCHER_TYPE", "none")
os.environ.setdefault("STREAMLIT_SERVER_RUN_ON_SAVE", "false")

import streamlit as st
from docxtpl import DocxTemplate


st.set_page_config(page_title="ID Card QR Reader", page_icon="🪪", layout="centered")
st.title("🪪 Đọc thông tin CCCD từ mã QR")
st.caption("Tải mẫu Word + ảnh thẻ CCCD (có QR code), xem kết quả, sau đó tải tất cả file kết quả.")
st.info("💡 Mã QR thường ở góc thẻ CCCD. Chụp rõ toàn bộ thẻ để detect tốt nhất.")


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


START_MARKER_TEXT = "Số TKHQ hàng hóa nhập khẩu đã thông quan"
END_MARKER_TEXT = "Tổng cộng"


def _normalize_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip().lower()


def extract_tkhq_numbers_from_excel(uploaded_excel, filename: str) -> list[dict]:
    load_workbook = importlib.import_module("openpyxl").load_workbook

    if hasattr(uploaded_excel, "seek"):
        uploaded_excel.seek(0)

    workbook = load_workbook(uploaded_excel, data_only=True, read_only=True)
    worksheet = workbook.active

    start_marker = _normalize_text(START_MARKER_TEXT)
    end_marker = _normalize_text(END_MARKER_TEXT)

    in_extract_range = False
    extracted_entries = []

    for row in worksheet.iter_rows(min_col=2, max_col=2):
        cell = row[0]
        cell_value = cell.value
        normalized = _normalize_text(cell_value)

        if not in_extract_range:
            if normalized == start_marker:
                in_extract_range = True
            continue

        if normalized == end_marker:
            break

        raw_text = str(cell_value).strip() if cell_value is not None else ""
        digits_only = "".join(char for char in raw_text if char.isdigit())
        if digits_only:
            extracted_entries.append(
                {
                    "number": digits_only,
                    "file": filename,
                    "sheet": worksheet.title,
                    "cell": f"B{cell.row}",
                }
            )

    workbook.close()

    if not in_extract_range:
        raise ValueError(f"Không tìm thấy dòng bắt đầu '{START_MARKER_TEXT}' trong file {filename}")
    if in_extract_range and len(extracted_entries) == 0:
        raise ValueError(f"Không đọc được số TKHQ nào giữa '{START_MARKER_TEXT}' và '{END_MARKER_TEXT}' trong file {filename}")

    return extracted_entries


if "batch_results" not in st.session_state:
    st.session_state["batch_results"] = []
if "show_usage_guide_cccd" not in st.session_state:
    st.session_state["show_usage_guide_cccd"] = False
if "show_usage_guide_excel" not in st.session_state:
    st.session_state["show_usage_guide_excel"] = False

tab_cccd, tab_excel = st.tabs(["🪪 Chức năng CCCD", "📊 Chức năng Excel"])

with tab_cccd:
    if st.button("📘 Cách dùng CCCD", key="guide_button_cccd"):
        st.session_state["show_usage_guide_cccd"] = not st.session_state["show_usage_guide_cccd"]

    if st.session_state["show_usage_guide_cccd"]:
        st.markdown(
            """
### Hướng dẫn chức năng CCCD
1. Tải file mẫu Word `.docx` có placeholder đúng định dạng `{{ ten_placeholder }}`.
2. Tải một hoặc nhiều ảnh CCCD (mặt có mã QR).
3. Nhấn **Đọc mã QR** để trích xuất dữ liệu.
4. Kiểm tra/chỉnh sửa thông tin ở từng ảnh.
5. Nhấn **Tạo file kết quả** để tải file `.zip`.

### Placeholder tiếng Việt hỗ trợ
- `{{ ho_va_ten }}`: Họ và tên
- `{{ so }}`: Số CCCD
- `{{ ngay_sinh }}`: Ngày sinh
- `{{ gioi_tinh }}`: Giới tính
- `{{ quoc_tich }}`: Quốc tịch
- `{{ noi_thuong_tru }}`: Nơi thường trú
- `{{ que_quan }}`: Quê quán
- `{{ co_gia_tri_den }}`: Có giá trị đến
            """
        )

    st.subheader("1) Tải mẫu Word")
    uploaded_template = st.file_uploader(
        "Tải file mẫu .docx (bắt buộc)",
        type=["docx"],
        key="template_required",
    )

    st.subheader("2) Tải ảnh thẻ CCCD")
    uploaded_images = st.file_uploader(
        "Tải lên một hoặc nhiều ảnh thẻ CCCD (chụp rõ mã QR)",
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
        st.info("Vui lòng tải lên ít nhất một ảnh thẻ CCCD để tiếp tục.")

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
        if st.button("Tạo file kết quả", key="generate_result_button"):
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

    st.caption("💡 Mẹo: Chụp rõ toàn bộ thẻ để mã QR dễ detect. Bạn có thể sửa từng trường trước khi tải file.")

with tab_excel:
    if st.button("📘 Cách dùng Excel", key="guide_button_excel"):
        st.session_state["show_usage_guide_excel"] = not st.session_state["show_usage_guide_excel"]

    if st.session_state["show_usage_guide_excel"]:
        st.markdown(
            """
### Hướng dẫn chức năng Excel
1. Tải file import hiện tại (.xlsx/.xlsm).
2. Tải các file ngày hôm trước (có thể nhiều file).
3. Nhấn **Rà soát trùng số TKHQ**.
4. Xem tổng hợp file nào trùng file nào.
5. Xem chi tiết từng ô trùng theo định dạng `Sheet!B...`.

### Quy tắc đọc dữ liệu
- Chỉ đọc cột B.
- Bắt đầu từ dòng sau: `Số TKHQ hàng hóa nhập khẩu đã thông quan`.
- Kết thúc trước dòng: `Tổng cộng`.
            """
        )

    st.subheader("Rà soát trùng số TKHQ từ Excel")
    st.caption("Đọc cột B từ dòng sau 'Số TKHQ hàng hóa nhập khẩu đã thông quan' đến trước 'Tổng cộng'.")

    current_import_excel = st.file_uploader(
        "File import hiện tại (.xlsx/.xlsm)",
        type=["xlsx", "xlsm"],
        key="current_import_excel",
    )

    previous_day_excels = st.file_uploader(
        "Các file ngày hôm trước (.xlsx/.xlsm, chọn nhiều file)",
        type=["xlsx", "xlsm"],
        accept_multiple_files=True,
        key="previous_day_excels",
    )

    can_check_duplicate = current_import_excel is not None and previous_day_excels
    if st.button("Rà soát trùng số TKHQ", disabled=not can_check_duplicate):
        with st.spinner("Đang đọc file Excel và đối soát dữ liệu..."):
            try:
                current_entries = extract_tkhq_numbers_from_excel(current_import_excel, current_import_excel.name)
                current_number_set = {entry["number"] for entry in current_entries}

                previous_entries = []
                for previous_file in previous_day_excels:
                    previous_entries.extend(extract_tkhq_numbers_from_excel(previous_file, previous_file.name))

                previous_index = {}
                for entry in previous_entries:
                    number = entry["number"]
                    if number not in previous_index:
                        previous_index[number] = []
                    previous_index[number].append(entry)

                duplicated_numbers = sorted(current_number_set.intersection(previous_index.keys()))

                match_rows = []
                file_pair_counter = {}

                for current_entry in current_entries:
                    number = current_entry["number"]
                    if number not in previous_index:
                        continue

                    for previous_entry in previous_index[number]:
                        match_rows.append(
                            {
                                "Số TKHQ": number,
                                "File hiện tại": current_entry["file"],
                                "Ô hiện tại": f"{current_entry['sheet']}!{current_entry['cell']}",
                                "File ngày trước": previous_entry["file"],
                                "Ô ngày trước": f"{previous_entry['sheet']}!{previous_entry['cell']}",
                            }
                        )

                        pair_key = (current_entry["file"], previous_entry["file"])
                        file_pair_counter[pair_key] = file_pair_counter.get(pair_key, 0) + 1

                st.write(f"Số TKHQ duy nhất trong file hiện tại: **{len(current_number_set)}**")
                st.write(f"Số TKHQ duy nhất bị trùng với file ngày trước: **{len(duplicated_numbers)}**")
                st.write(f"Tổng số lượt trùng theo từng ô (match records): **{len(match_rows)}**")

                if match_rows:
                    st.warning("Phát hiện số TKHQ bị trùng với dữ liệu ngày trước.")

                    summary_rows = []
                    for pair_key, count in sorted(file_pair_counter.items(), key=lambda item: (-item[1], item[0][1])):
                        summary_rows.append(
                            {
                                "File hiện tại": pair_key[0],
                                "File ngày trước": pair_key[1],
                                "Số lượt trùng": count,
                            }
                        )

                    st.markdown("**Tổng hợp file nào trùng file nào**")
                    st.dataframe(summary_rows, use_container_width=True)

                    st.markdown("**Chi tiết trùng theo ô**")
                    st.dataframe(match_rows, use_container_width=True)
                else:
                    st.success("Không phát hiện số TKHQ nào bị trùng với các file ngày hôm trước.")
            except Exception as error:
                st.error(f"Không thể rà soát dữ liệu: {error}")
