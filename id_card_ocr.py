import easyocr
import re
from collections import Counter
from docxtpl import DocxTemplate
import os
import unicodedata

class IDCardOCR:
    def __init__(self):
        print("Loading OCR Engine... (this may take a moment)")
        self.reader = easyocr.Reader(['vi', 'en'], gpu=False)
        self.cv2 = None
        try:
            import cv2 as cv2_module
            self.cv2 = cv2_module
        except Exception:
            self.cv2 = None

    @staticmethod
    def _normalize_text(text):
        text = unicodedata.normalize("NFKD", text)
        text = "".join(ch for ch in text if not unicodedata.combining(ch))
        text = text.lower().strip()
        text = re.sub(r'\s+', ' ', text)
        return text

    @staticmethod
    def _strip_known_labels(value):
        if not value:
            return ""
        value = re.sub(r'\s+', ' ', value).strip(" ;,:.-")
        value = re.sub(r'^[\d\W_]+', '', value)
        label_pattern = (
            r'^(?:ho\s*va\s*ten|ho\s*ten|full\s*name|'
            r'ngay\s*sinh|date\s*of\s*birth|'
            r'gioi\s*tinh|sex|'
            r'quoc\s*tich|nationality|'
            r'que\s*quan|place\s*of\s*origin|'
            r'noi\s*thuong\s*tru|thuong\s*tru|permanent\s*residence|residence|'
            r'id\s*no|no\.?|so\s*dinh\s*danh|so)\s*(?:/\s*)?'
        )
        normalized_value = IDCardOCR._normalize_text(value)
        while re.match(label_pattern, normalized_value):
            normalized_value = re.sub(label_pattern, '', normalized_value, count=1).strip(" ;,:.-")
            value = re.sub(r'^\s*[^:;\-]*[:;\-]?\s*', '', value, count=1).strip(" ;,:.-")
            if not value:
                break
            normalized_value = IDCardOCR._normalize_text(value)
        return value

    @staticmethod
    def _extract_segment_by_labels(raw_line, labels, stop_labels):
        normalized_line = IDCardOCR._normalize_text(raw_line)
        start = -1
        start_label = ""
        for label in labels:
            idx = normalized_line.find(label)
            if idx != -1 and (start == -1 or idx < start):
                start = idx
                start_label = label
        if start == -1:
            return ""

        segment = normalized_line[start + len(start_label):].strip(" ;,:.-/")
        stop_at = len(segment)
        for stop in stop_labels:
            stop_idx = segment.find(stop)
            if stop_idx != -1:
                stop_at = min(stop_at, stop_idx)
        segment = segment[:stop_at].strip(" ;,:.-/")
        segment = re.sub(r'^[\d\W_]+', '', segment)
        return segment

    @staticmethod
    def _extract_after_label(raw_line, labels):
        if not raw_line:
            return ""

        if ':' in raw_line:
            right = raw_line.split(':', 1)[1].strip(" ;,.-")
            cleaned = IDCardOCR._strip_known_labels(right)
            if cleaned:
                return cleaned

        if '-' in raw_line:
            right = raw_line.split('-', 1)[1].strip(" ;,.-")
            cleaned = IDCardOCR._strip_known_labels(right)
            if cleaned:
                return cleaned

        normalized_line = IDCardOCR._normalize_text(raw_line)
        for label in labels:
            if label in normalized_line:
                candidate = IDCardOCR._strip_known_labels(raw_line)
                if candidate and IDCardOCR._normalize_text(candidate) != normalized_line:
                    return candidate
        return ""

    @staticmethod
    def _clean_value(value):
        value = IDCardOCR._strip_known_labels(value)
        value = re.sub(r'\s+', ' ', value).strip(" ;,:.-")
        return value

    @staticmethod
    def _is_noise_or_label(value):
        if not value:
            return True
        normalized = IDCardOCR._normalize_text(value)
        if len(normalized) < 3:
            return True
        blocked = [
            "ho va ten", "full name", "ngay sinh", "date of birth", "gioi tinh", "sex",
            "quoc tich", "nationality", "que quan", "place of origin",
            "noi thuong tru", "permanent residence", "residence",
            "co gia tri", "valid until", "can cuoc", "cong hoa",
            "citizen identity card", "identity card", "can cuoc cong dan",
        ]
        if any(token in normalized for token in blocked):
            return True
        if normalized in {"quan", "pho", "of", "origin", "residence"}:
            return True
        if re.fullmatch(r'[a-z\s]+', normalized) and len(normalized.split()) <= 2 and normalized in {"of", "origin", "residence"}:
            return True
        return False

    @staticmethod
    def _normalize_date_text(raw):
        if not raw:
            return ""
        match = re.search(r'(\d{1,2})\s*[./-]\s*(\d{1,2})\s*[./-]\s*(\d{4})', raw)
        if not match:
            return ""
        day, month, year = match.groups()
        return f"{int(day):02d}/{int(month):02d}/{year}"

    @staticmethod
    def _normalize_compact_date(raw):
        if not raw:
            return ""
        match = re.search(r'\b(\d{8})\b', raw)
        if not match:
            return ""
        compact = match.group(1)
        day = int(compact[0:2])
        month = int(compact[2:4])
        year = int(compact[4:8])
        if 1 <= day <= 31 and 1 <= month <= 12 and 1900 <= year <= 2100:
            return f"{day:02d}/{month:02d}/{year:04d}"
        return ""

    @staticmethod
    def _choose_best_id(candidates):
        cleaned = [re.sub(r'\D', '', c) for c in candidates]
        cleaned = [c for c in cleaned if len(c) == 12]
        if not cleaned:
            return ""
        counts = Counter(cleaned)
        ranked = sorted(
            counts.items(),
            key=lambda pair: (
                pair[1],
                pair[0].startswith('0'),
            ),
            reverse=True,
        )
        return ranked[0][0]

    @staticmethod
    def _dedupe_place_text(value):
        if not value:
            return ""
        parts = re.split(r'[;,]+', value)
        seen = set()
        clean_parts = []
        for part in parts:
            part_clean = part.strip(" .,-")
            normalized = IDCardOCR._normalize_text(part_clean)
            if not part_clean or normalized in seen:
                continue
            if normalized in {"que", "quan", "pho", "thi"}:
                continue
            seen.add(normalized)
            clean_parts.append(part_clean)
        return ", ".join(clean_parts)

    def _extract_value_from_lines(self, lines, normalized_lines, labels):
        for index, nline in enumerate(normalized_lines):
            if any(label in nline for label in labels):
                inline = self._extract_after_label(lines[index], labels)
                if inline:
                    return self._clean_value(inline)
                if index + 1 < len(lines):
                    next_value = self._clean_value(lines[index + 1])
                    if next_value:
                        return next_value
        return ""

    @staticmethod
    def _line_has_any_label(normalized_line, labels):
        for label in labels:
            if label in normalized_line:
                return True
        return False

    def preprocess_image(self, image_path):
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"Cannot find image: {image_path}")

        if self.cv2 is None:
            return {
                "original_path": image_path,
            }

        img = self.cv2.imread(image_path)
        if img is None:
            raise FileNotFoundError(f"Cannot find image: {image_path}")

        gray = self.cv2.cvtColor(img, self.cv2.COLOR_BGR2GRAY)
        denoised = self.cv2.fastNlMeansDenoising(gray, h=10)
        clahe = self.cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        contrast = clahe.apply(denoised)
        thresh = self.cv2.adaptiveThreshold(
            contrast,
            255,
            self.cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            self.cv2.THRESH_BINARY,
            31,
            7,
        )

        return {
            "original": img,
            "processed": thresh,
            "gray": gray,
            "original_path": image_path,
        }

    def extract_text(self, image):
        variants = []
        if isinstance(image, dict):
            variants = [
                image.get("processed"),
                image.get("gray"),
                image.get("original"),
                image.get("original_path"),
            ]
        else:
            variants = [image]

        best_by_text = {}
        for variant in variants:
            if variant is None:
                continue
            result = self.reader.readtext(
                variant,
                detail=1,
                paragraph=False,
                text_threshold=0.6,
                low_text=0.2,
            )
            for item in result:
                if len(item) < 3:
                    continue
                text = str(item[1]).strip()
                conf = float(item[2])
                if not text or conf < 0.2:
                    continue
                key = self._normalize_text(text)
                existing = best_by_text.get(key)
                if existing is None or conf > float(existing[2]):
                    best_by_text[key] = item

        merged = list(best_by_text.values())

        def reading_order_key(item):
            bbox = item[0]
            xs = [point[0] for point in bbox]
            ys = [point[1] for point in bbox]
            return (min(ys), min(xs))

        merged.sort(key=reading_order_key)
        return merged

    def parse_data(self, text_list):
        lines = []
        if text_list and isinstance(text_list[0], (list, tuple)):
            for item in text_list:
                if len(item) >= 3 and item[2] >= 0.25:
                    lines.append(str(item[1]).strip())
        else:
            lines = [str(line).strip() for line in text_list]

        full_text = " ".join(lines)
        normalized_lines = [self._normalize_text(line) for line in lines]
        normalized_full_text = self._normalize_text(full_text)
        
        # Initialize with requested output fields
        data = {
            "fullname": "",
            "date_of_birth": "",
            "sex": "",
            "nationality": "",
            "place_of_origin": "",
            "no": "",
            "residence": "",
            "expiry_date": ""
        }

        # --- 1. No. ---
        no_labels = ["so", "so dinh danh", "id no", "no.", "no"]
        data["no"] = self._extract_value_from_lines(lines, normalized_lines, no_labels)
        id_candidates = re.findall(r'\d{12}', full_text)
        for line in lines:
            compact = re.sub(r'\D', '', line)
            if len(compact) == 12:
                id_candidates.append(compact)

        best_id = self._choose_best_id(id_candidates)
        digits_only = re.sub(r'\D', '', full_text)
        id_match = re.search(r'\d{12}', digits_only)
        if best_id:
            data["no"] = best_id
        elif id_match and (not data["no"] or len(re.sub(r'\D', '', data["no"])) < 9):
            data["no"] = id_match.group(0)

        # --- 2. Date of Birth ---
        date_pattern = r'(\d{1,2}\s*[/-]\s*\d{1,2}\s*[/-]\s*\d{4})'
        dob_labels = ["ngay sinh", "date of birth"]
        for index, nline in enumerate(normalized_lines):
            if any(label in nline for label in dob_labels):
                inline_date = self._normalize_date_text(lines[index])
                if inline_date:
                    data["date_of_birth"] = inline_date
                    break
                compact_inline = self._normalize_compact_date(lines[index])
                if compact_inline:
                    data["date_of_birth"] = compact_inline
                    break
                if index + 1 < len(lines):
                    next_line = self._normalize_date_text(lines[index + 1])
                    if next_line:
                        data["date_of_birth"] = next_line
                        break
                    compact_next = self._normalize_compact_date(lines[index + 1])
                    if compact_next:
                        data["date_of_birth"] = compact_next
                        break

        if not data["date_of_birth"]:
            data["date_of_birth"] = self._normalize_date_text(full_text)
        if not data["date_of_birth"]:
            data["date_of_birth"] = self._normalize_compact_date(full_text)

        # --- 3. Sex ---
        sex_line_value = self._extract_value_from_lines(lines, normalized_lines, ["gioi tinh", "sex"])
        normalized_sex = self._normalize_text(sex_line_value)
        if re.search(r'\bnam\b|\bmale\b', normalized_sex):
            data["sex"] = "Nam"
        elif re.search(r'\bnu\b|\bfemale\b', normalized_sex):
            data["sex"] = "Nữ"

        if re.search(r'\bnam\b|\bmale\b', normalized_full_text):
            data["sex"] = data["sex"] or "Nam"
        elif re.search(r'\bnu\b|\bfemale\b', normalized_full_text):
            data["sex"] = data["sex"] or "Nữ"

        # --- 4. Fullname ---
        name_labels = ["ho va ten", "ho ten", "full name"]
        for i, nline in enumerate(normalized_lines):
            if any(label in nline for label in name_labels):
                inline_name = self._extract_after_label(lines[i], name_labels)
                if inline_name:
                    candidate_name = self._clean_value(inline_name)
                    if not self._is_noise_or_label(candidate_name):
                        data["fullname"] = candidate_name
                        break
                if i + 1 < len(lines):
                    candidate = lines[i + 1].strip(" :-")
                    candidate_clean = self._clean_value(candidate)
                    if candidate_clean and not re.search(r'\d', candidate_clean) and not self._is_noise_or_label(candidate_clean):
                        data["fullname"] = candidate_clean
                        break

        if not data["fullname"]:
            uppercase_lines = [line for line in lines if len(line) > 4 and re.match(r'^[A-ZÀ-Ỹ\s]+$', line)]
            if uppercase_lines:
                data["fullname"] = uppercase_lines[0]

        # --- 5. Nationality ---
        nationality_labels = ["quoc tich", "nationality"]
        data["nationality"] = self._clean_value(self._extract_value_from_lines(lines, normalized_lines, nationality_labels))
        if not data["nationality"]:
            for line in lines:
                candidate = self._extract_segment_by_labels(
                    line,
                    nationality_labels,
                    ["que quan", "place of origin", "noi thuong tru", "residence", "gioi tinh", "sex", "ngay sinh", "date of birth", "ho va ten", "full name"],
                )
                if candidate:
                    data["nationality"] = self._clean_value(candidate)
                    break

        if data["nationality"]:
            nationality_norm = self._normalize_text(data["nationality"])
            match = re.search(
                r'(?:quoc\s*tich|nationality)\s*[:/\-]*\s*([a-z\s]{2,30})',
                nationality_norm,
            )
            if match:
                candidate = match.group(1).strip()
                candidate = re.sub(
                    r'\b(?:gioi\s*tinh|sex|que\s*quan|place\s*of\s*origin|noi\s*thuong\s*tru|residence|ngay\s*sinh|date\s*of\s*birth)\b.*$',
                    '',
                    candidate,
                ).strip()
                if candidate:
                    data["nationality"] = candidate.title()

        if data["nationality"]:
            normalized_nationality = self._normalize_text(data["nationality"])
            if "viet nam" in normalized_nationality or "vietnam" in normalized_nationality:
                data["nationality"] = "Việt Nam"
            elif (
                "quoc tich" in normalized_nationality
                or "nationality" in normalized_nationality
                or normalized_nationality in {"nam", "male"}
            ):
                data["nationality"] = "Việt Nam"

        if not data["nationality"] and "viet nam" in normalized_full_text:
            data["nationality"] = "Việt Nam"

        # --- 6. Place of origin ---
        origin_labels = ["que quan", "place of origin"]
        stop_labels = [
            "co gia tri den", "quoc tich", "gioi tinh", "ngay sinh", "ho va ten",
            "noi thuong tru", "thuong tru", "id no", "so dinh danh", "can cuoc"
        ]
        for i, nline in enumerate(normalized_lines):
            if any(label in nline for label in origin_labels):
                inline_origin = self._extract_after_label(lines[i], origin_labels)
                inline_origin_clean = self._clean_value(inline_origin)
                origin_parts = [inline_origin_clean] if inline_origin_clean and not self._is_noise_or_label(inline_origin_clean) else []

                j = i + 1
                while j < len(lines):
                    probe = normalized_lines[j]
                    if self._line_has_any_label(probe, stop_labels):
                        break
                    if re.search(date_pattern, lines[j]):
                        break
                    part = lines[j].strip(" ,.-")
                    cleaned_part = self._clean_value(part)
                    if cleaned_part and not self._is_noise_or_label(cleaned_part):
                        origin_parts.append(cleaned_part)
                    if len(origin_parts) >= 5:
                        break
                    j += 1

                if origin_parts:
                    combined_origin = self._clean_value(", ".join(origin_parts))
                    normalized_origin = self._normalize_text(combined_origin)
                    has_location_shape = any(ch in combined_origin for ch in [",", ";"]) or len(combined_origin) >= 10
                    if has_location_shape and "citizen identity card" not in normalized_origin:
                        data["place_of_origin"] = self._dedupe_place_text(combined_origin)
                break

        # --- Optional compatibility fields for old templates ---
        residence_labels = ["noi thuong tru", "thuong tru", "permanent residence", "residence"]
        residence_stop_labels = [
            "co gia tri den", "co gia tri", "quoc tich", "gioi tinh", "ngay sinh",
            "ho va ten", "que quan", "id no", "so dinh danh", "can cuoc"
        ]
        for i, nline in enumerate(normalized_lines):
            if any(label in nline for label in residence_labels):
                inline_address = self._extract_after_label(lines[i], residence_labels)
                inline_clean = self._clean_value(inline_address)
                address_parts = [inline_clean] if inline_clean and not self._is_noise_or_label(inline_clean) else []

                j = i + 1
                while j < len(lines):
                    probe = normalized_lines[j]
                    if self._line_has_any_label(probe, residence_stop_labels):
                        break
                    if re.search(date_pattern, lines[j]):
                        break
                    part = lines[j].strip(" ,.-")
                    cleaned_part = self._clean_value(part)
                    if cleaned_part and not self._is_noise_or_label(cleaned_part):
                        address_parts.append(cleaned_part)
                    if len(address_parts) >= 6:
                        break
                    j += 1

                if address_parts:
                    data["residence"] = self._clean_value(", ".join(address_parts))
                break

        if data["residence"]:
            data["residence"] = re.sub(
                r'^\s*\d*\s*(place\s*of\s*residence|permanent\s*residence|residence)\s*[:;,-]*\s*',
                '',
                data["residence"],
                flags=re.IGNORECASE,
            ).strip(" ;,:.-")

        dates = [self._normalize_date_text(d) for d in re.findall(date_pattern, full_text)]
        dates = [d for d in dates if d]
        for index, nline in enumerate(normalized_lines):
            if "co gia tri den" in nline or "valid until" in nline or "expiry" in nline:
                inline = self._normalize_date_text(lines[index])
                if inline:
                    data["expiry_date"] = inline
                    break
                compact_inline = self._normalize_compact_date(lines[index])
                if compact_inline:
                    data["expiry_date"] = compact_inline
                    break
                if index + 1 < len(lines):
                    next_line = self._normalize_date_text(lines[index + 1])
                    if next_line:
                        data["expiry_date"] = next_line
                        break
                    compact_next = self._normalize_compact_date(lines[index + 1])
                    if compact_next:
                        data["expiry_date"] = compact_next
                        break

        if not data["expiry_date"] and len(dates) > 1:
            for d in reversed(dates):
                if d != data["date_of_birth"]:
                    data["expiry_date"] = d
                    break

        if not data["residence"]:
            data["residence"] = data["place_of_origin"]

        data["full_name"] = data["fullname"]
        data["id_number"] = data["no"]
        data["dob"] = data["date_of_birth"]
        data["gender"] = data["sex"]
        data["ho_va_ten"] = data["fullname"]
        data["so"] = data["no"]
        data["ngay_sinh"] = data["date_of_birth"]
        data["gioi_tinh"] = data["sex"]
        data["quoc_tich"] = data["nationality"]
        data["que_quan"] = data["place_of_origin"]
        data["noi_thuong_tru"] = data["residence"]
        data["co_gia_tri_den"] = data["expiry_date"]

        return data

# --- NEW FUNCTION TO FILL WORD FILE ---
def fill_word_document(data, template_path, output_path):
    """
    Takes a dictionary of data and fills the Word template.
    """
    try:
        doc = DocxTemplate(template_path)

        template_context = {
            "fullname": data.get("fullname", ""),
            "date_of_birth": data.get("date_of_birth", ""),
            "sex": data.get("sex", ""),
            "nationality": data.get("nationality", ""),
            "place_of_origin": data.get("place_of_origin", ""),
            "no": data.get("no", ""),
            "full_name": data.get("full_name", data.get("fullname", "")),
            "id_number": data.get("id_number", data.get("no", "")),
            "dob": data.get("dob", data.get("date_of_birth", "")),
            "gender": data.get("gender", data.get("sex", "")),
            "residence": data.get("residence", data.get("place_of_origin", "")),
            "expiry_date": data.get("expiry_date", ""),
        }
        
        # The 'data' dictionary keys must match the {{ keys }} in the Word file
        doc.render(template_context)
        
        doc.save(output_path)
        print(f"✅ Successfully saved Word file to: {output_path}")
        
    except Exception as e:
        print(f"❌ Error creating Word file: {e}")

# --- MAIN EXECUTION ---
if __name__ == "__main__":
    # 1. Setup paths
    image_file = "images/01cf6f6eeb4665183c57.jpg"   # Your input image
    template_file = "template.docx" # Your Word template
    output_file = "result.docx"     # The file to be created

    # Check if template exists
    if not os.path.exists(template_file):
        print(f"Error: Please create '{template_file}' with placeholders first.")
    else:
        # 2. Run OCR
        ocr_engine = IDCardOCR()
        
        try:
            # Preprocess
            img = ocr_engine.preprocess_image(image_file)
            
            # Extract
            print("Extracting text...")
            raw_text = ocr_engine.extract_text(img)
            
            # Parse
            clean_data = ocr_engine.parse_data(raw_text)
            
            # Print to screen
            print("\n--- Extracted Data ---")
            for key, value in clean_data.items():
                print(f"{key}: {value}")
            
            # 3. Fill Word File
            print("\nFilling Word Document...")
            fill_word_document(clean_data, template_file, output_file)
            
        except Exception as e:
            print(f"An error occurred: {e}")