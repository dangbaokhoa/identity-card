import re
from docxtpl import DocxTemplate
import os

class IDCardQRReader:
    def __init__(self):
        print("[IDCardQRReader] Initializing QR Reader...")
        self.cv2 = None
        self.pyzbar = None
        
        try:
            print("[IDCardQRReader] Attempting to import cv2...")
            import cv2 as cv2_module
            self.cv2 = cv2_module
            print("[IDCardQRReader] ✓ CV2 module imported successfully")
        except Exception as e:
            print(f"[IDCardQRReader] ✗ Failed to import cv2: {e}")
            self.cv2 = None
        
        try:
            print("[IDCardQRReader] Attempting to import pyzbar...")
            from pyzbar import pyzbar
            self.pyzbar = pyzbar
            print("[IDCardQRReader] ✓ Pyzbar module imported successfully")
        except Exception as e:
            print(f"[IDCardQRReader] ✗ Failed to import pyzbar: {e}")
            self.pyzbar = None
        
        print("[IDCardQRReader] Initialization complete")
    
    @staticmethod
    def _format_date(date_str):
        """Convert DDMMYYYY to DD/MM/YYYY"""
        if not date_str or len(date_str) != 8:
            return ""
        try:
            day = date_str[0:2]
            month = date_str[2:4]
            year = date_str[4:8]
            return f"{day}/{month}/{year}"
        except:
            return ""
    
    def read_qr_code(self, image_path, debug=False):
        """Read QR code from ID card image with multiple preprocessing strategies"""
        print(f"[IDCardQRReader] Reading QR code from: {image_path}")
        
        if not os.path.exists(image_path):
            print(f"[IDCardQRReader] ✗ Image not found: {image_path}")
            raise FileNotFoundError(f"Cannot find image: {image_path}")
        
        if self.cv2 is None:
            raise RuntimeError("OpenCV không khả dụng. Cần cài đặt opencv-python-headless")
        
        # Read image
        img = self.cv2.imread(image_path)
        if img is None:
            raise FileNotFoundError(f"Cannot read image: {image_path}")
        
        h, w = img.shape[:2]
        print(f"[IDCardQRReader] Image size: {w}x{h}")
        
        print("[IDCardQRReader] Image loaded, detecting QR code...")
        
        # Strategy: Try multiple preprocessing variants
        variants = []
        
        # 1. Original image
        variants.append(("original", img))
        
        # 2. Grayscale
        gray = self.cv2.cvtColor(img, self.cv2.COLOR_BGR2GRAY)
        variants.append(("grayscale", gray))
        
        # 3. Enhanced contrast
        clahe = self.cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
        enhanced = clahe.apply(gray)
        variants.append(("enhanced", enhanced))
        
        # 4. Binary threshold
        _, binary = self.cv2.threshold(gray, 127, 255, self.cv2.THRESH_BINARY)
        variants.append(("binary", binary))
        
        # 5. Adaptive threshold
        adaptive = self.cv2.adaptiveThreshold(
            gray, 255, self.cv2.ADAPTIVE_THRESH_GAUSSIAN_C, 
            self.cv2.THRESH_BINARY, 11, 2
        )
        variants.append(("adaptive", adaptive))
        
        # 6. Upscaled version (for small QR codes)
        height, width = img.shape[:2]
        if max(height, width) < 1500:
            scale_factor = 1500 / max(height, width)
            new_width = int(width * scale_factor)
            new_height = int(height * scale_factor)
            upscaled = self.cv2.resize(img, (new_width, new_height), interpolation=self.cv2.INTER_CUBIC)
            variants.append(("upscaled", upscaled))
            
            upscaled_gray = self.cv2.cvtColor(upscaled, self.cv2.COLOR_BGR2GRAY)
            variants.append(("upscaled_gray", upscaled_gray))
        
        # Try pyzbar on all variants
        if self.pyzbar:
            for name, variant in variants:
                print(f"[IDCardQRReader] Trying pyzbar on {name}...")
                
                # Save debug images if requested
                if debug:
                    debug_path = f"debug_{name}.jpg"
                    self.cv2.imwrite(debug_path, variant)
                    print(f"[IDCardQRReader] Saved debug image: {debug_path}")
                
                decoded_objects = self.pyzbar.decode(variant)
                if decoded_objects:
                    qr_data = decoded_objects[0].data.decode('utf-8')
                    print(f"[IDCardQRReader] ✓ QR decoded with pyzbar ({name}): {qr_data[:50]}...")
                    return qr_data
        else:
            print("[IDCardQRReader] ⚠ pyzbar not available, will use cv2 only")
        
        # Try WeChat QR detector (more robust than default cv2 detector)
        try:
            wechat_detector = self.cv2.wechat_qrcode_WeChatQRCode()
            for name, variant in variants:
                print(f"[IDCardQRReader] Trying WeChat QR on {name}...")
                
                # Convert to BGR if grayscale
                if len(variant.shape) == 2:
                    variant_bgr = self.cv2.cvtColor(variant, self.cv2.COLOR_GRAY2BGR)
                else:
                    variant_bgr = variant
                
                res, points = wechat_detector.detectAndDecode(variant_bgr)
                if res and len(res) > 0 and res[0]:
                    qr_data = res[0]
                    print(f"[IDCardQRReader] ✓ QR decoded with WeChat ({name}): {qr_data[:50]}...")
                    return qr_data
        except Exception as e:
            print(f"[IDCardQRReader] WeChat QR detector not available: {e}")
        
        # Fallback to standard cv2 QR detector on all variants
        qr_detector = self.cv2.QRCodeDetector()
        for name, variant in variants:
            print(f"[IDCardQRReader] Trying cv2 standard QR on {name}...")
            # Convert to BGR if grayscale
            if len(variant.shape) == 2:
                variant = self.cv2.cvtColor(variant, self.cv2.COLOR_GRAY2BGR)
            
            data, bbox, _ = qr_detector.detectAndDecode(variant)
            if data:
                print(f"[IDCardQRReader] ✓ QR decoded with cv2 ({name}): {data[:50]}...")
                return data
        
        print("[IDCardQRReader] ✗ No QR code found after trying all variants")
        print("[IDCardQRReader] Suggestions:")
        print("  1. Đảm bảo ảnh chụp có mã QR (thẻ CCCD từ 2021 trở lên)")
        print("  2. Chụp rõ toàn bộ thẻ, không bị mờ hoặc nghiêng")
        print("  3. Thử chụp lại với ánh sáng tốt hơn")
        print("  4. Nếu có debug images (debug_*.jpg), kiểm tra xem QR có rõ không")
        raise ValueError("Không tìm thấy mã QR trên ảnh. Vui lòng kiểm tra:\n"
                        "✓ Ảnh chụp có mã QR (thẻ CCCD từ 2021+)\n"
                        "✓ Chụp rõ toàn bộ thẻ\n"
                        "✓ Ánh sáng tốt, không bị mờ/nghiêng")
    
    def parse_qr_data(self, qr_string):
        """
        Parse QR code data from Vietnamese ID card
        Format: 049205000868|206454491|Đặng Bảo Khoa|01072005|Nam|Hòa Nam, Tân Thạnh, Tam Kỳ, Quảng Nam|11042021
        Fields: cccd_no|old_cmnd|fullname|dob|sex|residence|issue_date
        """
        print("[IDCardQRReader] Parsing QR data...")
        
        parts = qr_string.split('|')
        
        if len(parts) < 7:
            print(f"[IDCardQRReader] ✗ Invalid QR format. Expected 7 fields, got {len(parts)}")
            raise ValueError(f"Dữ liệu QR không đúng định dạng. Cần 7 trường, nhận được {len(parts)}")
        
        data = {
            "no": parts[0].strip(),
            "old_id": parts[1].strip(),
            "fullname": parts[2].strip(),
            "date_of_birth": self._format_date(parts[3].strip()),
            "sex": parts[4].strip(),
            "residence": parts[5].strip(),
            "issue_date": self._format_date(parts[6].strip()) if len(parts) > 6 else "",
            "expiry_date": "",  # QR code không có field này, để trống
            "nationality": "Việt Nam",  # Mặc định
            "place_of_origin": parts[5].strip(),  # Dùng residence làm place_of_origin
        }
        
        # Add alias fields for template compatibility
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
        data["ngay_cap"] = data["issue_date"]
        data["co_gia_tri_den"] = data["expiry_date"]
        
        print("[IDCardQRReader] ✓ Parsing complete")
        print(f"[IDCardQRReader] Extracted: {data['fullname']} - {data['no']}")
        
        return data
    
    def process_image(self, image_path, debug=False):
        """
        Main method: read QR and parse data
        """
        print("[IDCardQRReader] Starting QR processing...")
        qr_data = self.read_qr_code(image_path, debug=debug)
        parsed_data = self.parse_qr_data(qr_data)
        print("[IDCardQRReader] ✓ Processing complete")
        return parsed_data


# --- FUNCTION TO FILL WORD FILE ---
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
            "issue_date": data.get("issue_date", ""),
            "old_id": data.get("old_id", ""),
        }
        
        doc.render(template_context)
        doc.save(output_path)
        print(f"✅ Successfully saved Word file to: {output_path}")
        
    except Exception as e:
        print(f"❌ Error creating Word file: {e}")


# --- MAIN EXECUTION ---
if __name__ == "__main__":
    # 1. Setup paths
    image_file = "images/01cf6f6eeb4665183c57.jpg"   # Your input image (back side with QR)
    template_file = "template.docx" # Your Word template
    output_file = "result.docx"     # The file to be created

    # Check if template exists
    if not os.path.exists(template_file):
        print(f"Error: Please create '{template_file}' with placeholders first.")
    else:
        # 2. Run QR Reader
        qr_reader = IDCardQRReader()
        
        try:
            # Read and parse QR
            print("Reading QR code...")
            clean_data = qr_reader.process_image(image_file, debug=True)  # Enable debug mode
            
            # Print to screen
            print("\n--- Extracted Data ---")
            for key, value in clean_data.items():
                print(f"{key}: {value}")
            
            # 3. Fill Word File
            print("\nFilling Word Document...")
            fill_word_document(clean_data, template_file, output_file)
            
        except Exception as e:
            print(f"An error occurred: {e}")
