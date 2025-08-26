import os
from fpdf import FPDF
from email import message_from_string
from email.policy import default
from email.utils import parsedate_to_datetime
from PIL import Image
from io import BytesIO
from tkinter import Tk, filedialog
import uuid

def safe_text(text):
    return text.encode("latin-1", "replace").decode("latin-1")

def format_email_date(raw_date):
    offset_map = {
        "-1200": "UTC-12:00 IDLW", "-1100": "UTC-11:00 SST", "-1000": "UTC-10:00 HST",
        "-0900": "UTC-09:00 AKST", "-0800": "UTC-08:00 PST", "-0700": "UTC-07:00 MST",
        "-0600": "UTC-06:00 CST", "-0500": "UTC-05:00 EST", "-0400": "UTC-04:00 AST",
        "-0330": "UTC-03:30 NST", "-0300": "UTC-03:00 BRT", "-0200": "UTC-02:00 FNT",
        "-0100": "UTC-01:00 AZOT", "+0000": "UTC+00:00 GMT", "+0100": "UTC+01:00 CET",
        "+0200": "UTC+02:00 EET", "+0300": "UTC+03:00 MSK", "+0330": "UTC+03:30 IRST",
        "+0400": "UTC+04:00 GST", "+0430": "UTC+04:30 AFT", "+0500": "UTC+05:00 PKT",
        "+0530": "UTC+05:30 IST", "+0545": "UTC+05:45 NPT", "+0600": "UTC+06:00 BST",
        "+0630": "UTC+06:30 MMT", "+0700": "UTC+07:00 ICT", "+0800": "UTC+08:00 CST",
        "+0900": "UTC+09:00 JST", "+0930": "UTC+09:30 ACST", "+1000": "UTC+10:00 AEST",
        "+1030": "UTC+10:30 LHST", "+1100": "UTC+11:00 AEDT", "+1200": "UTC+12:00 NZST"
    }
    try:
        dt = parsedate_to_datetime(raw_date)
        offset_str = dt.strftime("%z")
        label = offset_map.get(offset_str, f"UTC{offset_str[:3]}:{offset_str[3:]}")
        return dt.strftime("%Y-%m-%d %H:%M:%S") + " " + label
    except Exception:
        return raw_date


    try:
        dt = parsedate_to_datetime(raw_date)
        offset_str = dt.strftime("%z")
        label = offset_map.get(offset_str, f"UTC{offset_str[:3]}:{offset_str[3:]}")
        return dt.strftime("%Y-%m-%d %H:%M") + " " + label
    except Exception:
        return raw_date

class EmailPDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 12)
        if self.title:
            self.cell(0, 10, safe_text(self.title), ln=True, align="C")
        self.ln(5)

    def add_text(self, label, text):
        self.set_font("Arial", "B", 10)
        self.multi_cell(0, 10, safe_text(f"{label}: {text}"))
        self.set_font("Arial", "", 10)

    def add_body(self, body, image_map):
        self.set_font("Arial", "", 10)
        lines = body.splitlines()
        for line in lines:
            stripped = line.strip()
            if stripped.startswith("[image") or stripped.startswith("Image"):
                ref = stripped.split(",", 1)[-1].strip().split()[-1]
                match = next((k for k in image_map if ref in k), None)
                if match:
                    self.ln()
                    self.add_image(image_map[match], caption=f"Inline {ref}")
                    self.ln()
                    continue
            self.multi_cell(0, 10, safe_text(line))
        self.ln()

    def add_image(self, img_bytes, caption="Image"):
        try:
            buffer = BytesIO(img_bytes)
            buffer.seek(0)
            img = Image.open(buffer).convert("RGB")
            img.load()
            temp_path = f"_temp_{uuid.uuid4().hex}.png"
            img.save(temp_path)
            self.add_page()
            self.set_font("Arial", "I", 10)
            self.cell(0, 10, safe_text(caption), ln=True)
            self.image(temp_path, w=150)
            os.remove(temp_path)
        except Exception as e:
            self.cell(0, 10, safe_text(f"Could not render image: {e}"), ln=True)

def select_email_file():
    root = Tk()
    root.withdraw()
    return filedialog.askopenfilename(
        title="Select Raw Email .txt File",
        filetypes=[("Text Files", ".txt"), ("All Files", ".*")]
    )

def parse_email_txt(txt_path):
    with open(txt_path, "r", encoding="utf-8") as f:
        raw = f.read()

    msg = message_from_string(raw, policy=default)
    subject = msg.get("Subject", "No Subject")
    to = msg.get("To", "Unknown Recipient")
    from_ = msg.get("From", "Unknown Sender")
    cc = msg.get("Cc", "None")
    raw_date = msg.get("Date", "Unknown Date")
    formatted_date = format_email_date(raw_date)

    body = ""
    image_map = {}
    image_index = 0

    for part in msg.walk():
        content_type = part.get_content_type()
        disp = str(part.get("Content-Disposition", "")).lower()
        payload = part.get_payload(decode=True)
        filename = part.get_filename()

        if content_type == "text/plain" and "attachment" not in disp:
            body += part.get_content()
        elif content_type.startswith("image") and payload:
            base_name = filename or f"embedded_{image_index}.jpg"
            key = f"{image_index}_{base_name.split('.')[-1]}"
            image_map[key] = payload
            image_index += 1
        elif "attachment" in disp and payload:
            if filename and filename.lower().endswith((".png", ".jpg", ".jpeg", ".gif")):
                key = f"{image_index}_{filename.split('.')[-1]}"
                image_map[key] = payload
                image_index += 1

    pdf = EmailPDF()
    pdf.set_title("")
    pdf.add_page()
    pdf.add_text("From", from_)
    pdf.add_text("To", to)
    pdf.add_text("CC", cc)
    pdf.add_text("Date", formatted_date)
    pdf.add_text("Subject", subject)
    pdf.ln()
    pdf.add_body(body, image_map)

    used_keys = set()
    for line in body.splitlines():
        if "[image" in line or "Image" in line:
            ref = line.split(",", 1)[-1].strip().split()[-1]
            match = next((k for k in image_map if ref in k), None)
            if match:
                used_keys.add(match)

    for filename, data in image_map.items():
        if filename not in used_keys:
            pdf.add_image(data, caption=f"Attachment {filename}")

    output_dir = os.path.dirname(txt_path)
    output_name = os.path.splitext(os.path.basename(txt_path))[0] + ".pdf"
    output_path = os.path.join(output_dir, output_name)
    output_path_display = output_path.replace('/', '\\')
    pdf.output(output_path, "F")
    print(f"\033[92m✅ PDF successfully saved to {output_path_display}\033[0m")


if __name__ == "__main__":
    print("Welcome to Emailverter, select a valid file")
    selected_file = select_email_file()
    if selected_file:
        parse_email_txt(selected_file)
    else:
        print("\033[91m❌ Invalid selection\033[0m")

