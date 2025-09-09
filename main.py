import imaplib
import email
from email.utils import parseaddr, parsedate_to_datetime
from bs4 import BeautifulSoup
import base64
import os
import re
import requests
import cv2
import numpy as np
from PIL import Image

# === CONFIGURATION ===
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
FOLDER_NAME = os.getenv("FOLDER_NAME", '"Notif Report"')
SAVE_FOLDER = "ImageEmail"
os.makedirs(SAVE_FOLDER, exist_ok=True)

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
CHAT_IDS = os.getenv("CHAT_IDS", "").split(",")

def clean_filename(filename):
    return re.sub(r'[<>:"/\\|?*]', '_', filename)

def send_text_to_telegram(text):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    for chat_id in CHAT_IDS:
        data = {
            'chat_id': chat_id,
            'text': text,
            'parse_mode': 'HTML'
        }
        response = requests.post(url, data=data)
        print(f"Text sent to {chat_id}: {response.text}")

def send_document_to_telegram(file_path):
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendDocument"
    for chat_id in CHAT_IDS:
        with open(file_path, 'rb') as doc_file:
            files = {
                'document': (os.path.basename(file_path), doc_file)
            }
            data = {'chat_id': chat_id}
            response = requests.post(url, data=data, files=files)
            print(f"Document sent to {chat_id}: {response.text}")

def is_valid_image(filepath):
    try:
        with Image.open(filepath) as img:
            img.verify()
        return True
    except Exception as e:
        print(f"File bukan image valid: {e}")
        return False

def auto_crop_image(image_path, save_path=None):
    try:
        image = cv2.imread(image_path)
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        mask = gray < 245
        coords = np.argwhere(mask)

        if coords.size > 0:
            y0, x0 = coords.min(axis=0)
            y1, x1 = coords.max(axis=0) + 1
            cropped = image[y0:y1, x0:x1]

            if not save_path:
                save_path = image_path

            ext = os.path.splitext(save_path)[1].lower()
            if ext in ['.jpg', '.jpeg']:
                cv2.imwrite(save_path, cropped, [int(cv2.IMWRITE_JPEG_QUALITY), 95])
            else:
                cv2.imwrite(save_path, cropped)

            print(f"Gambar dicrop: {save_path}")
            return save_path
        else:
            print("Gambar kosong atau seluruhnya putih.")
            return image_path

    except Exception as e:
        print(f"Error saat cropping: {e}")
        return image_path

def check_email_and_send_inline_images():
    mail = imaplib.IMAP4_SSL("mail.allobank.com")
    mail.login(EMAIL_USER, EMAIL_PASSWORD)
    mail.select(FOLDER_NAME)

    status, messages = mail.search(None, 'UNSEEN')
    if status != "OK":
        print("Tidak ada email baru.")
        return

    for num in messages[0].split():
        status, msg_data = mail.fetch(num, '(RFC822)')
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                from_ = parseaddr(msg.get("From"))[1]
                if from_ != 'no-reply-reports@allobank.com':
                    continue

                subject = msg.get("Subject", "(No Subject)").strip()
                date = msg.get("Date")
                try:
                    parsed_date = parsedate_to_datetime(date)

                    # kalau timezone kosong, anggap UTC
                    if parsed_date.tzinfo is None:
                        parsed_date = parsed_date.replace(tzinfo=timezone.utc)

                    # konversi eksplisit ke WIB
                    jakarta_tz = pytz.timezone("Asia/Jakarta")
                    parsed_date = parsed_date.astimezone(jakarta_tz)

                    date_str = parsed_date.strftime("%d %b %Y %H:%M")
                except Exception as e:
                    print(f"Error parsing date: {e}")
                    date_str = date or "(Unknown Time)"

                message_text = f"<b>Subject:</b> {subject}\n<b>Waktu:</b> {date_str}"
                send_text_to_telegram(message_text)

                attachments = {}
                for part in msg.walk():
                    if part.get_content_maintype() == 'image':
                        content_id = part.get('Content-ID')
                        if content_id:
                            clean_cid = content_id.strip('<>')
                            attachments[clean_cid] = part.get_payload(decode=True)
                            print(f"Attachment ditemukan: CID={clean_cid}")

                for part in msg.walk():
                    if part.get_content_type() == "text/html":
                        html_body = part.get_payload(decode=True).decode(errors='ignore')
                        soup = BeautifulSoup(html_body, 'html.parser')
                        img_tags = soup.find_all('img')
                        print(f"Jumlah <img>: {len(img_tags)}")

                        for i, img in enumerate(img_tags):
                            src = img.get('src')
                            if not src:
                                continue

                            try:
                                if src.startswith('data:image'):
                                    header, encoded = src.split(',', 1)
                                    img_data = base64.b64decode(encoded)
                                    ext = header.split('/')[1].split(';')[0]
                                    filename = f'image_{i}_inline.{ext}'
                                    filepath = os.path.join(SAVE_FOLDER, filename)

                                elif src.startswith('cid:'):
                                    cid = src[4:]
                                    if cid in attachments:
                                        img_data = attachments[cid]
                                        filename = f'image_{i}_cid.jpg'
                                        filepath = os.path.join(SAVE_FOLDER, filename)
                                    else:
                                        print(f"CID {cid} tidak ditemukan.")
                                        continue

                                elif src.startswith('http'):
                                    response = requests.get(src)
                                    if response.status_code != 200:
                                        print(f"Download gagal: {src}")
                                        continue
                                    img_data = response.content
                                    filename = f'image_{i}_url.jpg'
                                    filepath = os.path.join(SAVE_FOLDER, filename)
                                else:
                                    print(f"Format img tidak dikenali: {src}")
                                    continue

                                with open(filepath, 'wb') as f:
                                    f.write(img_data)
                                print(f"Disimpan: {filepath}")

                                cropped_path = auto_crop_image(filepath)
                                if is_valid_image(cropped_path):
                                    send_document_to_telegram(cropped_path)
                                else:
                                    print("Gagal validasi gambar sebelum dikirim.")

                            except Exception as e:
                                print(f"Error proses gambar ke-{i}: {e}")

    mail.logout()

# === RUN ===
if __name__ == "__main__":
    check_email_and_send_inline_images()
    print("Proses selesai.")
