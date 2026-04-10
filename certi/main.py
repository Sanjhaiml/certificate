from PIL import Image, ImageDraw, ImageFont
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font as ExcelFont
from datetime import datetime
import qrcode
import json
import uuid
import re
import time
import socket
import webbrowser
from flask import Flask, render_template_string, request, send_file, redirect

# ==================== FLASK APP SETUP ====================
app = Flask(__name__)
app.secret_key = 'certificate-secret-key-2025'

# ==================== CONFIGURATION ====================
EXCEL_FILE = "CSE.xlsx"
TEMPLATE_PATH = "certificate_template.png"
BASE_OUTPUT_DIR = "Certificates"

# Get local IP address for network access
def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return "localhost"

LOCAL_IP = get_local_ip()
SERVER_URL = f"http://{LOCAL_IP}:8000"  # Use local IP for network access
# SERVER_URL = "http://localhost:8000"  # Use localhost for local access

print(f"\n🌐 Server will run at: {SERVER_URL}")

# ==================== PROFESSIONAL FONT CONFIGURATION ====================
FONT_CONFIG = {
    'name': {
        'fonts': [
            "C:/Windows/Fonts/timesbd.ttf",
            "C:/Windows/Fonts/georgiab.ttf",
            "C:/Windows/Fonts/garamond.ttf",
            "timesbd.ttf",
            "garamond.ttf",
            "/System/Library/Fonts/Times Bold.ttc",
        ],
        'size': 38,
        'bold': True
    },
    'title': {
        'fonts': [
            "C:/Windows/Fonts/times.ttf",
            "C:/Windows/Fonts/georgia.ttf",
            "times.ttf",
            "/System/Library/Fonts/Times.ttc",
        ],
        'size': 42,
        'bold': False
    },
    'college': {
        'fonts': [
            "C:/Windows/Fonts/timesbd.ttf",
            "C:/Windows/Fonts/georgiab.ttf",
            "timesbd.ttf",
            "/System/Library/Fonts/Times Bold.ttc",
        ],
        'size': 38,
        'bold': True
    }
}

def get_font_path(font_list):
    for font_path in font_list:
        if os.path.exists(font_path):
            return font_path
    return None

NAME_FONT_PATH = get_font_path(FONT_CONFIG['name']['fonts'])
TITLE_FONT_PATH = get_font_path(FONT_CONFIG['title']['fonts'])
COLLEGE_FONT_PATH = get_font_path(FONT_CONFIG['college']['fonts'])

print("\n" + "=" * 70)
print("📝 PROFESSIONAL FONT CONFIGURATION")
print("=" * 70)
print(f"✓ Name Font (Bold): {os.path.basename(NAME_FONT_PATH) if NAME_FONT_PATH else 'Default'} - Size: {FONT_CONFIG['name']['size']}pt")
print(f"✓ Title Font: {os.path.basename(TITLE_FONT_PATH) if TITLE_FONT_PATH else 'Default'} - Size: {FONT_CONFIG['title']['size']}pt")
print(f"✓ College Font (Bold): {os.path.basename(COLLEGE_FONT_PATH) if COLLEGE_FONT_PATH else 'Default'} - Size: {FONT_CONFIG['college']['size']}pt")
print("=" * 70)

# Get certificate dimensions
try:
    temp_img = Image.open(TEMPLATE_PATH)
    CERT_WIDTH, CERT_HEIGHT = temp_img.size
    temp_img.close()
    print(f"✓ Certificate loaded: {CERT_WIDTH} x {CERT_HEIGHT} pixels")
    print(f"✓ Center X position: {CERT_WIDTH // 2}")
except Exception as e:
    print(f"⚠ Error loading template: {e}")
    CERT_WIDTH = 1600
    CERT_HEIGHT = 1131

# ==================== TEXT POSITIONS ====================
NAME_Y_POSITION = 500
PAPER_TITLE_Y_POSITION = 600
COLLEGE_Y_POSITION = 550

NAME_X_OFFSET = 280
TITLE_X_OFFSET = 0
COLLEGE_X_OFFSET = -180

# QR Code settings
QR_SIZE = 130  # Slightly larger QR code
QR_MARGIN = 40
QR_ID_SPACING = 15

TEXT_COLOR = (0, 0, 0)

# Create directories
os.makedirs(BASE_OUTPUT_DIR, exist_ok=True)
os.makedirs("certificate_database", exist_ok=True)

CERTIFICATES_DB = os.path.join("certificate_database", "certificates_data.json")

# ==================== FONT LOADING FUNCTIONS ====================
def load_name_font(size):
    if NAME_FONT_PATH:
        try:
            return ImageFont.truetype(NAME_FONT_PATH, size)
        except:
            pass
    try:
        return ImageFont.truetype("C:/Windows/Fonts/timesbd.ttf", size)
    except:
        return ImageFont.load_default()

def load_title_font(size):
    if TITLE_FONT_PATH:
        try:
            return ImageFont.truetype(TITLE_FONT_PATH, size)
        except:
            pass
    try:
        return ImageFont.truetype("C:/Windows/Fonts/times.ttf", size)
    except:
        return ImageFont.load_default()

def load_college_font(size):
    if COLLEGE_FONT_PATH:
        try:
            return ImageFont.truetype(COLLEGE_FONT_PATH, size)
        except:
            pass
    try:
        return ImageFont.truetype("C:/Windows/Fonts/timesbd.ttf", size)
    except:
        return ImageFont.load_default()

# ==================== JSON DATABASE FUNCTIONS ====================
def load_certificates_data():
    if os.path.exists(CERTIFICATES_DB):
        with open(CERTIFICATES_DB, 'r') as f:
            return json.load(f)
    return {}

def save_certificates_data(data):
    with open(CERTIFICATES_DB, 'w') as f:
        json.dump(data, f, indent=2)

def generate_unique_id(author_name, paper_title, serial_no):
    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    unique_part = str(uuid.uuid4())[:8].upper()
    safe_title = re.sub(r'[^a-zA-Z0-9]', '', paper_title)[:15]
    unique_id = f"CERT-{serial_no:04d}-{unique_part}"
    return unique_id

def sanitize_folder_name(name):
    safe_name = re.sub(r'[<>:"/\\|?*]', '', name)
    safe_name = safe_name.strip()
    safe_name = safe_name[:100]
    return safe_name

# ==================== READ DATA FROM EXCEL ====================
def get_papers_from_excel(excel_file):
    try:
        if not os.path.exists(excel_file):
            print(f"✗ Error: Excel file '{excel_file}' not found!")
            return None
        
        try:
            df = pd.read_excel(excel_file, header=None, engine='openpyxl')
            print(f"✓ Successfully read Excel file: {excel_file}")
            print(f"  Total rows: {len(df)}")
            print(f"  Total columns: {len(df.columns)}")
        except PermissionError:
            print(f"\n✗ ERROR: Excel file is open!")
            print("  Please close the Excel file and try again.")
            return None
        
        start_row = 0
        for i in range(len(df)):
            has_data = False
            for j in range(len(df.columns)):
                val = df.iloc[i, j]
                if pd.notna(val) and str(val).strip() not in ['', 'nan', 'NaN']:
                    has_data = True
                    break
            if has_data:
                start_row = i
                break
        
        print(f"\n📋 Data starts at row {start_row}")
        
        print("\n📋 Sample data (first 5 rows):")
        for i in range(start_row, min(start_row + 5, len(df))):
            row_data = []
            for j in range(min(9, len(df.columns))):
                val = df.iloc[i, j]
                if pd.notna(val):
                    row_data.append(str(val)[:25])
                else:
                    row_data.append("")
            print(f"  Row {i}: {row_data}")
        
        print("\n" + "=" * 70)
        print("COLUMN MAPPING (based on your Excel structure):")
        print("=" * 70)
        print("  0: Team ID")
        print("  1: Paper Title")
        print("  2: Author 1")
        print("  3: Author 2")
        print("  4: Author 3")
        print("  5: Author 4")
        print("  6: Guide Name")
        print("  7: Guide Designation")
        print("  8: College Name")
        print("=" * 70)
        
        paper_col = int(input("\nEnter PAPER TITLE column number (0-8): "))
        author_input = input("Enter AUTHOR columns (comma-separated, e.g., 2,3,4,5): ")
        author_cols = [int(x.strip()) for x in author_input.split(',')]
        college_col = int(input("Enter COLLEGE NAME column number (0-8): "))
        
        papers_dict = {}
        
        for idx in range(start_row, len(df)):
            row = df.iloc[idx]
            
            paper_title = str(row[paper_col]).strip() if pd.notna(row[paper_col]) else ""
            if not paper_title or paper_title.lower() in ['nan', '']:
                continue
            
            college_name = str(row[college_col]).strip() if pd.notna(row[college_col]) else "Unknown_College"
            if college_name.lower() in ['nan', '']:
                college_name = "Unknown_College"
            
            authors = []
            for author_col in author_cols:
                if author_col < len(row):
                    author_name = str(row[author_col]).strip() if pd.notna(row[author_col]) else ""
                    if author_name and author_name.lower() not in ['nan', ''] and author_name not in authors:
                        authors.append(author_name)
            
            if not authors:
                continue
            
            if paper_title not in papers_dict:
                papers_dict[paper_title] = {
                    'college': college_name,
                    'authors': []
                }
            
            for author in authors:
                if author not in papers_dict[paper_title]['authors']:
                    papers_dict[paper_title]['authors'].append(author)
        
        if len(papers_dict) == 0:
            print("\n✗ No valid data found!")
            return None
        
        print(f"\n✓ Processed {len(papers_dict)} papers")
        for paper, data in papers_dict.items():
            print(f"  - {paper[:50]}: {len(data['authors'])} authors")
        
        return papers_dict
    
    except Exception as e:
        print(f"✗ Error: {e}")
        return None

# ==================== PROFESSIONAL TEXT DRAWING ====================
def draw_centered_text_professional(draw, text, y_position, font_func, font_size, image_width, x_offset=0, max_width_percentage=0.8):
    max_width = image_width * max_width_percentage
    current_size = font_size
    font = font_func(current_size)
    
    try:
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
    except:
        try:
            text_width = draw.textlength(text, font=font)
        except:
            text_width = len(text) * current_size * 0.6
    
    if text_width > max_width:
        scale_factor = max_width / text_width
        new_size = int(current_size * scale_factor)
        new_size = max(28, new_size)
        font = font_func(new_size)
        try:
            bbox = draw.textbbox((0, 0), text, font=font)
            text_width = bbox[2] - bbox[0]
        except:
            text_width = len(text) * new_size * 0.6
    
    center_x = image_width // 2
    text_x = center_x - (text_width // 2) + x_offset
    
    margin = 50
    if text_x < margin:
        text_x = margin
    elif text_x + text_width > image_width - margin:
        text_x = image_width - text_width - margin
    
    direction = "RIGHT" if x_offset > 0 else "LEFT" if x_offset < 0 else "CENTER"
    print(f"      Position: X={text_x}, Y={y_position} ({direction} by {abs(x_offset)}px)")
    
    draw.text((text_x, y_position), text, fill=TEXT_COLOR, font=font)
    return font, text_x

def wrap_text(text, font, max_width, draw):
    """Wrap text into multiple lines"""
    words = text.split()
    lines = []
    current_line = []
    
    for word in words:
        test_line = ' '.join(current_line + [word])
        try:
            bbox = draw.textbbox((0, 0), test_line, font=font)
            line_width = bbox[2] - bbox[0]
        except:
            line_width = len(test_line) * font.size * 0.6
        
        if line_width <= max_width:
            current_line.append(word)
        else:
            if current_line:
                lines.append(' '.join(current_line))
            current_line = [word]
    
    if current_line:
        lines.append(' '.join(current_line))
    
    return lines

def generate_certificate_with_qr(author_name, paper_title, college_name, serial_no, debug=False):
    try:
        if not os.path.exists(TEMPLATE_PATH):
            raise FileNotFoundError(f"Template '{TEMPLATE_PATH}' not found!")
        
        unique_id = generate_unique_id(author_name, paper_title, serial_no)
        verification_link = f"{SERVER_URL}/v/{unique_id}"
        
        print(f"\n  🔗 Verification Link: {verification_link}")
        
        # Generate QR code with error correction
        qr = qrcode.QRCode(
            version=3,  # Increased version for better error correction
            box_size=10,
            border=2,
            error_correction=qrcode.constants.ERROR_CORRECT_H  # High error correction
        )
        qr.add_data(verification_link)
        qr.make(fit=True)
        qr_img = qr.make_image(fill_color="black", back_color="white")
        
        # Open template
        certificate = Image.open(TEMPLATE_PATH)
        draw = ImageDraw.Draw(certificate)
        cert_width, cert_height = certificate.size
        
        print(f"\n  📝 Generating certificate for: {author_name}")
        
        # Draw Candidate Name
        name_font, name_x = draw_centered_text_professional(
            draw, author_name, NAME_Y_POSITION, load_name_font, 
            FONT_CONFIG['name']['size'], cert_width, 
            x_offset=NAME_X_OFFSET, max_width_percentage=0.7
        )
        
        # Draw Paper Title
        title_font, title_x = draw_centered_text_professional(
            draw, paper_title, PAPER_TITLE_Y_POSITION, load_title_font,
            FONT_CONFIG['title']['size'], cert_width,
            x_offset=TITLE_X_OFFSET, max_width_percentage=0.85
        )
        
        # Draw College Name
        college_font, college_x = draw_centered_text_professional(
            draw, college_name, COLLEGE_Y_POSITION, load_college_font,
            FONT_CONFIG['college']['size'], cert_width,
            x_offset=COLLEGE_X_OFFSET, max_width_percentage=0.8
        )
        
        # Add QR code at bottom right
        qr_resized = qr_img.resize((QR_SIZE, QR_SIZE), Image.Resampling.LANCZOS)
        qr_x = cert_width - QR_SIZE - QR_MARGIN
        qr_y = cert_height - QR_SIZE - QR_MARGIN - 60  # Moved up to make room for ID
        certificate.paste(qr_resized, (qr_x, qr_y))
        
        # ==================== ADD FULL CERTIFICATE ID BELOW QR CODE ====================
        try:
            # Load font for ID
            id_font = load_title_font(14)
            id_label_font = load_title_font(12)
            
            # Create ID text
            id_label = "Certificate ID:"
            id_value = unique_id
            
            # Calculate positions
            label_bbox = draw.textbbox((0, 0), id_label, font=id_label_font)
            label_width = label_bbox[2] - label_bbox[0]
            
            # Position label above value
            label_x = qr_x + (QR_SIZE - label_width) // 2
            label_y = qr_y + QR_SIZE + 5
            draw.text((label_x, label_y), id_label, fill=(80, 80, 80), font=id_label_font)
            
            # Wrap long ID if needed (max width = QR code width)
            max_id_width = QR_SIZE
            id_lines = wrap_text(id_value, id_font, max_id_width, draw)
            
            # Draw each line of the ID
            line_height = 18
            for i, line in enumerate(id_lines):
                line_bbox = draw.textbbox((0, 0), line, font=id_font)
                line_width = line_bbox[2] - line_bbox[0]
                line_x = qr_x + (QR_SIZE - line_width) // 2
                line_y = label_y + 20 + (i * line_height)
                
                # Make sure it doesn't go off the certificate
                if line_y + line_height < cert_height - 20:
                    draw.text((line_x, line_y), line, fill=(0, 0, 0), font=id_font)
            
            print(f"\n  🔑 Certificate ID: {unique_id}")
            print(f"     Added below QR code (wrapped to {len(id_lines)} lines)")
            
            # Add "Scan to Verify" label above QR code
            scan_label = "Scan to Verify"
            scan_font = load_title_font(11)
            scan_bbox = draw.textbbox((0, 0), scan_label, font=scan_font)
            scan_width = scan_bbox[2] - scan_bbox[0]
            scan_x = qr_x + (QR_SIZE - scan_width) // 2
            scan_y = qr_y - 18
            draw.text((scan_x, scan_y), scan_label, fill=(100, 100, 100), font=scan_font)
            
        except Exception as e:
            print(f"  ⚠ Could not add certificate ID: {e}")
        
        if debug:
            center_x = cert_width // 2
            draw.line([(center_x, 0), (center_x, cert_height)], fill=(255, 0, 0), width=2)
        
        # Create paper subfolder
        paper_folder_name = sanitize_folder_name(paper_title)
        paper_folder_path = os.path.join(BASE_OUTPUT_DIR, paper_folder_name)
        os.makedirs(paper_folder_path, exist_ok=True)
        
        # Save as PDF
        safe_author_name = "".join(c for c in author_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
        pdf_filename = f"{safe_author_name}.pdf"
        pdf_path = os.path.join(paper_folder_path, pdf_filename)
        
        if certificate.mode != 'RGB':
            certificate = certificate.convert('RGB')
        certificate.save(pdf_path, "PDF", resolution=100.0)
        print(f"\n  💾 Saved: {pdf_path}")
        
        # Save to database
        cert_data = load_certificates_data()
        if unique_id not in cert_data:
            cert_data[unique_id] = {
                "unique_id": unique_id,
                "author_name": author_name,
                "paper_title": paper_title,
                "college_name": college_name,
                "serial_no": int(serial_no),
                "pdf_path": pdf_path,
                "generated_on": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "verification_link": verification_link
            }
            save_certificates_data(cert_data)
        
        return True, pdf_path, unique_id
    
    except Exception as e:
        print(f"  ❌ Error: {e}")
        import traceback
        traceback.print_exc()
        return False, str(e), None

# ==================== FLASK ROUTES ====================
@app.route('/')
def home():
    return redirect('/v')

@app.route('/v')
def verify_form():
    return '''
    <!DOCTYPE html>
    <html>
    <head>
        <title>Verify Certificate</title>
        <style>
            body {
                font-family: 'Times New Roman', Georgia, serif;
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
                min-height: 100vh;
                display: flex;
                justify-content: center;
                align-items: center;
                margin: 0;
                padding: 20px;
            }
            .container {
                background: white;
                padding: 40px;
                border-radius: 20px;
                box-shadow: 0 20px 60px rgba(0,0,0,0.3);
                max-width: 500px;
                width: 100%;
                text-align: center;
            }
            h1 {
                background: linear-gradient(135deg, #667eea, #764ba2);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
                margin-bottom: 20px;
            }
            input {
                width: 100%;
                padding: 12px;
                margin: 10px 0;
                border: 2px solid #e0e0e0;
                border-radius: 10px;
                font-size: 14px;
                font-family: monospace;
            }
            button {
                background: linear-gradient(135deg, #667eea, #764ba2);
                color: white;
                padding: 12px 30px;
                border: none;
                border-radius: 10px;
                cursor: pointer;
                font-size: 16px;
                margin-top: 10px;
            }
            button:hover {
                transform: translateY(-2px);
            }
            .info {
                margin-top: 20px;
                color: #666;
                font-size: 14px;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>🔍 Verify Certificate</h1>
            <form method="POST">
                <input type="text" name="certificate_id" placeholder="Enter Certificate ID" required>
                <button type="submit">Verify</button>
            </form>
            <div class="info">
                Enter the Certificate ID shown on your certificate
            </div>
        </div>
    </body>
    </html>
    '''

@app.route('/v', methods=['POST'])
def verify_form_post():
    certificate_id = request.form.get('certificate_id')
    if certificate_id:
        return redirect(f'/v/{certificate_id}')
    return redirect('/v')

@app.route('/v/<certificate_id>')
def verify_direct(certificate_id):
    cert_data = load_certificates_data()
    
    if certificate_id in cert_data:
        return render_template_string(VERIFICATION_RESULT_PAGE, verified=True, cert_data=cert_data[certificate_id])
    else:
        return render_template_string(VERIFICATION_RESULT_PAGE, verified=False, cert_data=None)

VERIFICATION_RESULT_PAGE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Certificate Verification</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Times New Roman', Georgia, serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
            display: flex;
            justify-content: center;
            align-items: center;
        }
        .container {
            max-width: 700px;
            width: 100%;
            background: white;
            border-radius: 20px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #667eea, #764ba2);
            color: white;
            padding: 30px;
            text-align: center;
        }
        .header h1 { font-size: 2em; margin-bottom: 10px; }
        .content { padding: 30px; }
        .verified-badge {
            text-align: center;
            padding: 20px;
            background: #d4edda;
            border-radius: 15px;
            margin-bottom: 20px;
        }
        .verified-badge h2 { color: #155724; }
        .detail-card {
            background: #f8f9fa;
            border-radius: 15px;
            padding: 20px;
        }
        .detail-row {
            padding: 12px 0;
            border-bottom: 1px solid #e0e0e0;
            display: flex;
        }
        .detail-label {
            font-weight: bold;
            width: 140px;
        }
        .detail-value { flex: 1; }
        .download-btn {
            display: inline-block;
            margin-top: 20px;
            padding: 12px 24px;
            background: #28a745;
            color: white;
            text-decoration: none;
            border-radius: 8px;
            text-align: center;
        }
        .button-group { text-align: center; }
        .not-found { text-align: center; padding: 40px; }
        .not-found h2 { color: #721c24; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🎓 Certificate Verification System</h1>
        </div>
        <div class="content">
            {% if verified %}
            <div class="verified-badge">
                <h2>✓ CERTIFICATE VERIFIED</h2>
                <p>This is an authentic and valid certificate</p>
            </div>
            <div class="detail-card">
                <div class="detail-row">
                    <div class="detail-label">Certificate ID:</div>
                    <div class="detail-value"><code>{{ cert_data.unique_id }}</code></div>
                </div>
                <div class="detail-row">
                    <div class="detail-label">Certificate Holder:</div>
                    <div class="detail-value"><strong>{{ cert_data.author_name }}</strong></div>
                </div>
                <div class="detail-row">
                    <div class="detail-label">Paper Title:</div>
                    <div class="detail-value">{{ cert_data.paper_title }}</div>
                </div>
                <div class="detail-row">
                    <div class="detail-label">College:</div>
                    <div class="detail-value">{{ cert_data.college_name }}</div>
                </div>
                <div class="detail-row">
                    <div class="detail-label">Generated On:</div>
                    <div class="detail-value">{{ cert_data.generated_on }}</div>
                </div>
            </div>
            <div class="button-group">
                <a href="/download/{{ cert_data.unique_id }}" class="download-btn">📄 Download Certificate PDF</a>
            </div>
            {% else %}
            <div class="not-found">
                <h2>✗ CERTIFICATE NOT FOUND</h2>
                <p>The certificate ID you entered is not valid.</p>
            </div>
            {% endif %}
        </div>
    </div>
</body>
</html>
"""

@app.route('/download/<certificate_id>')
def download_certificate(certificate_id):
    cert_data = load_certificates_data()
    if certificate_id in cert_data:
        pdf_path = cert_data[certificate_id].get('pdf_path')
        if pdf_path and os.path.exists(pdf_path):
            return send_file(pdf_path, as_attachment=True, download_name=os.path.basename(pdf_path))
    return "Certificate not found", 404

# ==================== MAIN GENERATION ====================
def generate_certificates():
    print("\n" + "=" * 70)
    print("🎓 PROFESSIONAL CERTIFICATE GENERATION SYSTEM")
    print("=" * 70)
    
    papers_data = get_papers_from_excel(EXCEL_FILE)
    if not papers_data:
        return False
    
    total = sum(len(data['authors']) for data in papers_data.values())
    successful = 0
    failed = 0
    serial_counter = 1
    
    print(f"\n📊 Generating {total} professional certificates...\n")
    print(f"📝 Position Settings:")
    print(f"  👤 Name: Y={NAME_Y_POSITION}, X Offset +{NAME_X_OFFSET}px (RIGHT)")
    print(f"  📄 Title: Y={PAPER_TITLE_Y_POSITION}, CENTERED")
    print(f"  🏫 College: Y={COLLEGE_Y_POSITION}, X Offset {COLLEGE_X_OFFSET}px (LEFT)")
    print(f"  🔑 QR Code: Bottom right with wrapped Certificate ID below")
    print("-" * 70)
    
    for paper_title, paper_data in papers_data.items():
        college_name = paper_data['college']
        authors = paper_data['authors']
        
        print(f"\n📁 Paper: {paper_title[:60]}")
        print(f"   🏫 College: {college_name}")
        print(f"   👥 Authors: {len(authors)}")
        
        for author in authors:
            print(f"\n  👤 Generating for: {author}")
            success, _, unique_id = generate_certificate_with_qr(author, paper_title, college_name, serial_counter, debug=False)
            if success:
                successful += 1
                print(f"    ✅ Certificate generated! ID: {unique_id}")
            else:
                failed += 1
                print(f"    ❌ Failed")
            serial_counter += 1
    
    print("\n" + "=" * 70)
    print("✅ GENERATION SUMMARY")
    print("=" * 70)
    print(f"✅ Successful: {successful}")
    print(f"❌ Failed: {failed}")
    print(f"📁 Output Directory: {BASE_OUTPUT_DIR}/")
    print(f"\n🌐 To verify certificates, the server will run at: {SERVER_URL}")
    print("=" * 70)
    return successful > 0

def start_flask_server():
    print("\n" + "=" * 70)
    print("🌐 Starting Verification Server")
    print("=" * 70)
    print(f"✅ Local access: http://localhost:8000")
    print(f"✅ Network access: {SERVER_URL}")
    print("\n📌 QR codes on certificates contain the full verification link")
    print("📌 Certificate IDs are wrapped below QR codes for readability")
    print("\n⚠️  Keep this terminal open while verifying certificates")
    print("📌 Press Ctrl+C to stop the server")
    print("=" * 70)
    
    # Open browser automatically
    webbrowser.open(f"{SERVER_URL}/v")
    
    app.run(debug=False, host='0.0.0.0', port=8000)

# ==================== MAIN ====================
if __name__ == "__main__":
    print("\n" + "=" * 70)
    print("🎓 PROFESSIONAL CERTIFICATE GENERATOR")
    print("=" * 70)
    print("\n📌 OPTIONS:")
    print("   1. Generate Certificates & Start Server")
    print("   2. Exit")
    
    choice = input("\nChoice (1-2): ").strip()
    
    if choice == "1":
        print("\n⚠️  IMPORTANT: Close the Excel file if it's open!")
        input("Press Enter after closing Excel...")
        
        if generate_certificates():
            start_flask_server()
        else:
            input("\nPress Enter to exit...")
    else:
        print("Exiting...")