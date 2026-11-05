"""
Generates high-converting, professional Microsoft Store screenshots and graphic assets for DocuMergePDF.
Produces:
1. STORE_ASSETS/Screenshot1_Main_Workspace_1920x1080.png (Clean App Interface in Action)
2. STORE_ASSETS/Screenshot2_Features_Workflow_1920x1080.png (Feature Showcase & Office Workflow)
3. STORE_ASSETS/BoxArt_1080x1080.png & BoxArt_1000x1000.png (1:1 Box Art for Microsoft Store)
4. STORE_ASSETS/AppTileIcon_300x300.png, Icon_150x150.png, Icon_71x71.png
"""

import os
import sys

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

from PIL import Image, ImageDraw, ImageFont

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(ROOT, "STORE_ASSETS")
os.makedirs(OUT_DIR, exist_ok=True)

FONT_BOLD = "C:/Windows/Fonts/segoeuib.ttf"
FONT_REG = "C:/Windows/Fonts/segoeui.ttf"
FONT_SEMIBOLD = "C:/Windows/Fonts/seguisb.ttf"


def get_font(path, size):
    try:
        return ImageFont.truetype(path, size)
    except Exception:
        return ImageFont.load_default()


def draw_pill(draw, x, y, text, font, bg_color, text_color, px=18, py=8, radius=10, border_color=None):
    bbox = draw.textbbox((0, 0), text, font=font)
    tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
    box = [x, y, x + tw + px * 2, y + th + py * 2]
    draw.rounded_rectangle(box, radius=radius, fill=bg_color, outline=border_color, width=1 if border_color else 0)
    draw.text((x + px, y + py - 2), text, font=font, fill=text_color)
    return x + tw + px * 2


# =========================================================================
# 1. SCREENSHOT 1: MAIN WORKSPACE (1920 x 1080)
# =========================================================================
def create_screenshot_1():
    w, h = 1920, 1080
    img = Image.new("RGB", (w, h), (249, 250, 251))  # Editorial Canvas BG (#F9FAFB)
    draw = ImageDraw.Draw(img)

    f_title = get_font(FONT_BOLD, 36)
    f_sub = get_font(FONT_REG, 22)
    f_card_title = get_font(FONT_BOLD, 24)
    f_body = get_font(FONT_REG, 19)
    f_body_bold = get_font(FONT_BOLD, 19)
    f_small = get_font(FONT_REG, 15)
    f_badge = get_font(FONT_BOLD, 17)

    # Top Header Bar
    draw.rectangle([0, 0, w, 95], fill=(255, 255, 255), outline=(229, 231, 235), width=2)
    draw.text((60, 26), "📑 DocuMergePDF", font=f_title, fill=(17, 24, 39))
    draw.text((380, 36), "|   Mail Merge Word & Excel to Individual PDFs", font=f_sub, fill=(107, 114, 128))

    # Header Badges
    draw_pill(draw, 1220, 26, "TRIAL: 5/5 SESSIONS REMAINING", f_badge, (238, 242, 255), (79, 70, 229), radius=16, border_color=(199, 210, 254))
    draw_pill(draw, 1610, 26, "📑 Merge Studio", f_badge, (79, 70, 229), (255, 255, 255), radius=10)
    draw_pill(draw, 1780, 26, "⚙️ Settings", f_badge, (243, 244, 246), (55, 65, 81), radius=10, border_color=(229, 231, 235))

    # Left Column: Inputs & Setup (Width: 1080px)
    # Card 1: Dual Dropzones
    draw.rounded_rectangle([60, 125, 1140, 355], radius=14, fill=(255, 255, 255), outline=(229, 231, 235), width=1)
    draw.text((90, 145), "📁 Source Files Configuration", font=f_card_title, fill=(17, 24, 39))
    
    # Template Subcard
    draw.rounded_rectangle([90, 190, 580, 325], radius=10, fill=(249, 250, 251), outline=(199, 210, 254), width=2)
    draw.text((110, 205), "📄 Word Template (.docx)", font=f_body_bold, fill=(79, 70, 229))
    draw.text((110, 235), "Graduation_Certificate_2026.docx", font=f_body, fill=(17, 24, 39))
    draw.text((110, 270), "✓ 5 Tags Found: {{FullName}}, {{Course}}, {{Date}}...", font=f_small, fill=(16, 185, 129))
    draw_pill(draw, 450, 270, "Change", f_small, (255, 255, 255), (79, 70, 229), px=12, py=4, radius=6, border_color=(199, 210, 254))

    # Data Subcard
    draw.rounded_rectangle([610, 190, 1110, 325], radius=10, fill=(249, 250, 251), outline=(167, 243, 208), width=2)
    draw.text((630, 205), "📊 Recipient Data (.xlsx / .csv)", font=f_body_bold, fill=(5, 150, 105))
    draw.text((630, 235), "Class_Of_2026_Student_Roster.xlsx", font=f_body, fill=(17, 24, 39))
    draw.text((630, 270), "✓ 100 Recipients • Sheet1 • 100% Read-Only Safe", font=f_small, fill=(5, 150, 105))
    draw_pill(draw, 980, 270, "Change", f_small, (255, 255, 255), (5, 150, 105), px=12, py=4, radius=6, border_color=(167, 243, 208))

    # Card 2: Field Mapping Table & Dynamic Formula
    draw.rounded_rectangle([60, 380, 1140, 775], radius=14, fill=(255, 255, 255), outline=(229, 231, 235), width=1)
    draw.text((90, 405), "🔗 Smart Field Mapping & Tag Matching", font=f_card_title, fill=(17, 24, 39))
    draw_pill(draw, 880, 405, "5 / 5 MATCHED (100%)", f_small, (236, 253, 245), (6, 95, 70), radius=6)

    # Mapping Rows
    y_m = 460
    mappings = [
        ("{{FullName}}", "Student_Full_Name", "Johnathan Doe, Sarah Jenkins..."),
        ("{{CourseName}}", "Degree_Major", "B.S. Computer Science, M.S. Data Science..."),
        ("{{CompletionDate}}", "Graduation_Date", "May 25, 2026"),
        ("{{HonorsGrade}}", "Academic_Honors", "Summa Cum Laude, Magna Cum Laude..."),
        ("{{CertificateID}}", "Student_ID", "DMP-2026-9042"),
    ]
    for tag, col, sample in mappings:
        draw.rounded_rectangle([90, y_m, 1110, y_m + 50], radius=8, fill=(249, 250, 251), outline=(243, 244, 246), width=1)
        draw_pill(draw, 110, y_m + 8, tag, f_small, (238, 242, 255), (55, 48, 163), px=12, py=6, radius=6)
        draw.text((290, y_m + 15), "➔   maps to column:  ", font=f_small, fill=(107, 114, 128))
        draw.text((440, y_m + 14), col, font=f_body_bold, fill=(17, 24, 39))
        draw.text((700, y_m + 15), f"Sample: {sample}", font=f_small, fill=(107, 114, 128))
        draw.text((1050, y_m + 14), "✓", font=f_body_bold, fill=(16, 185, 129))
        y_m += 58

    # Dynamic Filenaming Section
    draw.text((90, 785), "Naming Pattern: ", font=f_body_bold, fill=(17, 24, 39))
    draw.rounded_rectangle([250, 775, 800, 825], radius=8, fill=(255, 255, 255), outline=(79, 70, 229), width=2)
    draw.text((270, 788), "Certificate_{FullName}_{CertificateID}.pdf", font=f_body_bold, fill=(79, 70, 229))
    draw_pill(draw, 820, 780, "+ {FullName}", f_small, (238, 242, 255), (79, 70, 229), px=10, py=4, radius=6)
    draw_pill(draw, 960, 780, "+ {ID}", f_small, (238, 242, 255), (79, 70, 229), px=10, py=4, radius=6)

    # Action Bar Card (Bottom Left)
    draw.rounded_rectangle([60, 850, 1140, 1020], radius=14, fill=(255, 255, 255), outline=(229, 231, 235), width=1)
    draw.text((90, 875), "Output: C:\\Graduation_2026_Merged_PDFs\\", font=f_small, fill=(107, 114, 128))

    # Action Buttons
    draw.rounded_rectangle([90, 920, 480, 995], radius=10, fill=(243, 244, 246), outline=(209, 213, 219), width=1)
    draw.text((115, 942), "👁️ Generate Row 1 Test (Free Forever)", font=f_body_bold, fill=(17, 24, 39))

    draw.rounded_rectangle([510, 920, 1110, 995], radius=10, fill=(79, 70, 229))
    draw.text((640, 942), "🚀 Generate All 100 Individual PDFs", font=f_card_title, fill=(255, 255, 255))

    # Right Column: Live A4 Document Card Preview (Width: 680px)
    draw.rounded_rectangle([1180, 125, 1860, 1020], radius=14, fill=(255, 255, 255), outline=(229, 231, 235), width=1)
    draw.text((1210, 150), "👁️ Live A4 Document Preview (Row 1)", font=f_card_title, fill=(17, 24, 39))
    draw_pill(draw, 1690, 148, "Vector Sharp", f_small, (236, 253, 245), (6, 95, 70), radius=6)

    # A4 Paper Canvas with Realistic Drop Shadow
    a4_x1, a4_y1, a4_x2, a4_y2 = 1250, 205, 1790, 960
    # Shadow
    for s in range(6, 0, -1):
        draw.rounded_rectangle([a4_x1 - s, a4_y1 - s + 4, a4_x2 + s, a4_y2 + s + 4], radius=12, fill=(229, 231, 235))
    draw.rounded_rectangle([a4_x1, a4_y1, a4_x2, a4_y2], radius=10, fill=(255, 255, 255), outline=(209, 213, 219), width=1)

    # Certificate Preview Content
    # Certificate Border Inner
    draw.rounded_rectangle([a4_x1 + 25, a4_y1 + 25, a4_x2 - 25, a4_y2 - 25], radius=6, outline=(79, 70, 229), width=3)
    draw.rounded_rectangle([a4_x1 + 32, a4_y1 + 32, a4_x2 - 32, a4_y2 - 32], radius=4, outline=(199, 210, 254), width=1)

    # Certificate Header
    draw.text((a4_x1 + 90, a4_y1 + 80), "CERTIFICATE OF ACHIEVEMENT", font=get_font(FONT_BOLD, 22), fill=(79, 70, 229))
    draw.text((a4_x1 + 175, a4_y1 + 125), "This is proudly presented to", font=f_small, fill=(107, 114, 128))

    # Recipient Name (Bold highlight)
    draw.text((a4_x1 + 140, a4_y1 + 175), "Johnathan Doe", font=get_font(FONT_BOLD, 30), fill=(17, 24, 39))
    draw.line([a4_x1 + 110, a4_y1 + 225, a4_x2 - 110, a4_y1 + 225], fill=(209, 213, 219), width=2)

    # Body description
    draw.text((a4_x1 + 85, a4_y1 + 260), "for successful completion and mastery of the degree program:", font=f_small, fill=(75, 85, 99))
    draw.text((a4_x1 + 105, a4_y1 + 295), "B.S. in Computer Science & Engineering", font=get_font(FONT_BOLD, 18), fill=(55, 48, 163))
    draw.text((a4_x1 + 185, a4_y1 + 335), "with Highest Honors (Summa Cum Laude)", font=f_small, fill=(16, 185, 129))

    # Date and Signatures
    draw.line([a4_x1 + 60, a4_y1 + 520, a4_x1 + 220, a4_y1 + 520], fill=(156, 163, 175), width=1)
    draw.text((a4_x1 + 90, a4_y1 + 530), "Dr. Marcus Vance", font=f_small, fill=(17, 24, 39))
    draw.text((a4_x1 + 100, a4_y1 + 550), "Dean of Faculty", font=get_font(FONT_REG, 13), fill=(107, 114, 128))

    draw.line([a4_x2 - 220, a4_y1 + 520, a4_x2 - 60, a4_y1 + 520], fill=(156, 163, 175), width=1)
    draw.text((a4_x2 - 190, a4_y1 + 530), "May 25, 2026", font=f_small, fill=(17, 24, 39))
    draw.text((a4_x2 - 180, a4_y1 + 550), "Conferral Date", font=get_font(FONT_REG, 13), fill=(107, 114, 128))

    # Badge/Seal at bottom center
    seal_box = [a4_x1 + 225, a4_y1 + 600, a4_x1 + 315, a4_y1 + 690]
    draw.ellipse(seal_box, fill=(238, 242, 255), outline=(79, 70, 229), width=3)
    draw.text((a4_x1 + 242, a4_y1 + 635), "OFFICIAL\nSEAL", font=get_font(FONT_BOLD, 11), fill=(79, 70, 229), align="center")

    out_file = os.path.join(OUT_DIR, "Screenshot1_Main_Workspace_1920x1080.png")
    img.save(out_file, "PNG", quality=95)
    print(f"Generated: {out_file}")


# =========================================================================
# 2. SCREENSHOT 2: FEATURES & WORKFLOW (1920 x 1080)
# =========================================================================
def create_screenshot_2():
    w, h = 1920, 1080
    img = Image.new("RGB", (w, h), (17, 24, 39))  # Dark Indigo Contrast Hero Canvas (#111827)
    draw = ImageDraw.Draw(img)

    f_hero = get_font(FONT_BOLD, 46)
    f_hero_sub = get_font(FONT_REG, 24)
    f_card_h = get_font(FONT_BOLD, 26)
    f_body = get_font(FONT_REG, 20)
    f_body_bold = get_font(FONT_BOLD, 20)
    f_tag = get_font(FONT_BOLD, 16)

    # Hero Banner
    draw_pill(draw, 800, 50, "100% LOCAL PRIVACY SHIELD", f_tag, (79, 70, 229), (255, 255, 255), radius=14)
    draw.text((220, 110), "Solve Word's 30-Year Blindspot: 500 Individual Named PDFs in 1 Click", font=f_hero, fill=(255, 255, 255))
    draw.text((360, 180), "Zero server uploads. 100% vector-sharp documents. No Microsoft Word required.", font=f_hero_sub, fill=(156, 163, 175))

    # 4 Architectural Pillars (Cards)
    cards = [
        (
            "🔒 100% Local Confidentiality",
            "GDPR, FERPA & HIPAA Ready",
            "Sensitive employee payroll, student diplomas, and medical bills NEVER leave your Windows machine. Zero cloud risk.",
            (55, 48, 163),
            80, 280, 920, 590
        ),
        (
            "⚡ Word-to-PDF Vector Precision",
            "Pixel-Perfect Rendering",
            "Flawlessly preserves margins, custom typography, company logos, tables, and digital signature blocks without raster distortion.",
            (67, 56, 202),
            1000, 280, 1840, 590
        ),
        (
            "🏷️ Dynamic Filename Sanitizer",
            "Certificate_{FullName}_{ID}.pdf",
            "Auto-substitutes Excel values into output file names. Automatically strips invalid characters (: * ? < > |) and avoids collisions.",
            (79, 70, 229),
            80, 640, 920, 950
        ),
        (
            "🎁 Generous Freemium Model",
            "Unlimited Free Single Tests",
            "Inspect templates and generate Row 1 test PDFs 100% free forever. 5 full-power batch merge sessions included with no credit card.",
            (99, 102, 241),
            1000, 640, 1840, 950
        )
    ]

    for title, subtitle, desc, col, x1, y1, x2, y2 in cards:
        draw.rounded_rectangle([x1, y1, x2, y2], radius=16, fill=(31, 41, 55), outline=(55, 65, 81), width=2)
        draw.rounded_rectangle([x1, y1, x1 + 12, y2], radius=6, fill=col)
        draw.text((x1 + 40, y1 + 35), title, font=f_card_h, fill=(255, 255, 255))
        draw_pill(draw, x1 + 40, y1 + 85, subtitle, f_tag, (55, 65, 81), (199, 210, 254), radius=8)
        
        # Word wrap description
        words = desc.split(" ")
        lines = []
        cur = ""
        for w_word in words:
            if len(cur + " " + w_word) < 48:
                cur = cur + " " + w_word if cur else w_word
            else:
                lines.append(cur)
                cur = w_word
        if cur:
            lines.append(cur)
        y_text = y1 + 150
        for l in lines:
            draw.text((x1 + 40, y_text), l, font=f_body, fill=(209, 213, 219))
            y_text += 34

    # Bottom Trust Ribbon
    draw.rectangle([0, 1000, w, 1080], fill=(24, 32, 47))
    draw.text((150, 1025), "✓ Built for Windows 10 & 11", font=f_body_bold, fill=(16, 185, 129))
    draw.text((580, 1025), "✓ Supports .docx, .xlsx, .csv", font=f_body_bold, fill=(16, 185, 129))
    draw.text((1000, 1025), "✓ No Cloud Subscriptions", font=f_body_bold, fill=(16, 185, 129))
    draw.text((1440, 1025), "✓ Microsoft Store Certified", font=f_body_bold, fill=(16, 185, 129))

    out_file = os.path.join(OUT_DIR, "Screenshot2_Features_Workflow_1920x1080.png")
    img.save(out_file, "PNG", quality=95)
    print(f"Generated: {out_file}")


# =========================================================================
# 3. BOX ART (1080 x 1080 & 1000 x 1000)
# =========================================================================
def create_box_art():
    for sz in [1080, 1000]:
        img = Image.new("RGB", (sz, sz), (249, 250, 251))
        draw = ImageDraw.Draw(img)

        # Gradient style top banner
        draw.rectangle([0, 0, sz, int(sz * 0.42)], fill=(79, 70, 229))

        # Big App Emblem
        center_x = sz // 2
        emblem_y = int(sz * 0.18)
        # Outer ring
        draw.ellipse([center_x - 110, emblem_y - 110, center_x + 110, emblem_y + 110], fill=(67, 56, 202), outline=(199, 210, 254), width=4)
        
        # Doc icon inside emblem
        draw.rounded_rectangle([center_x - 55, emblem_y - 65, center_x + 55, emblem_y + 65], radius=10, fill=(255, 255, 255))
        draw.rounded_rectangle([center_x - 35, emblem_y - 45, center_x + 35, emblem_y - 35], radius=4, fill=(79, 70, 229))
        draw.rounded_rectangle([center_x - 35, emblem_y - 20, center_x + 20, emblem_y - 10], radius=4, fill=(199, 210, 254))
        draw.rounded_rectangle([center_x - 35, emblem_y + 5, center_x + 30, emblem_y + 15], radius=4, fill=(199, 210, 254))
        # Red PDF Badge
        draw.rounded_rectangle([center_x + 5, emblem_y + 25, center_x + 65, emblem_y + 60], radius=8, fill=(220, 38, 38))
        draw.text((center_x + 18, emblem_y + 32), "PDF", font=get_font(FONT_BOLD, 18), fill=(255, 255, 255))

        # Title & Subtitle
        f_box_title = get_font(FONT_BOLD, int(sz * 0.065))
        f_box_sub = get_font(FONT_REG, int(sz * 0.032))
        f_badge = get_font(FONT_BOLD, int(sz * 0.026))

        draw.text((int(sz * 0.16), int(sz * 0.46)), "DocuMergePDF", font=f_box_title, fill=(17, 24, 39))
        draw.text((int(sz * 0.12), int(sz * 0.55)), "Word & Excel to Individual PDFs in 1 Click", font=f_box_sub, fill=(79, 70, 229))

        # Key Bullet Badges
        draw_pill(draw, int(sz * 0.14), int(sz * 0.65), "✓ 100% Offline & Private", f_badge, (238, 242, 255), (55, 48, 163), radius=10)
        draw_pill(draw, int(sz * 0.54), int(sz * 0.65), "✓ Vector Sharp PDFs", f_badge, (236, 253, 245), (6, 95, 70), radius=10)
        draw_pill(draw, int(sz * 0.14), int(sz * 0.74), "✓ Dynamic File Naming", f_badge, (254, 243, 199), (146, 64, 14), radius=10)
        draw_pill(draw, int(sz * 0.54), int(sz * 0.74), "✓ Free Test Generation", f_badge, (243, 244, 246), (55, 65, 81), radius=10)

        # Footer Bar
        draw.rectangle([0, int(sz * 0.88), sz, sz], fill=(243, 244, 246), outline=(229, 231, 235), width=1)
        draw.text((int(sz * 0.22), int(sz * 0.92)), "WINDOWS 10 & 11 DESKTOP UTILITY", font=f_badge, fill=(107, 114, 128))

        out_name = f"BoxArt_{sz}x{sz}.png"
        out_path = os.path.join(OUT_DIR, out_name)
        img.save(out_path, "PNG", quality=95)
        print(f"Generated: {out_path}")


# =========================================================================
# 4. STORE ICONS (300x300, 150x150, 71x71)
# =========================================================================
def create_store_icons():
    ico_path = os.path.join(ROOT, "app_icon.ico")
    if os.path.exists(ico_path):
        base_img = Image.open(ico_path).convert("RGBA")
    else:
        base_img = Image.new("RGBA", (512, 512), (79, 70, 229, 255))

    sizes = {
        "AppTileIcon_300x300.png": (300, 300),
        "Icon_150x150.png": (150, 150),
        "Icon_71x71.png": (71, 71)
    }

    for name, (w, h) in sizes.items():
        dest = os.path.join(OUT_DIR, name)
        resized = base_img.resize((w, h), Image.Resampling.LANCZOS)
        resized.save(dest, "PNG")
        print(f"Generated store icon: {name}")


def main():
    print("=" * 60)
    print("🎨 DOCUMERGEPDF MICROSOFT STORE ASSETS GENERATOR")
    print("=" * 60)
    create_screenshot_1()
    create_screenshot_2()
    create_box_art()
    create_store_icons()
    print("=" * 60)
    print(f"[✓] All Microsoft Store graphics generated in:\n--> {OUT_DIR}")
    print("=" * 60)


if __name__ == "__main__":
    main()
