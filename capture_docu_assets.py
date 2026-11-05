import os
import sys
import time
from PIL import Image, ImageDraw, ImageFont, ImageGrab

ROOT = r"C:\AllCode\Extension\Apps\Done\DocuMergePDF"
OUT_DIR = os.path.join(ROOT, "STORE_ASSETS")
os.makedirs(OUT_DIR, exist_ok=True)

if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

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

def create_box_art(size=1080):
    img = Image.new("RGB", (size, size), (15, 23, 42))  # Slate 900
    draw = ImageDraw.Draw(img)

    f_title = get_font(FONT_BOLD, int(size * 0.065))
    f_sub = get_font(FONT_REG, int(size * 0.032))
    f_badge = get_font(FONT_BOLD, int(size * 0.026))

    card_margin = int(size * 0.08)
    draw.rounded_rectangle([card_margin, card_margin, size - card_margin, size - card_margin], radius=int(size * 0.04), fill=(255, 255, 255))

    ico_path = os.path.join(ROOT, "app_icon.ico")
    if os.path.exists(ico_path):
        ico = Image.open(ico_path).convert("RGBA")
        ico_size = int(size * 0.32)
        ico_resized = ico.resize((ico_size, ico_size), Image.Resampling.LANCZOS)
        img.paste(ico_resized, ((size - ico_size) // 2, int(size * 0.16)), ico_resized)

    draw.text((size // 2, int(size * 0.58)), "DocuMergePDF", font=f_title, fill=(15, 23, 42), anchor="mm")
    draw.text((size // 2, int(size * 0.67)), "Word & Excel Mail Merge to Individual PDFs", font=f_sub, fill=(100, 116, 139), anchor="mm")
    draw.text((size // 2, int(size * 0.73)), "1-Click Merge • Smart Tag Mapping • 100% Offline", font=f_sub, fill=(79, 70, 229), anchor="mm")

    pill_y = int(size * 0.81)
    draw_pill(draw, int(size * 0.12), pill_y, "🔒 100% LOCAL PRIVACY", f_badge, (238, 242, 255), (67, 56, 202), px=16, py=8, radius=10)
    draw_pill(draw, int(size * 0.53), pill_y, "⚡ BATCH PDF CREATOR", f_badge, (243, 232, 255), (126, 34, 206), px=16, py=8, radius=10)

    out_file = os.path.join(OUT_DIR, f"BoxArt_{size}x{size}.png")
    img.save(out_file, "PNG")
    print(f"[✓] Created: {out_file}")

def create_poster_art(w=720, h=1080):
    img = Image.new("RGB", (w, h), (15, 23, 42))
    draw = ImageDraw.Draw(img)

    f_title = get_font(FONT_BOLD, 46)
    f_sub = get_font(FONT_REG, 24)
    f_feature = get_font(FONT_SEMIBOLD, 22)
    f_badge = get_font(FONT_BOLD, 18)

    card_left = 36
    card_top = 48
    card_right = w - 36
    card_bottom = h - 48
    draw.rounded_rectangle([card_left, card_top, card_right, card_bottom], radius=24, fill=(255, 255, 255))

    ico_path = os.path.join(ROOT, "app_icon.ico")
    if os.path.exists(ico_path):
        ico = Image.open(ico_path).convert("RGBA")
        ico_size = 240
        ico_resized = ico.resize((ico_size, ico_size), Image.Resampling.LANCZOS)
        img.paste(ico_resized, ((w - ico_size) // 2, 90), ico_resized)

    draw.text((w // 2, 380), "DocuMergePDF", font=f_title, fill=(15, 23, 42), anchor="mm")
    draw.text((w // 2, 430), "Mail Merge Word & Excel", font=f_sub, fill=(100, 116, 139), anchor="mm")
    draw.text((w // 2, 465), "to Individual PDF Documents", font=f_sub, fill=(100, 116, 139), anchor="mm")

    features = [
        "📑 Word (.docx) & Excel (.xlsx, .csv)",
        "🎯 Auto-Detect Merge Tags & Mapping",
        "⚡ Sub-Second Individual PDF Output",
        "📂 Custom Dynamic Filename Formulas",
        "🔒 100% Offline & Secure Mail Merge"
    ]

    fy = 530
    for feat in features:
        draw.rounded_rectangle([70, fy, w - 70, fy + 58], radius=12, fill=(248, 250, 252), outline=(226, 232, 240), width=1)
        draw.text((w // 2, fy + 29), feat, font=f_feature, fill=(30, 41, 59), anchor="mm")
        fy += 72

    draw.rounded_rectangle([70, 915, w - 70, 975], radius=12, fill=(79, 70, 229))
    draw.text((w // 2, 945), "WINDOWS 10 & 11 READY", font=f_badge, fill=(255, 255, 255), anchor="mm")

    out_file = os.path.join(OUT_DIR, f"PosterArt_{w}x{h}.png")
    img.save(out_file, "PNG")
    print(f"[✓] Created: {out_file}")

    img_2x = img.resize((1440, 2160), Image.Resampling.LANCZOS)
    out_2x = os.path.join(OUT_DIR, "PosterArt_1440x2160.png")
    img_2x.save(out_2x, "PNG")
    print(f"[✓] Created: {out_2x}")

def capture_real_screenshots():
    from gui.app import DocuMergePDFApp

    app = DocuMergePDFApp()
    app.state("zoomed")
    app.lift()
    app.attributes("-topmost", True)
    app.update()
    time.sleep(0.5)
    app.attributes("-topmost", False)

    def grab_window():
        app.update_idletasks()
        app.update()
        time.sleep(0.4)
        x = app.winfo_rootx()
        y = app.winfo_rooty()
        w = app.winfo_width()
        h = app.winfo_height()
        return ImageGrab.grab(bbox=(x, y, x + w, y + h))

    # 1. Merge Studio (Screenshot 1)
    app.switch_tab("merge")
    app.update()
    time.sleep(0.8)
    shot1 = grab_window()

    # 2. Live Preview Tab (Screenshot 2)
    app.switch_tab("preview")
    app.update()
    time.sleep(0.8)
    shot2 = grab_window()

    # 3. Settings & License Tab (Screenshot 3)
    app.switch_tab("settings")
    app.update()
    time.sleep(0.8)
    shot3 = grab_window()

    # 4. Back to Merge Studio (Screenshot 4)
    app.switch_tab("merge")
    app.update()
    time.sleep(0.5)
    shot4 = grab_window()

    app.destroy()

    shots = [
        (shot1, "Screenshot_1.png"),
        (shot2, "Screenshot_2.png"),
        (shot3, "Screenshot_3.png"),
        (shot4, "Screenshot_4.png")
    ]

    for img, fname in shots:
        resized = img.resize((1920, 1080), Image.Resampling.LANCZOS)
        out_p = os.path.join(OUT_DIR, fname)
        resized.save(out_p, "PNG")
        print(f"[✓] Saved {out_p}")

def clean_assets():
    keep_files = {
        "Screenshot_1.png",
        "Screenshot_2.png",
        "Screenshot_3.png",
        "Screenshot_4.png",
        "BoxArt_1080x1080.png",
        "PosterArt_720x1080.png",
        "AppTileIcon_300x300.png",
        "Icon_150x150.png",
        "Icon_71x71.png"
    }
    for item in os.listdir(OUT_DIR):
        if item not in keep_files and not item.startswith("PosterArt_1440"):
            p = os.path.join(OUT_DIR, item)
            if os.path.isfile(p):
                os.remove(p)
                print(f"Removed extra: {item}")

if __name__ == "__main__":
    create_box_art(1080)
    create_poster_art(720, 1080)
    capture_real_screenshots()
    clean_assets()
    print("[SUCCESS] All DocuMergePDF assets ready!")
