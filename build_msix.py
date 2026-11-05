"""
Automated MSIX Package Builder for DocuMergePDF.
Packages DocuMergePDF.exe into DocuMergePDF.msix matching Microsoft Partner Center Product Identity.
"""

import os
import sys
import shutil
import subprocess
from PIL import Image

ROOT = os.path.dirname(os.path.abspath(__file__))
BUILD_DIR = os.path.join(ROOT, "MSIX_PACKAGE_BUILD")
ASSETS_DIR = os.path.join(BUILD_DIR, "Assets")
OUT_MSIX = os.path.join(ROOT, "DocuMergePDF.msix")

# Microsoft Partner Center Identity values (from Product Identity)
IDENTITY_NAME = "SereneWorks.DocuMergePDF"
IDENTITY_PUBLISHER = "CN=EF5F47C7-3356-415D-9F39-F06A8678DF1C"
PUBLISHER_DISPLAY_NAME = "SereneWorks"
PACKAGE_VERSION = "1.0.0.0"

EXE_SRC = os.path.join(ROOT, "DocuMergePDF.exe")
if not os.path.exists(EXE_SRC):
    dist_exe = os.path.join(ROOT, "dist", "DocuMergePDF.exe")
    if os.path.exists(dist_exe):
        EXE_SRC = dist_exe

# Path to makeappx.exe
MAKEAPPX = r"C:\Program Files (x86)\Windows Kits\10\bin\10.0.26100.0\x64\makeappx.exe"
if not os.path.exists(MAKEAPPX):
    MAKEAPPX = r"C:\Program Files (x86)\Windows Kits\10\bin\10.0.22621.0\x64\makeappx.exe"
assert os.path.exists(MAKEAPPX), f"makeappx.exe not found!"


def create_icon_assets():
    os.makedirs(ASSETS_DIR, exist_ok=True)
    ico_path = os.path.join(ROOT, "app_icon.ico")
    
    if os.path.exists(ico_path):
        base_img = Image.open(ico_path).convert("RGBA")
    else:
        base_img = Image.new("RGBA", (512, 512), (79, 70, 229, 255))

    sizes = {
        "StoreLogo.png": (50, 50),
        "Square44x44Logo.png": (44, 44),
        "Square71x71Logo.png": (71, 71),
        "Square150x150Logo.png": (150, 150),
        "Wide310x150Logo.png": (310, 150),
        "SplashScreen.png": (620, 300),
    }

    for name, (w, h) in sizes.items():
        dest = os.path.join(ASSETS_DIR, name)
        canvas = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        min_dim = min(w, h)
        logo_sized = base_img.resize((min_dim, min_dim), Image.Resampling.LANCZOS)
        offset_x = (w - min_dim) // 2
        offset_y = (h - min_dim) // 2
        canvas.paste(logo_sized, (offset_x, offset_y), logo_sized)
        canvas.save(dest, "PNG")
        print(f"Generated asset: {name} ({w}x{h})")


def create_manifest():
    manifest_path = os.path.join(BUILD_DIR, "AppxManifest.xml")
    xml_content = f"""<?xml version="1.0" encoding="utf-8"?>
<Package
  xmlns="http://schemas.microsoft.com/appx/manifest/foundation/windows10"
  xmlns:uap="http://schemas.microsoft.com/appx/manifest/uap/windows10"
  xmlns:rescap="http://schemas.microsoft.com/appx/manifest/foundation/windows10/restrictedcapabilities"
  xmlns:desktop="http://schemas.microsoft.com/appx/manifest/desktop/windows10">

  <Identity
    Name="{IDENTITY_NAME}"
    Publisher="{IDENTITY_PUBLISHER}"
    Version="{PACKAGE_VERSION}"
    ProcessorArchitecture="x64" />

  <Properties>
    <DisplayName>DocuMergePDF</DisplayName>
    <PublisherDisplayName>{PUBLISHER_DISPLAY_NAME}</PublisherDisplayName>
    <Logo>Assets\\StoreLogo.png</Logo>
    <Description>Mail Merge Word &amp; Excel to Individual PDFs. Fast, secure, and 100% offline.</Description>
  </Properties>

  <Resources>
    <Resource Language="en-us" />
  </Resources>

  <Dependencies>
    <TargetDeviceFamily Name="Windows.Desktop" MinVersion="10.0.17763.0" MaxVersionTested="10.0.22621.0" />
  </Dependencies>

  <Capabilities>
    <rescap:Capability Name="runFullTrust" />
  </Capabilities>

  <Applications>
    <Application Id="DocuMergePDF"
      Executable="DocuMergePDF.exe"
      EntryPoint="Windows.FullTrustApplication">
      <uap:VisualElements
        DisplayName="DocuMergePDF"
        Description="DocuMergePDF - Mail Merge Word &amp; Excel to Individual PDFs"
        BackgroundColor="transparent"
        Square150x150Logo="Assets\\Square150x150Logo.png"
        Square44x44Logo="Assets\\Square44x44Logo.png">
        <uap:DefaultTile
          Wide310x150Logo="Assets\\Wide310x150Logo.png"
          Square71x71Logo="Assets\\Square71x71Logo.png" />
        <uap:SplashScreen Image="Assets\\SplashScreen.png" />
      </uap:VisualElements>
    </Application>
  </Applications>
</Package>
"""
    with open(manifest_path, "w", encoding="utf-8") as f:
        f.write(xml_content.strip())
    print(f"Created AppxManifest.xml for: {IDENTITY_NAME}")


def build_msix():
    assert os.path.exists(EXE_SRC), f"Executable not found at: {EXE_SRC}. Run build_exe.py first!"

    if os.path.exists(BUILD_DIR):
        shutil.rmtree(BUILD_DIR)
    os.makedirs(BUILD_DIR, exist_ok=True)

    # 1. Copy executable
    dest_exe = os.path.join(BUILD_DIR, "DocuMergePDF.exe")
    shutil.copy2(EXE_SRC, dest_exe)
    print(f"Copied executable: {dest_exe} ({os.path.getsize(dest_exe) / (1024*1024):.2f} MB)")

    # 2. Create assets & manifest
    create_icon_assets()
    create_manifest()

    # 3. Call makeappx.exe
    if os.path.exists(OUT_MSIX):
        try:
            os.remove(OUT_MSIX)
        except Exception:
            pass

    cmd = [
        MAKEAPPX,
        "pack",
        "/d", BUILD_DIR,
        "/p", OUT_MSIX,
        "/o"  # Overwrite
    ]
    print(f"Running makeappx: {' '.join(cmd)}")
    res = subprocess.run(cmd, capture_output=True, text=True)
    print(res.stdout)
    if res.returncode != 0:
        print(res.stderr)
        sys.exit(res.returncode)

    print("=" * 60)
    print(f"[SUCCESS] Built MSIX Package: {OUT_MSIX}")
    print(f"Package Size: {os.path.getsize(OUT_MSIX) / (1024*1024):.2f} MB")
    print("=" * 60)


if __name__ == "__main__":
    build_msix()

