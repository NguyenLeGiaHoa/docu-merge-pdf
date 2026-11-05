"""
DocuMergePDF - Desktop Application Entry Point
Standalone Portable Windows Utility: Mail Merge Word & Excel to Individual PDFs
"""

import sys
import os

# Ensure current project directory is in python path
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

# Enable high-DPI awareness on Windows 10 & 11 for razor-sharp typography
if sys.platform == "win32":
    try:
        import ctypes
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)  # Per-monitor v2
        except Exception:
            try:
                ctypes.windll.shcore.SetProcessDpiAwareness(1)  # Per-monitor v1
            except Exception:
                ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

from gui.app import DocuMergePDFApp


def main():
    app = DocuMergePDFApp()
    app.mainloop()


if __name__ == "__main__":
    main()

