# 📑 DocuMergePDF (Mail Merge to Individual PDFs)
> **The 1-Click Batch Document & Certificate Generator for Windows**  
> *Standalone Portable Windows App (.exe) • 100% Offline • Editorial Indigo Light UI*

---

## 📌 OVERVIEW
**DocuMergePDF** solves Microsoft Word's decades-old limitation: the inability to export Mail Merge records into separate, individually named PDF files.

Instead of generating one giant 500-page document or installing clunky, insecure VBA macros, users drop their Word template (`.docx`) and Excel spreadsheet (`.xlsx`), verify auto-mapped tags (`{{FullName}}`, `{{Course}}`), and generate hundreds of personalized, vector-sharp PDF documents in seconds.

---

## ⚡ KEY FEATURES
* **100% Local & Confidential**: Operates entirely offline on Windows. Zero student grades or employee payroll data ever touches the cloud.
* **Automatic Tag Matching**: Intelligently scans `{{mustache}}` placeholders and auto-matches them with Excel column headers.
* **Smart Filename Generation**: Automatically names each output PDF based on spreadsheet data (e.g., `Payslip_{Employee_ID}.pdf`).
* **Dual-Vault Freemium Licensing**:
  - Free Forever: Single-record Row 1 test generation & template tag inspection.
  - 5-Session Trial: Full batch generation of hundreds of separate PDF files.
  - $6.99 Lifetime License (or $1.99/mo): Unlimited commercial use via Gumroad API or Offline HWID Studio Key.
* **Editorial Indigo Light Interface**: A clean, accessible CustomTkinter light theme designed for educational, HR, and corporate environments.

---

## 🚀 QUICKSTART FOR DEVELOPERS
```powershell
cd c:\AllCode\Extension\Apps\Done\DocuMergePDF
pip install -r requirements.txt
python main.py
```
Build Standalone Executable:
```powershell
python build_exe.py
```

