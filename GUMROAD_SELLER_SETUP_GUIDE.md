# 🚀 DocuMergePDF — Gumroad & Microsoft Store Seller Setup Guide
> **Complete Commercial Blueprint for Publishing and Selling DocuMergePDF**

---

## 🏬 1. MICROSOFT STORE PUBLISHING WORKFLOW

### Product Identity in Microsoft Partner Center:
* **Package / Identity Name**: `SereneWorks.DocuMergePDF`
* **Publisher**: `CN=EF5F47C7-3356-415D-9F39-F06A8678DF1C`
* **Publisher Display Name**: `SereneWorks`
* **Product Name**: `DocuMergePDF`
* **Package Version**: `1.0.0.0`
* **Target OS**: Windows 10 (17763.0+) and Windows 11

### Upload Steps:
1. Run `python build_msix.py` to produce `DocuMergePDF.msix`.
2. Open [Microsoft Partner Center](https://partner.microsoft.com/dashboard).
3. Navigate to your app submission $\rightarrow$ **Packages**.
4. Drag and drop `DocuMergePDF.msix`.
5. Upload Store Assets from the `STORE_ASSETS/` directory:
   - `BoxArt_1080x1080.png` (1:1 Box Art)
   - `Screenshot1_Main_Workspace_1920x1080.png`
   - `Screenshot2_Features_Workflow_1920x1080.png`
   - `AppTileIcon_300x300.png`
6. Click **Submit to the Store**. Microsoft automatically signs the package and completes certification in 24–48 hours.

---

## 🛒 2. GUMROAD SELLER SETUP

### Product A: Free Trial Version (Lead Magnet)
* **Title**: `DocuMergePDF - Mail Merge Word & Excel to Individual PDFs (Free Trial)`
* **Price**: `$0+` (Pay what you want)
* **File to Upload**: `GUMROAD_UPLOAD_PACKAGES/Tier1_FREE_TRIAL_Upload_To_Gumroad.zip`
* **Permalink**: `documergepdf-trial`

### Product B: Lifetime Commercial License ($6.99)
* **Title**: `DocuMergePDF - Lifetime Commercial License`
* **Price**: `$6.99` (One-time)
* **File to Upload**: `GUMROAD_UPLOAD_PACKAGES/Tier2_LIFETIME_LICENSE_Upload_To_Gumroad.zip`
* **Permalink**: `documergepdf`
* **License Settings**:
  - Check **"Generate a unique license key for each sale"**.
  - Set uses count: 3 (allows users to activate on up to 3 personal/work PCs).

### Product C: 30-Day Monthly Project Pass ($1.99/mo)
* **Title**: `DocuMergePDF - 30-Day Project Pass`
* **Price**: `$1.99 / month` (Recurring subscription)
* **Permalink**: `documergepdf-monthly`
* **License Settings**:
  - Check **"Generate a unique license key for each sale"**.

---

## 🔑 3. AIR-GAPPED / OFFLINE CORPORATE ACTIVATION
For schools, government agencies, or corporate PCs without internet access:
1. Ask the customer for their **Hardware ID (HWID)** displayed in the DocuMergePDF **Settings & License** tab.
2. In your terminal, run the included seller tool:
   ```bash
   python generate_offline_key.py <CUSTOMER_HWID> [lifetime|monthly]
   ```
3. Copy the generated key (format `DMP-LIFETIME-XXXXXXXX-XXXXXXXXXXXXXXXX` or `DMP-MONTHLY-XXXXXXXX-XXXXXXXXXXXXXXXX`) and email it to the customer.
4. The key validates cryptographically on their machine with zero internet traffic required.

---

## 📧 4. INTERNATIONAL CUSTOMER SUPPORT TEMPLATES (CANNED REPLIES)

### Template A: Customer Asking How to Activate Gumroad License Key
```text
Hi [Customer Name],

Thank you for purchasing DocuMergePDF!

Here is how to activate your Commercial License:
1. Open DocuMergePDF on your Windows PC.
2. Click the "⚙️ Settings & License" tab (or click the trial status pill in the top header).
3. Paste your Gumroad license key into the license activation box.
4. Click "✓ Activate License".
5. The badge will instantly turn gold ("🏆 LIFETIME LICENSE ACTIVE" or "⚡ 30-DAY PASS ACTIVE").

All batch limits are now unlocked for your workstation!

Best regards,
DocuMergePDF Support Team
```

### Template B: Customer Asking How to Format Word Placeholders
```text
Hi [Customer Name],

DocuMergePDF makes placeholder templating seamless!

In your Microsoft Word (.docx) template, you can use any of the following standard placeholder styles:
- Double curly braces (Recommended): {{FullName}}, {{Address}}, {{Salary}}
- Standard Word MergeFields: Insert > Quick Parts > Field > MergeField
- Chevrons: «FullName», «Address»
- Single braces: {FullName}, {Date}

Make sure the names in your template match the column headers in your Excel spreadsheet (.xlsx, .csv). DocuMergePDF will automatically match them for you upon drag & drop!

Best regards,
DocuMergePDF Support Team
```

### Template C: Customer Requesting Offline Key for Air-Gapped Machine
```text
Hi [Customer Name],

Thank you for purchasing DocuMergePDF!

Here is your permanent Offline Studio License Key for your air-gapped PC:

  License Key: [INSERT_GENERATED_KEY]

Activation Instructions:
1. Open DocuMergePDF on your air-gapped PC.
2. Click the "⚙️ Settings & License" tab.
3. Paste the key above into the license box and click "✓ Activate License".
4. The license will validate cryptographically and unlock completely offline.

Best regards,
DocuMergePDF Support Team
```


