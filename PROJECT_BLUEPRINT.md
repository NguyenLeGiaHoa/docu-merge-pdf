# 📑 MASTER BLUEPRINT: DocuMergePDF (Mail Merge to Individual PDFs)
> **AI Pair Programming Blueprint for Antigravity (Advanced Agentic Coding Engine)**  
> **Target Production Model**: Ultra-Fast Standalone Windows Desktop Utility (Local-First, Zero Server Costs, Python + CustomTkinter Editorial Indigo UI)  
> **Commercial Model**: Freemium (Unlimited Single Document Test Generation Free Forever + 5 Full Batch Merge Sessions Free Trial + $6.99 One-Time Lifetime License or $1.99/mo Monthly Pass)  
> **Market Target**: 400,000+ Educators, HR Specialists, Event Coordinators, Billing Admins needing individual, personalized PDF documents from an Excel sheet.

---

## 📋 SECTION 1: INVOCATION PROMPT FOR ANTIGRAVITY
*(Copy and paste this exact prompt into a new conversation to build the entire production-ready application from scratch):*

```text
You are Antigravity, the Google DeepMind agentic AI coding assistant.
Read the master blueprints and architectural compliance rules located at:
1. `c:\AllCode\Extension\DocuMergePDF\PROJECT_BLUEPRINT.md`
2. `c:\AllCode\Extension\DocuMergePDF\TECHNICAL_ARCHITECTURE_RULES.md`
3. `c:\AllCode\Extension\DocuMergePDF\USER_GUIDE_FREE_TRIAL.md`
4. `c:\AllCode\Extension\DocuMergePDF\USER_GUIDE_LIFETIME.md`

Your task is to build the complete, production-ready, commercial-grade Windows desktop application "DocuMergePDF" in this workspace.

MANDATORY DIRECTIVES BEFORE WRITING ANY LINE OF CODE:
1. NATIVE WORD-TO-PDF MERGE PIPELINE: Non-tech users want pixel-perfect document rendering. The application must support standard Word templates (.docx) with mustache tags {{FullName}}, {{Grade}}, {{Salary}}, or standard Word merge fields, merging them with Excel (.xlsx) / CSV rows, and rendering each row into an independent, vector-sharp PDF file.
2. ZERO CLOUD & ZERO SERVER RISK: All document generation and PDF rendering must happen 100% locally on the user's machine (using docx-mailmerge / python-docx + local headless LibreOffice or Word COM bridge). Confidential employee payslips and student grades must NEVER leave the machine.
3. EXCLUSIVE EDITORIAL INDIGO & PAPER LIGHT UI: Adhere strictly to the "Editorial Indigo & Paper Light" theme specified in TECHNICAL_ARCHITECTURE_RULES.md. Use clean paper backgrounds (#F9FAFB canvas, #FFFFFF card containers), authoritative indigo accents (#4F46E5 / #6366F1), high-contrast slate text (#0F172A), and a split-pane layout showing field mapping on the left and a live A4 preview card on the right.
4. FREEMIUM & DUAL-VAULT LICENSING:
   - Free Forever Tier: Generating a single test record (Row 1 test export) and inspecting template tags {{Placeholder}} is 100% UNLIMITED and FREE FOREVER.
   - Trial Tier: 5 Full-Power Batch Merge sessions (generating up to 1,000+ separate named PDFs in one click).
   - Lifetime Tier: $6.99 one-time purchase (or $1.99/mo) unlocked via Gumroad API or air-gapped Offline Studio Key. Dual-vault persistence across Windows Registry and LocalAppData.
5. ADVANCED DYNAMIC FILENAMING: Allow users to name output files dynamically from any column value, e.g. `Certificate_{Student_Name}_{ID}.pdf` or `Payslip_{Employee_ID}_{Month}.pdf` with automatic sanitization of illegal characters.

Execute all stages autonomously without stopping:
1. Initialize project workspace, virtual environment configs, and requirements.txt.
2. Build core merge engine (template placeholder detector, row-by-row document renderer, PDF export pipeline).
3. Build dynamic filename generator with regex validation and collision avoidance.
4. Build dual-vault licensing engine (Registry + LocalAppData, HWID binding, 5-session trial counter, Gumroad API).
5. Build Editorial Indigo Light GUI (CustomTkinter light mode, dual-dropzone for Word + Excel, field mapping list, A4 preview widget).
6. Build automated test suite covering 500-recipient batch runs, missing placeholders, unicode names, and trial defenses.
7. Package standalone portable executable (build_exe.py) and Gumroad distribution packages.
```

---

## 🎯 SECTION 2: EXECUTIVE SUMMARY & COMMERCIAL POSITIONING

### 2.1 The Market Pain Point: Microsoft Word's 30-Year-Old Blindspot
Microsoft Word has included "Mail Merge" since the 1990s. However, Word suffers from one glaring, frustrating limitation that has plagued office workers for decades:
* Word's built-in Mail Merge only outputs **one single, massive 500-page document**, or sends directly to a physical printer!
* Word **CANNOT natively export 500 separate, individual PDF files** named dynamically by recipient (e.g. `Payslip_JohnDoe.pdf`, `Payslip_JaneSmith.pdf`).

**Who Struggles With This?**
1. **Human Resources**: Monthly payslips, bonus letters, employment contracts. Sending one giant 200-page PDF means employees could see each other's private salaries! HR needs 200 individual password-protectable files.
2. **Schools & Universities**: Generating 500 graduation diplomas, completion certificates, and student report cards.
3. **Event Planners & Non-Profits**: Personalized invitations, ticketing badges, donor thank-you letters.
4. **Sales & Freelancers**: Monthly retainers, billing statements, customized quotes.

**Current Inadequate Alternatives**:
* *Doug Robbins' MergeTools*: A legacy Word macro written in VBA. Requires installing into hidden Word startup folders, triggers macro security warnings, and breaks whenever Office updates.
* *Enterprise Add-Ins (Mail Merge Toolkit)*: Cost **$99 – $199/year**, require Outlook enterprise integrations, and overwhelm non-tech users with complex setup dialogs.

### 2.2 The DocuMergePDF Solution
DocuMergePDF makes batch individual PDF creation effortless in 3 simple steps:
1. **Drop Word Template & Excel Data**: Drag and drop your `.docx` certificate or letter template and your `.xlsx` contact list.
2. **Review Auto-Mapped Fields**: DocuMergePDF instantly scans tags like `{{Name}}`, `{{Course}}`, `{{Date}}` and auto-matches them with Excel column headers.
3. **Set Filename & Merge**: Choose how to name files (e.g., `Certificate_{Name}.pdf`) and click **"Generate All PDFs"**. 500 individual, high-resolution PDFs are generated into a neat folder in seconds.

### 2.3 Pricing & Commercial Math
* **Price Point**: **$6.99 One-Time Lifetime License** (or **$1.99/mo**).
  - Highly accessible: Educators and HR clerks will gladly pay $6.99 once (or $1.99 for a project) to eliminate 4 hours of tedious manual cutting and saving.
* **The Freemium Hook**:
  - Free Forever: Inspecting templates, auto-mapping fields, and exporting **Row 1 Test PDF** is 100% free with unlimited runs. Users can test and verify their design forever.
  - Paywall Trigger: Batch exporting rows 2 to 500+ unlocks via the 5-session Free Trial, then prompts for the $6.99 Lifetime License.
* **Sales Math**:
  - **12 to 20 sales per month** = **$300 – $500 MRR**.
  - High organic search volume across Google and Microsoft Store for keywords like `mail merge to individual pdfs`, `split mail merge into separate pdf files`, `excel to individual pdfs`.

---

## 🎨 SECTION 3: VISUAL IDENTITY & UI DESIGN (EDITORIAL INDIGO & PAPER)

* **Canvas Background**: `#F9FAFB` (Gray-50, soft paper canvas).
* **Card & Panel Background**: `#FFFFFF` (Pure white containers with `#E5E7EB` borders).
* **Primary Accent / Action Color**: `#4F46E5` (Authoritative Indigo-600, evoking professional typography and trust).
* **Hover State**: `#4338CA` (Indigo-700).
* **Secondary Highlights**: `#6366F1` (Indigo-500) and `#EEF2FF` (Indigo-50 soft pill backgrounds).
* **Text Primary**: `#111827` (Gray-900, dark crisp editorial readability).
* **Text Secondary**: `#6B7280` (Gray-500, clean metadata labels).
* **A4 Preview Card**: Styled with realistic document drop-shadow (`box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1)`).

---

## 🏗️ SECTION 4: SYSTEM ARCHITECTURE & FILE TREE

```
c:\AllCode\Extension\DocuMergePDF\
├── PROJECT_BLUEPRINT.md               # Master blueprint and AI instructions
├── TECHNICAL_ARCHITECTURE_RULES.md    # Strict compliance, document fidelity, UI rules
├── USER_GUIDE_FREE_TRIAL.md           # End-user onboarding and trial guide
├── USER_GUIDE_LIFETIME.md             # Commercial activation and license perks
├── README.md                          # Quickstart developer overview
├── requirements.txt                   # python-docx, docx-mailmerge, pandas, openpyxl, customtkinter
├── build_exe.py                       # Standalone Windows executable compiler
├── package_gumroad.py                 # Automated packaging script for Gumroad/Store
├── reset_trial.py                     # Developer utility to reset 5-run trial
├── core/
│   ├── __init__.py
│   ├── template_inspector.py          # Detects {{mustache}} and Word MergeField tags
│   ├── data_reader.py                 # Reads Excel/CSV rows with data type normalization
│   ├── document_merger.py             # Replaces tags, clones formatting, handles tables
│   ├── pdf_converter.py               # Local headless PDF renderer (Word COM or local engine)
│   └── licensing.py                   # Dual-Vault (Registry + LocalAppData), HWID, Gumroad API
├── gui/
│   ├── __init__.py
│   ├── app.py                         # Main CustomTkinter light window
│   ├── views_merge.py                 # Dual dropzones, mapping table, progress bar
│   ├── views_preview.py               # Live A4 page visualizer showing Row 1 sample
│   ├── views_settings.py              # Dynamic naming builder, license management
│   └── styles.py                      # Editorial Indigo Light theme constants
├── demo_data/
│   ├── Certificate_Template.docx      # Sample graduation diploma template with {{tags}}
│   └── Student_Roster_100.xlsx        # Sample spreadsheet with 100 recipient rows
└── tests/
    ├── test_template_inspector.py     # Tag detection and missing variable alerts
    ├── test_document_merger.py        # Replacement fidelity and table row replication
    └── test_licensing.py              # 5-session trial bounds and anti-rollback checks
```

