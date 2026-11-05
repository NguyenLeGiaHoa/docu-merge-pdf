# 🛡️ DOCUMERGEPDF - TECHNICAL ARCHITECTURE & AI COMPLIANCE RULES
> **MANDATORY INSTRUCTION FILE FOR ALL FUTURE AI AGENTS & DEVELOPERS**  
> *If you are an AI assistant (Antigravity, Claude, ChatGPT, or other LLMs) modifying or implementing this codebase, you MUST strictly adhere to the architecture, rules, and invariants documented in this file. DO NOT REFACTOR OR VIOLATE THESE CORE PRINCIPLES.*

---

## 🏛️ 1. CORE ARCHITECTURAL INVARIANTS & PRIVACY SHIELDS

### Rule 1: 100% Local Confidential Document Processing
* **The Reason**: User documents frequently contain private personal information: Employee salaries, Social Security numbers, student grades, or medical details. Any upload to a remote cloud server violates GDPR, FERPA, and corporate non-disclosure agreements.
* **Invariant**:
  - All document parsing, placeholder replacement, and PDF rendering **MUST OCCUR 100% LOCALLY** on the user's Windows PC.
  - Zero document contents, recipient names, or text streams may ever be transmitted across the network.

### Rule 2: Non-Destructive Source File Protection
* **The Reason**: Users value their original master Word templates and Excel rosters.
* **Invariant**:
  - The source `.docx` and `.xlsx` files must be opened strictly in read-only mode.
  - Generated PDFs must be saved into a dedicated user-selected output directory (e.g. `[TemplateName]_Merged_[Timestamp]/`).
  - Filename collisions must automatically append `_01`, `_02` instead of silently overwriting existing files.

### Rule 3: Graceful Handling of Missing Template Placeholders
* **The Reason**: In real-world rosters, some rows have blank cells (e.g., middle names, optional apartment numbers, missing phone numbers). If the engine crashes or leaves raw `{{MiddleName}}` tags, the document looks unprofessional.
* **Invariant**:
  - If a spreadsheet cell is empty or NaN, the engine must replace the placeholder with an empty string `""` by default, or an optional user-defined fallback (e.g. `N/A`).
  - If the template contains a tag that does NOT exist in the spreadsheet columns, the app must present a friendly visual warning dialog highlighting the unmapped tag before merging begins.

### Rule 4: Dynamic Filename Sanitization & Length Clamping
* **The Reason**: When naming files from user data (e.g. `Certificate_{FullName}.pdf`), a name might contain illegal characters like `Dr. Smith / Cardiologist` or quotation marks.
* **Invariant**:
  - Strip or replace illegal characters `\ / : * ? " < > |` with clean underscores `_`.
  - Enforce maximum filename length of 120 characters to guarantee Windows path length compliance.

---

## 🎨 2. UI DESIGN & VISUAL THEME: EDITORIAL INDIGO & PAPER LIGHT

DocuMergePDF features a refined, distraction-free **Editorial Indigo & Paper Light** theme:

```python
# Editorial Indigo Light Theme Palette Constants
CANVAS_BG = "#F9FAFB"        # Gray-50: Main application canvas
CARD_BG = "#FFFFFF"          # Pure White: Content containers and dropzones
CARD_BORDER = "#E5E7EB"      # Gray-200: Subtle card outlines
PRIMARY_ACCENT = "#4F46E5"   # Indigo-600: Primary action buttons and focus rings
PRIMARY_HOVER = "#4338CA"    # Indigo-700: Button hover state
SECONDARY_ACCENT = "#6366F1"# Indigo-500: Active tabs and toggles
SOFT_PILL_BG = "#EEF2FF"     # Indigo-50: Tag background for detected placeholders
SOFT_PILL_TEXT = "#3730A3"   # Indigo-900: Tag text
TEXT_PRIMARY = "#111827"     # Gray-900: High-contrast primary typography
TEXT_SECONDARY = "#6B7280"   # Gray-500: Subtitles, helper text, and placeholders
WARNING_ACCENT = "#F59E0B"   # Amber-500: Trial counter badge and missing field alerts
SUCCESS_ACCENT = "#10B981"   # Emerald-500: Completed batch notifications
FONT_FAMILY = "Segoe UI"     # Native Windows 10/11 system font
```

---

## 🔐 3. FREEMIUM & DUAL-VAULT LICENSING SPECIFICATIONS

### The Free Tier (Never Expires):
* **Single Document Test Generation**: Generating **Row 1 Test PDF** is 100% FREE FOREVER with unlimited runs. Users can test and verify their Word template formatting, fonts, margins, and tag mappings without ever paying.
* **Template Tag Inspector**: Unlimited drag-and-drop analysis of Word templates to preview detected `{{tags}}`.

### The Paid Trigger & Free Trial (5 Full-Power Sessions):
* **Batch Document Merging**: Merging rows 2 to 1,000+ into separate individual files triggers the **5-session Free Trial**.
* Each trial session allows batch generating **unlimited individual PDF files** in a single run.
* **Trial Counter Badge**: Visible in the top header:
  `TRIAL: 5/5 BATCH RUNS LEFT` $\rightarrow$ `4/5` $\rightarrow$ `3/5`...
* After 5 sessions, batch merging opens the purchase modal for the **$6.99 Lifetime License** (or $1.99/mo).

### Dual-Vault Security Architecture:
1. **Vault 1**: Windows Registry `HKCU\Software\Microsoft\Windows\CurrentVersion\DocuMergePDF\LicenseState`
2. **Vault 2**: File System `%LOCALAPPDATA%\DocuMergePDF\.security\vault.dat` (AES-256 encrypted with HWID salt).
* **Anti-Rollback Defense**: Validates monotonic timestamp progression against system clock manipulation.
* **Offline Key Support**: Format `DMP-LIFETIME-[HWID_PREFIX]-[HMAC_16]` enables instant corporate air-gapped activation.

