# 📜 IraPrep: Universal IRaMuTeQ Corpus Studio

> **Transform any document (Excel, Word, PDF, CSV, Text) into an error-free, publication-ready IRaMuTeQ corpus in seconds.**
> 100% Local, zero cloud leakage, IRB/Ethics compliant.

---

## 🚀 Quick Start (Zero Code)

### On macOS / Linux:
Double-click `start_iraprep.sh` in Finder, or run in Terminal:
```bash
./start_iraprep.sh
```

### On Windows or any system:
```bash
python3 start_iraprep.py
```
Your default web browser will automatically open to the **IraPrep Studio** (`http://localhost:8501`).

---

## 🌟 Why IraPrep?

IRaMuTeQ requires a very strict text format (`**** *var1_val1 *var2_val2`) and crashes easily if text contains:
* Rogue asterisks (`*`)
* 4-byte emojis or pictographs
* Curly/smart quotes (`“ ” ‘ ’`)
* Unspaced punctuation (`word.next`)
* Invalid variable names (accents, spaces, capitals, hyphens)

**IraPrep handles all of this automatically in the background.**

---

## 🛠️ Supported Formats & Customizations

### 1. 📊 Spreadsheets (`.xlsx`, `.csv`, `.tsv`)
* **Text Selection:** Pick one or multiple columns to act as the text body.
* **Variable Mapping:** Select any number of metadata columns (e.g. `Author`, `Condition`, `Gender`, `Score`).
* **Global Variables:** Add batch variables like `*date_2024_11` or `*type_survey` across all rows.

### 2. 📄 Documents & Transcripts (`.docx`, `.pdf`, `.txt`)
* **Batch Ingestion:** Upload 1 or 50 files simultaneously.
* **Flexible Sectioning Options:**
  * **Whole Document = 1 UCI:** Perfect for a folder of individual participant transcripts.
  * **By Heading Level:** Split by Word headings (`Heading 1`, `Heading 2`).
  * **By Speaker Turns:** Split dialogues using regex (e.g. `Participant:`, `Interviewer:`, `Q1:`).
  * **By Custom Delimiters:** Split by `---`, `###`, or custom markers.
  * **By Paragraphs / Blank lines.**
* **Filename Metadata Extraction:** Auto-extract variables from filenames (e.g. `P01_Female_30.docx` $\to$ `*filetag_1_p01 *filetag_2_female *filetag_3_30`).

### 3. ✍️ Quick Paste
* Paste raw text directly, choose a section delimiter (e.g. `=== Section ===`), and specify date/type variables.

---

## 🔒 100% Privacy & Data Ethics
* **No Closed-Weight APIs:** Does not send data to OpenAI, Google Cloud, or Anthropic.
* **Strict Confidentiality:** All parsing, cleaning, and regex operations run locally in memory on your computer.
* **Completely safe for sensitive data:** Qualitative interviews, addiction studies, patient records, and confidential research.

---

## 💾 Save & Share "Recipes"
Configured a complex set of compound words (e.g. `mental_health`, `harm_reduction`) and cleaning rules?
* Click **"Download Current Recipe"** in the sidebar to export a `.json` settings file.
* Share this recipe with colleagues so everyone produces identically formatted corpora.
