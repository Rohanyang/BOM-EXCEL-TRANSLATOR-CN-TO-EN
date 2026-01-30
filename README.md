# BOM Excel Translator (CN → EN)

A lightweight, company-friendly BOM translation tool designed for tooling / stamping / manufacturing projects.

This project translates Chinese text in Excel BOM files into English using a **customizable glossary (`WORD LIST.xlsx`)**, and automatically outputs a QA list for untranslated terms to support continuous glossary improvement.

---

## ✨ Key Features

- 🔁 **Excel → Excel translation**
- 📘 Uses a **user-maintained glossary** (`WORD LIST.xlsx`)
- 🎯 **Exact match + mixed-text replacement** (safe for BOM context)
- 🧾 Automatically generates **QA file** for untranslated Chinese terms
- 🖥️ Works as:
  - Python script (`.py`)
  - Standalone Windows executable (`.exe`, built with PyInstaller)

---

## 📂 Project Structure
.
├─ run_gui.py # GUI entry point
├─ translate_bom.py # Core translation logic
├─ rules.py # Translation rules (if applicable)
├─ WORD LIST.xlsx # Sample glossary (CN / EN)
├─ requirements.txt # Python dependencies
├─ README.md # Project overview (this file)


---

## 🧠 Design Philosophy

This tool is **not a generic AI translator**.

It is designed to:
- Preserve BOM structure and formatting
- Avoid incorrect substring translations
- Support **team-level glossary accumulation**
- Reduce repetitive manual translation work in engineering projects

The QA output is intentionally kept simple so teams can quickly copy new terms back into `WORD LIST.xlsx` and iterate.

---

## 🚀 Getting Started (Developer)

```bash
pip install -r requirements.txt
python run_gui.py


