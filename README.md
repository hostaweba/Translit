# 📝 Translit - Document Maker

A professional, high-performance Rich Text Editor tailored specifically for seamless Hindi typing. Built with Python and PySide6, this application combines real-time phonetic transliteration, voice typing, advanced dictionary management, and native DOCX/PDF handling into a single, offline-capable productivity suite.

## ✨ Key Features

### ⌨️ Typing & Input

* **Real-time Transliteration:** Type phonetically in English (Latin) and watch it instantly convert to Devanagari (Hindi) without lag.
* **Voice Typing (Speech-to-Text):** Hold `Ctrl+L` to dictate in Hindi using the Google Speech Recognition API. Release to seamlessly insert the transliterated text.
* **Smart Dictionary Training:** Easily add custom Latin-to-Devanagari word mappings. Scan entire `.txt`, `.docx`, or `.xlsx` files to automatically extract new words and train your custom dictionary.
* **Quick Phrases & Templates:** A dedicated sidebar to save and insert frequently used phrases, plus built-in professional templates (Question Papers, Applications, Declarations).

### 📄 Document Formatting & View

* **Rich Text Editing:** Full support for Bold, Italic, Underline, Strikethrough, Text Color, Highlighting, Bullet/Numbered Lists, and Tables.
* **Print Layout View:** Toggle a true-to-life "Print Layout" view that simulates an A4 (or custom-sized) physical paper on a desk.
* **Dark Mode & Full Screen:** Built-in Dark Mode that intuitively handles page colors, and a distraction-free Full Screen mode (`F11`).
* **Page Setup & Print Preview:** Configure margins, paper size, and orientation with native Print Preview support.

### 💾 File Management & Export

* **Native DOCX Export:** Generates clean MS Word (`.docx`) files, preserving inline images, custom fonts, tables, and rich text styles.
* **PDF Export:** Direct, high-resolution PDF generation.
* **Multi-Format Import:** Open and read `.docx`, `.html`, `.txt`, `.csv`, `.xlsx`, and `.ods` files directly into the editor.
* **Continuous Autosave:** Never lose your work. Prompts you to enable background autosaving (customizable intervals) on new documents.
* **Word Integration:** Direct "Insert into Word" functionality via OLE automation (Windows only).

---

## 🛠️ Prerequisites & Installation

This application requires **Python 3.8+**. It heavily relies on several third-party libraries for its extended feature set.

### 1. Install Dependencies

Open your terminal or command prompt and run the following command to install all required packages:

```bash
pip install PySide6 indic-transliteration python-docx pandas openpyxl odfpy SpeechRecognition PyAudio Pillow pywin32

```

*Note: `PyAudio` is required for microphone input. If you are on Windows and face issues installing PyAudio, you may need to use `pipwin` or install the pre-compiled wheel.*
*Note: `pywin32` is for Windows users only.*

### 2. Running the Application

Save the main script as `main.py` and run it via Python:

```bash
python main.py

```

---

## ⌨️ Keyboard Shortcuts

| Shortcut | Action |
| --- | --- |
| **`Ctrl + L`** | **Hold** to start voice typing (Release to insert) |
| **`Ctrl + T`** | Toggle Phonetic Transliteration (On/Off) |
| **`Ctrl + Shift + C`** | Open Dictionary / Corrections Editor |
| **`Ctrl + D`** | Toggle Dark Mode |
| **`F11`** | Toggle Full Screen |
| **`Ctrl + + / -`** | Zoom In / Zoom Out |
| **`Ctrl + Shift + D`** | Insert Current Date & Time |
| **`Ctrl + F`** | Find and Replace |
| **`Ctrl + P`** | Print / Print Preview |
| **`Ctrl + N` / `O` / `S**` | New, Open, Save Document |

---

## 🗂️ Supported File Formats

**Read / Import:**

* Text / Web: `.txt`, `.html`, `.htm`
* Word Processors: `.docx` (via `python-docx`)
* Spreadsheets: `.csv`, `.xlsx`, `.ods` (via `pandas` & `openpyxl`)

**Write / Export:**

* Documents: `.docx`, `.pdf`, `.html`, `.txt`

---

## 🧠 Dictionary File Scanner Note

If you have existing documents full of English/Latin abbreviations or custom names that you want to map to Devanagari, open the **Dictionary (Ctrl+Shift+C)** and click **"Scan/Extract Words from File..."**. The app will parse the document, extract all non-numeric tokens, and add the missing ones to your dictionary grid so you can define their Hindi translations rapidly.
