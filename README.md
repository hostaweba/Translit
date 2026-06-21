# Translit - Document Maker

**Translit** is a professional, high-performance WYSIWYG rich-text editor built with PySide6. Designed specifically for bilingual (Hindi/English) workflows, it features real-time phonetic (ITRANS) to Devanagari transliteration, integrated voice typing, and an advanced templating system.

With native DOCX and PDF export, highly customizable UI themes, and smart adaptive dictionaries, Translit bridges the gap between simple text editors and full-fledged office suites.

---

## 🌟 Key Features

### 🔠 Intelligent Typing & Transliteration

* **Real-Time Transliteration:** Type phonetically in Latin (e.g., *kya* -> *क्या*) with instant conversion.
* **English/Hinglish Toggle:** Instantly switch between Hindi transliteration and standard English typing.
* **Smart Suggestions:** Auto-complete suggestions available via an inline popup or a dedicated sidebar dock.
* **Voice Typing (Speech-to-Text):** Hold a hotkey to dictate in Hindi directly into the editor.

### 📄 Professional Document Management

* **Rich Format Support:** Open, edit, and save as native `.docx`, `.pdf`, `.html`, and `.txt`.
* **Print & Page Layouts:** Switch between Web layout, Portrait (Canvas), and Landscape views for accurate print representation.
* **Advanced Formatting:** Bold, italics, lists (bulleted/numbered), superscript, subscript, text/highlight colors, and dynamic table insertions.
* **Image Support:** Insert, scale, and align images directly within the document.

### 🛠️ Advanced Tools & Automation

* **Dictionary Training:** Manually train custom Latin-to-Devanagari mappings for specific vocabulary.
* **Document Scanner:** Extract unique words from PDFs, Word files, and Spreadsheets (CSV/XLSX) to populate your custom suggestion database.
* **Template Manager:** Create, edit, and reuse HTML-based document templates (e.g., Question Papers, Applications, Declarations).
* **Phrase Manager:** Save frequently used phrases or greetings and insert them with a single click.

### 🎨 Customizable UI

* **Chameleon Theme Engine:** Seamlessly toggle between Ultra-Compact Dark and Light modes.
* **Customizable Popups:** Tailor the suggestion popup's style (Classic, Modern, Neon, Google Search), grid layout, spacing, and colors.
* **Productivity Timers:** Built-in status bar trackers monitor app uptime and active document session times.

---

## ⚙️ Requirements & Installation

Translit requires **Python 3.8+**. It heavily relies on `PySide6` for the interface and uses several optional libraries to unlock its full potential.

### 1. Install Core Dependencies

```bash
pip install PySide6 indic-transliteration

```

### 2. Install Feature-Specific Dependencies (Highly Recommended)

To enable all features (DOCX export, Speech Recognition, PDF parsing, Excel scanning), install the following:

```bash
# For native DOCX support
pip install python-docx

# For Speech Recognition (requires PyAudio)
pip install SpeechRecognition PyAudio

# For parsing PDFs and Spreadsheets to train the dictionary
pip install PyPDF2 pandas openpyxl odfpy

```

### 3. Run the Application

```bash
python main.py

```

*(Replace `main.py` with the actual filename of your script).*
**
---

## ⌨️ Keyboard Shortcuts

| Shortcut | Action |
| --- | --- |
| **`Ctrl` + `Space`** | Toggle English (Hinglish) / Hindi Transliteration mode |
| Hold **`Ctrl` + `L`** | Activate Speech-to-Text (Voice Typing) |
| **`Ctrl` + `T`** | Toggle Transliteration Engine ON/OFF completely |
| **`Ctrl` + `D`** | Toggle Dark / Light Theme |
| **`Ctrl` + `N` / `O` / `S`** | New, Open, Save Document |
| **`Ctrl` + `P`** | Open Print / PDF Export Preview |
| **`Ctrl` + `F`** | Find and Replace |
| **`Ctrl` + `Shift` + `C`** | Open Correction Dictionary |
| **`Ctrl` + `Shift` + `D`** | Insert Current Date/Time |
| **`Ctrl` + `1-9, 0`** | Insert the corresponding numbered suggestion from the popup |
| **`Ctrl` + `↑` / `↓`** | Navigate the suggestion list in the Sidebar Dock |
| **`Tab`** | Insert the selected suggestion (Dock or Inline) |

---

## 📂 User Data Directory

Translit automatically saves user preferences, custom dictionaries, phrases, and templates in the standard application data folder for your OS.

* **Windows:** `C:\Users\<User>\AppData\Roaming\.hindi_office_wysiwyg\`
* **macOS / Linux:** `~/.hindi_office_wysiwyg/`

**Generated JSON files include:**

* `user_translit.json` (Your custom typing mappings)
* `suggestions.json` (Your extracted autocomplete database)
* `phrases.json` (Quick-insert phrases)
* `templates.json` (Saved document templates)
