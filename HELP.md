
# Translit - Document Maker

**Translit** is a professional, high-performance WYSIWYG rich-text editor built with PySide6. Designed specifically for bilingual (Hindi/English) workflows, it features real-time phonetic (ITRANS) to Devanagari transliteration, context-aware typing predictions, integrated voice typing, and an advanced templating system.

With native DOCX and PDF export, a fully-featured Advanced Image Studio, highly customizable UI themes, and smart adaptive dictionaries, Translit bridges the gap between simple text editors and full-fledged office suites.

---

## 🌟 Key Features

### 🔠 Intelligent Typing & Transliteration
* **Real-Time Transliteration:** Type phonetically in Latin (e.g., *kya* -> *क्या*) with instant conversion.
* **AI Next-Word Predictions:** The editor learns your conversational flow and intelligently predicts the next word based on context.
* **English/Hinglish Toggle:** Instantly switch between Hindi transliteration and standard English typing.
* **In-Place Translation:** Highlight any English text and translate it directly to Hindi.
* **Voice Typing (Speech-to-Text):** Hold a hotkey to dictate in Hindi directly into the editor.

### 📄 Professional Document Management
* **Rich Format Support:** Open, edit, and save as native `.docx`, `.pdf`, `.html`, and `.txt`.
* **Print & Page Layouts:** Switch between Web layout, Portrait (Canvas), and Landscape views for accurate print representation.
* **Advanced Formatting:** Bold, italics, multi-style lists, superscript, subscript, text/highlight colors, and dynamic table insertions.
* **Advanced Image Studio:** A built-in image processor featuring a drag-to-crop visual tool, exact pixel/cm/inch resizing, rounded corners, borders, drop shadows, rotation, and photo filters (Grayscale, Sepia, Invert, Tinting).

### 🛠️ Advanced Tools & Automation
* **Dictionary Training:** Manually train custom Latin-to-Devanagari mappings and shortcuts for specific vocabulary.
* **Prediction Model & Vocabulary Scanners:** Extract unique words and bigrams (word pairs) from PDFs, Word files, and Spreadsheets (CSV/XLSX) to populate your custom AI memory.
* **Template Manager:** Create, edit, and reuse HTML-based document templates (e.g., Question Papers, Applications, Declarations).
* **Phrase Manager:** Save frequently used phrases or greetings and insert them instantly via a dedicated sidebar dock.

### 🎨 Customizable UI
* **Chameleon Theme Engine:** Seamlessly toggle between Ultra-Compact Dark and Light modes with dynamically styled tabs and menus.
* **Customizable Popups:** Tailor the suggestion popup's visual style (Classic, Modern, Neon, Google Search), grid layout, line spacing, and text colors.
* **Productivity Timers:** Built-in status bar trackers monitor app uptime, active document session times, word counts, and cursor positions.

---

## ⚙️ Requirements & Installation

Translit requires **Python 3.8+**. It heavily relies on `PySide6` for the interface and uses several optional libraries to unlock its full potential.

### 1. Install Core Dependencies

```bash
pip install PySide6 indic-transliteration

```

### 2. Install Feature-Specific Dependencies (Highly Recommended)

To enable all advanced features (DOCX export, Speech Recognition, PDF parsing, Excel scanning), install the following:

```bash
# For native DOCX saving and importing
pip install python-docx

# For Speech Recognition (requires PyAudio)
pip install SpeechRecognition PyAudio

# For parsing PDFs and Spreadsheets to train the AI & Dictionary
pip install pypdf pandas openpyxl odfpy

```

### 3. Run the Application

```bash
python main.py

```

*(Replace `main.py` with the actual filename of your script).*

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
| **`Ctrl` + `Shift` + `C`** | Open Correction Dictionary Manager |
| **`Ctrl` + `Shift` + `E`** | Translate Selected Text to Hindi |
| **`Ctrl` + `Shift` + `D`** | Insert Current Date/Time |
| **`Alt` + `P`** | Focus the Phrases Sidebar |
| **`Ctrl` + `1-9, 0`** | Insert the corresponding numbered suggestion from the popup |
| **`Ctrl` + `↑` / `↓`** | Navigate the suggestion list in the Sidebar Dock |
| **`Tab`** | Insert the selected suggestion (Dock or Inline) |

---

## 📂 User Data Directory

Translit automatically saves user preferences, custom dictionaries, AI prediction models, phrases, and templates in the standard application data folder for your OS.

* **Windows:** `C:\Users\<User>\AppData\Roaming\.hindi_office_wysiwyg\`
* **macOS / Linux:** `~/.hindi_office_wysiwyg/`

**Generated JSON files include:**

* `user_translit.json` (Your custom typing mappings & macros)
* `suggestions.json` (Your extracted autocomplete vocabulary)
* `next_words.json` (Your AI predictive flow model)
* `phrases.json` (Quick-insert text snippets)
* `templates.json` (Saved HTML document templates)
