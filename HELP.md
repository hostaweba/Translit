# Translit - Complete Application Guide & Use Cases

Translit is a professional, bilingual (Hindi/English) rich-text editor designed to eliminate the friction of typing in Devanagari while providing full word-processing capabilities. This guide covers every feature, how to use it, and real-world use cases.

---

## 1. Core Typing & Transliteration

### Phonetic Typing (ITRANS)

Type phonetically in Latin characters, and Translit will automatically convert it to Devanagari in real-time.

* **Use Case:** Typing Hindi documents quickly without needing a specialized Hindi keyboard layout.
* **Halant (Half-Letters):** Type a tilde `~` (or `0`) immediately after a letter to force a half-letter (e.g., `k~` becomes **क्**).

### English / Hinglish Toggle (`Ctrl + Space`)

Instantly pause the transliteration engine to type in standard English.

* **Use Case:** Inserting English names, email addresses, technical terms, or URLs into a Hindi document. A badge in the bottom-right corner confirms your current mode.

### Voice Typing (Speech-to-Text)

Dictate text directly into the editor using your microphone.

* **How to use:** Press and **hold `Ctrl + L**`. Speak your sentence in Hindi, then release the keys. The app will process the audio and insert the Devanagari text.
* **Use Case:** Rapidly drafting long paragraphs, taking quick notes, or resting your hands during heavy typing sessions.

---

## 2. Smart Suggestions & Vocabulary Building

Translit predicts what you are typing and offers auto-completions to save keystrokes.

### Using the Suggestion UI

You can display suggestions in an **Inline Popup** (under your cursor), a **Sidebar Dock**, or **Both** (configurable via the *View > Suggestion UI Mode* menu).

* **Inline Popup:** Press `Up`/`Down` arrows to navigate, `Enter` or `Tab` to insert. Press `Ctrl + [1-9]` to instantly insert a numbered suggestion.
* **Sidebar Dock:** Hold `Ctrl` and press `Up`/`Down` arrows to navigate the dock while typing. Press `Tab` to insert the selected dock word.

### Suggestion Database Manager (Vocabulary Builder)

Accessible via *Tools > Suggestion Database Manager*. This is where you manage the words Translit predicts.

* **Document Scanner:** Click **Scan/Extract Words from File(s)** to bulk-import vocabulary. You can select PDF, DOCX, TXT, Excel (XLSX/CSV) files.
* **Use Case:** If you are a legal typist or a teacher, scan your past documents. Translit will extract all unique words and add them to your suggestion database, instantly learning your specialized vocabulary.

---

## 3. Custom Macros & Correction Dictionary

Accessible via *Tools > Correction Dictionary* (`Ctrl + Shift + C`). This powerful tool lets you override the default transliteration engine.

### Full-Word Shortcuts (Text Expanders)

Map a lowercase Latin string to a specific Devanagari word or phrase.

* **Rule:** Must be typed as an isolated word.
* **Use Case:** Map `he` to `Mahatama Gandhi`. Map `.sig` to `भवदीय, आपका नाम`. This acts as a rapid text-expander for frequent phrases or hard-to-type names.

### In-Word Modifiers

Map an uppercase string or symbol to a specific phonetic output.

* **Rule:** The engine will replace these characters *inside* of other words.
* **Use Case:** Map `Ksh` to `क्ष्`. Typing `Kshatriya` will guarantee the word starts with the correct conjunct.

### Context Menu Training

* **How to use:** Highlight a wrongly transliterated word in the editor, right-click, and select **Train map from selection**. You will be prompted to enter the Latin shortcut you want to assign to that text.

---

## 4. Advanced Templating & Phrases

### Advanced Template Manager

Accessible via the *File* toolbar icon or *Templates* menu. Translit supports reusable HTML-based document structures.

* **Features:** Insert templates, save your current editor canvas as a new template, preview HTML, or export/import your entire template library as a JSON file to share with colleagues.
* **Use Case:** Creating standardized Question Papers, Leave Applications, Legal Declarations, or Office Memos without re-typing the layout every time.

### Phrase Manager Dock

A sidebar dock displaying pre-saved text snippets.

* **How to use:** Double-click any phrase in the dock to instantly insert it at your cursor. Right-click the dock to Add, Edit, or Remove phrases.
* **Use Case:** Storing standard greetings ("नमस्ते,"), sign-offs, or recurring boilerplate text.

---

## 5. Document Formatting & Layout

### Visual Layout Modes

Change how the page looks via the *View > Page Layout View* menu.

* **Web Layout (Default):** Edge-to-edge typing, ideal for drafting and screen reading.
* **Portrait / Landscape Canvas:** Simulates an actual physical page (like MS Word). It respects printer margins, making it perfect for visualizing how the document will look when printed.

### Rich Text Formatting

* **Tools:** Standard controls for Bold, Italic, Underline, Strikethrough, Subscript, Superscript, Text Color, and Background Highlight.
* **Lists:** Bullet points and multi-style numbering (1,2,3... a,b,c... I,II,III...).
* **Tables:** Insert tables, and use the right-click context menu inside a table to change its alignment (Left, Center, Right) on the page.
* **Images:** Insert images with exact pixel widths and alignment controls.

---

## 6. Document Management & Export

### Supported Formats

* **Save/Open Native Formats:** Fully supports `.docx`, `.html`, `.txt`.
* **Spreadsheet Import:** Open `.csv`, `.xlsx`, or `.ods` files to instantly render spreadsheet data as an editable HTML table inside your document.
* **PDF Export:** Export your exact layout to `.pdf` via *File > Export as PDF*.

### Safety & Tracking

* **Autosave:** When creating a new file or saving for the first time, you are prompted to enable Autosave. Configure the interval via *Tools > Set Autosave Interval*.
* **Productivity Timers:** The status bar tracks exactly how long the app has been open ("App") and how long you've spent actively editing the current document ("Doc").

---

## 7. UI Customization & Themes

### Dark Mode / Light Mode (`Ctrl + D`)

Instantly toggle between an ultra-compact Dark Theme (gentle on the eyes) and a crisp Light Theme.

### Suggestion Popup Design

Accessible via *View > Suggestion Popup Settings*. Completely customize the look of the autocomplete menu.

* **Visual Themes:** Choose from OS Native, Classic, Modern, Minimalist, Neon, or Google Search style.
* **Grid Layout:** Transform the vertical list into a multi-column grid by setting exact row and column counts.
* **Typography:** Change font size, text color, bold formatting, line spacing, and toggle the visibility of shortcut numbers `[1]`, `[2]`.

---

## Master Keyboard Shortcuts Cheat Sheet

| Shortcut | Action |
| --- | --- |
| **`Ctrl` + `Space`** | Toggle English (Hinglish) / Hindi Transliteration mode |
| **Hold `Ctrl` + `L`** | Activate Speech-to-Text (Voice Typing) |
| **`Ctrl` + `T`** | Toggle Transliteration Engine ON/OFF completely |
| **`Ctrl` + `D`** | Toggle Dark / Light Theme |
| **`Ctrl` + `Shift` + `C`** | Open Correction Dictionary Manager |
| **`Ctrl` + `Shift` + `D`** | Insert Current Date/Time |
| **`Ctrl` + `1-9, 0`** | Insert the corresponding numbered suggestion from the inline popup |
| **`Ctrl` + `↑` / `↓`** | Navigate the suggestion list in the Sidebar Dock |
| **`Tab`** | Insert the currently selected suggestion (Inline or Dock) |
| **`Ctrl` + `N` / `O` / `S`** | New, Open, Save Document |
| **`Ctrl` + `P`** | Open Print Preview |
| **`Ctrl` + `F`** | Find and Replace |
| **`Ctrl` + `B` / `I` / `U`** | Bold, Italic, Underline |
| **`Ctrl` + `=`** | Subscript (X₂) |
| **`Ctrl` + `Shift` + `+`** | Superscript (X²) |
