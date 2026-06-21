#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Translit - Document Maker
Professional, high-performance editor with Native DOCX, PDF Export, Print Layout, Page Setup, and Advanced Templates.
"""

from __future__ import annotations

import json
import logging
import os
import re
import sys
import tempfile
from datetime import datetime
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Set

from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtGui import QIcon, QTextCharFormat, QFont
from PySide6.QtCore import Qt, QTimer
from PySide6.QtPrintSupport import QPrinter, QPrintPreviewDialog, QPageSetupDialog

logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

# Dependencies
try:
    from indic_transliteration.sanscript import transliterate  # type: ignore
    HAS_TRANSLIT = True
except Exception:
    transliterate = None
    HAS_TRANSLIT = False

try:
    import docx
    from docx.shared import Pt, RGBColor, Inches
    from docx.enum.text import WD_COLOR_INDEX, WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    HAS_DOCX = True
except Exception:
    docx = None
    HAS_DOCX = False

try:
    import pandas as pd
    HAS_PANDAS = True
except Exception:
    pd = None
    HAS_PANDAS = False

try:
    import speech_recognition as sr  # type: ignore
    HAS_SR = True
except Exception:
    sr = None
    HAS_SR = False

APP_NAME = "Translit"
USER_DIR = Path(
    QtCore.QStandardPaths.writableLocation(QtCore.QStandardPaths.StandardLocation.AppDataLocation)
    or Path.home() / ".hindi_office_wysiwyg"
)
USER_DIR.mkdir(parents=True, exist_ok=True)
USER_DICT_PATH = USER_DIR / "user_translit.json"
PHRASES_PATH = USER_DIR / "phrases.json"
SUGGESTIONS_PATH = USER_DIR / "suggestions.json"
TEMPLATES_PATH = USER_DIR / "templates.json"

DEFAULT_PHRASES = [
    "नमस्ते,\nआशा है आप स्वस्थ हैं।",
    "धन्यवाद,\n[आपका नाम]",
    "सादर,\n[आपका नाम]\n[पद]",
]

DEFAULT_DICT = {
    "yahan": "यहाँ", "yah": "यह", "rha": "रहा", "narkhi": "नारखी",
    "kya": "क्या", "koko": "कोमल", "karne": "करने", "hi": "हाय",
    "hemlata": "हेमलता", "hello": "हेलो", "han": "हाँ", "etah": "एटा",
    "dr.": "डॉ.", "aj": "आज", "main": "मैं", "hoon": "हूँ", "aur": "और",
    "ki": "कि", "kee": "की", "hota": "होता", "hoti": "होती", "hote": "होते",
    "bahut": "बहुत", "Ri": "ऋ", "R": "ृ", "O": "ॉ", "M": "ँ", "H": "ः",
    "^": "ं", "*": ".", "!": "ॠ", "Aum": "ॐ", "Z": "़",
    "Gy": "ज्ञ्", "Ksh": "क्ष्", "Tr": "त्र्", "Shr": "श्र्",
    "Dhy": "ध्य्", "Dy": "द्य्", "Tt": "ट्ट्", "Dd": "ड्ड्", "Tth": "ट्ठ्",
    "Ru": "रु", "Roo": "रू"
}

U_LINE = "<p align='center'>__________________________________________________________________________________</p>"

TPL_QUESTION_PAPER = f"""
<h2 align="center">विद्यालय / संस्था का नाम</h2>
<h3 align="center">वार्षिक परीक्षा २०२४-२५</h3>
<table width="100%" border="0" cellpadding="4">
  <tr>
    <td align="left"><b>कक्षा:</b> ............</td>
    <td align="right"><b>विषय:</b> ............</td>
  </tr>
  <tr>
    <td align="left"><b>समय:</b> ३ घंटे</td>
    <td align="right"><b>पूर्णांक:</b> १००</td>
  </tr>
</table>
{U_LINE}
<p><b>निर्देश:</b> सभी प्रश्न अनिवार्य हैं।</p>
<ol>
  <li>पहला प्रश्न यहाँ लिखें। (५ अंक)</li>
  <li>दूसरा प्रश्न यहाँ लिखें। (५ अंक)</li>
</ol>
"""

TPL_APPLICATION = """
<p>सेवा में,</p>
<p>श्रीमान प्रधानाचार्य महोदय,<br>
[विद्यालय/संस्था का नाम],<br>
[शहर का नाम]</p>
<p><b>विषय:</b> [अवकाश/अन्य हेतु आवेदन पत्र]</p>
<p>महोदय,</p>
<p>सविनय निवेदन है कि [यहाँ अपना कारण लिखें]। इसलिए मुझे [तारीख] से [तारीख] तक अवकाश प्रदान करने की कृपा करें।</p>
<p>आपकी अति कृपा होगी।</p>
<table width="100%" border="0" cellpadding="4">
  <tr>
    <td align="left"><b>दिनांक:</b> ............</td>
    <td align="right"><b>आपका आज्ञाकारी शिष्य/शिष्या,</b><br>नाम: ............<br>कक्षा: ............</td>
  </tr>
</table>
"""

TPL_DECLARATION = """
<h2 align="center">घोषणा पत्र</h2>
<p>मैं, <b>[आपका नाम]</b>, पुत्र/पुत्री श्री <b>[पिता का नाम]</b>, निवासी <b>[पूरा पता]</b>, यह घोषणा करता/करती हूँ कि:</p>
<ol>
  <li>मेरे द्वारा दी गई सभी जानकारी मेरी जानकारी और विश्वास के अनुसार सत्य और सही है।</li>
  <li>मैंने कोई भी तथ्य नहीं छुपाया है।</li>
</ol>
<p>यदि भविष्य में कोई भी जानकारी असत्य पाई जाती है, तो इसके लिए मैं स्वयं जिम्मेदार रहूँगा/रहूँगी।</p>
<br><br>
<table width="100%" border="0" cellpadding="4">
  <tr>
    <td align="left"><b>स्थान:</b> ............<br><b>दिनांक:</b> ............</td>
    <td align="right"><b>हस्ताक्षर:</b> ____________________<br><b>नाम:</b> ............</td>
  </tr>
</table>
"""

@dataclass
class AppState:
    transliteration_enabled: bool = True
    font_family: str = "Nirmala UI"
    font_size: int = 16
    user_dict: Dict[str, str] = field(default_factory=lambda: DEFAULT_DICT.copy())
    phrases: List[str] = field(default_factory=lambda: DEFAULT_PHRASES.copy())
    suggestion_words: Set[str] = field(default_factory=set)
    templates: Dict[str, str] = field(default_factory=lambda: {
        "Question Paper": TPL_QUESTION_PAPER,
        "Application Form": TPL_APPLICATION,
        "Declaration Letter": TPL_DECLARATION
    })
    active_dict_name: str = "Default (user_translit.json)"
    active_sugg_name: str = "Default (suggestions.json)"

    def load(self):
        if USER_DICT_PATH.exists():
            try:
                with open(USER_DICT_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict) and data: self.user_dict = data
            except Exception as e: logging.error(f"Error loading dict: {e}")
        else:
            self.save_user_dict() 

        if PHRASES_PATH.exists():
            try:
                with open(PHRASES_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list): self.phrases = data
            except Exception as e: logging.error(f"Error loading phrases: {e}")
            
        if SUGGESTIONS_PATH.exists():
            try:
                with open(SUGGESTIONS_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list): self.suggestion_words = set(data)
            except Exception as e: logging.error(f"Error loading suggestions: {e}")

        if TEMPLATES_PATH.exists():
            try:
                with open(TEMPLATES_PATH, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict) and data: self.templates = data
            except Exception as e: logging.error(f"Error loading templates: {e}")
        else:
            self.save_templates()

    def save_user_dict(self, path: Optional[Path] = None) -> bool:
        try:
            with open(path or USER_DICT_PATH, "w", encoding="utf-8") as f:
                json.dump(self.user_dict, f, ensure_ascii=False, indent=2)
            return True
        except Exception: return False

    def save_phrases(self):
        try:
            with open(PHRASES_PATH, "w", encoding="utf-8") as f:
                json.dump(self.phrases, f, ensure_ascii=False, indent=2)
            return True
        except Exception: return False

    def save_suggestions(self, path: Optional[Path] = None) -> bool:
        try:
            with open(path or SUGGESTIONS_PATH, "w", encoding="utf-8") as f:
                json.dump(list(self.suggestion_words), f, ensure_ascii=False, indent=2)
            return True
        except Exception: return False

    def save_templates(self, path: Optional[Path] = None) -> bool:
        try:
            with open(path or TEMPLATES_PATH, "w", encoding="utf-8") as f:
                json.dump(self.templates, f, ensure_ascii=False, indent=2)
            return True
        except Exception: return False


class AdaptiveTransliterator:
    def __init__(self, state: AppState):
        if not HAS_TRANSLIT:
            raise RuntimeError("Missing indic-transliteration.")
        self.src = "itrans"
        self.dst = "devanagari"
        self.state = state
        self._reverse_cache = {}
        self._cached_dict_keys = []
        self._modifier_map = {}
        self._modifier_pattern = None
        self._dict_hash = 0
        self._update_cache()

    def _get_itrans(self, dev_str: str) -> str:
        if dev_str not in self._reverse_cache:
            try:
                self._reverse_cache[dev_str] = transliterate(dev_str, self.dst, self.src)
            except Exception:
                self._reverse_cache[dev_str] = dev_str
        return self._reverse_cache[dev_str]

    def _update_cache(self):
        current_hash = hash(frozenset(self.state.user_dict.items()))
        if current_hash != self._dict_hash:
            self._cached_dict_keys = sorted(self.state.user_dict.keys(), key=len, reverse=True)
            
            # Build a safe mapping and regex pattern for modifiers
            self._modifier_map = {}
            for k in self._cached_dict_keys:
                is_modifier = (not k.isalpha()) or any(c.isupper() for c in k)
                if is_modifier:
                    dev_val = self.state.user_dict[k]
                    self._modifier_map[k] = self._get_itrans(dev_val)
                    
            if self._modifier_map:
                # Compile a regex pattern to replace everything in a single, safe pass
                sorted_map_keys = sorted(self._modifier_map.keys(), key=len, reverse=True)
                self._modifier_pattern = re.compile('|'.join(map(re.escape, sorted_map_keys)))
            else:
                self._modifier_pattern = None
                
            self._dict_hash = current_hash

    def translit_token(self, latin: str) -> str:
        if not latin: return ""
        # 1. Check exact matches first
        if latin in self.state.user_dict: return self.state.user_dict[latin]
        if latin.lower() in self.state.user_dict: return self.state.user_dict[latin.lower()]
            
        self._update_cache()
        
        # 2. Apply safe one-pass replacement using regex
        if self._modifier_pattern:
            temp_latin = self._modifier_pattern.sub(lambda m: self._modifier_map[m.group(0)], latin)
        else:
            temp_latin = latin
                
        # 3. Halant forced check
        force_halant = temp_latin.endswith('~')
        if force_halant:
            temp_latin = temp_latin[:-1]
                
        try: 
            res = transliterate(temp_latin, self.src, self.dst)
            # Remove trailing halant for natural Hindi typing (schwa omission)
            if not force_halant and res.endswith('्'):
                res = res[:-1]
            return res
        except Exception: 
            return temp_latin

    def translit_full(self, text: str) -> str:
        if not text: return ""
        words = text.split()
        return " ".join([self.translit_token(w) for w in words])


# ---------------------------------------------------------
# NATIVE MEMORY EXTRACTOR (QTextFrame -> DOCX)
# ---------------------------------------------------------
def export_native_to_docx(qdoc: QtGui.QTextDocument, filepath: str):
    if not HAS_DOCX: raise Exception("python-docx is not installed.")
    
    doc = docx.Document()
    style = doc.styles['Normal']
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)
    style.paragraph_format.line_spacing = 1.0

    def process_paragraph(block, docx_parent=None, p=None):
        if block.text().strip() == "" and not block.begin().fragment().charFormat().isImageFormat():
            if p is None and docx_parent is not None:
                docx_parent.add_paragraph()
            return

        if p is None and docx_parent is not None:
            p = docx_parent.add_paragraph()
            
        if p is None: return

        align = block.blockFormat().alignment()
        if align & Qt.AlignmentFlag.AlignHCenter: p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align & Qt.AlignmentFlag.AlignRight: p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        elif align & Qt.AlignmentFlag.AlignJustify: p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

        it = block.begin()
        while not it.atEnd():
            fragment = it.fragment()
            if fragment.isValid():
                fmt = fragment.charFormat()
                if fmt.isImageFormat():
                    img_fmt = fmt.toImageFormat()
                    img_name = img_fmt.name()
                    img = qdoc.resource(QtGui.QTextDocument.ResourceType.ImageResource, QtCore.QUrl(img_name))
                    if img and not img.isNull():
                        fd, tmp = tempfile.mkstemp(suffix=".png")
                        os.close(fd)
                        try:
                            img.save(tmp, "PNG")
                            w = img_fmt.width()
                            if w > 0:
                                run = p.add_run()
                                run.add_picture(tmp, width=Inches(w / 96.0))
                        finally:
                            # CRITICAL: Fixes memory leak if Word rendering fails
                            if os.path.exists(tmp): os.remove(tmp)
                else:
                    txt = fragment.text().replace('\u2028', '\n')
                    if txt:
                        run = p.add_run(txt)
                        font = fmt.font()
                        
                        run.font.bold = font.bold() or (fmt.fontWeight() >= QFont.Weight.Bold)
                        run.font.italic = font.italic() or fmt.fontItalic()
                        run.font.underline = font.underline() or fmt.fontUnderline()
                        run.font.strike = font.strikeOut() or fmt.fontStrikeOut()
                        
                        if fmt.verticalAlignment() == QTextCharFormat.VerticalAlignment.AlignSubScript:
                            run.font.subscript = True
                        elif fmt.verticalAlignment() == QTextCharFormat.VerticalAlignment.AlignSuperScript:
                            run.font.superscript = True
                        
                        f_size = font.pointSize()
                        if f_size <= 0:
                            p_size = font.pixelSize()
                            if p_size > 0: f_size = p_size * 0.75
                        if f_size > 0: run.font.size = Pt(f_size)
                        
                        fg_brush = fmt.foreground()
                        if fg_brush.style() != Qt.BrushStyle.NoBrush:
                            c = fg_brush.color()
                            if c.isValid() and c.alpha() > 0 and c.name().lower() != "#000000":
                                run.font.color.rgb = RGBColor(c.red(), c.green(), c.blue())
                        
                        bg_brush = fmt.background()
                        if bg_brush.style() != Qt.BrushStyle.NoBrush:
                            c = bg_brush.color()
                            if c.isValid() and c.alpha() > 0 and c.name().lower() != "#ffffff":
                                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            it += 1

    def process_table(table: QtGui.QTextTable, docx_parent):
        docx_table = docx_parent.add_table(rows=table.rows(), cols=table.columns())
        docx_table.style = 'Table Grid'
        
        docx_table.autofit = False
        try:
            total_width = Inches(6.5)
            col_width = total_width / table.columns()
            for col in docx_table.columns:
                col.width = col_width
            for row in docx_table.rows:
                for cell in row.cells:
                    cell.width = col_width
        except Exception: pass
        
        if WD_TABLE_ALIGNMENT is not None:
            t_align = table.format().alignment()
            if t_align & Qt.AlignmentFlag.AlignHCenter: docx_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            elif t_align & Qt.AlignmentFlag.AlignRight: docx_table.alignment = WD_TABLE_ALIGNMENT.RIGHT
            else: docx_table.alignment = WD_TABLE_ALIGNMENT.LEFT

        for r in range(table.rows()):
            for c in range(table.columns()):
                cell = table.cellAt(r, c)
                docx_cell = docx_table.cell(r, c)
                
                is_first_para = True
                it = cell.begin()
                while not it.atEnd():
                    child_frame = it.currentFrame()
                    child_block = it.currentBlock()
                    if child_frame:
                        if isinstance(child_frame, QtGui.QTextTable): process_table(child_frame, docx_cell)
                        else: process_frame(child_frame, docx_cell)
                    elif child_block.isValid(): 
                        if is_first_para and len(docx_cell.paragraphs) > 0:
                            process_paragraph(child_block, docx_parent=None, p=docx_cell.paragraphs[0])
                            is_first_para = False
                        else:
                            process_paragraph(child_block, docx_parent=docx_cell)
                    it += 1

    def process_frame(frame, docx_parent):
        it = frame.begin()
        while not it.atEnd():
            child_frame = it.currentFrame()
            child_block = it.currentBlock()
            if child_frame:
                if isinstance(child_frame, QtGui.QTextTable): process_table(child_frame, docx_parent)
                else: process_frame(child_frame, docx_parent)
            elif child_block.isValid(): process_paragraph(child_block, docx_parent=docx_parent)
            it += 1

    process_frame(qdoc.rootFrame(), doc)
    try: doc.save(filepath)
    except PermissionError: raise PermissionError(f"Cannot save to {filepath}. Please close the document if it is open in another program.")


# ---------------------------
# ADVANCED TEMPLATE MANAGER
# ---------------------------
class TemplateManagerDialog(QtWidgets.QDialog):
    def __init__(self, state: AppState, editor_ref: QtWidgets.QTextEdit, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Advanced Template Manager")
        self.resize(900, 600)
        self.state = state
        self.editor_ref = editor_ref

        layout = QtWidgets.QVBoxLayout(self)
        main_h = QtWidgets.QHBoxLayout()
        
        left_v = QtWidgets.QVBoxLayout()
        self.list_widget = QtWidgets.QListWidget()
        self.list_widget.currentTextChanged.connect(self._on_select)
        left_v.addWidget(QtWidgets.QLabel("<b>Your Templates</b>"))
        left_v.addWidget(self.list_widget)
        
        self.btn_insert = QtWidgets.QPushButton("Insert to Editor")
        self.btn_insert.setStyleSheet("background-color: #4CAF50; color: white; font-weight:bold;")
        self.btn_add_current = QtWidgets.QPushButton("Save Editor as Template")
        self.btn_delete = QtWidgets.QPushButton("Delete Selected")
        
        left_v.addWidget(self.btn_insert)
        left_v.addWidget(self.btn_add_current)
        left_v.addWidget(self.btn_delete)
        main_h.addLayout(left_v, 1)

        right_v = QtWidgets.QVBoxLayout()
        right_v.addWidget(QtWidgets.QLabel("<b>Preview / Edit Template HTML</b>"))
        self.preview_editor = QtWidgets.QTextEdit()
        self.preview_editor.setAcceptRichText(True)
        right_v.addWidget(self.preview_editor)
        
        self.btn_save_edit = QtWidgets.QPushButton("Update Template")
        right_v.addWidget(self.btn_save_edit)
        main_h.addLayout(right_v, 3)

        layout.addLayout(main_h)
        
        bot_h = QtWidgets.QHBoxLayout()
        self.btn_export = QtWidgets.QPushButton("Export Group to File...")
        self.btn_import = QtWidgets.QPushButton("Import Group from File...")
        self.btn_close = QtWidgets.QPushButton("Close")
        bot_h.addWidget(self.btn_export)
        bot_h.addWidget(self.btn_import)
        bot_h.addStretch()
        bot_h.addWidget(self.btn_close)
        layout.addLayout(bot_h)

        self.btn_insert.clicked.connect(self._insert_template)
        self.btn_add_current.clicked.connect(self._add_from_editor)
        self.btn_delete.clicked.connect(self._delete_template)
        self.btn_save_edit.clicked.connect(self._update_template)
        self.btn_export.clicked.connect(self._export_group)
        self.btn_import.clicked.connect(self._import_group)
        self.btn_close.clicked.connect(self.accept)

        self._refresh_list()

    def _refresh_list(self):
        self.list_widget.clear()
        self.list_widget.addItems(sorted(self.state.templates.keys()))

    def _on_select(self, name: str):
        if name and name in self.state.templates:
            self.preview_editor.setHtml(self.state.templates[name])

    def _insert_template(self):
        name = self.list_widget.currentItem().text() if self.list_widget.currentItem() else None
        if name:
            if hasattr(self.editor_ref, "_commit_composing"): self.editor_ref._commit_composing()
            self.editor_ref.insertHtml(self.state.templates[name])
            self.accept()

    def _add_from_editor(self):
        name, ok = QtWidgets.QInputDialog.getText(self, "Save Template", "Template Name:")
        if ok and name.strip():
            self.state.templates[name.strip()] = self.editor_ref.toHtml()
            self.state.save_templates()
            self._refresh_list()
            self.list_widget.setCurrentItem(self.list_widget.findItems(name.strip(), Qt.MatchFlag.MatchExactly)[0])
            QtWidgets.QMessageBox.information(self, "Success", "Template saved!")

    def _delete_template(self):
        item = self.list_widget.currentItem()
        if item:
            name = item.text()
            ans = QtWidgets.QMessageBox.question(self, "Delete", f"Delete template '{name}'?")
            if ans == QtWidgets.QMessageBox.StandardButton.Yes:
                del self.state.templates[name]
                self.state.save_templates()
                self._refresh_list()
                self.preview_editor.clear()

    def _update_template(self):
        item = self.list_widget.currentItem()
        if item:
            name = item.text()
            self.state.templates[name] = self.preview_editor.toHtml()
            self.state.save_templates()
            QtWidgets.QMessageBox.information(self, "Success", f"Template '{name}' updated.")

    def _export_group(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Export Template Group", "", "JSON Files (*.json)")
        if path:
            try:
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(self.state.templates, f, ensure_ascii=False, indent=2)
                QtWidgets.QMessageBox.information(self, "Exported", f"Templates exported to {path}")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", str(e))

    def _import_group(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Import Template Group", "", "JSON Files (*.json)")
        if path:
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        self.state.templates.update(data)
                        self.state.save_templates()
                        self._refresh_list()
                        QtWidgets.QMessageBox.information(self, "Imported", "Templates loaded successfully!")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", str(e))


# ---------------------------
# Speech worker (QThread)
# ---------------------------
class SpeechWorker(QtCore.QThread):
    partial = QtCore.Signal(str)
    finished_text = QtCore.Signal(str)
    error = QtCore.Signal(str)

    def __init__(self, lang: str = "hi-IN", parent=None):
        super().__init__(parent)
        self._stop = False
        self.lang = lang
        self._accum: List[str] = []

    def stop(self):
        self._stop = True

    def run(self):
        if not HAS_SR:
            self.error.emit("SpeechRecognition not installed.")
            return
        try:
            rec = sr.Recognizer()
            with sr.Microphone() as source:
                try: rec.adjust_for_ambient_noise(source, duration=0.4)
                except Exception: pass
                while not self._stop:
                    try: 
                        # Use short timeouts so thread can exit cleanly
                        audio = rec.listen(source, timeout=0.5, phrase_time_limit=5)
                    except sr.WaitTimeoutError:
                        continue
                    except Exception:
                        if self._stop: break
                        continue
                    try:
                        text = rec.recognize_google(audio, language=self.lang)
                        self._accum.append(text)
                        self.partial.emit(text)
                    except Exception: continue
        except Exception as e: self.error.emit(str(e))
        finally:
            self.finished_text.emit(" ".join(self._accum).strip())


# ---------------------------
# Suggestions Database Manager
# ---------------------------
class SuggestionsManagerDialog(QtWidgets.QDialog):
    def __init__(self, state: AppState, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Suggestions Database Manager")
        self.resize(750, 600)
        self.state = state

        layout = QtWidgets.QVBoxLayout(self)
        
        self.lbl_dict = QtWidgets.QLabel(f"<b>Active Database:</b> {self.state.active_sugg_name} | <b>Total Words:</b> {len(self.state.suggestion_words)}")
        self.lbl_dict.setStyleSheet("font-size: 14px; padding: 5px; color: #333; background: #e0e0e0;")
        layout.addWidget(self.lbl_dict)

        self.table = QtWidgets.QTableWidget(0, 1)
        self.table.setHorizontalHeaderLabels(["Suggestion Word"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSortingEnabled(True)
        layout.addWidget(self.table)

        btns_top = QtWidgets.QHBoxLayout()
        self.btn_add = QtWidgets.QPushButton("Add Word")
        self.btn_remove = QtWidgets.QPushButton("Remove Selected")
        self.btn_apply = QtWidgets.QPushButton("Apply (runtime)")
        btns_top.addWidget(self.btn_add); btns_top.addWidget(self.btn_remove); btns_top.addWidget(self.btn_apply)
        layout.addLayout(btns_top)

        btns_mid = QtWidgets.QHBoxLayout()
        self.btn_load_def = QtWidgets.QPushButton("Load Default")
        self.btn_save_def = QtWidgets.QPushButton("Save Default")
        self.btn_load_man = QtWidgets.QPushButton("Load Manual File...")
        self.btn_save_man = QtWidgets.QPushButton("Save Manual File...")
        btns_mid.addWidget(self.btn_load_def); btns_mid.addWidget(self.btn_save_def)
        btns_mid.addWidget(self.btn_load_man); btns_mid.addWidget(self.btn_save_man)
        layout.addLayout(btns_mid)

        btns_bot = QtWidgets.QHBoxLayout()
        self.btn_scan = QtWidgets.QPushButton("Scan/Extract Words from File...")
        self.btn_scan.setStyleSheet("background-color: #4CAF50; color: white; font-weight: bold;")
        self.btn_close = QtWidgets.QPushButton("Close")
        btns_bot.addWidget(self.btn_scan)
        btns_bot.addStretch()
        btns_bot.addWidget(self.btn_close)
        layout.addLayout(btns_bot)

        self.btn_add.clicked.connect(self.add_row)
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_apply.clicked.connect(self.apply_runtime)
        self.btn_load_def.clicked.connect(self.load_default)
        self.btn_save_def.clicked.connect(lambda: self.save_to_file(SUGGESTIONS_PATH))
        self.btn_load_man.clicked.connect(self.load_manual)
        self.btn_save_man.clicked.connect(lambda: self.save_to_file(None))
        self.btn_scan.clicked.connect(self.scan_file)
        self.btn_close.clicked.connect(self.accept)

        self.reload_table()

    def reload_table(self, data_set=None):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)
        source = data_set if data_set is not None else self.state.suggestion_words
        
        sample = list(source)[:2000]
        for w in sample:
            r = self.table.rowCount()
            self.table.insertRow(r)
            self.table.setItem(r, 0, QtWidgets.QTableWidgetItem(w))
            
        self.table.setSortingEnabled(True)
        self.lbl_dict.setText(f"<b>Active Database:</b> {self.state.active_sugg_name} | <b>Total Words:</b> {len(source)}")

    def add_row(self):
        self.table.setSortingEnabled(False)
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QtWidgets.QTableWidgetItem(""))
        self.table.editItem(self.table.item(r, 0))
        self.table.setSortingEnabled(True)

    def remove_selected(self):
        rows = sorted({i.row() for i in self.table.selectedIndexes()}, reverse=True)
        for r in rows: self.table.removeRow(r)

    def apply_runtime(self):
        new = set()
        for r in range(self.table.rowCount()):
            w = self.table.item(r, 0)
            if w and w.text().strip(): new.add(w.text().strip())
        self.state.suggestion_words = new
        self.state.active_sugg_name = "Runtime / Unsaved"
        self.lbl_dict.setText(f"<b>Active Database:</b> {self.state.active_sugg_name} | <b>Total Words:</b> {len(self.state.suggestion_words)}")
        QtWidgets.QMessageBox.information(self, "Applied", "Suggestions applied to runtime memory.")

    def save_to_file(self, target_path=None):
        if not target_path:
            target_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save Manual Suggestions", "", "JSON Files (*.json)")
            if not target_path: return
        self.apply_runtime()
        try:
            with open(target_path, "w", encoding="utf-8") as f:
                json.dump(list(self.state.suggestion_words), f, ensure_ascii=False, indent=2)
            self.state.active_sugg_name = os.path.basename(str(target_path))
            self.lbl_dict.setText(f"<b>Active Database:</b> {self.state.active_sugg_name} | <b>Total Words:</b> {len(self.state.suggestion_words)}")
            QtWidgets.QMessageBox.information(self, "Saved", f"Saved successfully to: {target_path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Save failed", str(e))

    def load_default(self):
        if SUGGESTIONS_PATH.exists():
            try:
                with open(SUGGESTIONS_PATH, "r", encoding="utf-8") as f: data = json.load(f)
                if isinstance(data, list):
                    self.state.suggestion_words = set(data)
                    self.state.active_sugg_name = "Default (suggestions.json)"
                    self.reload_table()
            except Exception as e: QtWidgets.QMessageBox.critical(self, "Load failed", str(e))

    def load_manual(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Load Manual Suggestions", "", "JSON Files (*.json)")
        if path:
            try:
                with open(path, "r", encoding="utf-8") as f: data = json.load(f)
                if isinstance(data, list):
                    self.state.suggestion_words = set(data)
                    self.state.active_sugg_name = os.path.basename(path)
                    self.reload_table()
            except Exception as e: QtWidgets.QMessageBox.critical(self, "Error", str(e))

    def scan_file(self):
        filters = "Documents (*.txt *.csv *.xlsx *.ods *.docx);;All Files (*.*)"
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Scan Words from File", "", filters)
        if not path: return
        
        ext = path.lower().split('.')[-1]
        extracted_words: Set[str] = set()
        
        try:
            if ext == "txt":
                with open(path, "r", encoding="utf-8") as f:
                    extracted_words = set(re.findall(r'[a-zA-Z]+|[\u0900-\u097F]+', f.read()))
            elif ext == "docx" and HAS_DOCX:
                doc = docx.Document(path)
                full_text = " ".join([p.text for p in doc.paragraphs])
                extracted_words = set(re.findall(r'[a-zA-Z]+|[\u0900-\u097F]+', full_text))
            elif ext in ["csv", "xlsx", "ods"]:
                if not HAS_PANDAS:
                    QtWidgets.QMessageBox.warning(self, "Pandas Required", "Please install pandas and openpyxl to scan spreadsheets.")
                    return
                if ext == "xlsx":
                    df = pd.read_excel(path, engine="openpyxl")
                elif ext == "ods":
                    df = pd.read_excel(path, engine="odf")
                else:
                    df = pd.read_csv(path)
                    
                df = df.fillna("") 
                full_text = " ".join([str(val) for val in df.values.flatten() if str(val).strip()])
                extracted_words = set(re.findall(r'[a-zA-Z]+|[\u0900-\u097F]+', full_text))
            else:
                QtWidgets.QMessageBox.warning(self, "Unsupported", "File format not supported for scanning.")
                return

            added = 0
            for w in extracted_words:
                if len(w) > 1 and w not in self.state.suggestion_words:
                    self.state.suggestion_words.add(w)
                    added += 1
                    
            self.reload_table()
            QtWidgets.QMessageBox.information(self, "Scan Complete", f"Extracted and added {added} NEW unique words.")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Scan Error", f"Failed to extract words: {str(e)}")


# ---------------------------
# Dictionary UI (Corrections)
# ---------------------------
class CorrectionsDialog(QtWidgets.QDialog):
    def __init__(self, state: AppState, editor_ref: Optional[QtWidgets.QTextEdit] = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Dictionary Training & File Scanner")
        self.resize(950, 650)
        self.state = state
        self.editor_ref = editor_ref

        layout = QtWidgets.QVBoxLayout(self)
        
        self.lbl_dict = QtWidgets.QLabel(f"<b>Active Dictionary:</b> {self.state.active_dict_name}")
        self.lbl_dict.setStyleSheet("font-size: 14px; padding: 5px; color: #333; background: #e0e0e0;")
        layout.addWidget(self.lbl_dict)

        self.table = QtWidgets.QTableWidget(0, 2)
        self.table.setHorizontalHeaderLabels(["Shortcut / Latin (key)", "Devanagari / Value (value)"])
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setSortingEnabled(True)
        layout.addWidget(self.table)

        btns_top = QtWidgets.QHBoxLayout()
        self.btn_add = QtWidgets.QPushButton("Add Row")
        self.btn_remove = QtWidgets.QPushButton("Remove Selected")
        self.btn_apply = QtWidgets.QPushButton("Apply (runtime)")
        btns_top.addWidget(self.btn_add); btns_top.addWidget(self.btn_remove); btns_top.addWidget(self.btn_apply)
        layout.addLayout(btns_top)

        btns_mid = QtWidgets.QHBoxLayout()
        self.btn_load_def = QtWidgets.QPushButton("Load Default")
        self.btn_save_def = QtWidgets.QPushButton("Save Default")
        self.btn_load_man = QtWidgets.QPushButton("Load Manual File...")
        self.btn_save_man = QtWidgets.QPushButton("Save Manual File...")
        btns_mid.addWidget(self.btn_load_def); btns_mid.addWidget(self.btn_save_def)
        btns_mid.addWidget(self.btn_load_man); btns_mid.addWidget(self.btn_save_man)
        layout.addLayout(btns_mid)

        btns_bot = QtWidgets.QHBoxLayout()
        self.btn_scan = QtWidgets.QPushButton("Scan/Extract Words from File...")
        self.btn_scan.setStyleSheet("background-color: #2196F3; color: white; font-weight: bold;")
        self.btn_apply_sel = QtWidgets.QPushButton("Apply mapping to selection")
        self.btn_close = QtWidgets.QPushButton("Close")
        btns_bot.addWidget(self.btn_scan)
        btns_bot.addWidget(self.btn_apply_sel)
        btns_bot.addStretch()
        btns_bot.addWidget(self.btn_close)
        layout.addLayout(btns_bot)

        self.btn_add.clicked.connect(self.add_row)
        self.btn_remove.clicked.connect(self.remove_selected)
        self.btn_apply.clicked.connect(self.apply_runtime)
        self.btn_load_def.clicked.connect(self.load_default)
        self.btn_save_def.clicked.connect(lambda: self.save_to_file(USER_DICT_PATH))
        self.btn_load_man.clicked.connect(self.load_manual)
        self.btn_save_man.clicked.connect(lambda: self.save_to_file(None))
        self.btn_scan.clicked.connect(self.scan_file)
        self.btn_apply_sel.clicked.connect(self.apply_mapping_to_selection)
        self.btn_close.clicked.connect(self.accept)

        self.reload_table()

    def reload_table(self, data_dict=None):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)
        source = data_dict if data_dict is not None else self.state.user_dict
        for k, v in source.items():
            r = self.table.rowCount()
            self.table.insertRow(r)
            self.table.setItem(r, 0, QtWidgets.QTableWidgetItem(k))
            self.table.setItem(r, 1, QtWidgets.QTableWidgetItem(v))
        self.table.setSortingEnabled(True)

    def add_row(self):
        self.table.setSortingEnabled(False)
        r = self.table.rowCount()
        self.table.insertRow(r)
        self.table.setItem(r, 0, QtWidgets.QTableWidgetItem(""))
        self.table.setItem(r, 1, QtWidgets.QTableWidgetItem(""))
        self.table.editItem(self.table.item(r, 0))
        self.table.setSortingEnabled(True)

    def remove_selected(self):
        rows = sorted({i.row() for i in self.table.selectedIndexes()}, reverse=True)
        for r in rows: self.table.removeRow(r)

    def apply_runtime(self):
        new = {}
        for r in range(self.table.rowCount()):
            k, v = self.table.item(r, 0), self.table.item(r, 1)
            if k and v and k.text().strip(): new[k.text().strip()] = v.text()
        self.state.user_dict = new
        self.state.active_dict_name = "Runtime / Unsaved"
        self.lbl_dict.setText(f"<b>Active Dictionary:</b> {self.state.active_dict_name}")
        QtWidgets.QMessageBox.information(self, "Applied", "Corrections applied to runtime memory.")

    def save_to_file(self, target_path=None):
        if not target_path:
            target_path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save Manual Dictionary", "", "JSON Files (*.json)")
            if not target_path: return
        new = {}
        for r in range(self.table.rowCount()):
            k, v = self.table.item(r, 0), self.table.item(r, 1)
            if k and v and k.text().strip(): new[k.text().strip()] = v.text()
        try:
            with open(target_path, "w", encoding="utf-8") as f:
                json.dump(new, f, ensure_ascii=False, indent=2)
            self.state.user_dict = new
            self.state.active_dict_name = os.path.basename(str(target_path))
            self.lbl_dict.setText(f"<b>Active Dictionary:</b> {self.state.active_dict_name}")
            QtWidgets.QMessageBox.information(self, "Saved", f"Saved successfully to: {target_path}")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Save failed", str(e))

    def load_default(self):
        if USER_DICT_PATH.exists():
            try:
                with open(USER_DICT_PATH, "r", encoding="utf-8") as f: data = json.load(f)
                if isinstance(data, dict):
                    self.state.user_dict = data
                    self.state.active_dict_name = "Default (user_translit.json)"
                    self.lbl_dict.setText(f"<b>Active Dictionary:</b> {self.state.active_dict_name}")
                    self.reload_table()
            except Exception as e: QtWidgets.QMessageBox.critical(self, "Load failed", str(e))

    def load_manual(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Load Manual Dictionary", "", "JSON Files (*.json)")
        if path:
            try:
                with open(path, "r", encoding="utf-8") as f: data = json.load(f)
                if isinstance(data, dict):
                    self.state.user_dict = data
                    self.state.active_dict_name = os.path.basename(path)
                    self.lbl_dict.setText(f"<b>Active Dictionary:</b> {self.state.active_dict_name}")
                    self.reload_table()
            except Exception as e: QtWidgets.QMessageBox.critical(self, "Error", str(e))

    def scan_file(self):
        filters = "Documents (*.txt *.csv *.xlsx *.ods *.docx);;All Files (*.*)"
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Scan Words from File", "", filters)
        if not path: return
        
        ext = path.lower().split('.')[-1]
        extracted_words: Set[str] = set()
        
        try:
            if ext == "txt":
                with open(path, "r", encoding="utf-8") as f:
                    extracted_words = set(re.findall(r'[a-zA-Z]+|[\u0900-\u097F]+', f.read()))
            elif ext == "docx" and HAS_DOCX:
                doc = docx.Document(path)
                full_text = " ".join([p.text for p in doc.paragraphs])
                extracted_words = set(re.findall(r'[a-zA-Z]+|[\u0900-\u097F]+', full_text))
            elif ext in ["csv", "xlsx", "ods"]:
                if not HAS_PANDAS:
                    QtWidgets.QMessageBox.warning(self, "Pandas Required", "Please install pandas and openpyxl to scan spreadsheets.")
                    return
                if ext == "xlsx":
                    df = pd.read_excel(path, engine="openpyxl")
                elif ext == "ods":
                    df = pd.read_excel(path, engine="odf")
                else:
                    df = pd.read_csv(path)
                    
                df = df.fillna("") 
                full_text = " ".join([str(val) for val in df.values.flatten() if str(val).strip()])
                extracted_words = set(re.findall(r'[a-zA-Z]+|[\u0900-\u097F]+', full_text))
            else:
                QtWidgets.QMessageBox.warning(self, "Unsupported", "File format not supported for scanning.")
                return

            existing_words = set()
            for r in range(self.table.rowCount()):
                item = self.table.item(r, 0)
                if item and item.text().strip():
                    existing_words.add(item.text().strip().lower())

            added = 0
            self.table.setSortingEnabled(False)
            for w in extracted_words:
                w_lower = w.lower()
                if w_lower not in existing_words and not w_lower.isnumeric():
                    r = self.table.rowCount()
                    self.table.insertRow(r)
                    
                    is_devanagari = bool(re.search(r"[\u0900-\u097F]", w_lower))
                    if is_devanagari:
                        self.table.setItem(r, 0, QtWidgets.QTableWidgetItem(""))
                        self.table.setItem(r, 1, QtWidgets.QTableWidgetItem(w_lower))
                    else:
                        self.table.setItem(r, 0, QtWidgets.QTableWidgetItem(w_lower))
                        self.table.setItem(r, 1, QtWidgets.QTableWidgetItem(""))
                        
                    existing_words.add(w_lower)
                    added += 1
                    
            self.table.setSortingEnabled(True)
            QtWidgets.QMessageBox.information(self, "Scan Complete", f"Extracted and added {added} NEW unique words.")
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Scan Error", f"Failed to extract words: {str(e)}")

    def apply_mapping_to_selection(self):
        if not self.editor_ref: return
        cur = self.editor_ref.textCursor()
        if not cur.hasSelection(): cur.select(QtGui.QTextCursor.SelectionType.WordUnderCursor)
        sel = cur.selectedText()
        if not sel: return
        
        mappings = [(self.table.item(r, 0).text().strip(), self.table.item(r, 1).text()) 
                    for r in range(self.table.rowCount()) if self.table.item(r, 0) and self.table.item(r, 1)]
        if not mappings: return
        
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Choose mapping")
        dlg.resize(480, 360)
        v = QtWidgets.QVBoxLayout(dlg)
        lw = QtWidgets.QListWidget(dlg)
        for k, v_ in mappings: lw.addItem(f"{k}  →  {v_}")
        v.addWidget(lw)
        h = QtWidgets.QHBoxLayout()
        b_apply = QtWidgets.QPushButton("Apply")
        b_apply.clicked.connect(lambda: self._do_apply(lw, mappings, cur, dlg))
        h.addWidget(b_apply); h.addWidget(QtWidgets.QPushButton("Cancel", clicked=dlg.reject))
        v.addLayout(h)
        dlg.exec()

    def _do_apply(self, lw, mappings, cur, dlg):
        idx = lw.currentRow()
        if idx >= 0:
            cur.insertText(mappings[idx][1])
            dlg.accept()


# ---------------------------
# Find & Replace UI
# ---------------------------
class FindReplaceDialog(QtWidgets.QDialog):
    def __init__(self, editor: QtWidgets.QTextEdit, parent=None):
        super().__init__(parent)
        self.editor = editor
        self.setWindowTitle("Find & Replace")
        self.resize(520, 180)
        v = QtWidgets.QVBoxLayout(self)
        form = QtWidgets.QFormLayout()
        self.find_edit = QtWidgets.QLineEdit()
        self.replace_edit = QtWidgets.QLineEdit()
        form.addRow("Find:", self.find_edit)
        form.addRow("Replace:", self.replace_edit)
        v.addLayout(form)
        opt = QtWidgets.QHBoxLayout()
        self.ck_case = QtWidgets.QCheckBox("Match case")
        self.ck_word = QtWidgets.QCheckBox("Whole word")
        opt.addWidget(self.ck_case); opt.addWidget(self.ck_word); opt.addStretch(1)
        v.addLayout(opt)
        h = QtWidgets.QHBoxLayout()
        self.btn_prev = QtWidgets.QPushButton("Find Prev")
        self.btn_next = QtWidgets.QPushButton("Find Next")
        self.btn_rep = QtWidgets.QPushButton("Replace")
        self.btn_all = QtWidgets.QPushButton("Replace All")
        self.btn_close = QtWidgets.QPushButton("Close")
        for b in (self.btn_prev, self.btn_next, self.btn_rep, self.btn_all, self.btn_close): h.addWidget(b)
        v.addLayout(h)
        self.btn_close.clicked.connect(self.reject)
        self.btn_prev.clicked.connect(lambda: self._find(backward=True))
        self.btn_next.clicked.connect(lambda: self._find(backward=False))
        self.btn_rep.clicked.connect(self._replace_once)
        self.btn_all.clicked.connect(self._replace_all)

    def _flags(self):
        flags = QtGui.QTextDocument.FindFlag(0)
        if self.ck_case.isChecked(): flags |= QtGui.QTextDocument.FindFlag.FindCaseSensitively
        if self.ck_word.isChecked(): flags |= QtGui.QTextDocument.FindFlag.FindWholeWords
        return flags

    def _find(self, backward=False):
        needle = self.find_edit.text()
        if not needle: return
        flags = self._flags()
        if backward: flags |= QtGui.QTextDocument.FindFlag.FindBackward
        found = self.editor.find(needle, flags)
        if not found:
            c = self.editor.textCursor()
            c.movePosition(QtGui.QTextCursor.MoveOperation.End if backward else QtGui.QTextCursor.MoveOperation.Start)
            self.editor.setTextCursor(c)
            self.editor.find(needle, flags)

    def _replace_once(self):
        needle = self.find_edit.text()
        rep = self.replace_edit.text()
        if not needle: return
        cur = self.editor.textCursor()
        if cur.hasSelection() and cur.selectedText() == needle: cur.insertText(rep)
        self._find(backward=False)

    def _replace_all(self):
        needle = self.find_edit.text()
        rep = self.replace_edit.text()
        if not needle: return
        flags = self._flags()
        doc = self.editor.document()
        c = QtGui.QTextCursor(doc)
        c.beginEditBlock()
        c.movePosition(QtGui.QTextCursor.MoveOperation.Start)
        self.editor.setTextCursor(c)
        count = 0
        while self.editor.find(needle, flags):
            cur = self.editor.textCursor()
            cur.insertText(rep)
            count += 1
        c.endEditBlock()
        QtWidgets.QMessageBox.information(self, "Replace All", f"Replaced {count} occurrence(s).")


# ---------------------------
# Core Editor
# ---------------------------
class HindiEditor(QtWidgets.QTextEdit):
    countsChanged = QtCore.Signal(int, int)
    cursorPositionChangedDetailed = QtCore.Signal(int, int)
    contextActionTriggered = QtCore.Signal(str) 
    prefixChanged = QtCore.Signal(str)
    navigateSuggestion = QtCore.Signal(int)
    insertSuggestionTrigger = QtCore.Signal()

    def __init__(self, translit: AdaptiveTransliterator, state: AppState, parent=None):
        super().__init__(parent)
        self.translit = translit
        self.state = state
        self.setAcceptRichText(True)
        f = QtGui.QFont(self.state.font_family, self.state.font_size)
        self.setFont(f)
        self.document().setDefaultFont(f)
        self.setPlaceholderText("Type phonetically (Latin) — Devanagari will appear")

        self._composing_latin = ""
        self._composing_start_pos: Optional[int] = None
        self._composing_display_len = 0
        self._ignore_cursor_move = False
        self._has_suggestions = False
        
        self.textChanged.connect(self._on_text_changed)
        self.cursorPositionChanged.connect(self._on_cursor_moved)

    def set_has_suggestions(self, val: bool):
        self._has_suggestions = val

    def keyPressEvent(self, ev: QtGui.QKeyEvent):
        key, mods = ev.key(), ev.modifiers()
        
        if key == Qt.Key.Key_Tab and self._has_suggestions:
            self.insertSuggestionTrigger.emit()
            return

        if (key == Qt.Key.Key_Up or key == Qt.Key.Key_Down) and (mods & Qt.KeyboardModifier.ControlModifier):
            step = -1 if key == Qt.Key.Key_Up else 1
            self.navigateSuggestion.emit(step)
            return
            
        if mods & Qt.KeyboardModifier.ControlModifier: 
            return super().keyPressEvent(ev)

        if key in {Qt.Key.Key_Left, Qt.Key.Key_Right, Qt.Key.Key_Up, Qt.Key.Key_Down, Qt.Key.Key_Home, Qt.Key.Key_End}:
            if self._composing_latin: self._commit_composing()
            return super().keyPressEvent(ev)

        if key == Qt.Key.Key_Backspace:
            if self._composing_latin:
                self._composing_latin = self._composing_latin[:-1]
                if self._composing_latin: 
                    self._update_composing()
                    self.prefixChanged.emit(self.translit.translit_token(self._composing_latin) if self.state.transliteration_enabled else self._composing_latin)
                else: 
                    self._remove_composing()
                    self.prefixChanged.emit("")
                return
            return super().keyPressEvent(ev)

        if key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            if self._composing_latin: self._commit_composing()
            return super().keyPressEvent(ev)

        txt = ev.text()
        if not txt or re.search(r"[\u0900-\u097F]", txt):
            if self._composing_latin: self._commit_composing()
            return super().keyPressEvent(ev)

        if txt.isprintable() and len(txt) == 1 and not txt.isspace() and ord(txt) < 128:
            cur = self.textCursor()
            if not self._composing_latin:
                if cur.hasSelection(): cur.removeSelectedText()
                start = cur.position()
                self._composing_start_pos = start
                self._composing_latin = txt
                
                translit_text = self.translit.translit_token(self._composing_latin) if self.state.transliteration_enabled else self._composing_latin
                fmt = cur.charFormat()
                
                self._ignore_cursor_move = True
                cur.insertText(translit_text, fmt)
                self._composing_display_len = len(translit_text)
                self.setTextCursor(cur)
                self._ignore_cursor_move = False
                
                self.prefixChanged.emit(translit_text)
                return
            else:
                self._composing_latin += txt
                translit_text = self.translit.translit_token(self._composing_latin) if self.state.transliteration_enabled else self._composing_latin
                self._update_composing_with_text(translit_text)
                
                self.prefixChanged.emit(translit_text)
                return

        if self._composing_latin: self._commit_composing()
        return super().keyPressEvent(ev)

    def _on_text_changed(self): self._emit_counts()

    def _on_cursor_moved(self):
        if self._ignore_cursor_move: return
        if self._composing_latin: self._commit_composing()
        self._emit_prefix()
        
        cur = self.textCursor()
        line = cur.blockNumber() + 1
        col = cur.positionInBlock()
        self.cursorPositionChangedDetailed.emit(line, col)

    def _emit_prefix(self):
        if self._ignore_cursor_move: return
        if self._composing_latin:
            prefix = self.translit.translit_token(self._composing_latin) if self.state.transliteration_enabled else self._composing_latin
            self.prefixChanged.emit(prefix)
        else:
            cur = self.textCursor()
            cur.movePosition(QtGui.QTextCursor.MoveOperation.StartOfWord, QtGui.QTextCursor.MoveMode.KeepAnchor)
            self.prefixChanged.emit(cur.selectedText().strip())

    def insert_suggestion(self, word: str):
        cur = self.textCursor()
        if self._composing_latin: self._remove_composing()
        else:
            cur.movePosition(QtGui.QTextCursor.MoveOperation.StartOfWord, QtGui.QTextCursor.MoveMode.KeepAnchor)
            cur.removeSelectedText()
            
        cur.insertText(word + " ")
        self.setTextCursor(cur)
        self.setFocus()
        self._has_suggestions = False
        self.prefixChanged.emit("")

    def _update_composing_with_text(self, translit_text: str):
        if self._composing_start_pos is None: return
        self._ignore_cursor_move = True
        cur = self.textCursor()
        cur.setPosition(self._composing_start_pos)
        cur.setPosition(self._composing_start_pos + self._composing_display_len, QtGui.QTextCursor.MoveMode.KeepAnchor)
        fmt = cur.charFormat()
        cur.insertText(translit_text, fmt)
        self._composing_display_len = len(translit_text)
        self.setTextCursor(cur)
        self._ignore_cursor_move = False

    def _update_composing(self):
        t_text = self.translit.translit_token(self._composing_latin) if self.state.transliteration_enabled else self._composing_latin
        self._update_composing_with_text(t_text)

    def _remove_composing(self):
        if self._composing_start_pos is None: return
        self._ignore_cursor_move = True
        cur = self.textCursor()
        cur.setPosition(self._composing_start_pos)
        cur.setPosition(self._composing_start_pos + self._composing_display_len, QtGui.QTextCursor.MoveMode.KeepAnchor)
        cur.removeSelectedText()
        self._composing_latin = ""
        self._composing_start_pos = None
        self._composing_display_len = 0
        self.setTextCursor(cur)
        self._ignore_cursor_move = False

    def _commit_composing(self):
        if not self._composing_latin: return
        self._composing_latin = ""
        self._composing_start_pos = None
        self._composing_display_len = 0
        self._emit_counts()

    def paste(self):
        md = QtWidgets.QApplication.clipboard().mimeData()
        text = md.text() if md else ""
        if not text: return super().paste()
        if re.search(r"[\u0900-\u097F]", text):
            if self._composing_latin: self._commit_composing()
            return super().paste()
        parts = re.split(r"(\s+)", text)
        for p in parts:
            if p.strip() == "": super().insertPlainText(p)
            else:
                chunk = self.translit.translit_token(p) if self.state.transliteration_enabled else p
                super().insertPlainText(chunk)
        self._emit_counts()

    def _emit_counts(self):
        txt = self.toPlainText()
        w = len([w for w in re.split(r"\s+", txt.strip()) if w]) if txt.strip() else 0
        self.countsChanged.emit(w, len(txt))

    def contextMenuEvent(self, e):
        menu = self.createStandardContextMenu()
        menu.addSeparator()
        
        cur = self.textCursor()
        table = cur.currentTable()
        if table:
            t_menu = menu.addMenu("Table Alignment (Page)")
            t_menu.addAction("Align Left", lambda: self.contextActionTriggered.emit("table_left"))
            t_menu.addAction("Align Center", lambda: self.contextActionTriggered.emit("table_center"))
            t_menu.addAction("Align Right", lambda: self.contextActionTriggered.emit("table_right"))
            menu.addSeparator()

        menu.addAction("Train map from selection", lambda: self.contextActionTriggered.emit("train_sel"))
        menu.addAction("Edit Correction Dictionary", lambda: self.contextActionTriggered.emit("edit_dict"))
        menu.addSeparator()
        menu.addAction("New Document", lambda: self.contextActionTriggered.emit("new_doc"))
        menu.addAction("Open...", lambda: self.contextActionTriggered.emit("open_doc"))
        menu.addAction("Export as PDF", lambda: self.contextActionTriggered.emit("export_pdf"))
        
        sel = cur.selectedText()
        if sel:
            matches = [(k, v) for k, v in self.state.user_dict.items() if k in sel]
            if matches:
                menu.addSeparator()
                sub = menu.addMenu("Apply correction")
                for k, v in matches:
                    a = sub.addAction(f"{k} → {v}")
                    a.triggered.connect(lambda _, kk=k, vv=v: self._apply_correction_in_selection(kk, vv))
        menu.exec(e.globalPos())

    def _apply_correction_in_selection(self, key, val):
        cur = self.textCursor()
        if not cur.hasSelection(): cur.select(QtGui.QTextCursor.SelectionType.WordUnderCursor)
        text = cur.selectedText()
        cur.insertText(text.replace(key, val))

    def _align_table(self, alignment):
        cur = self.textCursor()
        table = cur.currentTable()
        if table:
            fmt = table.format()
            fmt.setAlignment(alignment)
            table.setFormat(fmt)
            self.textChanged.emit()


# ---------------------------
# Main Window UI & Logic
# ---------------------------
class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1200, 820)
        self.setWindowIcon(QIcon("icons/icon.png"))        
        self.state = AppState()
        self.state.load()
        
        self.current_filepath: Optional[str] = None
        self.dark_mode_enabled = False
        self.print_layout_enabled = False
        
        self.printer = QPrinter(QPrinter.PrinterMode.HighResolution)
        
        self.autosave_enabled: bool = False
        self.autosave_timer = QTimer(self)
        self.autosave_timer.setInterval(30000) 
        self.autosave_timer.timeout.connect(self._perform_autosave)
        self.is_dirty = False

        try: self.translit = AdaptiveTransliterator(self.state)
        except RuntimeError as e:
            QtWidgets.QMessageBox.critical(None, "Missing dependency", str(e))
            raise SystemExit(1)

        self.editor = HindiEditor(self.translit, self.state, self)
        self.setCentralWidget(self.editor)
        self.editor.countsChanged.connect(self._update_counts)
        self.editor.cursorPositionChangedDetailed.connect(self._update_cursor_position_label)
        self.editor.contextActionTriggered.connect(self._handle_context_action)
        self.editor.textChanged.connect(self._mark_dirty)

        self.finddlg = FindReplaceDialog(self.editor, self)
        
        self.sr_worker: Optional[SpeechWorker] = None
        self.sr_listening = False
        self._ctrl_l_down = False
        QtWidgets.QApplication.instance().installEventFilter(self)
        
        self._setup_actions() 
        self._build_toolbars()
        self._build_menus()
        
        self._build_sidebar()
        self._build_suggestions_sidebar()
        
        self.editor.prefixChanged.connect(self._update_suggestions)
        self.editor.navigateSuggestion.connect(self._navigate_suggestion_list)
        self.editor.insertSuggestionTrigger.connect(self._insert_selected_suggestion)

        self._build_statusbar()
        self._update_translit_badge()
        
        QTimer.singleShot(100, lambda: self._new_file(prompt_autosave=True))

    def eventFilter(self, watched, event):
        try:
            if event.type() == QtCore.QEvent.Type.KeyPress:
                key_event: QtGui.QKeyEvent = event
                if key_event.key() == Qt.Key.Key_F11:
                    if self.isFullScreen(): self.showNormal()
                    else: self.showFullScreen()
                    return True
                    
                if key_event.key() == Qt.Key.Key_L and (key_event.modifiers() & Qt.KeyboardModifier.ControlModifier) and not key_event.isAutoRepeat():
                    if not self.sr_listening:
                        self._start_sr_listening()
                    self._ctrl_l_down = True
                    return True
            elif event.type() == QtCore.QEvent.Type.KeyRelease:
                key_event: QtGui.QKeyEvent = event
                if key_event.key() == Qt.Key.Key_L and not key_event.isAutoRepeat():
                    if self.sr_listening and self._ctrl_l_down:
                        self._stop_sr_listening_and_insert()
                    self._ctrl_l_down = False
                    return True
        except Exception:
            pass
        return super().eventFilter(watched, event)

    def _start_sr_listening(self):
        if not HAS_SR:
            QtWidgets.QMessageBox.information(self, "Speech missing", "Install SpeechRecognition and PyAudio: pip install SpeechRecognition pyaudio")
            return
        if self.sr_worker and self.sr_worker.isRunning():
            return
        self.sr_worker = SpeechWorker(lang="hi-IN", parent=self)
        self.sr_worker.partial.connect(self._update_speech_label_partial)
        self.sr_worker.finished_text.connect(self._sr_finished)
        self.sr_worker.error.connect(lambda m: self.status.showMessage(f"Speech error: {m}", 4000))
        self.sr_worker.start()
        self.sr_listening = True
        self.speech_label.setText("Listening... release Ctrl+L to stop")

    @QtCore.Slot(str)
    def _update_speech_label_partial(self, s: str):
        if s: self.speech_label.setText(f"Listening... '{s[:30]}...'")
        else: self.speech_label.setText("Listening...")

    def _stop_sr_listening_and_insert(self):
        try:
            if self.sr_worker:
                self.sr_worker.stop()
                self.sr_worker.wait(3000)
        except Exception: pass
        self.sr_listening = False
        self.speech_label.setText("Processing...")

    @QtCore.Slot(str)
    def _sr_finished(self, text: str):
        if text:
            try: devanagari = self.translit.translit_full(text)
            except Exception: devanagari = text
            cur = self.editor.textCursor()
            cur.insertText(devanagari)
            self.status.showMessage("Inserted speech", 3000)
        else:
            self.status.showMessage("No speech recognized", 3000)
        self.sr_worker = None
        self.speech_label.setText("Hold Ctrl+L to speak")

    def _mark_dirty(self):
        self.is_dirty = True

    def _perform_autosave(self):
        if self.autosave_enabled and self.current_filepath and self.is_dirty:
            try:
                ext = self.current_filepath.lower().split('.')[-1]
                if ext == "docx" and HAS_DOCX: 
                    self._save_docx_silent(self.current_filepath)
                elif ext in ["html", "htm"]:
                    self._save_html_silent(self.current_filepath)
                else:
                    with open(self.current_filepath, "w", encoding="utf-8") as f:
                        f.write(self.editor.toPlainText())
                        
                self.is_dirty = False
                self.status.showMessage(f"Autosaved to {os.path.basename(self.current_filepath)}", 2000)
            except Exception: pass

    def _handle_context_action(self, action: str):
        if action == "train_sel": self._train_manual()
        elif action == "edit_dict": self._open_corrections()
        elif action == "new_doc": self.act_new.trigger()
        elif action == "open_doc": self.act_open.trigger()
        elif action == "save_docx": self.act_save_docx.trigger()
        elif action == "export_pdf": self.act_export_pdf.trigger()
        elif action == "table_left": self.editor._align_table(Qt.AlignmentFlag.AlignLeft)
        elif action == "table_center": self.editor._align_table(Qt.AlignmentFlag.AlignHCenter)
        elif action == "table_right": self.editor._align_table(Qt.AlignmentFlag.AlignRight)

    def _setup_actions(self):
        self.act_new = QtGui.QAction("New", self, shortcut="Ctrl+N", triggered=lambda: self._new_file(prompt_autosave=True))
        self.act_open = QtGui.QAction("Open", self, shortcut="Ctrl+O", triggered=self._open_file)
        self.act_save = QtGui.QAction("Save", self, shortcut="Ctrl+S", triggered=self._save_manual)
        self.act_save_html = QtGui.QAction("Save as HTML", self, triggered=self._save_html)
        self.act_save_docx = QtGui.QAction("Save as DOCX", self, triggered=self._save_docx)
        self.act_export_pdf = QtGui.QAction("Export as PDF", self, triggered=self._export_pdf)
        self.act_page_setup = QtGui.QAction("Page Setup...", self, triggered=self._page_setup)
        self.act_print = QtGui.QAction("Print (Preview)", self, shortcut="Ctrl+P", triggered=self._print_doc)
        
        self.act_undo = QtGui.QAction("Undo", self, shortcut="Ctrl+Z", triggered=self.editor.undo)
        self.act_redo = QtGui.QAction("Redo", self, shortcut="Ctrl+Y", triggered=self.editor.redo)
        self.act_find = QtGui.QAction("Find & Replace", self, shortcut="Ctrl+F", triggered=self.finddlg.show)

        self.act_dict = QtGui.QAction("Correction Dictionary", self, shortcut="Ctrl+Shift+C", triggered=self._open_corrections)
        self.act_sugg_manager = QtGui.QAction("Suggestion Database Manager...", self, triggered=self._open_suggestion_trainer)
        self.act_translit = QtGui.QAction("Toggle Translit", self, shortcut="Ctrl+T", triggered=self._toggle_transliteration)
        self.act_theme = QtGui.QAction("Toggle Dark Mode", self, shortcut="Ctrl+D", triggered=self._toggle_theme)
        self.act_print_layout = QtGui.QAction("Print Layout View", self, checkable=True, triggered=self._toggle_print_layout)
        
        self.act_zoom_in = QtGui.QAction("Zoom In", self, shortcut="Ctrl++", triggered=lambda: self.editor.zoomIn(1))
        self.act_zoom_out = QtGui.QAction("Zoom Out", self, shortcut="Ctrl+-", triggered=lambda: self.editor.zoomOut(1))
        self.act_insert_date = QtGui.QAction("Insert Date/Time", self, shortcut="Ctrl+Shift+D", triggered=self._insert_datetime)
        self.act_insert_line = QtGui.QAction("Insert Horizontal Line", self, triggered=self._insert_horizontal_line)
        self.act_word_wrap = QtGui.QAction("Toggle Word Wrap", self, triggered=self._toggle_wrap)

    def _build_toolbars(self):
        tb_file = QtWidgets.QToolBar("File")
        self.addToolBar(Qt.ToolBarArea.TopToolBarArea, tb_file)

        tb_file.addAction(self.act_new)
        tb_file.addAction(self.act_open)
        tb_file.addAction(self.act_save)
        tb_file.addAction(self.act_print)
        tb_file.addSeparator()

        self.tpl_btn = QtWidgets.QToolButton()
        self.tpl_btn.setText("Templates")
        self.tpl_btn.setPopupMode(QtWidgets.QToolButton.ToolButtonPopupMode.InstantPopup)
        self.tpl_menu = QtWidgets.QMenu(self)
        self.tpl_btn.setMenu(self.tpl_menu)
        tb_file.addWidget(self.tpl_btn)
        self.tpl_menu.aboutToShow.connect(self._populate_template_menu)

        tb_file.addSeparator()
        tb_file.addAction(self.act_undo)
        tb_file.addAction(self.act_redo)
        tb_file.addAction(self.act_theme)
        tb_file.addAction(self.act_print_layout)

        self.addToolBarBreak() 
        tb_fmt = QtWidgets.QToolBar("Format")
        self.addToolBar(Qt.ToolBarArea.TopToolBarArea, tb_fmt)

        def add_fmt(text, slot, shortcut=None):
            a = QtGui.QAction(text, self)
            if shortcut: a.setShortcut(shortcut)
            a.triggered.connect(slot)
            tb_fmt.addAction(a)

        self.font_combo = QtWidgets.QFontComboBox()
        try: self.font_combo.setCurrentFont(QtGui.QFont(self.state.font_family))
        except Exception: pass
        self.font_combo.currentFontChanged.connect(self._font_changed)
        tb_fmt.addWidget(self.font_combo)

        self.font_size_spin = QtWidgets.QSpinBox()
        self.font_size_spin.setRange(8, 72); self.font_size_spin.setValue(self.state.font_size)
        self.font_size_spin.valueChanged.connect(self._font_size_changed)
        tb_fmt.addWidget(self.font_size_spin)

        tb_fmt.addSeparator()
        add_fmt("Color", self._choose_text_color)
        add_fmt("Highlight", self._choose_highlight)
        tb_fmt.addSeparator()

        add_fmt("B", lambda: self._toggle_format('bold'), "Ctrl+B")
        add_fmt("I", lambda: self._toggle_format('italic'), "Ctrl+I")
        add_fmt("U", lambda: self._toggle_format('underline'), "Ctrl+U")
        add_fmt("S", lambda: self._toggle_format('strike'))
        
        tb_fmt.addSeparator()
        add_fmt("X₂", lambda: self._toggle_format('subscript'), "Ctrl+=")
        add_fmt("X²", lambda: self._toggle_format('superscript'), "Ctrl+Shift++")
        
        tb_fmt.addSeparator()
        add_fmt("Left", lambda: self.editor.setAlignment(Qt.AlignmentFlag.AlignLeft))
        add_fmt("Center", lambda: self.editor.setAlignment(Qt.AlignmentFlag.AlignCenter))
        add_fmt("Right", lambda: self.editor.setAlignment(Qt.AlignmentFlag.AlignRight))
        add_fmt("Justify", lambda: self.editor.setAlignment(Qt.AlignmentFlag.AlignJustify))

        tb_fmt.addSeparator()
        add_fmt("Bullet", self._insert_bullet)
        add_fmt("Numbered", self._insert_numbered)
        tb_fmt.addSeparator()
        add_fmt("Table", self._insert_table)
        add_fmt("Image", self._insert_image)
        
        tb_fmt.addSeparator()
        self.speech_label = QtWidgets.QLabel("Hold Ctrl+L to speak")
        self.speech_label.setStyleSheet("padding-left: 10px; color: #888;")
        tb_fmt.addWidget(self.speech_label)

    def _populate_template_menu(self):
        self.tpl_menu.clear()
        for name, html in self.state.templates.items():
            self.tpl_menu.addAction(name, lambda h=html: self._insert_template_safe(h))
        self.tpl_menu.addSeparator()
        self.tpl_menu.addAction("Manage Advanced Templates...", self._open_template_manager)

    def _open_suggestion_trainer(self):
        dlg = SuggestionsManagerDialog(self.state, self)
        dlg.exec()

    def _open_template_manager(self):
        dlg = TemplateManagerDialog(self.state, self.editor, self)
        dlg.exec()

    def _build_menus(self):
        men = self.menuBar()
        filem = men.addMenu("&File")
        filem.addAction(self.act_new)
        filem.addAction(self.act_open)
        filem.addAction(self.act_save)
        filem.addAction(self.act_save_docx)
        filem.addAction(self.act_save_html)
        filem.addAction(self.act_export_pdf)
        filem.addSeparator()
        filem.addAction(self.act_page_setup)
        filem.addAction(self.act_print)
        
        editm = men.addMenu("&Edit")
        editm.addAction(self.act_undo)
        editm.addAction(self.act_redo)
        editm.addSeparator()
        editm.addAction(self.act_find)
        editm.addSeparator()
        editm.addAction(self.act_insert_date)
        editm.addAction(self.act_insert_line) 
        
        viewm = men.addMenu("&View")
        viewm.addAction(self.act_print_layout)
        viewm.addSeparator()
        viewm.addAction(self.act_zoom_in)
        viewm.addAction(self.act_zoom_out)
        viewm.addAction(self.act_word_wrap)

        toolsm = men.addMenu("&Tools")
        toolsm.addAction(self.act_dict)
        toolsm.addAction(self.act_sugg_manager)
        toolsm.addSeparator()
        toolsm.addAction(self.act_translit)
        toolsm.addAction(self.act_theme)
        toolsm.addSeparator()
        toolsm.addAction("Set Autosave Interval...", self._set_autosave_interval)

    def _build_sidebar(self):
        self.dock = QtWidgets.QDockWidget("Phrases", self)
        self.dock.setAllowedAreas(Qt.DockWidgetArea.LeftDockWidgetArea | Qt.DockWidgetArea.RightDockWidgetArea)
        w = QtWidgets.QWidget()
        v = QtWidgets.QVBoxLayout(w)
        v.setContentsMargins(0, 0, 0, 0)
        
        self.ph_list = QtWidgets.QListWidget()
        self.ph_list.addItems(self.state.phrases)
        self.ph_list.itemDoubleClicked.connect(lambda i: self.editor.textCursor().insertText(i.text()))
        
        self.ph_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.ph_list.customContextMenuRequested.connect(self._show_phrase_context_menu)
        
        v.addWidget(self.ph_list)
        w.setLayout(v)
        self.dock.setWidget(w)
        self.addDockWidget(Qt.DockWidgetArea.RightDockWidgetArea, self.dock)

    def _show_phrase_context_menu(self, pos):
        menu = QtWidgets.QMenu(self.ph_list)
        act_add = menu.addAction("Add new phrase...")
        act_rem = menu.addAction("Remove selected phrase")
        
        if self.ph_list.currentRow() < 0:
            act_rem.setEnabled(False)
            
        action = menu.exec(self.ph_list.mapToGlobal(pos))
        if action == act_add:
            self._add_phrase()
        elif action == act_rem:
            self._remove_phrase()

    def _build_suggestions_sidebar(self):
        self.sugg_dock = QtWidgets.QDockWidget("Suggestions (Ctrl+↑/↓, Tab to insert)", self)
        self.sugg_dock.setAllowedAreas(Qt.DockWidgetArea.LeftDockWidgetArea | Qt.DockWidgetArea.RightDockWidgetArea)
        w = QtWidgets.QWidget()
        v = QtWidgets.QVBoxLayout(w)
        v.setContentsMargins(0, 0, 0, 0)
        
        self.sugg_list = QtWidgets.QListWidget()
        self.sugg_list.setFocusPolicy(Qt.FocusPolicy.NoFocus) 
        self.sugg_list.itemClicked.connect(lambda i: self.editor.insert_suggestion(i.data(Qt.ItemDataRole.UserRole)))
        v.addWidget(self.sugg_list)
        
        w.setLayout(v)
        self.sugg_dock.setWidget(w)
        self.addDockWidget(Qt.DockWidgetArea.LeftDockWidgetArea, self.sugg_dock)

    @QtCore.Slot(str)
    def _update_suggestions(self, prefix: str):
        self.sugg_list.clear()
        if not prefix or len(prefix) < 1:
            self.editor.set_has_suggestions(False)
            return
            
        prefix_lower = prefix.lower()
        matches = []
        
        for word in self.state.suggestion_words:
            if word.lower().startswith(prefix_lower):
                matches.append(word)
                    
        matches.sort(key=len)

        for word in matches[:100]:
            item = QtWidgets.QListWidgetItem(word)
            item.setData(Qt.ItemDataRole.UserRole, word)
            self.sugg_list.addItem(item)
            
        if self.sugg_list.count() > 0:
            self.sugg_list.setCurrentRow(0)
            self.editor.set_has_suggestions(True)
        else:
            self.editor.set_has_suggestions(False)

    @QtCore.Slot(int)
    def _navigate_suggestion_list(self, step: int):
        c = self.sugg_list.count()
        if c == 0: return
        curr = self.sugg_list.currentRow()
        new_row = (curr + step) % c
        self.sugg_list.setCurrentRow(new_row)

    @QtCore.Slot()
    def _insert_selected_suggestion(self):
        row = self.sugg_list.currentRow()
        if row >= 0:
            word = self.sugg_list.item(row).data(Qt.ItemDataRole.UserRole)
            self.editor.insert_suggestion(word)

    def _build_statusbar(self):
        self.status = self.statusBar()
        
        self.lbl_cursor_pos = QtWidgets.QLabel("Ln 1, Col 0")
        self.lbl_counts = QtWidgets.QLabel("0 words • 0 chars")
        self.lbl_translit_badge = QtWidgets.QLabel()
        
        self.lbl_cursor_pos.setStyleSheet("padding-right: 15px;")
        self.lbl_counts.setStyleSheet("padding-right: 15px;")
        
        self.status.addPermanentWidget(self.lbl_cursor_pos)
        self.status.addPermanentWidget(self.lbl_counts)
        self.status.addPermanentWidget(self.lbl_translit_badge)

    def _update_translit_badge(self):
        if self.state.transliteration_enabled:
            self.lbl_translit_badge.setText("TRANSLIT: ON")
            self.lbl_translit_badge.setStyleSheet("color: white; background-color: #4CAF50; padding: 2px 5px; border-radius: 3px; font-weight: bold;")
        else:
            self.lbl_translit_badge.setText("TRANSLIT: OFF")
            self.lbl_translit_badge.setStyleSheet("color: white; background-color: #F44336; padding: 2px 5px; border-radius: 3px; font-weight: bold;")

    def _page_setup(self):
        dialog = QPageSetupDialog(self.printer, self)
        if dialog.exec() == QtWidgets.QDialog.DialogCode.Accepted:
            self._update_view_mode()

    def _toggle_print_layout(self):
        self.print_layout_enabled = self.act_print_layout.isChecked()
        self._update_view_mode()

    def _update_view_mode(self):
        fmt = self.editor.document().rootFrame().frameFormat()
        
        if self.print_layout_enabled:
            layout = self.printer.pageLayout()
            try:
                full_rect = layout.fullRectPixels(96)
                margins = layout.marginsPixels(96)
                self.editor.document().setPageSize(QtCore.QSizeF(full_rect.width(), full_rect.height()))
                fmt.setTopMargin(margins.top())
                fmt.setBottomMargin(margins.bottom())
                fmt.setLeftMargin(margins.left())
                fmt.setRightMargin(margins.right())
            except Exception:
                self.editor.document().setPageSize(QtCore.QSizeF(794, 1123))
                fmt.setTopMargin(96); fmt.setBottomMargin(96); fmt.setLeftMargin(96); fmt.setRightMargin(96)

            fmt.setBackground(QtGui.QColor("#ffffff"))
            
            if self.dark_mode_enabled:
                self.editor.setStyleSheet("QTextEdit { background-color: #2b2b2b; color: #000000; border: none; }")
            else:
                self.editor.setStyleSheet("QTextEdit { background-color: #e0e0e0; color: #000000; border: none; }")
        else:
            self.editor.document().setPageSize(QtCore.QSizeF(-1, -1))
            fmt.setTopMargin(8)
            fmt.setBottomMargin(8)
            fmt.setLeftMargin(8)
            fmt.setRightMargin(8)
            fmt.clearBackground()
            
            if self.dark_mode_enabled:
                self.editor.setStyleSheet("QTextEdit { background-color: #1e1e1e; color: #ffffff; border: 1px solid #555; }")
            else:
                self.editor.setStyleSheet("")
                
        self.editor.document().rootFrame().setFormat(fmt)

    def _toggle_theme(self):
        self.dark_mode_enabled = not self.dark_mode_enabled
        if self.dark_mode_enabled:
            self.setStyleSheet("""
                QMainWindow, QDialog, QDockWidget, QMenuBar, QMenu { background-color: #2b2b2b; color: #ffffff; }
                QToolBar { background-color: #333333; border: none; }
                QPushButton, QToolButton { background-color: #444; color: white; border-radius: 3px; padding: 4px; }
                QPushButton:hover, QToolButton:hover { background-color: #555; }
                QTableWidget, QListWidget, QLineEdit { background-color: #1e1e1e; color: #ffffff; border: 1px solid #555; }
            """)
        else:
            self.setStyleSheet("")
        self._update_view_mode()

    def _insert_template_safe(self, html: str):
        self.editor._commit_composing()
        self.editor.insertHtml(html)

    def _toggle_format(self, fmt_type):
        cur = self.editor.textCursor()
        fmt = cur.charFormat()
        
        if fmt_type == 'bold': fmt.setFontWeight(QtGui.QFont.Weight.Bold if fmt.fontWeight() != QtGui.QFont.Weight.Bold else QtGui.QFont.Weight.Normal)
        elif fmt_type == 'italic': fmt.setFontItalic(not fmt.fontItalic())
        elif fmt_type == 'underline': fmt.setFontUnderline(not fmt.fontUnderline())
        elif fmt_type == 'strike': fmt.setFontStrikeOut(not fmt.fontStrikeOut())
        elif fmt_type == 'subscript':
            if fmt.verticalAlignment() == QTextCharFormat.VerticalAlignment.AlignSubScript:
                fmt.setVerticalAlignment(QTextCharFormat.VerticalAlignment.AlignNormal)
            else:
                fmt.setVerticalAlignment(QTextCharFormat.VerticalAlignment.AlignSubScript)
        elif fmt_type == 'superscript':
            if fmt.verticalAlignment() == QTextCharFormat.VerticalAlignment.AlignSuperScript:
                fmt.setVerticalAlignment(QTextCharFormat.VerticalAlignment.AlignNormal)
            else:
                fmt.setVerticalAlignment(QTextCharFormat.VerticalAlignment.AlignSuperScript)
            
        if cur.hasSelection(): cur.mergeCharFormat(fmt)
        else: self.editor.mergeCurrentCharFormat(fmt)
        self.editor.setFocus()

    def _font_changed(self, qfont: QtGui.QFont):
        self.state.font_family = qfont.family()
        fmt = QtGui.QTextCharFormat()
        fmt.setFont(QtGui.QFont(self.state.font_family, self.state.font_size))
        cur = self.editor.textCursor()
        if cur.hasSelection(): cur.mergeCharFormat(fmt)
        else: self.editor.mergeCurrentCharFormat(fmt)
        self.editor.setFocus()

    def _font_size_changed(self, size: int):
        self.state.font_size = size
        fmt = QtGui.QTextCharFormat()
        fmt.setFontPointSize(size)
        cur = self.editor.textCursor()
        if cur.hasSelection(): cur.mergeCharFormat(fmt)
        else: self.editor.mergeCurrentCharFormat(fmt)
        self.editor.setFocus()

    def _choose_text_color(self):
        col = QtWidgets.QColorDialog.getColor(parent=self)
        if col.isValid():
            fmt = QtGui.QTextCharFormat()
            fmt.setForeground(col)
            cur = self.editor.textCursor()
            if cur.hasSelection(): cur.mergeCharFormat(fmt)
            else: self.editor.mergeCurrentCharFormat(fmt)
            self.editor.setFocus()

    def _choose_highlight(self):
        col = QtWidgets.QColorDialog.getColor(parent=self)
        if col.isValid():
            fmt = QtGui.QTextCharFormat()
            fmt.setBackground(col)
            cur = self.editor.textCursor()
            if cur.hasSelection(): cur.mergeCharFormat(fmt)
            else: self.editor.mergeCurrentCharFormat(fmt)
            self.editor.setFocus()

    def _insert_bullet(self):
        self.editor.textCursor().insertList(QtGui.QTextListFormat.Style.ListDisc)

    def _insert_numbered(self):
        fmt = QtGui.QTextListFormat()
        fmt.setStyle(QtGui.QTextListFormat.Style.ListDecimal)
        self.editor.textCursor().createList(fmt)

    def _insert_table(self):
        self.editor._commit_composing()
        rows, ok1 = QtWidgets.QInputDialog.getInt(self, "Insert Table", "Rows:", 3, 1, 100)
        if ok1:
            cols, ok2 = QtWidgets.QInputDialog.getInt(self, "Insert Table", "Columns:", 3, 1, 100)
            if ok2:
                fmt = QtGui.QTextTableFormat()
                fmt.setBorder(1); fmt.setCellPadding(4); fmt.setCellSpacing(0)
                self.editor.textCursor().insertTable(rows, cols, fmt)

    def _insert_image(self):
        self.editor._commit_composing()
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Insert Image", "", "Images (*.png *.jpg *.bmp *.gif *.jpeg)")
        if not path: return
        
        width, ok = QtWidgets.QInputDialog.getInt(self, "Image Size", "Width (pixels):", 400, 50, 1000)
        if not ok: return
        
        align_str, ok2 = QtWidgets.QInputDialog.getItem(self, "Alignment", "Choose Alignment:", ["Left", "Center", "Right"], 1, False)
        if not ok2: return

        cur = self.editor.textCursor()
        try:
            img = QtGui.QImage(path)
            if img.isNull(): raise ValueError("Cannot load image")
            img = img.scaledToWidth(width, Qt.TransformationMode.SmoothTransformation)
            
            fmt = QtGui.QTextImageFormat()
            url = QtCore.QUrl.fromLocalFile(os.path.abspath(path)).toString()
            fmt.setName(url)
            fmt.setWidth(width)
            fmt.setHeight(img.height())
            
            bf = cur.blockFormat()
            if align_str == "Center": bf.setAlignment(Qt.AlignmentFlag.AlignCenter)
            elif align_str == "Right": bf.setAlignment(Qt.AlignmentFlag.AlignRight)
            else: bf.setAlignment(Qt.AlignmentFlag.AlignLeft)
            cur.setBlockFormat(bf)
            
            cur.insertImage(fmt)
            self.status.showMessage(f"Inserted image ({width}px, {align_str})", 2000)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Insert image failed", str(e))

    def _set_autosave_interval(self):
        current_sec = self.autosave_timer.interval() // 1000
        if current_sec == 0: current_sec = 30
        sec, ok = QtWidgets.QInputDialog.getInt(self, "Autosave Interval", "Interval in seconds:", current_sec, 5, 3600)
        if ok:
            self.autosave_timer.setInterval(sec * 1000)
            if self.autosave_enabled:
                self.status.showMessage(f"Autosave interval updated to {sec} seconds", 3000)
                
    def _insert_datetime(self):
        dt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.editor.textCursor().insertText(dt)
        
    def _insert_horizontal_line(self):
        self.editor._commit_composing()
        self.editor.insertHtml(U_LINE)
        self.editor.setFocus()
        
    def _toggle_wrap(self):
        if self.editor.lineWrapMode() == QtWidgets.QTextEdit.LineWrapMode.NoWrap:
            self.editor.setLineWrapMode(QtWidgets.QTextEdit.LineWrapMode.WidgetWidth)
            self.status.showMessage("Word Wrap: ON", 2000)
        else:
            self.editor.setLineWrapMode(QtWidgets.QTextEdit.LineWrapMode.NoWrap)
            self.status.showMessage("Word Wrap: OFF", 2000)

    def _new_file(self, prompt_autosave=True):
        if self.is_dirty or self.editor.toPlainText().strip():
            reply = QtWidgets.QMessageBox.question(
                self, "Save Current Document?",
                "Do you want to save the current document before starting a new one?",
                QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No | QtWidgets.QMessageBox.StandardButton.Cancel
            )
            if reply == QtWidgets.QMessageBox.StandardButton.Yes:
                self._save_manual()
                if self.is_dirty: return 
            elif reply == QtWidgets.QMessageBox.StandardButton.Cancel:
                return

        self.editor._commit_composing()
        self.editor.clear()
        self.current_filepath = None
        self.autosave_enabled = False
        self.autosave_timer.stop()
        self.is_dirty = False
        
        if prompt_autosave:
            msg = QtWidgets.QMessageBox(self)
            msg.setWindowTitle("Enable Autosave")
            msg.setText("Would you like to save this new document now to enable continuous auto-saving?")
            btn_save = msg.addButton("Save Now", QtWidgets.QMessageBox.ButtonRole.AcceptRole)
            btn_later = msg.addButton("Later", QtWidgets.QMessageBox.ButtonRole.RejectRole)
            msg.exec()
            
            if msg.clickedButton() == btn_save:
                path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Initialize Document", "", "Word Document (*.docx);;HTML Web Page (*.html);;Text File (*.txt)")
                if path:
                    self.current_filepath = path
                    self.autosave_enabled = True
                    interval = self.autosave_timer.interval()
                    if interval == 0: interval = 30000
                    self.autosave_timer.start(interval)
                    self.status.showMessage(f"Auto-save enabled (Every {interval//1000}s).")

    def _open_file(self):
        self.editor._commit_composing()
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Open", "", "Documents (*.docx *.html *.htm *.txt *.csv *.xlsx *.ods);;All Files (*.*)")
        if not path: return
        ext = path.lower().split('.')[-1]
        try:
            if ext == "docx" and HAS_DOCX:
                doc = docx.Document(path)
                html_fragments = []
                for p in doc.paragraphs:
                    p_html = "<p"
                    if p.alignment and WD_ALIGN_PARAGRAPH:
                        if p.alignment == WD_ALIGN_PARAGRAPH.CENTER: p_html += ' align="center"'
                        elif p.alignment == WD_ALIGN_PARAGRAPH.RIGHT: p_html += ' align="right"'
                        elif p.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY: p_html += ' align="justify"'
                    p_html += ">"
                    for run in p.runs:
                        t = run.text.replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br>")
                        if run.bold: t = f"<b>{t}</b>"
                        if run.italic: t = f"<i>{t}</i>"
                        if run.underline: t = f"<u>{t}</u>"
                        if run.font.strike: t = f"<s>{t}</s>"
                        p_html += t
                    p_html += "</p>"
                    html_fragments.append(p_html)
                    
                for table in doc.tables:
                    html_fragments.append("<table border='1' cellspacing='0' cellpadding='4'>")
                    for row in table.rows:
                        html_fragments.append("<tr>")
                        for cell in row.cells: html_fragments.append(f"<td>{cell.text}</td>")
                        html_fragments.append("</tr>")
                    html_fragments.append("</table><br>")
                self.editor.setHtml("".join(html_fragments))
            elif ext in ["html", "htm"]:
                with open(path, "r", encoding="utf-8") as f: self.editor.setHtml(f.read())
            elif ext in ["csv", "xlsx", "ods"]:
                if ext == "xlsx":
                    df = pd.read_excel(path, engine="openpyxl")
                elif ext == "ods":
                    df = pd.read_excel(path, engine="odf")
                else:
                    df = pd.read_csv(path)
                self.editor.insertHtml(df.to_html(index=False, border=1) + "<br>")
            else:
                with open(path, "r", encoding="utf-8") as f: self.editor.setPlainText(f.read())
            
            self.current_filepath = path
            self.status.showMessage(f"Opened: {path}", 3000)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Open failed", str(e))

    def _save_manual(self):
        if self.current_filepath:
            ext = self.current_filepath.lower().split('.')[-1]
            try:
                if ext == "docx" and HAS_DOCX: 
                    self._save_docx_silent(self.current_filepath)
                elif ext in ["html", "htm"]:
                    self._save_html_silent(self.current_filepath)
                else:
                    with open(self.current_filepath, "w", encoding="utf-8") as f: 
                        f.write(self.editor.toPlainText())
                        
                self.is_dirty = False
                self.status.showMessage("Saved successfully.")
            except PermissionError:
                QtWidgets.QMessageBox.warning(self, "File Locked", f"Cannot save to {self.current_filepath}. Close the document in Word first.")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Save Error", str(e))
        else:
            self._save_docx()

    def _save_docx(self):
        if not HAS_DOCX:
            QtWidgets.QMessageBox.warning(self, "Missing", "python-docx required")
            return
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save Document", "", "Word Document (*.docx)")
        if path:
            try:
                export_native_to_docx(self.editor.document(), path)
                self.current_filepath = path
                self.is_dirty = False
                
                self.autosave_enabled = True
                if not self.autosave_timer.isActive():
                    self.autosave_timer.start(self.autosave_timer.interval() or 30000)

                self.status.showMessage(f"Saved: {path}")
            except PermissionError as e:
                QtWidgets.QMessageBox.warning(self, "File Locked", str(e))
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Save Error", str(e))

    def _save_html(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Save as HTML", "", "HTML Web Page (*.html *.htm)")
        if path:
            try:
                self._save_html_silent(path)
                self.status.showMessage(f"Exported HTML to: {path}")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Save Error", str(e))

    def _save_html_silent(self, path):
        html_content = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
    body {{
        font-family: '{self.state.font_family}', sans-serif;
        font-size: {self.state.font_size}pt;
        line-height: 1.5;
        margin: auto;
        background-color: #fcfcfc;
    }}
    .document {{

    }}
    table {{ width: 100%; border-collapse: collapse; }}
    td {{ padding: 5px; }}
</style>
</head>
<body>
<div class="document">
{self.editor.toHtml()}
</div>
</body>
</html>"""
        with open(path, "w", encoding="utf-8") as f:
            f.write(html_content)

    def _save_docx_silent(self, path):
        export_native_to_docx(self.editor.document(), path)

    def _export_pdf(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "Export PDF", "", "PDF Document (*.pdf)")
        if path:
            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setPageLayout(self.printer.pageLayout())
            printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
            printer.setOutputFileName(path)
            self.editor.document().print_(printer)
            self.status.showMessage(f"Exported to PDF: {path}")

    def _print_doc(self):
        title = APP_NAME + " Document"
        if self.current_filepath:
            title = os.path.basename(self.current_filepath)
        self.printer.setDocName(title)
        
        preview = QPrintPreviewDialog(self.printer, self)
        preview.setWindowTitle("Print Preview - " + APP_NAME)
        preview.paintRequested.connect(self.editor.print_)
        preview.exec()

    def _open_corrections(self):
        dlg = CorrectionsDialog(self.state, editor_ref=self.editor, parent=self)
        dlg.exec()

    def _train_manual(self):
        cur = self.editor.textCursor()
        if not cur.hasSelection():
            QtWidgets.QMessageBox.information(self, "Info", "Select Devanagari text in the editor first.")
            return
        sel = cur.selectedText()
        latin, ok = QtWidgets.QInputDialog.getText(self, "Train mapping", f"Map Latin shortcut for:\n{sel}")
        if ok and latin.strip():
            self.state.user_dict[latin.strip()] = sel
            self.status.showMessage(f"Mapped {latin.strip()} -> {sel}")

    def _toggle_transliteration(self):
        self.editor._commit_composing()
        self.state.transliteration_enabled = not self.state.transliteration_enabled
        self._update_translit_badge()
        st = "ON" if self.state.transliteration_enabled else "OFF"
        self.status.showMessage(f"Transliteration {st}", 4000)

    def _update_counts(self, w: int, c: int):
        self.lbl_counts.setText(f"{w} words • {c} chars")

    def _update_cursor_position_label(self, line: int, col: int):
        self.lbl_cursor_pos.setText(f"Ln {line}, Col {col}")

    def _add_phrase(self):
        txt, ok = QtWidgets.QInputDialog.getMultiLineText(self, "Add phrase", "Phrase (Devanagari or Latin):")
        if ok and txt.strip():
            self.state.phrases.append(txt)
            self.ph_list.addItem(txt)
            self.state.save_phrases()

    def _remove_phrase(self):
        r = self.ph_list.currentRow()
        if r >= 0:
            self.state.phrases.pop(r)
            self.ph_list.takeItem(r)
            self.state.save_phrases()


def main():
    app = QtWidgets.QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    
    # Modern font fallback
    available_families = QtGui.QFontDatabase.families()
    target_font = "Arial"
    
    for f in ["Nirmala UI", "Mangal", "Noto Sans Devanagari", "Arial Unicode MS"]:
        if f in available_families:
            target_font = f
            break
            
    if target_font:
        app.setFont(QtGui.QFont(target_font, 10))
    
    if not HAS_TRANSLIT:
        QtWidgets.QMessageBox.critical(None, "Missing dependency", "Please install indic-transliteration")
        return
        
    win = MainWindow()
    win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
