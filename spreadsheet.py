import string, re, json, urllib.request, urllib.parse
from PySide6 import QtWidgets, QtCore, QtGui
from PySide6.QtCore import Qt

class SmartTableWidget(QtWidgets.QTableWidget):
    """A custom table handling Excel-like Copy, Paste, Cut, Delete, and Context Menus."""
    def __init__(self, rows, cols, parent=None):
        super().__init__(rows, cols, parent)
        
    def keyPressEvent(self, event):
        key = event.key()
        mods = event.modifiers()
        
        if mods & Qt.KeyboardModifier.ControlModifier:
            if key == Qt.Key.Key_C:
                self.copy_selection()
                return
            elif key == Qt.Key.Key_V:
                self.paste_selection()
                return
            elif key == Qt.Key.Key_X:
                self.cut_selection()
                return
                
        # --- NEW: Delete / Backspace clears selected cells ---
        if key in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            self.clear_selection()
            return
                
        if key in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            current = self.currentIndex()
            if current.isValid():
                next_row = current.row() + 1
                if next_row < self.rowCount():
                    self.setCurrentCell(next_row, current.column())
                return
                
        super().keyPressEvent(event)

    def contextMenuEvent(self, event):
        menu = QtWidgets.QMenu(self)
        
        # --- NEW: Merge Cells Option ---
        act_merge = menu.addAction("Merge / Unmerge Selected Cells")
        act_merge.triggered.connect(self.toggle_merge)
        menu.addSeparator()
        
        act_serial = menu.addAction("1️⃣ Insert Serial Numbers (1, 2, 3...)")
        act_serial.triggered.connect(self.parent().insert_serial_numbers)
        menu.addSeparator()
        
        act_hi = menu.addAction("Translate Selection to Hindi")
        act_hi.triggered.connect(lambda: self.parent()._translate_cells("hi"))
        
        act_en = menu.addAction("Translate Selection to English")
        act_en.triggered.connect(lambda: self.parent()._translate_cells("en"))
        
        menu.exec(event.globalPos())

    def toggle_merge(self):
        ranges = self.selectedRanges()
        if not ranges: return
        rng = ranges[0]
        # If it is already spanning multiple rows/cols, unmerge it. Otherwise, merge it.
        if self.rowSpan(rng.topRow(), rng.leftColumn()) > 1 or self.columnSpan(rng.topRow(), rng.leftColumn()) > 1:
            self.setSpan(rng.topRow(), rng.leftColumn(), 1, 1) 
        else:
            self.setSpan(rng.topRow(), rng.leftColumn(), rng.rowCount(), rng.columnCount())

    def clear_selection(self):
        selection = self.selectedIndexes()
        self.blockSignals(True)
        for idx in selection:
            item = self.item(idx.row(), idx.column())
            if item: item.setText("")
        self.blockSignals(False)
        main_win = QtWidgets.QApplication.activeWindow()
        if hasattr(main_win, '_mark_dirty'): main_win._mark_dirty()

    def copy_selection(self):
        selection = self.selectedIndexes()
        if not selection: return
        rows = sorted(list(set(idx.row() for idx in selection)))
        cols = sorted(list(set(idx.column() for idx in selection)))
        copy_text = ""
        for r in rows:
            row_data = []
            for c in cols:
                item = self.item(r, c)
                row_data.append(item.text() if item else "")
            copy_text += "\t".join(row_data) + "\n"
        QtWidgets.QApplication.clipboard().setText(copy_text)

    def cut_selection(self):
        self.copy_selection()
        self.clear_selection()

    def paste_selection(self):
        text = QtWidgets.QApplication.clipboard().text()
        if not text: return
        selection = self.selectedIndexes()
        if not selection: return
        start_row, start_col = selection[0].row(), selection[0].column()
        lines = text.strip('\n').split('\n')
        self.blockSignals(True)
        for i, line in enumerate(lines):
            cells = line.split('\t')
            for j, cell_text in enumerate(cells):
                target_row = start_row + i
                target_col = start_col + j
                if target_row < self.rowCount() and target_col < self.columnCount():
                    item = self.item(target_row, target_col)
                    if not item:
                        item = QtWidgets.QTableWidgetItem()
                        self.setItem(target_row, target_col, item)
                    item.setText(cell_text)
        self.blockSignals(False)
        main_win = QtWidgets.QApplication.activeWindow()
        if hasattr(main_win, '_mark_dirty'): main_win._mark_dirty()


class HindiCellEditor(QtWidgets.QLineEdit):
    def __init__(self, translit, state, parent=None):
        super().__init__(parent)
        self.translit = translit
        self.state = state
        self._english_mode = False
        self._composing_latin = ""
        self._composing_start = 0
        self._last_word = ""

    def keyPressEvent(self, ev: QtGui.QKeyEvent):
        key = ev.key()
        mods = ev.modifiers()
        main_win = self.window()

        # --- NEW: TAB COMPLETION FOR SUGGESTIONS ---
        if key == Qt.Key.Key_Tab and hasattr(main_win, 'sugg_list') and main_win.sugg_list.count() > 0:
            main_win._insert_selected_dock_suggestion()
            return

        if key in (Qt.Key.Key_Tab, Qt.Key.Key_Return, Qt.Key.Key_Enter, Qt.Key.Key_Up, Qt.Key.Key_Down):
            self._update_last_word()
            self._composing_latin = ""
            super().keyPressEvent(ev)
            if hasattr(main_win, '_populate_dock_suggestions'): main_win._populate_dock_suggestions([])
            return

        if key == Qt.Key.Key_Space:
            if mods & Qt.KeyboardModifier.ControlModifier:
                self._english_mode = not self._english_mode
                
                # --- FIX: Update the visual badge in the Main Window instantly! ---
                main_win = self.window()
                if hasattr(main_win, 'lbl_translit_badge'):
                    if self._english_mode:
                        main_win.lbl_translit_badge.setText("CELL EN / HINGLISH")
                        main_win.lbl_translit_badge.setStyleSheet("color: white; background-color: rgba(30, 58, 138, 0.8); padding: 2px 6px; border-radius: 4px; font-weight: bold;")
                    else:
                        main_win.lbl_translit_badge.setText("CELL TRANSLIT: ON")
                        main_win.lbl_translit_badge.setStyleSheet("color: white; background-color: rgba(4, 120, 87, 0.8); padding: 2px 6px; border-radius: 4px; font-weight: bold;")
              
                return
                
            self._update_last_word()
            self._composing_latin = ""
            super().keyPressEvent(ev)
            self._notify_main_window_for_suggestions()
            return

        if self._english_mode or not self.state.transliteration_enabled:
            super().keyPressEvent(ev)
            self._notify_main_window_for_suggestions()
            return

        txt = ev.text()
        if len(txt) == 1 and txt.isprintable() and not txt.isspace() and ord(txt) < 128:
            if not self._composing_latin:
                self._composing_start = self.cursorPosition()
                self._composing_latin = txt
            else:
                self.setSelection(self._composing_start, self.cursorPosition() - self._composing_start)
                self.del_()
                self._composing_latin += txt

            translit_text = self.translit.translit_token(self._composing_latin)
            self.insert(translit_text)
            self._notify_main_window_for_suggestions()
            return

        if key in (Qt.Key.Key_Backspace, Qt.Key.Key_Delete):
            if self._composing_latin:
                self._composing_latin = self._composing_latin[:-1]
                self.setSelection(self._composing_start, self.cursorPosition() - self._composing_start)
                self.del_()
                if self._composing_latin:
                    self.insert(self.translit.translit_token(self._composing_latin))
                self._notify_main_window_for_suggestions()
                return
            super().keyPressEvent(ev)
            self._notify_main_window_for_suggestions()
            return

        super().keyPressEvent(ev)

    def _update_last_word(self):
        text = self.text()
        if text.strip(): self._last_word = text.split()[-1]

    def _notify_main_window_for_suggestions(self):
        main_win = self.window()
        if not main_win or not hasattr(main_win, '_populate_dock_suggestions'): return

        text = self.text()
        pos = self.cursorPosition()
        start = pos - 1
        while start >= 0 and not text[start].isspace(): start -= 1
        start += 1
        current_prefix = text[start:pos].lower()
        
        # SMART PREDICTIONS based on last word
        predicted_words = set()
        if self._last_word:
            last_lower = self._last_word.lower()
            if last_lower in self.state.next_words:
                predicted_words = set(self.state.next_words[last_lower].keys())

        if current_prefix:
            all_words = list(self.state.suggestion_words)
            matches = [w for w in all_words if w.lower().startswith(current_prefix)]
            matches.sort(key=lambda w: (
                w.lower() not in predicted_words, 
                w.lower() != current_prefix,        
                len(w),                             
                w.lower()                           
            ))
            main_win._populate_dock_suggestions(matches[:50])
        elif predicted_words:
            sorted_preds = sorted(predicted_words, key=lambda w: self.state.next_words.get(self._last_word.lower(), {}).get(w, 0), reverse=True)
            main_win._populate_dock_suggestions(sorted_preds[:50])
        else:
            main_win._populate_dock_suggestions([])


class HindiItemDelegate(QtWidgets.QStyledItemDelegate):
    def __init__(self, translit, state, parent=None):
        super().__init__(parent)
        self.translit = translit
        self.state = state

    def createEditor(self, parent, option, index):
        return HindiCellEditor(self.translit, self.state, parent)


class SpreadsheetWidget(QtWidgets.QWidget):
    def __init__(self, translit, state, parent=None):
        super().__init__(parent)
        self.translit = translit
        self.state = state
        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)

        formula_layout = QtWidgets.QHBoxLayout()
        formula_layout.addWidget(QtWidgets.QLabel("  fx  "))
        self.formula_bar = QtWidgets.QLineEdit()
        self.formula_bar.returnPressed.connect(self._commit_formula)
        formula_layout.addWidget(self.formula_bar)
        self.layout.addLayout(formula_layout)

        self.table = SmartTableWidget(100, 26, self) 
        self.table.setAlternatingRowColors(True)
        self.table.setItemDelegate(HindiItemDelegate(self.translit, self.state, self.table))
        
        self.col_headers = list(string.ascii_uppercase)
        self.table.setHorizontalHeaderLabels(self.col_headers)
        self.table.setVerticalHeaderLabels([str(i+1) for i in range(100)])

        self.layout.addWidget(self.table)
        self.formulas = {} 

        self.table.itemSelectionChanged.connect(self._update_formula_bar)
        self.table.itemChanged.connect(self._on_cell_changed)
        self._is_updating = False

    def insert_serial_numbers(self):
        ranges = self.table.selectedRanges()
        if not ranges: return
        counter = 1
        self._is_updating = True 
        for rng in ranges:
            for row in range(rng.topRow(), rng.bottomRow() + 1):
                for col in range(rng.leftColumn(), rng.rightColumn() + 1):
                    item = self.table.item(row, col)
                    if not item:
                        item = QtWidgets.QTableWidgetItem()
                        self.table.setItem(row, col, item)
                    item.setText(str(counter))
                    counter += 1
        self._is_updating = False
        if self.window(): self.window()._mark_dirty()

    def _translate_cells(self, target_lang):
        items = self.table.selectedItems()
        if not items: return
        if self.window(): self.window().status.showMessage("Translating cells...")
        QtWidgets.QApplication.processEvents()
        
        self._is_updating = True
        try:
            for item in items:
                text = item.text().strip()
                if not text: continue
                url = f"https://translate.googleapis.com/translate_a/single?client=gtx&sl=auto&tl={target_lang}&dt=t&q=" + urllib.parse.quote(text)
                req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                response = urllib.request.urlopen(req, timeout=3)
                data = json.loads(response.read().decode('utf-8'))
                translated_text = "".join([sentence[0] for sentence in data[0]])
                item.setText(translated_text)
                
            if self.window(): 
                self.window().status.showMessage("Cells translated successfully.", 3000)
                self.window()._mark_dirty()
        except Exception as e:
            if self.window(): self.window().status.showMessage("Translation failed. Check internet.", 3000)
        self._is_updating = False

    def _update_formula_bar(self):
        items = self.table.selectedItems()
        if items:
            item = items[0]
            cell_id = f"{self.col_headers[item.column()]}{item.row()+1}"
            self.formula_bar.setText(self.formulas.get(cell_id, item.text()))

    def _commit_formula(self):
        items = self.table.selectedItems()
        if items:
            item = items[0]
            item.setText(self.formula_bar.text())

    def _on_cell_changed(self, item):
        if self._is_updating: return
        self._is_updating = True
        
        main_win = QtWidgets.QApplication.activeWindow()
        if main_win and hasattr(main_win, '_mark_dirty'):
            main_win._mark_dirty()

        cell_id = f"{self.col_headers[item.column()]}{item.row()+1}"
        raw_text = item.text()

        if raw_text.startswith("="):
            self.formulas[cell_id] = raw_text
            item.setText(str(self._evaluate_formula(raw_text)))
        else:
            if cell_id in self.formulas:
                del self.formulas[cell_id]
        self._is_updating = False

    def _get_cell_value(self, cell_id):
        try:
            col = self.col_headers.index(cell_id[0].upper())
            row = int(cell_id[1:]) - 1
            item = self.table.item(row, col)
            if item:
                val = item.text()
                try: return float(val)
                except ValueError: return val
        except Exception: return 0
        return 0

    def _get_range_values(self, range_str):
        try:
            start, end = range_str.split(":")
            start_col, start_row = self.col_headers.index(start[0].upper()), int(start[1:]) - 1
            end_col, end_row = self.col_headers.index(end[0].upper()), int(end[1:]) - 1
            vals = []
            for r in range(min(start_row, end_row), max(start_row, end_row) + 1):
                for c in range(min(start_col, end_col), max(start_col, end_col) + 1):
                    item = self.table.item(r, c)
                    if item:
                        txt = item.text()
                        try: vals.append(float(txt))
                        except ValueError: vals.append(txt)
            return vals
        except Exception: return []

    def _evaluate_formula(self, formula):
        try:
            formula = formula[1:].upper().strip() 
            if formula.startswith("SUM(") and formula.endswith(")"):
                vals = [v for v in self._get_range_values(formula[4:-1]) if isinstance(v, float)]
                return sum(vals)
            elif formula.startswith("COUNT(") and formula.endswith(")"):
                vals = [v for v in self._get_range_values(formula[6:-1]) if isinstance(v, float)]
                return len(vals)
            elif formula.startswith("CONCATENATE(") and formula.endswith(")"):
                vals = [str(v) for v in self._get_range_values(formula[12:-1])]
                return "".join(vals)
                
            parsed_math = formula
            for col in self.col_headers:
                for row in range(1, 101):
                    cell = f"{col}{row}"
                    if cell in parsed_math:
                        parsed_math = parsed_math.replace(cell, str(self._get_cell_value(cell)))
            return eval(parsed_math, {"__builtins__": None}, {})
        except Exception:
            return "#ERROR!"

    def load_file(self, filepath):
        import pandas as pd 
        ext = filepath.lower().split('.')[-1]
        try:
            if ext in ["xlsx", "xls"]: 
                df = pd.read_excel(filepath, header=None)
            elif ext == "ods": 
                df = pd.read_excel(filepath, header=None, engine="odf")
            else: 
                df = pd.read_csv(filepath, header=None)

            df = df.fillna("")
            self.table.setRowCount(max(100, len(df)))
            self.table.setColumnCount(max(26, len(df.columns)))

            self._is_updating = True
            self.table.clearContents()
            for r in range(len(df)):
                for c in range(len(df.columns)):
                    item = QtWidgets.QTableWidgetItem(str(df.iat[r, c]))
                    self.table.setItem(r, c, item)
            self._is_updating = False
            return True
        except Exception as e:
            raise e

    def save_file(self, filepath):
        import pandas as pd
        rows = self.table.rowCount()
        cols = self.table.columnCount()
        
        max_r, max_c = 0, 0
        data = []
        for r in range(rows):
            row_data = []
            has_data = False
            for c in range(cols):
                item = self.table.item(r, c)
                val = item.text() if item else ""
                if val: has_data = True
                row_data.append(val)
            data.append(row_data)
            if has_data: max_r = r

        for r in data[:max_r+1]:
            for c, val in enumerate(r):
                if val: max_c = max(max_c, c)

        trimmed_data = [r[:max_c+1] for r in data[:max_r+1]]
        df = pd.DataFrame(trimmed_data)

        ext = filepath.lower().split('.')[-1]
        if ext in ["xlsx", "xls"]: 
            df.to_excel(filepath, index=False, header=False)
        elif ext == "ods": 
            df.to_excel(filepath, index=False, header=False, engine="odf")
        else: 
            df.to_csv(filepath, index=False, header=False)