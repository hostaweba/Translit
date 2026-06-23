# style.py
"""
CSS Stylesheets and Theming Engine for Translit
"""

def get_main_theme(is_dark: bool) -> str:
    if is_dark:
        return """
            /* --- ULTRA-COMPACT DARK THEME --- */
            QMainWindow, QDialog, QMessageBox, QInputDialog, QFileDialog, QPageSetupDialog { 
                background-color: #111827; color: #f9fafb; 
            }
            QLabel, QCheckBox, QRadioButton, QGroupBox { color: #f9fafb; }
            
            /* --- TAB WIDGET (DARK) --- */
            QTabWidget::pane { border: 1px solid #374151; border-radius: 4px; background-color: #1f2937; }
            QTabBar::tab { background-color: #111827; color: #d1d5db; border: 1px solid #374151; padding: 6px 12px; border-top-left-radius: 4px; border-top-right-radius: 4px; margin-right: 2px; }
            QTabBar::tab:selected { background-color: #1f2937; color: #10b981; border-bottom-color: #1f2937; font-weight: bold; }
            QTabBar::tab:hover:!selected { background-color: #374151; color: #ffffff; }
            QTabWidget QWidget { background-color: transparent; color: #f9fafb; }
            
            QToolBar { 
                background-color: #1f2937; border-bottom: 1px solid #374151; 
                padding: 2px; spacing: 2px; 
            }
            
            QToolButton { 
                background-color: transparent; color: #d1d5db; 
                border: 1px solid transparent; border-radius: 3px; 
                padding: 2px 6px; font-weight: normal; height: 20px; 
            }
            QToolButton:hover { background-color: #374151; color: #ffffff; border: 1px solid #4b5563; }
            QToolButton:pressed, QToolButton:checked { background-color: #111827; border: 1px solid #4b5563; color: #10b981; }
            
            QToolButton::menu-indicator { right: 2px; width: 6px; subcontrol-origin: padding; subcontrol-position: center right; }
            QToolButton[popupMode="1"] { padding-right: 14px; }
            
            QMenuBar { background-color: #1f2937; border-bottom: 1px solid #374151; min-height: 22px; }
            QMenuBar::item { padding: 3px 8px; border-radius: 3px; color: #f9fafb; }
            QMenuBar::item:selected { background-color: #374151; }
            
            QMenu { background-color: #1f2937; color: #f9fafb; border: 1px solid #4b5563; border-radius: 4px; padding: 2px; }
            QMenu::item { padding: 4px 20px 4px 12px; border-radius: 3px; background-color: transparent; }
            QMenu::item:selected { background-color: #059669; color: white; }
            QMenu::separator { background-color: #4b5563; height: 1px; margin: 2px 0px; }
            
            QLineEdit, QSpinBox, QDoubleSpinBox, QFontComboBox, QComboBox { 
                background-color: #111827; color: #f9fafb; 
                border: 1px solid #4b5563; border-radius: 3px; 
                padding: 1px 6px; height: 20px; margin: 0px 2px; 
            }
            QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QFontComboBox:focus, QComboBox:focus { border: 1px solid #3b82f6; background-color: #1f2937; }
            
            QSpinBox::up-button, QSpinBox::down-button, QDoubleSpinBox::up-button, QDoubleSpinBox::down-button { width: 14px; background: transparent; border-left: 1px solid #4b5563; }
            QSpinBox::up-button:hover, QSpinBox::down-button:hover, QDoubleSpinBox::up-button:hover, QDoubleSpinBox::down-button:hover { background: #374151; }
            
            QPushButton { 
                background-color: #374151; color: #f9fafb; 
                border: 1px solid #4b5563; border-radius: 3px; 
                padding: 3px 10px; font-weight: bold; height: 20px;
            }
            QPushButton:hover { background-color: #4b5563; }
            QPushButton:pressed { background-color: #1f2937; }
            
            QDockWidget { color: #f9fafb; font-weight: bold; }
            QDockWidget::title { background-color: #1f2937; padding: 4px; text-align: left; }
            QDockWidget > QWidget { background-color: #1f2937; border: 1px solid #374151; border-radius: 4px; margin: 2px; }
            
            QListWidget, QTableWidget { background-color: #1f2937; border: 1px solid #374151; color: #f9fafb; border-radius: 4px; }
            QListWidget::item, QTableWidget::item { padding: 4px; border-radius: 3px; }
            QListWidget::item:selected, QTableWidget::item:selected { background-color: #059669; color: #ffffff; }
            
            QScrollArea { background-color: transparent; border: none; }
            QFrame#page_frame { background-color: #1f2937; border: 1px solid #374151; border-radius: 4px; }
            
            QStatusBar { background-color: #1f2937; color: #d1d5db; border-top: 1px solid #374151; min-height: 20px; }
            QStatusBar QLabel { color: #d1d5db; background: transparent; padding: 0px; }
            QTextEdit { background-color: transparent; color: #f9fafb; border: none; }
            
            QLabel#TopClock { color: #4ade80; font-weight: 800; font-size: 18px; padding-right: 12px; margin: 0px; }
            
            QTableWidget QLineEdit { background-color: #374151; color: #ffffff; border: none; padding: 0px; margin: 0px; outline: none; }
        """
    else:
        return """
            /* --- ULTRA-COMPACT LIGHT THEME --- */
            QMainWindow, QDialog, QMessageBox, QInputDialog, QFileDialog { 
                background-color: #f3f4f6; color: #111827; 
            }
            QLabel, QCheckBox, QRadioButton, QGroupBox { color: #111827; }
            
            /* --- TAB WIDGET (LIGHT) --- */
            QTabWidget::pane { border: 1px solid #d1d5db; border-radius: 4px; background-color: #ffffff; }
            QTabBar::tab { background-color: #f3f4f6; color: #4b5563; border: 1px solid #d1d5db; padding: 6px 12px; border-top-left-radius: 4px; border-top-right-radius: 4px; margin-right: 2px; }
            QTabBar::tab:selected { background-color: #ffffff; color: #059669; border-bottom-color: #ffffff; font-weight: bold; }
            QTabBar::tab:hover:!selected { background-color: #e5e7eb; color: #111827; }
            QTabWidget QWidget { background-color: transparent; color: #111827; }
            
            QToolBar { 
                background-color: #ffffff; border-bottom: 1px solid #e5e7eb; 
                padding: 2px; spacing: 2px; 
            }
            
            QToolButton { 
                background-color: transparent; color: #4b5563; 
                border: 1px solid transparent; border-radius: 3px; 
                padding: 2px 6px; font-weight: normal; height: 20px; 
            }
            QToolButton:hover { background-color: #f3f4f6; color: #111827; border: 1px solid #d1d5db; }
            QToolButton:pressed, QToolButton:checked { background-color: #e5e7eb; border: 1px solid #d1d5db; color: #059669; }
            
            QToolButton::menu-indicator { right: 2px; width: 6px; subcontrol-origin: padding; subcontrol-position: center right; }
            QToolButton[popupMode="1"] { padding-right: 14px; }
            
            QMenuBar { background-color: #ffffff; border-bottom: 1px solid #e5e7eb; min-height: 22px; }
            QMenuBar::item { padding: 3px 8px; border-radius: 3px; color: #111827; }
            QMenuBar::item:selected { background-color: #f3f4f6; }
            
            QMenu { background-color: #ffffff; color: #111827; border: 1px solid #d1d5db; border-radius: 4px; padding: 2px; }
            QMenu::item { padding: 4px 20px 4px 12px; border-radius: 3px; background-color: transparent; }
            QMenu::item:selected { background-color: #059669; color: white; }
            QMenu::separator { background-color: #e5e7eb; height: 1px; margin: 2px 0px; }
            
            QLineEdit, QSpinBox, QDoubleSpinBox, QFontComboBox, QComboBox { 
                background-color: #ffffff; color: #111827; 
                border: 1px solid #d1d5db; border-radius: 3px; 
                padding: 1px 6px; height: 20px; margin: 0px 2px; 
            }
            QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QFontComboBox:focus, QComboBox:focus { border: 1px solid #3b82f6; background-color: #ffffff; }
            
            QSpinBox::up-button, QSpinBox::down-button, QDoubleSpinBox::up-button, QDoubleSpinBox::down-button { width: 14px; background: transparent; border-left: 1px solid #d1d5db; }
            QSpinBox::up-button:hover, QSpinBox::down-button:hover, QDoubleSpinBox::up-button:hover, QDoubleSpinBox::down-button:hover { background: #f3f4f6; }
            
            QPushButton { 
                background-color: #ffffff; color: #111827; 
                border: 1px solid #d1d5db; border-radius: 3px; 
                padding: 3px 10px; font-weight: bold; height: 20px;
            }
            QPushButton:hover { background-color: #f3f4f6; }
            QPushButton:pressed { background-color: #e5e7eb; }
            
            QDockWidget { color: #111827; font-weight: bold; }
            QDockWidget::title { background-color: #e5e7eb; padding: 4px; text-align: left; }
            QDockWidget > QWidget { background-color: #ffffff; border: 1px solid #e5e7eb; border-radius: 4px; margin: 2px; }
            
            QListWidget, QTableWidget { background-color: #ffffff; border: 1px solid #d1d5db; color: #111827; border-radius: 4px; }
            QListWidget::item, QTableWidget::item { padding: 4px; border-radius: 3px; }
            QListWidget::item:selected, QTableWidget::item:selected { background-color: #059669; color: #ffffff; }
            
            QScrollArea { background-color: transparent; border: none; }
            QFrame#page_frame { background-color: #ffffff; border: 1px solid #e5e7eb; border-radius: 4px; }
            
            QPrintPreviewDialog { background-color: #e5e7eb; }
            QPrintPreviewDialog QToolBar { background-color: #1f2937; border-bottom: 1px solid #111827; padding: 4px; }
            QPrintPreviewDialog QToolButton { background-color: transparent; color: #f9fafb; border: 1px solid transparent; border-radius: 3px; padding: 4px; margin: 0px 2px; height: 20px;}
            QPrintPreviewDialog QToolButton:hover { background-color: #374151; border: 1px solid #4b5563; }
            QPrintPreviewDialog QToolButton:pressed { background-color: #111827; }
            
            QStatusBar { background-color: #ffffff; color: #374151; border-top: 1px solid #e5e7eb; min-height: 20px; }
            QStatusBar QLabel { color: #374151; background: transparent; padding: 0px; }
            QTextEdit { background-color: transparent; color: #000000; border: none; }
            
            QLabel#TopClock { color: darkgreen; font-weight: 800; font-size: 18px; padding-right: 12px; margin: 0px; }
            
            QTableWidget QLineEdit { background-color: #ffffff; color: #000000; border: none; padding: 0px; margin: 0px; outline: none; }
        """

def get_popup_style(sugg_style: str, is_dark: bool, sugg_text_color: str, size: int, is_bold: bool, spacing_level: str) -> str:
    bg = "rgba(31, 41, 55, 245)" if is_dark else "rgba(255, 255, 255, 250)"
    border = "#4b5563" if is_dark else "#d1d5db"
    fg = sugg_text_color if sugg_text_color else ("#f9fafb" if is_dark else "#111827")
    fw = "bold" if is_bold else "normal"

    if spacing_level == "Compact":
        pad = "4px 8px"
    elif spacing_level == "Relaxed":
        pad = "12px 16px"
    else: 
        pad = "8px 12px"

    if sugg_style == "Google (Search Style)":
        g_bg = "#202124" if is_dark else "#ffffff"
        g_fg = "#e8eaed" if is_dark else "#202124"
        g_sel = "#3c4043" if is_dark else "#f1f3f4"
        g_hov = "rgba(255, 255, 255, 0.1)" if is_dark else "rgba(0, 0, 0, 0.05)"
        g_border = "#5f6368" if is_dark else "#dfe1e5"
        
        return f"""
            QListWidget {{ background-color: {g_bg}; color: {g_fg}; border: 1px solid {g_border}; border-radius: 8px; font-size: {size}px; font-weight: {fw}; outline: none; }}
            QListWidget::item {{ padding: {pad}; border-radius: 0px; }}
            QListWidget::item:hover {{ background-color: {g_hov}; }}
            QListWidget::item:selected {{ background-color: {g_sel}; color: {g_fg}; }}
        """
    elif sugg_style == "Modern":
        sel_bg = "#3b82f6"
        hov_bg = "rgba(59, 130, 246, 0.2)"
        return f"""
            QListWidget {{ background-color: {bg}; color: {fg}; border: 1px solid {border}; border-radius: 12px; font-size: {size}px; font-weight: {fw}; outline: none; }}
            QListWidget::item {{ padding: {pad}; border-radius: 6px; }}
            QListWidget::item:hover {{ background-color: {hov_bg}; }}
            QListWidget::item:selected {{ background-color: {sel_bg}; color: white; }}
        """
    elif sugg_style == "Minimalist":
        accent = "#10b981"
        sel_bg = "rgba(16, 185, 129, 0.15)" if is_dark else "rgba(16, 185, 129, 0.2)"
        hov_bg = "rgba(16, 185, 129, 0.08)"
        return f"""
            QListWidget {{ background-color: {bg}; color: {fg}; border-left: 4px solid {accent}; border-top: 1px solid {border}; border-right: 1px solid {border}; border-bottom: 1px solid {border}; font-size: {size}px; font-weight: {fw}; outline: none; }}
            QListWidget::item {{ padding: {pad}; border-bottom: 1px solid transparent; }}
            QListWidget::item:hover {{ background-color: {hov_bg}; }}
            QListWidget::item:selected {{ background-color: {sel_bg}; color: {accent}; }}
        """
    elif sugg_style == "Neon":
        accent = "#38bdf8" if is_dark else "#0ea5e9"
        neon_fg = sugg_text_color if sugg_text_color else accent
        hov_bg = "rgba(56, 189, 248, 0.15)" if is_dark else "rgba(14, 165, 233, 0.15)"
        return f"""
            QListWidget {{ background-color: #0f172a; color: {neon_fg}; border: 2px solid {accent}; border-radius: 8px; font-size: {size}px; font-weight: {fw}; outline: none; }}
            QListWidget::item {{ padding: {pad}; border-radius: 4px; }}
            QListWidget::item:hover {{ background-color: {hov_bg}; }}
            QListWidget::item:selected {{ background-color: {accent}; color: #0f172a; font-weight: bold; }}
        """
    else: # Classic
        sel_bg = "#059669"
        hov_bg = "rgba(5, 150, 105, 0.15)"
        return f"""
            QListWidget {{ background-color: {bg}; color: {fg}; border: 1px solid {border}; border-radius: 8px; font-size: {size}px; font-weight: {fw}; outline: none; }}
            QListWidget::item {{ padding: {pad}; border-radius: 4px; }}
            QListWidget::item:hover {{ background-color: {hov_bg}; }}
            QListWidget::item:selected {{ background-color: {sel_bg}; color: white; }}
        """

def get_popup_style(sugg_style: str, is_dark: bool, sugg_text_color: str, size: int, is_bold: bool, spacing_level: str) -> str:
    bg = "rgba(31, 41, 55, 245)" if is_dark else "rgba(255, 255, 255, 250)"
    border = "#4b5563" if is_dark else "#d1d5db"
    fg = sugg_text_color if sugg_text_color else ("#f9fafb" if is_dark else "#111827")
    fw = "bold" if is_bold else "normal"

    if spacing_level == "Compact":
        pad = "4px 8px"
    elif spacing_level == "Relaxed":
        pad = "12px 16px"
    else: 
        pad = "8px 12px"

    if sugg_style == "Google (Search Style)":
        g_bg = "#202124" if is_dark else "#ffffff"
        g_fg = "#e8eaed" if is_dark else "#202124"
        g_sel = "#3c4043" if is_dark else "#f1f3f4"
        g_hov = "rgba(255, 255, 255, 0.1)" if is_dark else "rgba(0, 0, 0, 0.05)"
        g_border = "#5f6368" if is_dark else "#dfe1e5"
        
        return f"""
            QListWidget {{ background-color: {g_bg}; color: {g_fg}; border: 1px solid {g_border}; border-radius: 8px; font-size: {size}px; font-weight: {fw}; outline: none; }}
            QListWidget::item {{ padding: {pad}; border-radius: 0px; }}
            QListWidget::item:hover {{ background-color: {g_hov}; }}
            QListWidget::item:selected {{ background-color: {g_sel}; color: {g_fg}; }}
        """
    elif sugg_style == "Modern":
        sel_bg = "#3b82f6"
        hov_bg = "rgba(59, 130, 246, 0.2)"
        return f"""
            QListWidget {{ background-color: {bg}; color: {fg}; border: 1px solid {border}; border-radius: 12px; font-size: {size}px; font-weight: {fw}; outline: none; }}
            QListWidget::item {{ padding: {pad}; border-radius: 6px; }}
            QListWidget::item:hover {{ background-color: {hov_bg}; }}
            QListWidget::item:selected {{ background-color: {sel_bg}; color: white; }}
        """
    elif sugg_style == "Minimalist":
        accent = "#10b981"
        sel_bg = "rgba(16, 185, 129, 0.15)" if is_dark else "rgba(16, 185, 129, 0.2)"
        hov_bg = "rgba(16, 185, 129, 0.08)"
        return f"""
            QListWidget {{ background-color: {bg}; color: {fg}; border-left: 4px solid {accent}; border-top: 1px solid {border}; border-right: 1px solid {border}; border-bottom: 1px solid {border}; font-size: {size}px; font-weight: {fw}; outline: none; }}
            QListWidget::item {{ padding: {pad}; border-bottom: 1px solid transparent; }}
            QListWidget::item:hover {{ background-color: {hov_bg}; }}
            QListWidget::item:selected {{ background-color: {sel_bg}; color: {accent}; }}
        """
    elif sugg_style == "Neon":
        accent = "#38bdf8" if is_dark else "#0ea5e9"
        neon_fg = sugg_text_color if sugg_text_color else accent
        hov_bg = "rgba(56, 189, 248, 0.15)" if is_dark else "rgba(14, 165, 233, 0.15)"
        return f"""
            QListWidget {{ background-color: #0f172a; color: {neon_fg}; border: 2px solid {accent}; border-radius: 8px; font-size: {size}px; font-weight: {fw}; outline: none; }}
            QListWidget::item {{ padding: {pad}; border-radius: 4px; }}
            QListWidget::item:hover {{ background-color: {hov_bg}; }}
            QListWidget::item:selected {{ background-color: {accent}; color: #0f172a; font-weight: bold; }}
        """
    else: # Classic
        sel_bg = "#059669"
        hov_bg = "rgba(5, 150, 105, 0.15)"
        return f"""
            QListWidget {{ background-color: {bg}; color: {fg}; border: 1px solid {border}; border-radius: 8px; font-size: {size}px; font-weight: {fw}; outline: none; }}
            QListWidget::item {{ padding: {pad}; border-radius: 4px; }}
            QListWidget::item:hover {{ background-color: {hov_bg}; }}
            QListWidget::item:selected {{ background-color: {sel_bg}; color: white; }}
        """
