# splash.py
"""
Chameleon Premium Splash Screen for Translit
"""

from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtCore import Qt, QTimer
import os

class PremiumSplash(QtWidgets.QSplashScreen):
    def __init__(self):
        # A cinematic 720x480 widescreen ratio
        pixmap = QtGui.QPixmap(720, 480)
        pixmap.fill(Qt.GlobalColor.transparent)
        super().__init__(pixmap)

        self.setWindowOpacity(0.0)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        # --- UI Layout ---
        layout = QtWidgets.QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignHCenter)
        layout.setContentsMargins(0, 0, 0, 35) # Spacing from the bottom edge

        self.loading_label = QtWidgets.QLabel("Initializing Translit...")
        self.loading_label.setStyleSheet("""
            color: #bef264; 
            font-family: 'Segoe UI', system-ui, sans-serif; 
            font-size: 14px; 
            font-weight: 600; 
            letter-spacing: 1.5px;
            background: transparent;
        """)
        self.loading_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.loading_label)

        # --- Smooth 60FPS Spinner State ---
        self.spinner_angle = 0
        self.spin_timer = QTimer(self)
        self.spin_timer.timeout.connect(self.update_spinner)
        self.spin_timer.start(16) # ~60 FPS

    def update_spinner(self):
        """Advances the rotation angle and triggers a repaint."""
        self.spinner_angle = (self.spinner_angle + 5) % 360
        self.repaint() # Calls drawContents automatically

    def update_text(self, message: str):
        self.loading_label.setText(message)
        # Force the UI to repaint immediately every time the text changes
        QtWidgets.QApplication.processEvents()

    def drawContents(self, painter: QtGui.QPainter):
        # Turn on all high-quality rendering flags
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)
        painter.setRenderHint(QtGui.QPainter.RenderHint.TextAntialiasing)
        painter.setRenderHint(QtGui.QPainter.RenderHint.SmoothPixmapTransform)

        w, h = self.width(), self.height()
        
        # 1. Base Glass Background (Deep Navy/Ocean Blue)
        base_path = QtGui.QPainterPath()
        base_path.addRoundedRect(0, 0, w, h, 20, 20)
        
        bg_grad = QtGui.QLinearGradient(0, 0, w, h)
        bg_grad.setColorAt(0.0, QtGui.QColor("#020617")) # Deep midnight dark blue
        bg_grad.setColorAt(1.0, QtGui.QColor("#0f766e")) # Deep teal/cyan
        painter.fillPath(base_path, bg_grad)
        
        # 2. Modern Aurora Depth Orbs (Greenish-Yellow & Sky Blue)
        painter.setClipPath(base_path) 
        
        # Top right orb: Greenish Yellow (Lime)
        orb1_grad = QtGui.QRadialGradient(w * 0.85, h * 0.1, w * 0.5)
        orb1_grad.setColorAt(0.0, QtGui.QColor(163, 230, 53, 90)) # #a3e635 with opacity
        orb1_grad.setColorAt(1.0, QtGui.QColor(163, 230, 53, 0))
        painter.fillRect(0, 0, w, h, orb1_grad)
        
        # Bottom left orb: Bright Blue/Cyan
        orb2_grad = QtGui.QRadialGradient(w * 0.1, h * 0.9, w * 0.6)
        orb2_grad.setColorAt(0.0, QtGui.QColor(14, 165, 233, 85)) # #0ea5e9 with opacity
        orb2_grad.setColorAt(1.0, QtGui.QColor(14, 165, 233, 0))
        painter.fillRect(0, 0, w, h, orb2_grad)

        # 3. Draw Transparent PNG Logo
        # Make sure "splash_logo.png" exists in your project folder!
        logo_pixmap = QtGui.QPixmap("icons/splash.png")
        if not logo_pixmap.isNull():
            # Scale it nicely (max 180x180) to leave room for text
            logo = logo_pixmap.scaled(180, 180, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
            logo_x = int((w - logo.width()) / 2)
            logo_y = int(h * 0.12) # Distance from top
            painter.drawPixmap(logo_x, logo_y, logo)

        # 4. Typography (Clean, Bold, Sans-serif)
        font = QtGui.QFont("Segoe UI", 58, QtGui.QFont.Weight.ExtraBold)
        font.setLetterSpacing(QtGui.QFont.SpacingType.AbsoluteSpacing, -1.0)
        painter.setFont(font)
        
        # Shift text down slightly to anchor below the logo
        text_rect = QtCore.QRect(0, 0, w, h + 80)
        
        # Clean Drop Shadow for text
        painter.setPen(QtGui.QColor(0, 0, 0, 150))
        painter.drawText(text_rect.translated(2, 4), Qt.AlignmentFlag.AlignCenter, "Translit")
        
        # Text Fill (Metallic Gradient matching the theme)
        text_grad = QtGui.QLinearGradient(text_rect.left(), text_rect.top(), text_rect.right(), text_rect.bottom())
        text_grad.setColorAt(0.0, QtGui.QColor("#ffffff"))
        text_grad.setColorAt(0.5, QtGui.QColor("#fef08a")) # Soft yellow
        text_grad.setColorAt(1.0, QtGui.QColor("#38bdf8")) # Sky blue
        
        painter.setPen(QtGui.QPen(QtGui.QBrush(text_grad), 1))
        painter.drawText(text_rect, Qt.AlignmentFlag.AlignCenter, "Translit")
        
        # Subtitle
        sub_font = QtGui.QFont("Segoe UI", 11, QtGui.QFont.Weight.Medium)
        sub_font.setLetterSpacing(QtGui.QFont.SpacingType.AbsoluteSpacing, 5)
        painter.setFont(sub_font)
        painter.setPen(QtGui.QColor(226, 232, 240, 200)) # Light slate
        painter.drawText(text_rect.translated(0, 60), Qt.AlignmentFlag.AlignCenter, "DOCUMENT STUDIO")

        # 5. Sleek Animated Loading Spinner
        spinner_size = 28
        spinner_rect = QtCore.QRectF((w - spinner_size) / 2, h - 90, spinner_size, spinner_size)
        
        # Static background track
        pen_track = QtGui.QPen(QtGui.QColor(255, 255, 255, 30))
        pen_track.setWidth(3)
        pen_track.setCapStyle(Qt.PenCapStyle.RoundCap)
        painter.setPen(pen_track)
        painter.drawArc(spinner_rect, 0, 360 * 16)
        
        # Rotating neon head
        pen_head = QtGui.QPen(QtGui.QColor("#bef264")) # Bright greenish yellow
        pen_head.setWidth(3)
        pen_head.setCapStyle(Qt.PenCapStyle.RoundCap)
        painter.setPen(pen_head)
        painter.drawArc(spinner_rect, -self.spinner_angle * 16, 110 * 16) 

    def start_fade_in(self):
        self.show()
        self.fade_in = QtCore.QPropertyAnimation(self, b"windowOpacity")
        self.fade_in.setDuration(1000)
        self.fade_in.setStartValue(0.0)
        self.fade_in.setEndValue(1.0)
        self.fade_in.start()

    def start_fade_out(self):
        self.fade_out = QtCore.QPropertyAnimation(self, b"windowOpacity")
        self.fade_out.setDuration(700)
        self.fade_out.setStartValue(1.0)
        self.fade_out.setEndValue(0.0)
        self.fade_out.finished.connect(self.close)
        self.fade_out.start()