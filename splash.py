# splash.py
"""
Chameleon Premium Splash Screen for Translit
"""

from PySide6 import QtCore, QtGui, QtWidgets
from PySide6.QtCore import Qt

class PremiumSplash(QtWidgets.QSplashScreen):
    def __init__(self):
        # Fallback placeholder matching your 720x740 dimensions
        pixmap = QtGui.QPixmap("splash.png")
        if pixmap.isNull():
            pixmap = QtGui.QPixmap(720, 740)
            pixmap.fill(Qt.GlobalColor.transparent)
            painter = QtGui.QPainter(pixmap)
            painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)
            
            grad = QtGui.QLinearGradient(0, 0, 720, 740)
            grad.setColorAt(0, QtGui.QColor("#0f172a"))
            grad.setColorAt(1, QtGui.QColor("#1e293b"))
            painter.setBrush(grad)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawRoundedRect(0, 0, 720, 740, 24, 24)
            painter.end()

        super().__init__(pixmap)
        self.setWindowOpacity(0.0)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

    def drawContents(self, painter):
        super().drawContents(painter)
        
        # 1. Subtle dark vignette at the bottom so the glowing text always pops
        bottom_grad = QtGui.QLinearGradient(0, self.height() - 250, 0, self.height())
        bottom_grad.setColorAt(0.0, QtGui.QColor(0, 0, 0, 0))    
        bottom_grad.setColorAt(1.0, QtGui.QColor(0, 0, 0, 160))  
        painter.fillRect(0, self.height() - 250, self.width(), 250, bottom_grad)

        # 2. Cursive Font Setup
        font = QtGui.QFont()
        font.setFamilies(["Brush Script MT", "Segoe Script", "Lucida Handwriting", "Monotype Corsiva", "cursive"])
        font.setPointSize(95) 
        font.setBold(True)
        font.setItalic(True)
        painter.setFont(font)
        
        rect = self.rect().adjusted(0, 0, 0, -35)
        
        # 3. Outer Glow Aura (Lime/Yellow mix)
        aura_color = QtGui.QColor(163, 230, 53, 12) 
        painter.setPen(QtGui.QPen(aura_color))
        for spread in range(2, 14, 2): 
            for dx, dy in [(-1,0), (1,0), (0,-1), (0,1), (-1,-1), (1,1), (-1,1), (1,-1)]:
                offset_rect = rect.translated(dx * spread, dy * spread)
                painter.drawText(offset_rect, Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignHCenter, "Translit")

        # 4. Inner Concentrated Glow (Bright Yellow)
        inner_glow = QtGui.QColor(253, 224, 71, 50) 
        painter.setPen(QtGui.QPen(inner_glow))
        for spread in range(1, 4):
            for dx, dy in [(-1,0), (1,0), (0,-1), (0,1)]:
                offset_rect = rect.translated(dx * spread, dy * spread)
                painter.drawText(offset_rect, Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignHCenter, "Translit")
        
        # 5. Core Chameleon Gradient Fill
        grad = QtGui.QLinearGradient(rect.left(), rect.bottom() - 140, rect.right(), rect.bottom())
        grad.setColorAt(0.0, QtGui.QColor("#fef08a")) # Neon Yellow
        grad.setColorAt(0.5, QtGui.QColor("#bef264")) # Shifting Yellow-Green
        grad.setColorAt(1.0, QtGui.QColor("#10b981")) # Deep Emerald

        painter.setPen(QtGui.QPen(QtGui.QBrush(grad), 1))
        painter.drawText(rect, Qt.AlignmentFlag.AlignBottom | Qt.AlignmentFlag.AlignHCenter, "Translit")

    def start_fade_in(self):
        self.show()
        self.fade_in = QtCore.QPropertyAnimation(self, b"windowOpacity")
        self.fade_in.setDuration(1000)
        self.fade_in.setStartValue(0.0)
        self.fade_in.setEndValue(1.0)
        self.fade_in.start()

    def start_fade_out(self):
        self.fade_out = QtCore.QPropertyAnimation(self, b"windowOpacity")
        self.fade_out.setDuration(800)
        self.fade_out.setStartValue(1.0)
        self.fade_out.setEndValue(0.0)
        self.fade_out.finished.connect(self.close)
        self.fade_out.start()