import sys
import socket
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QHBoxLayout, QLabel
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QColor, QPainter

SERVER_IP = '127.0.0.1'  # <-- CHANGE THIS to your server's IP!
SERVER_PORT = 5050

def check_server():
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.settimeout(0.2)
        sock.sendto(b'PING', (SERVER_IP, SERVER_PORT))
        return True
    except Exception:
        return False

def send_command(cmd):
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.sendto(cmd.encode(), (SERVER_IP, SERVER_PORT))
    except Exception:
        pass

class StatusLED(QLabel):
    def __init__(self, color='red', parent=None):
        super().__init__(parent)
        self.color = color
        self.setFixedSize(22, 22)
    def setColor(self, color):
        self.color = color
        self.update()
    def paintEvent(self, event):
        qp = QPainter(self)
        qp.setRenderHint(QPainter.Antialiasing)
        qp.setBrush(QColor(self.color))
        qp.setPen(Qt.gray)
        qp.drawEllipse(3, 3, 16, 16)

class PPTControl(QWidget):
    def __init__(self):
        super().__init__()
        self.offset = None
        self.initUI()

    def initUI(self):
        self.setWindowTitle('PPT Remote')
        self.setWindowFlags(Qt.WindowStaysOnTopHint | Qt.FramelessWindowHint)
        self.setFixedSize(360, 60)
        self.setStyleSheet('background:#f1f5f9; border-radius:18px;')
        screen = QApplication.primaryScreen().geometry()
        self.move(screen.width() - 400, 20)

        hbox = QHBoxLayout()
        hbox.setSpacing(10)

        # LED
        self.led = StatusLED('red', self)
        hbox.addWidget(self.led)

        # Button style with pressed effect
        button_style = '''
            QPushButton {
                font-size:18px; padding:12px 18px;
                background:#f3f4f6; border-radius:8px; border: none;
            }
            QPushButton:pressed {
                background:#dbeafe;
            }
        '''

        btn_prev = QPushButton('⟵ Prev', self)
        btn_prev.setStyleSheet(button_style)
        btn_prev.clicked.connect(lambda: send_command('PREV'))

        btn_next = QPushButton('Next ⟶', self)
        btn_next.setStyleSheet(button_style)
        btn_next.clicked.connect(lambda: send_command('NEXT'))

        hbox.addWidget(btn_prev)
        hbox.addWidget(btn_next)

        hbox.addStretch()

        # Close button (right)
        btn_close = QPushButton('✕', self)
        btn_close.setStyleSheet('font-size:18px; background:transparent; border:none; color:#888;')
        btn_close.setFixedSize(32, 32)
        btn_close.clicked.connect(self.close)
        hbox.addWidget(btn_close)

        self.setLayout(hbox)

        # Timer to check server status every second
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_status)
        self.timer.start(1000)
        self.update_status()

    def update_status(self):
        if check_server():
            self.led.setColor('green')
        else:
            self.led.setColor('red')

    # Dragging events
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.offset = event.globalPos() - self.frameGeometry().topLeft()
            event.accept()
    def mouseMoveEvent(self, event):
        if self.offset is not None and event.buttons() & Qt.LeftButton:
            self.move(event.globalPos() - self.offset)
            event.accept()
    def mouseReleaseEvent(self, event):
        self.offset = None
        event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = PPTControl()
    win.show()
    sys.exit(app.exec_())
