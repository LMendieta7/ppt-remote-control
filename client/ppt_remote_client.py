# === ppt_remote_client.py ===
import sys
import socket
import threading
import time
import keyboard
import win32com.client
import pygetwindow as gw

from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QHBoxLayout, QLabel
from PyQt5.QtCore import Qt, QTimer
from PyQt5.QtGui import QColor, QPainter

# --- CONFIG ---
SERVER_IP = '192.168.1.100'  # Change this to your server's IP
SERVER_PORT = 5051

sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
ppt = win32com.client.Dispatch("PowerPoint.Application")
ppt.Visible = True

# --- Slide Sync ---
def go_to_slide(index):
    try:
        slide_show = ppt.SlideShowWindows(1)
        slide_show.View.GotoSlide(index)
        print(f"[PPT] Moved to slide {index}")
    except Exception as e:
        print(f"[PPT ERROR] {e}")

def get_server_slide():
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.settimeout(2.0)
        sock.sendto(b'GET_SLIDE', (SERVER_IP, SERVER_PORT))
        data, _ = sock.recvfrom(1024)
        if data.decode().startswith("SLIDE:"):
            index = int(data.decode().split(":")[1])
            print(f"[SYNC] Server slide: {index}")
            go_to_slide(index)
    except Exception as e:
        print(f"[SYNC ERROR] {e}")

# --- Send UDP Command ---
def send_command(cmd):
    try:
        sock.sendto(cmd.encode(), (SERVER_IP, SERVER_PORT))
        print(f"[SEND] {cmd}")
    except Exception as e:
        print(f"[UDP ERROR] {e}")

# --- Check if PowerPoint is in presentation mode ---
def is_powerpoint_in_slideshow():
    try:
        win = gw.getActiveWindow()
        return win and 'PowerPoint Slide Show' in win.title
    except:
        return False

# --- Global Arrow Key Listener ---
def listen_for_keys():
    def on_key(event):
        if is_powerpoint_in_slideshow():
            if event.name == 'left':
                send_command('PREV')
                get_server_slide()
            elif event.name == 'right':
                send_command('NEXT')
                get_server_slide()

    keyboard.on_press(on_key)
    while True:
        time.sleep(1)

# --- LED Widget ---
class StatusLED(QLabel):
    def __init__(self, color='green', parent=None):
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

# --- GUI ---
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

        self.led = StatusLED('red', self)
        hbox.addWidget(self.led)

        style = '''
            QPushButton {
                font-size:18px; padding:12px 18px;
                background:#f3f4f6; border-radius:8px; border: none;
            }
            QPushButton:hover {
                background:#e0f2fe;
            }
            QPushButton:pressed {
                background:#dbeafe;
            }
        '''

        btn_prev = QPushButton('⟵ Prev', self)
        btn_prev.setStyleSheet(style)
        btn_prev.clicked.connect(lambda: self.try_send('PREV'))
        btn_prev.setFocusPolicy(Qt.NoFocus)

        btn_next = QPushButton('Next ⟶', self)
        btn_next.setStyleSheet(style)
        btn_next.clicked.connect(lambda: self.try_send('NEXT'))
        btn_next.setFocusPolicy(Qt.NoFocus)

        hbox.addWidget(btn_prev)
        hbox.addWidget(btn_next)
        hbox.addStretch()

        btn_close = QPushButton('✕', self)
        btn_close.setStyleSheet('font-size:18px; background:transparent; border:none; color:#888;')
        btn_close.setFixedSize(32, 32)
        btn_close.clicked.connect(self.close)
        btn_close.setFocusPolicy(Qt.NoFocus)

        hbox.addWidget(btn_close)
        self.setLayout(hbox)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_status)
        self.timer.start(1000)
        self.update_status()

    def update_status(self):
        if is_powerpoint_in_slideshow():
            self.led.setColor('green')
        else:
            self.led.setColor('red')

    def try_send(self, cmd):
        if is_powerpoint_in_slideshow():
            send_command(cmd)
            get_server_slide()

    def enterEvent(self, event):
        self.show()

    def leaveEvent(self, event):
        pass

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

# --- App Entry Point ---
if __name__ == '__main__':
    threading.Thread(target=listen_for_keys, daemon=True).start()
    app = QApplication(sys.argv)
    win = PPTControl()
    win.show()
    sys.exit(app.exec_())
