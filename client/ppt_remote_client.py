import socket
import time
import keyboard
import pythoncom
import win32com.client
from datetime import datetime

# === CONFIGURATION ===
SERVER_IP = '192.168.1.100'   # <-- Change this to your actual server IP
SERVER_PORT = 5051
LOG_FILE = "ppt_key_test.log"

# === Logger ===
def log(msg):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}"
    print(line)
    with open(LOG_FILE, "a") as f:
        f.write(line + "\n")

# === UDP Command Function ===
def send_command(cmd):
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.sendto(cmd.encode(), (SERVER_IP, SERVER_PORT))
        log(f"[SEND] {cmd}")
    except Exception as e:
        log(f"[ERROR] {e}")

# === PowerPoint Check (per-thread safe) ===
def is_powerpoint_in_slideshow(ppt_instance):
    try:
        return ppt_instance.SlideShowWindows.Count > 0
    except Exception as e:
        log(f"[CHECK ERROR] {e}")
        return False

# === Global Key Listener ===
def run_key_listener():
    pythoncom.CoInitialize()
    ppt = win32com.client.Dispatch("PowerPoint.Application")

    def on_key(event):
        log(f"[KEY] {event.name}")
        if is_powerpoint_in_slideshow(ppt):
            if event.name == 'left':
                send_command("PREV")
            elif event.name == 'right':
                send_command("NEXT")
        else:
            log("[INFO] Not in slideshow mode")

    keyboard.on_press(on_key)
    log("[READY] Listening for arrow keys (left/right)...")
    while True:
        time.sleep(1)

# === MAIN ===
if __name__ == "__main__":
    run_key_listener()
