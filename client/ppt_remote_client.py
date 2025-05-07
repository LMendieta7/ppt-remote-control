import socket
import threading
import time
import keyboard
import win32com.client
import pythoncom

# --- CONFIG ---
SERVER_IP = '192.168.1.100'  # CHANGE THIS
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

# --- Send Command ---
def send_command(cmd):
    try:
        sock.sendto(cmd.encode(), (SERVER_IP, SERVER_PORT))
        print(f"[SEND] {cmd}")
    except Exception as e:
        print(f"[UDP ERROR] {e}")

# --- Keyboard Listener ---
def listen_for_keys():
    pythoncom.CoInitialize()
    try:
        ppt_thread = win32com.client.GetActiveObject("PowerPoint.Application")
    except Exception as e:
        print(f"[COM ERROR] {e}")
        return

    def on_key(event):
        print(f"[KEY] {event.name}")
        if event.name == 'left':
            send_command('PREV')
            get_server_slide()
        elif event.name == 'right':
            send_command('NEXT')
            get_server_slide()

    keyboard.on_press(on_key)
    while True:
        time.sleep(1)

# --- MAIN ---
if __name__ == '__main__':
    print("[CLIENT] PPT arrow-key control started")
    threading.Thread(target=listen_for_keys, daemon=True).start()
    while True:
        time.sleep(1)
