import socket
import threading
import keyboard
import time
import win32com.client
import pythoncom

# --- CONFIG ---
SERVER_IP = "10.0.0.2"  # Replace with your server's IP
SERVER_PORT = 5051

ppt = win32com.client.Dispatch("PowerPoint.Application")
ppt.Visible = True

def go_to_slide(index):
    try:
        if ppt.SlideShowWindows.Count > 0:
            ppt.SlideShowWindows(1).View.GotoSlide(index)
            print(f"[CLIENT] Synced to slide {index}")
    except Exception as e:
        print(f"[CLIENT SYNC ERROR] {e}")

def send_command_and_sync(cmd):
    try:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.settimeout(2.0)
        sock.sendto(cmd.encode(), (SERVER_IP, SERVER_PORT))
        data, _ = sock.recvfrom(1024)
        if data.decode().startswith("SLIDE:"):
            index = int(data.decode().split(":")[1])
            go_to_slide(index)
    except Exception as e:
        print(f"[CLIENT NETWORK ERROR] {e}")

def listen_for_keys():
    pythoncom.CoInitialize()
    def on_key(event):
        if event.name == 'left':
            send_command_and_sync("PREV")
        elif event.name == 'right':
            send_command_and_sync("NEXT")
    keyboard.on_press(on_key)
    while True:
        time.sleep(1)

if __name__ == "__main__":
    threading.Thread(target=listen_for_keys, daemon=True).start()
    print("[CLIENT] Listening for arrow keys to control server...")
    while True:
        time.sleep(1)
