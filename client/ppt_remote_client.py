import socket
import keyboard
import threading
import win32com.client
import time
import queue
import pythoncom
import tkinter as tk
from discovery_helper import wait_for_server
from gui_helper import FloatingControl

# === CONFIGURATION ===
SERVER_IP = wait_for_server()
UDP_PORT = 5005
POLL_INTERVAL = 3  # seconds

# === SETUP ===
sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
sock.bind(('', 5006))
sock.settimeout(0.5)

# === State Tracking ===
slide_queue = queue.Queue()
sync_request_flag = threading.Event()
last_manual_time = time.time()
current_slide = 0

# === BACKGROUND THREAD: Poll server every 3 seconds or on sync request ===
def poll_slide_sync():
    global last_manual_time
    while True:
        if time.time() - last_manual_time >= POLL_INTERVAL or sync_request_flag.is_set():
            try:
                sock.sendto(b'GET_SLIDE', (SERVER_IP, UDP_PORT))
                data, _ = sock.recvfrom(1024)
                message = data.decode().strip()
                if message == "NOSHOW":
                    print("[CLIENT] Slideshow not active.")
                elif message.startswith("SLIDE:"):
                    slide_num = int(message.split(":")[1])
                    slide_queue.put(slide_num)
                    print(f"[CLIENT] Synced slide from server: {slide_num}")
            except socket.timeout:
                pass
            except Exception as e:
                print(f"[CLIENT] Poll error: {e}")
            last_manual_time = time.time()
            sync_request_flag.clear()
        time.sleep(0.05)

threading.Thread(target=poll_slide_sync, daemon=True).start()

# === BACKGROUND THREAD: Monitor and control local PowerPoint slideshow ===
def monitor_ppt_slideshow():
    pythoncom.CoInitialize()
    slideshow_active = False

    while True:
        try:
            ppt = win32com.client.Dispatch("PowerPoint.Application")

            while not slide_queue.empty():
                slide_num = slide_queue.get()
                if ppt.SlideShowWindows.Count > 0:
                    ppt.SlideShowWindows(1).View.GotoSlide(slide_num)
                    print(f"[CLIENT] Slide updated from queue: {slide_num}")

            if ppt.SlideShowWindows.Count > 0:
                if not slideshow_active:
                    time.sleep(1)
                    if ppt.SlideShowWindows.Count > 0:
                        print("[CLIENT] Slideshow started. Syncing to server...")
                        sock.sendto(b'GET_SLIDE', (SERVER_IP, UDP_PORT))
                        data, _ = sock.recvfrom(1024)
                        message = data.decode().strip()
                        if message.startswith("SLIDE:"):
                            slide_num = int(message.split(":")[1])
                            ppt.SlideShowWindows(1).View.GotoSlide(slide_num)
                            slideshow_active = True
            else:
                slideshow_active = False

        except Exception as e:
            print(f"[CLIENT] monitor_ppt_slideshow error: {e}")
            slideshow_active = False

        time.sleep(1)

threading.Thread(target=monitor_ppt_slideshow, daemon=True).start()

# === GUI WINDOW ===
def start_gui():
    root = tk.Tk()

    def on_prev():
        sock.sendto(b'PREV', (SERVER_IP, UDP_PORT))
        global last_manual_time
        last_manual_time = time.time()
        time.sleep(0.3)
        sync_request_flag.set()

    def on_next():
        sock.sendto(b'NEXT', (SERVER_IP, UDP_PORT))
        global last_manual_time
        last_manual_time = time.time()
        time.sleep(0.3)
        sync_request_flag.set()

    def on_close():
        root.destroy()

    FloatingControl(root, on_prev, on_next, on_close)
    root.mainloop()

threading.Thread(target=start_gui, daemon=True).start()

# === MAIN LOOP: Keyboard control only ===
print("[CLIENT] Press ← or → to control slides. Ctrl+C to quit.")
while True:
    try:
        if keyboard.is_pressed('right'):
            sock.sendto(b'NEXT', (SERVER_IP, UDP_PORT))
            last_manual_time = time.time()
            sync_request_flag.set()
            while keyboard.is_pressed('right'): pass
        elif keyboard.is_pressed('left'):
            sock.sendto(b'PREV', (SERVER_IP, UDP_PORT))
            last_manual_time = time.time()
            sync_request_flag.set()
            while keyboard.is_pressed('left'): pass
    except:
        pass

    time.sleep(0.05)
