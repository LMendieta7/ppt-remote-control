import threading
import time
import sys
import itertools

# Spinner thread function
# === Spinner setup ===
spinner_running = True
def show_spinner(message="Launching PowerPoint Remote, please wait..."):
    spinner = itertools.cycle(['|', '/', '-', '\\'])
    while spinner_running:
        sys.stdout.write(f"\r{message} {next(spinner)}")
        sys.stdout.flush()
        time.sleep(0.1)
    print("\r" + " " * 50 + "\r", end="")  # Clear line

spinner_thread = threading.Thread(target=show_spinner)
spinner_thread.start()
# Simulate load time / init logic

import socket
import keyboard
import win32com.client
import queue
import pythoncom
import tkinter as tk
import ctypes
from discovery_helper import wait_for_server
from gui_helper import FloatingControl

SERVER_IP = wait_for_server()
UDP_PORT = 5005
POLL_INTERVAL = 3

sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
sock.bind(('', 5006))
sock.settimeout(1.0)

slide_queue = queue.Queue()
sync_request_flag = threading.Event()
last_manual_time = time.time()
current_slide = 0
running = True  # For clean exit control

def hide_console():
    if sys.platform == "win32":
        ctypes.windll.user32.ShowWindow(
            ctypes.windll.kernel32.GetConsoleWindow(), 0  # 0 = hide
        )


def poll_slide_sync():
    global last_manual_time, current_slide
    while running:
        if time.time() - last_manual_time >= POLL_INTERVAL or sync_request_flag.is_set():
            try:
                sock.sendto(b'GET_SLIDE', (SERVER_IP, UDP_PORT))
                data, _ = sock.recvfrom(1024)
                message = data.decode().strip()
                if message.startswith("SLIDE:"):
                    slide_num = int(message.split(":")[1])
                    if slide_num != current_slide:
                        slide_queue.put(slide_num)
                        current_slide = slide_num
                        print(f"[CLIENT] Synced slide from server: {slide_num}")
            except:
                pass
            last_manual_time = time.time()
            sync_request_flag.clear()
        time.sleep(0.05)

threading.Thread(target=poll_slide_sync, daemon=True).start()

def monitor_ppt_slideshow():
    global current_slide
    pythoncom.CoInitialize()
    slideshow_active = False

    while running:
        try:
            ppt = win32com.client.Dispatch("PowerPoint.Application")
            while not slide_queue.empty():
                slide_num = slide_queue.get()
                if ppt.SlideShowWindows.Count > 0:
                    ppt.SlideShowWindows(1).View.GotoSlide(slide_num)
                    current_slide = slide_num
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
                            current_slide = slide_num
                            slideshow_active = True
            else:
                slideshow_active = False

        except Exception as e:
            print(f"[CLIENT] monitor_ppt_slideshow error: {e}")
            slideshow_active = False

        time.sleep(0.5)

threading.Thread(target=monitor_ppt_slideshow, daemon=True).start()

def keyboard_loop():
    global last_manual_time, running
    while running:
        try:
            if keyboard.is_pressed('right'):
                sock.sendto(b'NEXT', (SERVER_IP, UDP_PORT))
                try:
                    data, _ = sock.recvfrom(1024)
                    message = data.decode().strip()
                    if message.startswith("ACK:"):
                        print(f"[CLIENT] Server acknowledged {message}")
                except:
                    pass
                last_manual_time = time.time()
                sync_request_flag.clear()
                while keyboard.is_pressed('right'): pass

            elif keyboard.is_pressed('left'):
                sock.sendto(b'PREV', (SERVER_IP, UDP_PORT))
                try:
                    data, _ = sock.recvfrom(1024)
                    message = data.decode().strip()
                    if message.startswith("ACK:"):
                        print(f"[CLIENT] Server acknowledged {message}")
                except:
                    pass
                last_manual_time = time.time()
                sync_request_flag.clear()
                while keyboard.is_pressed('left'): pass

        except:
            pass

        time.sleep(0.04)

keyboard_thread = threading.Thread(target=keyboard_loop)
keyboard_thread.start()

def start_gui():
    root = tk.Tk()

    def send_and_wait(command):
        global last_manual_time, current_slide
        sock.sendto(command.encode(), (SERVER_IP, UDP_PORT))
        try:
            data, _ = sock.recvfrom(1024)
            message = data.decode().strip()
            if message.startswith("ACK:"):
                print(f"[CLIENT] Server acknowledged {message}")
        except:
            pass
        last_manual_time = time.time()
        sync_request_flag.clear()

    def on_prev():
        send_and_wait("PREV")

    def on_next():
        send_and_wait("NEXT")

    def on_close():
        global running
        print("[CLIENT] Exiting cleanly...")
        running = False
        try:
            sock.close()
        except:
            pass
        root.quit()
        root.destroy()
        sys.exit(0)

    FloatingControl(root, on_prev, on_next, on_close)
     # Stop spinner and hide console after GUI loads
    global spinner_running
    spinner_running = False
    time.sleep(1)
    hide_console()
    root.mainloop()

# Run GUI in main thread to keep terminal responsive and allow clean exit
start_gui()

print("[CLIENT] Client is running. Use arrow keys or GUI to control slides.")
