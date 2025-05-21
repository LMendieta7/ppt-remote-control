import socket
import keyboard
import threading
import win32com.client
import time
import queue
from discovery_helper import wait_for_server

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
last_manual_time = time.time()
current_slide = 0

# === PowerPoint COM must stay in main thread ===
ppt = win32com.client.Dispatch("PowerPoint.Application")
presentation = ppt.Presentations(1)

# Wait for slideshow to start
while True:
    try:
        slide_show = presentation.SlideShowWindow.View
        print("[CLIENT] Slideshow detected. Ready to sync.")
        break
    except:
        print("[CLIENT] Waiting for slideshow to start...")
        time.sleep(1)

# === BACKGROUND THREAD: Receive slide sync from server ===
def udp_listener():
    while True:
        try:
            data, _ = sock.recvfrom(1024)
            message = data.decode().strip()
            if message.startswith("SLIDE:"):
                slide_num = int(message.split(":")[1])
                slide_queue.put(slide_num)
        except socket.timeout:
            continue
        except Exception as e:
            print(f"[CLIENT] Listener error: {e}")

threading.Thread(target=udp_listener, daemon=True).start()

# === BACKGROUND THREAD: Poll server every 4 seconds ===
def poll_slide_sync():
    global last_manual_time
    while True:
        if time.time() - last_manual_time >= POLL_INTERVAL:
            try:
                sock.sendto(b'GET_SLIDE', (SERVER_IP, UDP_PORT))
            except Exception as e:
                print(f"[CLIENT] Poll error: {e}")
            last_manual_time = time.time()  # reset even after poll
        time.sleep(0.05)

threading.Thread(target=poll_slide_sync, daemon=True).start()

# === MAIN LOOP: Process slide sync + listen for keys ===
print("[CLIENT] Press ← or → to control slides. Press Q to quit.")
while True:
    while not slide_queue.empty():
        try:
            slide_num = slide_queue.get()
            if slide_num != current_slide:
                slide_show = presentation.SlideShowWindow.View
                slide_show.GotoSlide(slide_num)
                current_slide = slide_num
                print(f"[CLIENT] Synced to slide {slide_num}")
        except Exception as e:
            print(f"[CLIENT] Slide sync error: {e}")

    try:
        if keyboard.is_pressed('right'):
            sock.sendto(b'NEXT', (SERVER_IP, UDP_PORT))
            last_manual_time = time.time()
            while keyboard.is_pressed('right'): pass

        elif keyboard.is_pressed('left'):
            sock.sendto(b'PREV', (SERVER_IP, UDP_PORT))
            last_manual_time = time.time()
            while keyboard.is_pressed('left'): pass

    
    except:
        pass  # keyboard lib may throw if not available

    time.sleep(0.05)
