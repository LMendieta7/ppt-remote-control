import socket
import keyboard
import threading
import win32com.client
import time
import queue
from discovery_helper import wait_for_server

# === CONFIGURATION ===
SERVER_IP =  wait_for_server() # Replace with actual IP
UDP_PORT = 5005

# === SETUP ===
sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
sock.bind(('', 5006))  # local port for receiving

# === Queue for inter-thread slide sync ===
slide_queue = queue.Queue()

# === PowerPoint COM must stay in main thread ===
ppt = win32com.client.Dispatch("PowerPoint.Application")
presentation = ppt.Presentations(1)

# Wait until slideshow is started manually
while True:
    try:
        slide_show = presentation.SlideShowWindow.View
        print("[CLIENT] Slideshow detected. Ready to sync.")
        break
    except:
        print("[CLIENT] Waiting for slideshow to start...")
        time.sleep(1)

# === BACKGROUND THREAD to receive slide sync ===
def udp_listener():
    while True:
        try:
            data, _ = sock.recvfrom(1024)
            message = data.decode().strip()
            if message.startswith("SLIDE:"):
                slide_num = int(message.split(":")[1])
                print(f"[CLIENT] Received slide number: {slide_num}")
                slide_queue.put(slide_num)
        except Exception as e:
            print(f"[CLIENT] Listener error: {e}")

threading.Thread(target=udp_listener, daemon=True).start()

# === MAIN LOOP: send commands + apply slide sync ===
print("[CLIENT] Press ← or → to control slides. Press Q to quit.")
while True:
    # Process queued slide syncs (safe in main thread)
    while not slide_queue.empty():
        try:
            slide_num = slide_queue.get()
            slide_show = presentation.SlideShowWindow.View
            slide_show.GotoSlide(slide_num)
            print(f"[CLIENT] Synced to slide {slide_num}")
        except Exception as e:
            print(f"[CLIENT] Slide sync error: {e}")

    # Send NEXT/PREV via UDP
    if keyboard.is_pressed('right'):
        sock.sendto(b'NEXT', (SERVER_IP, UDP_PORT))
        while keyboard.is_pressed('right'): pass

    elif keyboard.is_pressed('left'):
        sock.sendto(b'PREV', (SERVER_IP, UDP_PORT))
        while keyboard.is_pressed('left'): pass

    elif keyboard.is_pressed('q'):
        sock.sendto(b'EXIT', (SERVER_IP, UDP_PORT))
        print("[CLIENT] Exiting...")
        break

    time.sleep(0.05)  # keep CPU load low