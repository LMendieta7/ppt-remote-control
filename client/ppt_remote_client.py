import socket
import threading
import win32com.client
import time
import queue
import sys

# === CONFIG ===
UDP_PORT = 5005
DISCOVERY_PORT = 5001
slide_queue = queue.Queue()
exit_requested = threading.Event()
SERVER_IP = None

# === Server Discovery ===
def listen_for_discovery():
    global SERVER_IP
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    sock.bind(('', DISCOVERY_PORT))
    print("[CLIENT] Listening for server broadcast...")

    while SERVER_IP is None and not exit_requested.is_set():
        try:
            data, addr = sock.recvfrom(1024)
            if data.decode().startswith("DISCOVER:PPT_SERVER"):
                SERVER_IP = addr[0]
                print(f"[CLIENT] Discovered server at {SERVER_IP}")
                break
        except Exception as e:
            print(f"[CLIENT] Discovery error: {e}")

    sock.close()

listen_for_discovery()

if SERVER_IP is None:
    print("[CLIENT] No server discovered. Exiting.")
    sys.exit()

# === UDP Command Socket ===
sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
sock.bind(('', 5006))

# === PowerPoint COM Setup ===
ppt = win32com.client.Dispatch("PowerPoint.Application")
presentation = ppt.Presentations(1)
while True:
    try:
        slide_show = presentation.SlideShowWindow.View
        print("[CLIENT] Slideshow detected.")
        break
    except:
        print("[CLIENT] Waiting for slideshow...")
        time.sleep(1)

# === Background listener for slide sync ===
def udp_listener():
    while not exit_requested.is_set():
        try:
            data, _ = sock.recvfrom(1024)
            if data.decode().startswith("SLIDE:"):
                slide_num = int(data.decode().split(":")[1])
                slide_queue.put(slide_num)
        except Exception as e:
            print(f"[CLIENT] Listener error: {e}")

threading.Thread(target=udp_listener, daemon=True).start()

# === Main loop (no GUI, no hotkey) ===
print("[CLIENT] Use ← / → arrow keys to control slides. Press Ctrl+C in terminal to quit.")

while not exit_requested.is_set():
    while not slide_queue.empty():
        try:
            slide_num = slide_queue.get()
            slide_show = presentation.SlideShowWindow.View
            slide_show.GotoSlide(slide_num)
        except Exception as e:
            print(f"[CLIENT] Slide sync error: {e}")

    try:
        import keyboard
        if keyboard.is_pressed('right'):
            sock.sendto(b'NEXT', (SERVER_IP, UDP_PORT))
            while keyboard.is_pressed('right'): pass

        elif keyboard.is_pressed('left'):
            sock.sendto(b'PREV', (SERVER_IP, UDP_PORT))
            while keyboard.is_pressed('left'): pass

    except:
        pass  # Keyboard may not be installed or usable in some environments

    time.sleep(0.05)
