import socket
import keyboard
import time
import threading
import pythoncom
import win32com.client

# === UDP SETUP ===
UDP_PORT = 5005
BUFFER_SIZE = 1024

sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
sock.bind(('', UDP_PORT))
print(f"[SERVER] Listening for UDP commands on port {UDP_PORT}...")

# === GLOBAL CLIENT ADDRESS ===
client_address = None
exit_requested = threading.Event()   # You can use this flag to stop the thread

def broadcast_discovery():
    discovery_socket = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    discovery_socket.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)

    while not exit_requested.is_set():
        try:
            message = b"DISCOVER:PPT_SERVER"
            # Send to broadcast address on port 5001
            discovery_socket.sendto(message, ('255.255.255.255', 5001))
            print("[SERVER] Broadcasting discovery message...")
        except Exception as e:
            print(f"[SERVER] Broadcast error: {e}")

        time.sleep(3)  # Broadcast every 3 seconds

# === BACKGROUND THREAD: SEND SLIDE NUMBERS CONTINUOUSLY ===
def send_slide_number_loop():
    global client_address
    pythoncom.CoInitialize()

    ppt = win32com.client.Dispatch("PowerPoint.Application")
    time.sleep(1)  # Let PowerPoint initialize

    while True:
        if client_address:
            try:
                if ppt.SlideShowWindows.Count > 0:
                    slide_show = ppt.SlideShowWindows(1).View
                    current_slide = slide_show.CurrentShowPosition
                    message = f"SLIDE:{current_slide}"
                    sock.sendto(message.encode(), client_address)
                    print(f"[SERVER] Sent slide number: {current_slide}")
                else:
                    print("[SERVER] No active slideshow detected.")
            except Exception as e:
                print(f"[SERVER] Slide send failed: {e}")
        time.sleep(3)

# Start slide number sync thread
threading.Thread(target=send_slide_number_loop, daemon=True).start()
threading.Thread(target=broadcast_discovery, daemon=True).start()

# === MAIN LOOP: RECEIVE COMMANDS ===
while True:
    data, addr = sock.recvfrom(BUFFER_SIZE)
    message = data.decode().strip().upper()
    client_address = addr  # Save latest sender
    print(f"[SERVER] Received '{message}' from {addr}")

    if message == 'NEXT':
        keyboard.press_and_release('right')

    elif message == 'PREV':
        keyboard.press_and_release('left')

    elif message == 'EXIT':
        keyboard.press_and_release('esc')
        print("[SERVER] Exit command received. Shutting down.")
        break

sock.close()
print("[SERVER] Socket closed.")
