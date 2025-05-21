import socket
import keyboard
import time
import threading
import pythoncom
import win32com.client

# === UDP SETUP ===
UDP_PORT = 5005
BUFFER_SIZE = 1024
DISCOVERY_PORT = 5001
DISCOVERY_MESSAGE = b"DISCOVER_PPT_SERVER"
RESPONSE_MESSAGE = b"PPT_SERVER_HERE"

sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
sock.bind(('', UDP_PORT))
sock.settimeout(1.0)  # <- Key addition: avoid blocking forever
print(f"[SERVER] Listening for UDP commands on port {UDP_PORT}...")

# === GLOBAL CLIENT ADDRESS ===
client_address = None

def start_discovery_server():
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.bind(("0.0.0.0", DISCOVERY_PORT))
    print(f"[SERVER] Discovery service running on UDP port {DISCOVERY_PORT}")

    while True:
        try:
            data, addr = sock.recvfrom(1024)
            if data == DISCOVERY_MESSAGE:
                print(f"[SERVER] Received discovery request from {addr[0]}")
                sock.sendto(RESPONSE_MESSAGE, addr)
        except Exception as e:
            print(f"[SERVER] Error: {e}")

# === BACKGROUND THREAD: SEND SLIDE NUMBERS CONTINUOUSLY ===
# === SLIDE SYNC THREAD ===
def send_slide_number_loop():
    global client_address
    pythoncom.CoInitialize()

    try:
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        time.sleep(1)

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
                        print("[SERVER] No active slideshow.")
                except Exception as e:
                    print(f"[SERVER] Slide send error: {e}")
            time.sleep(3)

    finally:
        pythoncom.CoUninitialize()


# Start slide number sync thread
threading.Thread(target=start_discovery_server, daemon=True).start()
threading.Thread(target=send_slide_number_loop, daemon=True).start()



# === MAIN LOOP: RECEIVE COMMANDS ===

try:
    while True:
        try:
            data, addr = sock.recvfrom(BUFFER_SIZE)
            message = data.decode().strip().upper()
            client_address = addr
            print(f"[SERVER] Received '{message}' from {addr}")

            if message == 'NEXT':
                keyboard.press_and_release('right')

            elif message == 'PREV':
                keyboard.press_and_release('left')

            elif message == 'EXIT':
                keyboard.press_and_release('esc')
                print("[SERVER] Exit command received.")
                break

        except socket.timeout:
            continue  # just loop and check again

except KeyboardInterrupt:
    print("\n[SERVER] Ctrl+C detected. Shutting down...")

finally:
    sock.close()
    print("[SERVER] Socket closed.")
