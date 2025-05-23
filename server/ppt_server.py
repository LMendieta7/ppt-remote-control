import socket
import threading
import time
import queue
import pythoncom
import win32com.client

UDP_PORT = 5005
DISCOVERY_PORT = 5001
DISCOVERY_MESSAGE = b"DISCOVER_PPT_SERVER"
RESPONSE_MESSAGE = b"PPT_SERVER_HERE"
BUFFER_SIZE = 1024

sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
sock.bind(('', UDP_PORT))
sock.settimeout(1.0)

client_address = None
current_slide = 0
slide_tracker_queue = queue.Queue()

print(f"[SERVER] Listening for UDP commands on port {UDP_PORT}...")

# Background thread to track slide changes
def track_slide_loop():
    while True:
        slide_tracker_queue.put("check_slide")
        time.sleep(0.5)

threading.Thread(target=track_slide_loop, daemon=True).start()

# Discovery service
def start_discovery_server():
    disc_sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    disc_sock.bind(("0.0.0.0", DISCOVERY_PORT))
    print(f"[SERVER] Discovery service running on UDP port {DISCOVERY_PORT}")
    while True:
        try:
            data, addr = disc_sock.recvfrom(1024)
            if data == DISCOVERY_MESSAGE:
                print(f"[SERVER] Discovery request from {addr[0]}")
                disc_sock.sendto(RESPONSE_MESSAGE, addr)
        except Exception as e:
            print(f"[SERVER] Discovery error: {e}")

threading.Thread(target=start_discovery_server, daemon=True).start()

# Main COM thread
try:
    pythoncom.CoInitialize()
    ppt = win32com.client.Dispatch("PowerPoint.Application")

    while True:
        # Slide tracking (manual updates)
        while not slide_tracker_queue.empty():
            _ = slide_tracker_queue.get()
            try:
                if ppt.SlideShowWindows.Count > 0:
                    view = ppt.SlideShowWindows(1).View
                    slide_pos = view.CurrentShowPosition
                    if slide_pos != current_slide:
                        current_slide = slide_pos
                        print(f"[SERVER] Manual slide change detected: {current_slide}")
            except Exception as e:
                print(f"[SERVER] Slide tracking error: {e}")

        # Handle incoming commands
        try:
            data, addr = sock.recvfrom(BUFFER_SIZE)
            message = data.decode().strip().upper()
            client_address = addr
            print(f"[SERVER] Received '{message}' from {addr}")

            if message == 'NEXT':
                view = ppt.SlideShowWindows(1).View
                view.Next()
                time.sleep(0.3)
                current_slide = view.CurrentShowPosition

            elif message == 'PREV':
                view = ppt.SlideShowWindows(1).View
                view.Previous()
                time.sleep(0.3)
                current_slide = view.CurrentShowPosition

            elif message == 'GET_SLIDE':
                if ppt.SlideShowWindows.Count > 0:
                    view = ppt.SlideShowWindows(1).View
                    current_slide = view.CurrentShowPosition
                    response = f"SLIDE:{current_slide}"
                else:
                    response = "NOSHOW"

                if client_address:
                    sock.sendto(response.encode(), client_address)

        except socket.timeout:
            continue

except KeyboardInterrupt:
    print("\n[SERVER] Shutting down.")

finally:
    sock.close()
    print("[SERVER] Socket closed.")
