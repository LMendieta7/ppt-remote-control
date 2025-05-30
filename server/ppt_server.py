import socket
import threading
import time
import queue
import pythoncom
import win32com.client
import web_server  # ✅ Import the web app module
import qrcode

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

def track_slide_loop():
    while True:
        slide_tracker_queue.put("check_slide")
        time.sleep(0.2)

threading.Thread(target=track_slide_loop, daemon=True).start()

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

def print_qr_to_terminal(port=8080):
    try:
        # Get local IP
        hostname = socket.gethostname()
        local_ip = socket.gethostbyname(hostname)
        url = f"http://{local_ip}:{port}"

        # Generate and print QR code in terminal
        qr = qrcode.QRCode()
        qr.add_data(url)
        qr.make()
        print(f"\n[SERVER] Web remote available at: {url}")
        print("[SERVER] Scan this QR code to open on your phone:")
        print()
        print()
        qr.print_ascii()
        print()
    except Exception as e:
        print(f"[QR ERROR] Could not generate QR: {e}")

# ✅ Start web server in background
web_server.run()
print_qr_to_terminal(port=8080)

try:
    pythoncom.CoInitialize()
    ppt = win32com.client.Dispatch("PowerPoint.Application")

    while True:
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

        try:
            data, addr = sock.recvfrom(BUFFER_SIZE)
            message = data.decode().strip().upper()
            client_address = addr
            print(f"[SERVER] Received '{message}' from {addr}")

            if ppt.SlideShowWindows.Count == 0:
                continue

            view = ppt.SlideShowWindows(1).View

            if message == 'NEXT':
                view.Next()
                sock.sendto(b"ACK:NEXT", client_address)

            elif message == 'PREV':
                view.Previous()
                sock.sendto(b"ACK:PREV", client_address)

            elif message == 'GET_SLIDE':
                current_slide = view.CurrentShowPosition
                response = f"SLIDE:{current_slide}"
                sock.sendto(response.encode(), client_address)

        except socket.timeout:
            continue

except KeyboardInterrupt:
    print("\n[SERVER] Shutting down.")

finally:
    sock.close()
    print("[SERVER] Socket closed.")