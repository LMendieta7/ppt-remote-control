import socket
import threading
import pyautogui
import time
import win32com.client

# --- CONFIG ---
UDP_IP = "0.0.0.0"
UDP_PORT = 5051
SYNC_PORT = 6060
last_known_client_ip = None

# --- PowerPoint COM ---
ppt = win32com.client.Dispatch("PowerPoint.Application")
ppt.Visible = True

def get_slide_index():
    try:
        return ppt.SlideShowWindows(1).View.CurrentShowPosition
    except:
        return None

# --- UDP Listener ---
def udp_listener():
    global last_known_client_ip
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.bind((UDP_IP, UDP_PORT))
    print(f"[SERVER] Listening on {UDP_PORT}...")

    while True:
        try:
            data, addr = sock.recvfrom(1024)
            cmd = data.decode().strip().upper()
            print(f"[CMD] {cmd} from {addr}")

            if cmd == "NEXT":
                pyautogui.press('right')
                last_known_client_ip = addr[0]

            elif cmd == "PREV":
                pyautogui.press('left')
                last_known_client_ip = addr[0]

            elif cmd == "GET_SLIDE":
                index = get_slide_index()
                if index:
                    sock.sendto(f"SLIDE:{index}".encode(), addr)
                    last_known_client_ip = addr[0]
        except Exception as e:
            print(f"[SERVER ERROR] {e}")

# --- Slide Sync Broadcaster ---
def broadcast_slide_changes():
    global last_known_client_ip
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    last_slide = None

    while True:
        try:
            current = get_slide_index()
            if current and current != last_slide:
                last_slide = current
                if last_known_client_ip:
                    msg = f"SLIDE:{current}"
                    sock.sendto(msg.encode(), (last_known_client_ip, SYNC_PORT))
                    print(f"[BROADCAST] {msg} → {last_known_client_ip}:{SYNC_PORT}")
        except Exception as e:
            print(f"[SYNC ERROR] {e}")
        time.sleep(0.5)

# --- MAIN ---
if __name__ == "__main__":
    threading.Thread(target=udp_listener, daemon=True).start()
    threading.Thread(target=broadcast_slide_changes, daemon=True).start()
    print("[SERVER] Ready.")
    while True:
        time.sleep(1)
