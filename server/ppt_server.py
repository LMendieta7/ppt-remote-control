import socket
import threading
import pyautogui
import time
import win32com.client

# --- CONFIG ---
UDP_IP = "0.0.0.0"
UDP_PORT = 5051
last_known_client = None

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
    global last_known_client
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
                last_known_client = addr

            elif cmd == "PREV":
                pyautogui.press('left')
                last_known_client = addr

            elif cmd == "GET_SLIDE":
                index = get_slide_index()
                if index:
                    sock.sendto(f"SLIDE:{index}".encode(), addr)
                    last_known_client = addr
        except Exception as e:
            print(f"[SERVER ERROR] {e}")

# --- MAIN ---
if __name__ == "__main__":
    threading.Thread(target=udp_listener, daemon=True).start()
    print("[SERVER] Ready.")
    while True:
        time.sleep(1)
