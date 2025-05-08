import socket
import threading
import pyautogui
import win32com.client

# --- CONFIG ---
UDP_IP = "0.0.0.0"
UDP_PORT = 5051

ppt = win32com.client.Dispatch("PowerPoint.Application")
ppt.Visible = True

def get_current_slide():
    try:
        return ppt.SlideShowWindows(1).View.CurrentShowPosition
    except:
        return 0

def udp_listener():
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.bind((UDP_IP, UDP_PORT))
    print(f"[SERVER] Listening on {UDP_PORT}...")

    while True:
        try:
            data, addr = sock.recvfrom(1024)
            cmd = data.decode().strip().upper()
            print(f"[SERVER] Command: {cmd} from {addr}")

            if cmd == "NEXT":
                pyautogui.press("right")
            elif cmd == "PREV":
                pyautogui.press("left")

            current = get_current_slide()
            sock.sendto(f"SLIDE:{current}".encode(), addr)

        except Exception as e:
            print(f"[SERVER ERROR] {e}")

if __name__ == "__main__":
    threading.Thread(target=udp_listener, daemon=True).start()
    print("[SERVER] Ready.")
    while True:
        pass
