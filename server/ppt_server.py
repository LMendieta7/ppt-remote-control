import socket
import win32com.client
import pythoncom
import keyboard
import time

# === CONFIG ===
LISTEN_IP = '0.0.0.0'
LISTEN_PORT = 505

# === Initialize COM and PowerPoint ===
pythoncom.CoInitialize()
ppt = win32com.client.Dispatch("PowerPoint.Application")

def get_current_slide():
    try:
        return ppt.SlideShowWindows(1).View.CurrentShowPosition
    except:
        return -1  # Not in slideshow

# === Set up UDP server ===
sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
sock.bind((LISTEN_IP, LISTEN_PORT))

print(f"Listening for commands on {LISTEN_IP}:{LISTEN_PORT}...")

while True:
    try:
        data, addr = sock.recvfrom(1024)
        command = data.decode().strip().upper()

        if command == "NEXT":
            keyboard.press_and_release('right')
        elif command == "PREV":
            keyboard.press_and_release('left')
        else:
            print(f"Unknown command from {addr}: {command}")
            continue

        # Wait briefly to let slide change register
        time.sleep(0.2)
        slide = get_current_slide()
        response = str(slide)
        sock.sendto(response.encode(), addr)
        print(f"{command} → slide {slide} sent to {addr}")

    except KeyboardInterrupt:
        print("Server stopped by user.")
        break
    except Exception as e:
        print(f"Error: {e}")
