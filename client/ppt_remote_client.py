import socket
import keyboard
import time
import win32com.client

# === CONFIG ===
SERVER_IP = '192.168.1.100'  # Update with server's IP
SERVER_PORT = 505

# === Setup ===
ppt = win32com.client.Dispatch("PowerPoint.Application")

# === UDP Communication ===
def send_command(command):
    with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
        sock.sendto(command.encode(), (SERVER_IP, SERVER_PORT))

# === Main Loop ===
print("Ready. Press LEFT or RIGHT arrows to control slides. Press ESC to exit.")
while True:
    try:
        if keyboard.is_pressed('right'):
            send_command("NEXT")
            keyboard.press_and_release('right')
            time.sleep(0.3)
        elif keyboard.is_pressed('left'):
            send_command("PREV")
            keyboard.press_and_release('left')
            time.sleep(0.3)
        elif keyboard.is_pressed('esc'):
            print("Exiting...")
            break
    except:
        break
