import socket
import keyboard
import time
import win32com.client

# === CONFIG ===
SERVER_IP = '192.168.1.100'  # Change to your server IP
SERVER_PORT = 505

# === Setup PowerPoint ===
ppt = win32com.client.Dispatch("PowerPoint.Application")

def get_local_slide():
    try:
        return ppt.SlideShowWindows(1).View.CurrentShowPosition
    except:
        return -1  # Not in slideshow mode

def sync_slide(target_slide):
    current = get_local_slide()
    if current == -1:
        print("PowerPoint not in slideshow mode.")
        return

    while current < target_slide:
        keyboard.press_and_release('right')
        time.sleep(0.2)
        current = get_local_slide()
    while current > target_slide:
        keyboard.press_and_release('left')
        time.sleep(0.2)
        current = get_local_slide()

    print(f"Synced to slide {current}")

def send_command(command):
    with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
        sock.settimeout(1.0)
        sock.sendto(command.encode(), (SERVER_IP, SERVER_PORT))
        try:
            response, _ = sock.recvfrom(1024)
            target_slide = int(response.decode().strip())
            print(f"Server says current slide: {target_slide}")
            sync_slide(target_slide)
        except socket.timeout:
            print("No response from server.")
        except Exception as e:
            print(f"Error syncing: {e}")

print("Ready. Press LEFT/RIGHT to control slides. ESC to quit.")

while True:
    try:
        if keyboard.is_pressed('right'):
            send_command("NEXT")
            time.sleep(0.3)
        elif keyboard.is_pressed('left'):
            send_command("PREV")
            time.sleep(0.3)
        elif keyboard.is_pressed('esc'):
            print("Exiting...")
            break
    except:
        break
