import socket
import win32com.client
import pythoncom
import keyboard

# === CONFIG ===
LISTEN_IP = '0.0.0.0'
LISTEN_PORT = 505

# === Initialize COM for PowerPoint control ===
pythoncom.CoInitialize()
ppt = win32com.client.Dispatch("PowerPoint.Application")

# === UDP Server Setup ===
sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
sock.bind((LISTEN_IP, LISTEN_PORT))

print(f"Listening for slide commands on UDP {LISTEN_IP}:{LISTEN_PORT}...")

while True:
    try:
        data, addr = sock.recvfrom(1024)
        command = data.decode().strip().upper()

        if command == "NEXT":
            print("Received NEXT -> Simulating Right Arrow")
            keyboard.press_and_release('right')
        elif command == "PREV":
            print("Received PREV -> Simulating Left Arrow")
            keyboard.press_and_release('left')
        else:
            print(f"Unknown command from {addr}: {command}")
    except KeyboardInterrupt:
        print("Server stopped.")
        break
    except Exception as e:
        print(f"Error: {e}")
