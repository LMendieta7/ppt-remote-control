# discovery_helper.py

import socket
import time

DISCOVERY_PORT = 5001
DISCOVERY_MESSAGE = b"DISCOVER_PPT_SERVER"
RESPONSE_MESSAGE = b"PPT_SERVER_HERE"

def get_local_subnet():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(('8.8.8.8', 80))
        local_ip = s.getsockname()[0]
    finally:
        s.close()
    return '.'.join(local_ip.split('.')[:3])

def get_server_ip():
    subnet = get_local_subnet()
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.settimeout(0.3)

    for i in range(1, 255):
        ip = f"{subnet}.{i}"
        try:
            sock.sendto(DISCOVERY_MESSAGE, (ip, DISCOVERY_PORT))
            data, addr = sock.recvfrom(1024)
            if data == RESPONSE_MESSAGE:
                return addr[0]
        except:
            continue
    return None

def wait_for_server(retry_delay=2):
    while True:
        ip = get_server_ip()
        if ip:
            print(f"[DISCOVERY] Server found at {ip}")
            return ip
        print("[DISCOVERY] Server not found. Retrying...")
        time.sleep(retry_delay)
