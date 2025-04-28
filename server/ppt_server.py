import socket
import threading
import pyautogui
from flask import Flask, request, render_template_string, jsonify

# --- UDP Listener ---
UDP_IP = "0.0.0.0"
UDP_PORT = 5050

def udp_listener():
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.bind((UDP_IP, UDP_PORT))
    print(f"Listening for UDP commands on port {UDP_PORT}...")
    while True:
        data, addr = sock.recvfrom(1024)
        cmd = data.decode().strip().upper()
        print(f"UDP command: {cmd} from {addr}")
        if cmd == "NEXT":
            pyautogui.press('right')
        elif cmd == "PREV":
            pyautogui.press('left')
        elif cmd == "PING":
            sock.sendto(b"PONG", addr)

udp_thread = threading.Thread(target=udp_listener, daemon=True)
udp_thread.start()

# --- Flask Web App ---
app = Flask(__name__)

HTML = '''
<!doctype html>
<html>
<head>
    <title>PPT Remote</title>
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <style>
    body { background: #e0e7ef; font-family: Arial, sans-serif; height:100vh; margin:0; }
    .remote-box {
        background: #f1f5f9;
        margin: 70px auto;
        max-width: 320px;
        border-radius: 28px;
        box-shadow: 0 8px 30px rgba(0,0,0,0.10);
        display: flex;
        flex-direction: column;
        align-items: center;
        padding: 32px 0 28px 0;
    }
    .remote-title {
        font-size: 1.5em;
        color: #22223b;
        margin-bottom: 22px;
        font-weight: bold;
        letter-spacing: 1px;
    }
    .remote-row {
        display: flex;
        flex-direction: row;
        gap: 22px;
        margin: 0 0 0 0;
    }
    .remote-btn {
        font-size: 1.4em;
        border: none;
        border-radius: 16px;
        background: #2563eb;
        color: #fff;
        padding: 24px 34px;
        margin: 0 0 0 0;
        box-shadow: 0 2px 10px rgba(37,99,235,0.10);
        transition: background 0.12s;
        font-weight: 600;
    }
    .remote-btn:active {
        background: #1e40af;
    }
    </style>
</head>
<body>
  <div class="remote-box">
    <div class="remote-title">PPT Remote</div>
    <div class="remote-row">
      <button class="remote-btn" onclick="fetch('/prev', {method:'POST'})">&#8592; Prev</button>
      <button class="remote-btn" onclick="fetch('/next', {method:'POST'})">Next &#8594;</button>
    </div>
  </div>
</body>
</html>
'''

@app.route("/")
def home():
    return render_template_string(HTML)

@app.route("/next", methods=["POST"])
def next_slide():
    pyautogui.press('right')
    return jsonify(ok=True)

@app.route("/prev", methods=["POST"])
def prev_slide():
    pyautogui.press('left')
    return jsonify(ok=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)