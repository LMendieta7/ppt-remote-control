import socket
import threading
import pyautogui
from flask import Flask, request, render_template_string

# --- CONFIG ---
UDP_IP = "0.0.0.0"
UDP_PORT = 5051

# --- UDP LISTENER FUNCTION ---
def udp_listener():
    sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    sock.bind((UDP_IP, UDP_PORT))
    print(f"[UDP] Listening on port {UDP_PORT}...")

    while True:
        try:
            data, addr = sock.recvfrom(1024)
            cmd = data.decode().strip().upper()
            print(f"[UDP] Received: {cmd} from {addr}")

            if cmd == "NEXT":
                pyautogui.press('right')
            elif cmd == "PREV":
                pyautogui.press('left')
            elif cmd == "GET_SLIDE":
                try:
                    import win32com.client
                    ppt = win32com.client.Dispatch("PowerPoint.Application")
                    slide = ppt.SlideShowWindows(1).View.CurrentShowPosition
                    sock.sendto(f"SLIDE:{slide}".encode(), addr)
                    print(f"[SLIDE] Sent slide number {slide} to {addr}")
                except Exception as e:
                    print(f"[SLIDE ERROR] {e}")
                    sock.sendto(b"SLIDE:0", addr)
        except Exception as e:
            print(f"[UDP ERROR] {e}")

# --- Start UDP Listener in a Thread ---
udp_thread = threading.Thread(target=udp_listener, daemon=True)
udp_thread.start()

# --- SIMPLE FLASK WEB INTERFACE ---
app = Flask(__name__)

HTML = '''
<!doctype html>
<html>
<head>
  <title>PPT Remote</title>
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <style>
    body { background: #f1f5f9; font-family: sans-serif; text-align: center; padding-top: 50px; }
    h1 { font-size: 2em; margin-bottom: 30px; }
    button {
      padding: 16px 32px;
      margin: 16px;
      font-size: 1.4em;
      background: #2563eb;
      color: white;
      border: none;
      border-radius: 10px;
    }
    button:active {
      background: #1e40af;
    }
  </style>
</head>
<body>
  <h1>PPT Remote Control</h1>
  <form method="POST" action="/prev">
    <button>&larr; Prev</button>
  </form>
  <form method="POST" action="/next">
    <button>Next &rarr;</button>
  </form>
</body>
</html>
'''

@app.route("/")
def home():
    return render_template_string(HTML)

@app.route("/next", methods=["POST"])
def next_slide():
    pyautogui.press('right')
    return home()

@app.route("/prev", methods=["POST"])
def prev_slide():
    pyautogui.press('left')
    return home()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
