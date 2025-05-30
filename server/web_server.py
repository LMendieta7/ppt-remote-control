import threading
from flask import Flask, render_template, redirect, url_for, send_file
import os
from PIL import Image
import win32com.client
import pythoncom
from flask import jsonify

# Ensure correct paths to templates and static files
base_dir = os.path.dirname(os.path.abspath(__file__))
template_dir = os.path.join(base_dir, "templates")
static_dir = os.path.join(base_dir, "static")

app = Flask(__name__, template_folder=template_dir, static_folder=static_dir)

def export_current_slide_as_image(output_path=os.path.join(static_dir, "preview.jpg")):
    try:
        pythoncom.CoInitialize()
        ppt = win32com.client.Dispatch("PowerPoint.Application")

        if ppt.SlideShowWindows.Count == 0:
            raise Exception("No slideshow running")

        slide = ppt.SlideShowWindows(1).View.Slide
        temp_path = os.path.join(base_dir, "temp_slide.jpg")
        slide.Export(temp_path, "JPG")
                           
        img = Image.open(temp_path)
        img.thumbnail((480, 360))
        img.save(output_path)
        os.remove(temp_path)
        print("[WEB] Slide preview updated.")

    except Exception as e:
        print(f"[WEB] Using placeholder: {e}")
        placeholder = os.path.join(static_dir, "placeholder.jpg")
        if os.path.exists(placeholder):
            Image.open(placeholder).save(output_path)

@app.route('/slide_info')
def slide_info():
    try:
        pythoncom.CoInitialize()
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        if ppt.SlideShowWindows.Count == 0:
            return jsonify(current=None, total=None)

        view = ppt.SlideShowWindows(1).View
        current = view.CurrentShowPosition
        total = ppt.ActivePresentation.Slides.Count
        return jsonify(current=current, total=total)
    
    except Exception as e:
        print(f"[WEB] Slide info error: {e}")
        return jsonify(current=None, total=None)

@app.route('/')
def index():
    export_current_slide_as_image()
    return render_template("index.html", ts=os.path.getmtime(os.path.join(static_dir, "preview.jpg")))

@app.route('/next')
def next_slide():
    try:
        pythoncom.CoInitialize()
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        if ppt.SlideShowWindows.Count > 0:
            ppt.SlideShowWindows(1).View.Next()
    except Exception as e:
        print(f"[WEB] Error on next: {e}")
    return redirect(url_for('index'))

@app.route('/prev')
def prev_slide():
    try:
        pythoncom.CoInitialize()
        ppt = win32com.client.Dispatch("PowerPoint.Application")
        if ppt.SlideShowWindows.Count > 0:
            ppt.SlideShowWindows(1).View.Previous()
    except Exception as e:
        print(f"[WEB] Error on prev: {e}")
    return redirect(url_for('index'))

@app.route('/preview.jpg')
def serve_preview():
    export_current_slide_as_image()  # Always try to update before serving
    return send_file(os.path.join(static_dir, "preview.jpg"), mimetype='image/jpeg')

def run():
    thread = threading.Thread(target=lambda: app.run(host='0.0.0.0', port=8080), daemon=True)
    thread.start()
    print("[WEB] Flask web server started on port 8080")
