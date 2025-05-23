import tkinter as tk

class FloatingControl:
    def __init__(self, master, on_prev, on_next, on_close):
        self.master = master
        self.master.overrideredirect(True)  # Frameless window
        self.master.attributes("-topmost", True)
        self.master.geometry("+1200+50")  # Adjust screen position

        self.master.bind("<ButtonPress-1>", self.start_move)
        self.master.bind("<B1-Motion>", self.do_move)

        self.frame = tk.Frame(master, bg="lightgray", bd=1, relief="solid")
        self.frame.pack()

        self.btn_prev = self.create_button("←", on_prev)
        self.btn_prev.grid(row=0, column=0, padx=2, pady=2)

        self.btn_next = self.create_button("→", on_next)
        self.btn_next.grid(row=0, column=1, padx=2, pady=2)

        self.btn_close = self.create_button("X", on_close, fg="red", width=2)
        self.btn_close.grid(row=0, column=2, padx=2, pady=2)

    def create_button(self, text, command, fg="black", width=4):
        btn = tk.Button(self.frame, text=text, width=width, fg=fg, bg="white",
                        activebackground="#cccccc", command=command, relief="flat")
        btn.bind("<Enter>", lambda e: btn.config(bg="#e0e0e0"))
        btn.bind("<Leave>", lambda e: btn.config(bg="white"))
        btn.bind("<ButtonPress>", lambda e: btn.config(bg="#4a90e2"))  # blue on press
        btn.bind("<ButtonRelease>", lambda e: btn.config(bg="#e0e0e0"))
        return btn

    def start_move(self, event):
        self.x = event.x
        self.y = event.y

    def do_move(self, event):
        dx = event.x - self.x
        dy = event.y - self.y
        x = self.master.winfo_x() + dx
        y = self.master.winfo_y() + dy
        self.master.geometry(f"+{x}+{y}")
