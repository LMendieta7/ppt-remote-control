import tkinter as tk

class FloatingControl:
    def __init__(self, master, on_prev, on_next, on_close):
        self.master = master
        self.master.overrideredirect(True)
        self.master.attributes("-topmost", True)
        self.master.geometry("+1200+50")

        self.master.bind("<ButtonPress-1>", self.start_move)
        self.master.bind("<B1-Motion>", self.do_move)

        self.frame = tk.Frame(master, bg="lightgray", bd=1, relief="solid", padx=10, pady=5)
        self.frame.pack()

        self.btn_prev = self.create_button("←", on_prev)
        self.btn_prev.grid(row=0, column=0, padx=(5, 10), pady=2)

        self.btn_next = self.create_button("→", on_next)
        self.btn_next.grid(row=0, column=1, padx=(5, 20), pady=2)

        self.btn_close = self.create_button("X", on_close, fg="red", width=3)
        self.btn_close.grid(row=0, column=2, padx=(20, 5), pady=2)

    def create_button(self, text, command, fg="black", width=4):
        btn = tk.Button(self.frame, text=text, width=width, fg=fg, bg="white",
                        activebackground="#4a90e2", command=command, relief="flat", font=("Arial", 12))

        # State tracking
        def on_enter(e): btn.config(bg="#cce6ff")
        def on_leave(e): btn.config(bg="white")
        def on_press(e): btn.config(bg="#4a90e2", fg="white")
        def on_release(e): btn.config(bg="#cce6ff", fg=fg)

        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        btn.bind("<ButtonPress-1>", on_press)
        btn.bind("<ButtonRelease-1>", on_release)

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
