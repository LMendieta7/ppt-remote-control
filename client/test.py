import keyboard

keyboard.add_hotkey('left', lambda: print("LEFT pressed"))
keyboard.add_hotkey('right', lambda: print("RIGHT pressed"))

print("Waiting for arrow keys. Press ESC to quit.")
keyboard.wait('esc')
