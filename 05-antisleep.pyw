# import pyautogui
import time
import threading
import ctypes
import random
import tkinter as tk
# import keyboard


try:
    import pyautogui
    import keyboard
except ImportError as e:
    print(f"Missing package: {e}")
    print("Run: pip install pyautogui keyboard")
    exit()
# --- CONFIG ---
IDLE_THRESHOLD = 60

running = False
last_mouse_pos = pyautogui.position()
last_activity_time = time.time()

# --- Windows API ---
ES_CONTINUOUS = 0x80000000
ES_SYSTEM_REQUIRED = 0x00000001
ES_DISPLAY_REQUIRED = 0x00000002


def prevent_sleep():
    ctypes.windll.kernel32.SetThreadExecutionState(
        ES_CONTINUOUS | ES_SYSTEM_REQUIRED | ES_DISPLAY_REQUIRED
    )


# --- Track activity ---
def track_activity():
    global last_mouse_pos, last_activity_time

    while True:
        pos = pyautogui.position()
        if pos != last_mouse_pos:
            last_mouse_pos = pos
            last_activity_time = time.time()
        time.sleep(2)


# --- Smart keep awake ---
def keep_awake():
    global running

    while True:
        if running:
            idle = time.time() - last_activity_time

            if idle > IDLE_THRESHOLD:
                try:
                    prevent_sleep()

                    # Random movement
                    dx = random.randint(-3, 3)
                    dy = random.randint(-3, 3)
                    pyautogui.moveRel(dx, dy, duration=0.2)

                    # Random key
                    if random.random() > 0.5:
                        pyautogui.press('shift')

                except Exception as e:
                    print("Error:", e)

        time.sleep(random.randint(15, 25))


# --- GUI ---
def start_gui():
    global running

    root = tk.Tk()
    root.title("Anti-Sleep Bot")

    root.geometry("220x120")
    root.attributes('-topmost', True)

    label = tk.Label(root, text="❎ Stopped", font=("Arial", 12))  #❌
    label.pack(pady=10)

    def toggle():
        global running

        if not running:
            running = True
            label.config(text="✅ Running")
            btn_toggle.config(text="Stop")
        else:
            running = False
            ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)
            label.config(text="❎ Stopped")
            btn_toggle.config(text="Start")

    btn_toggle = tk.Button(root, text="Start", width=10, command=toggle)
    btn_toggle.pack(pady=5)

    root.mainloop()


# --- Hotkeys ---
keyboard.add_hotkey("ctrl+alt+s", lambda: print("Hotkey Start") or globals().update(running=True))
keyboard.add_hotkey("ctrl+alt+x", lambda: print("Hotkey Stop") or globals().update(running=False))

# --- Threads ---
threading.Thread(target=track_activity, daemon=True).start()
threading.Thread(target=keep_awake, daemon=True).start()

# --- Launch GUI ---
start_gui()