import keyboard
import subprocess
import ctypes
import sys
import threading
import json
import os
import time
import tkinter as tk
from tkinter import ttk

# Use AppData for config so it persists with exe
CONFIG_FILE = os.path.join(os.environ.get('APPDATA', os.path.dirname(os.path.abspath(__file__))), "QuickDC", "config.json")

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {"hotkey": "alt+\\", "duration_ms": 5000, "spam_interval_ms": 50, "spam_after_ms": 3000, "spam_before_ms": 100, "spam_enabled": True}

def save_config(config):
    os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f)

class QuickDCApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QuickDC")
        self.root.geometry("320x280")
        self.root.resizable(False, False)

        self.config = load_config()
        self.hotkey_registered = None
        self.recording = False

        row = 0

        # Hotkey row
        ttk.Label(root, text="Hotkey:").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        self.hotkey_var = tk.StringVar(value=self.config["hotkey"])
        self.hotkey_entry = ttk.Entry(root, textvariable=self.hotkey_var, width=15, state="readonly")
        self.hotkey_entry.grid(row=row, column=1, padx=5, pady=5)
        self.record_btn = ttk.Button(root, text="Set", width=6, command=self.start_recording)
        self.record_btn.grid(row=row, column=2, padx=5, pady=5)
        row += 1

        # Disconnect duration row (milliseconds)
        ttk.Label(root, text="Disconnect (ms):").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        self.duration_var = tk.StringVar(value=str(self.config.get("duration_ms", 5000)))
        self.duration_entry = ttk.Entry(root, textvariable=self.duration_var, width=15)
        self.duration_entry.grid(row=row, column=1, padx=5, pady=5)
        row += 1

        # Spam E checkbox
        self.spam_enabled_var = tk.BooleanVar(value=self.config.get("spam_enabled", True))
        self.spam_checkbox = ttk.Checkbutton(root, text="Spam E", variable=self.spam_enabled_var)
        self.spam_checkbox.grid(row=row, column=0, columnspan=2, padx=10, pady=5, sticky="w")
        row += 1

        # Spam interval row (milliseconds)
        ttk.Label(root, text="E interval (ms):").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        self.spam_interval_var = tk.StringVar(value=str(self.config.get("spam_interval_ms", 50)))
        self.spam_interval_entry = ttk.Entry(root, textvariable=self.spam_interval_var, width=15)
        self.spam_interval_entry.grid(row=row, column=1, padx=5, pady=5)
        row += 1

        # Spam before reconnect (milliseconds)
        ttk.Label(root, text="E before (ms):").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        self.spam_before_var = tk.StringVar(value=str(self.config.get("spam_before_ms", 100)))
        self.spam_before_entry = ttk.Entry(root, textvariable=self.spam_before_var, width=15)
        self.spam_before_entry.grid(row=row, column=1, padx=5, pady=5)
        row += 1

        # Spam after reconnect duration row (milliseconds)
        ttk.Label(root, text="E after (ms):").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        self.spam_after_var = tk.StringVar(value=str(self.config.get("spam_after_ms", 3000)))
        self.spam_after_entry = ttk.Entry(root, textvariable=self.spam_after_var, width=15)
        self.spam_after_entry.grid(row=row, column=1, padx=5, pady=5)
        row += 1

        # Status
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = ttk.Label(root, textvariable=self.status_var, foreground="green")
        self.status_label.grid(row=row, column=0, columnspan=3, pady=5)
        row += 1

        # Start button
        self.start_btn = ttk.Button(root, text="Start", command=self.toggle_listening)
        self.start_btn.grid(row=row, column=0, columnspan=3, pady=8)

        self.listening = False
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def start_recording(self):
        self.recording = True
        self.record_btn.config(text="...")
        self.hotkey_var.set("Press keys...")
        self.root.bind("<KeyPress>", self.on_key_press)
        self.root.focus_set()

    def on_key_press(self, event):
        if not self.recording:
            return

        parts = []
        if event.state & 0x4:  # Control
            parts.append("ctrl")
        if event.state & 0x20000 or event.state & 0x8:  # Alt
            parts.append("alt")
        if event.state & 0x1:  # Shift
            parts.append("shift")

        key = event.keysym.lower()
        modifier_keys = ['control_l', 'control_r', 'alt_l', 'alt_r', 'shift_l', 'shift_r', 'meta_l', 'meta_r']
        if key not in modifier_keys:
            parts.append(key)
            hotkey = "+".join(parts)
            self.hotkey_var.set(hotkey)
            self.record_btn.config(text="Set")
            self.recording = False
            self.root.unbind("<KeyPress>")
            self.save_current_config()

    def save_current_config(self):
        try:
            duration_ms = int(self.duration_var.get())
        except ValueError:
            duration_ms = 5000
        try:
            spam_interval_ms = int(self.spam_interval_var.get())
        except ValueError:
            spam_interval_ms = 50
        try:
            spam_before_ms = int(self.spam_before_var.get())
        except ValueError:
            spam_before_ms = 100
        try:
            spam_after_ms = int(self.spam_after_var.get())
        except ValueError:
            spam_after_ms = 3000
        self.config = {
            "hotkey": self.hotkey_var.get(),
            "duration_ms": duration_ms,
            "spam_enabled": self.spam_enabled_var.get(),
            "spam_interval_ms": spam_interval_ms,
            "spam_before_ms": spam_before_ms,
            "spam_after_ms": spam_after_ms
        }
        save_config(self.config)

    def reset_wifi(self):
        self.status_var.set("WiFi Off...")
        self.status_label.config(foreground="orange")
        self.root.update()

        def do_reset():
            try:
                duration_ms = int(self.duration_var.get())
            except ValueError:
                duration_ms = 5000

            spam_enabled = self.spam_enabled_var.get()

            try:
                spam_interval_ms = int(self.spam_interval_var.get())
            except ValueError:
                spam_interval_ms = 50
            try:
                spam_before_ms = int(self.spam_before_var.get())
            except ValueError:
                spam_before_ms = 100
            try:
                spam_after_ms = int(self.spam_after_var.get())
            except ValueError:
                spam_after_ms = 3000

            spam_interval = spam_interval_ms / 1000.0
            spam_before = spam_before_ms / 1000.0

            # Get current SSID before disconnecting
            result = subprocess.run(
                ["netsh", "wlan", "show", "interfaces"],
                capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW
            )
            ssid = None
            for line in result.stdout.splitlines():
                if "SSID" in line and "BSSID" not in line:
                    ssid = line.split(":", 1)[1].strip()
                    break

            # Disconnect WiFi
            subprocess.run(
                ["netsh", "wlan", "disconnect"],
                creationflags=subprocess.CREATE_NO_WINDOW
            )

            if spam_enabled:
                # Wait until spam_before_ms before reconnect time
                wait_time = (duration_ms / 1000.0) - spam_before
                if wait_time > 0:
                    time.sleep(wait_time)

                # Spam before reconnect
                pre_reconnect_end = time.time() + spam_before
                while time.time() < pre_reconnect_end:
                    keyboard.press_and_release('e')
                    time.sleep(spam_interval)

                # Hit E right at reconnect moment
                keyboard.press_and_release('e')
            else:
                # Just wait the full duration
                time.sleep(duration_ms / 1000.0)

            # Reconnect to the same network
            if ssid:
                subprocess.run(
                    ["netsh", "wlan", "connect", f"name={ssid}"],
                    creationflags=subprocess.CREATE_NO_WINDOW
                )

            if spam_enabled:
                # Hit E right after reconnect command
                keyboard.press_and_release('e')

                # Keep spamming E for configured time after reconnect
                spam_end = time.time() + (spam_after_ms / 1000.0)
                while time.time() < spam_end:
                    keyboard.press_and_release('e')
                    time.sleep(spam_interval)

            self.status_var.set("Listening...")
            self.status_label.config(foreground="green")

        threading.Thread(target=do_reset, daemon=True).start()

    def toggle_listening(self):
        if self.listening:
            if self.hotkey_registered:
                keyboard.remove_hotkey(self.hotkey_registered)
                self.hotkey_registered = None
            self.listening = False
            self.start_btn.config(text="Start")
            self.status_var.set("Ready")
            self.status_label.config(foreground="gray")
        else:
            self.save_current_config()
            hotkey = self.hotkey_var.get()
            self.hotkey_registered = keyboard.add_hotkey(hotkey, self.reset_wifi)
            self.listening = True
            self.start_btn.config(text="Stop")
            self.status_var.set("Listening...")
            self.status_label.config(foreground="green")

    def on_close(self):
        self.save_current_config()
        if self.hotkey_registered:
            keyboard.remove_hotkey(self.hotkey_registered)
        self.root.destroy()

if __name__ == "__main__":
    if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", sys.executable, " ".join(sys.argv), None, 1
        )
        sys.exit()

    root = tk.Tk()
    app = QuickDCApp(root)
    root.mainloop()
