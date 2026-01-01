import keyboard
import subprocess
import ctypes
from ctypes import wintypes
import sys
import threading
import json
import os
import time
import tkinter as tk
from tkinter import ttk

# Direct SendInput for fastest key press
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
VK_E = 0x45

class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong))
    ]

class INPUT(ctypes.Structure):
    _fields_ = [
        ("type", wintypes.DWORD),
        ("ki", KEYBDINPUT),
        ("padding", ctypes.c_ubyte * 8)
    ]

def fast_press_e():
    """Press and release E using direct SendInput - fastest method"""
    inputs = (INPUT * 2)()
    inputs[0].type = INPUT_KEYBOARD
    inputs[0].ki.wVk = VK_E
    inputs[0].ki.dwFlags = 0
    inputs[1].type = INPUT_KEYBOARD
    inputs[1].ki.wVk = VK_E
    inputs[1].ki.dwFlags = KEYEVENTF_KEYUP
    ctypes.windll.user32.SendInput(2, ctypes.byref(inputs), ctypes.sizeof(INPUT))

def precise_sleep(seconds):
    """High-precision sleep using busy-wait for short durations"""
    if seconds <= 0:
        return
    end = time.perf_counter() + seconds
    while time.perf_counter() < end:
        pass

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# Use AppData for config
CONFIG_FILE = os.path.join(os.environ.get('APPDATA', os.path.dirname(os.path.abspath(__file__))), "QuickDC", "config.json")

def load_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    return {
        "hotkey": "alt+\\",
        "mode": "wifi",
        "auto_reconnect": True,
        "duration_ms": 5000,
        "spam_enabled": True,
        "spam_interval_ms": 50,
        "spam_before_ms": 100,
        "spam_after_ms": 3000
    }

def save_config(config):
    os.makedirs(os.path.dirname(CONFIG_FILE), exist_ok=True)
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f)

class QuickDCApp:
    def __init__(self, root):
        self.root = root
        self.root.title("QuickDC")
        self.root.geometry("340x320")
        self.root.resizable(False, False)

        self.config = load_config()
        self.hotkey_registered = None
        self.recording = False
        self.disconnected = False
        self.ssid = None

        row = 0

        # Hotkey row
        ttk.Label(root, text="Hotkey:").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        self.hotkey_var = tk.StringVar(value=self.config["hotkey"])
        self.hotkey_entry = ttk.Entry(root, textvariable=self.hotkey_var, width=15, state="readonly")
        self.hotkey_entry.grid(row=row, column=1, padx=5, pady=5)
        self.record_btn = ttk.Button(root, text="Set", width=6, command=self.start_recording)
        self.record_btn.grid(row=row, column=2, padx=5, pady=5)
        row += 1

        # Mode selection row (WiFi vs Ethernet)
        ttk.Label(root, text="Mode:").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        self.mode_var = tk.StringVar(value=self.config.get("mode", "wifi"))
        mode_frame = ttk.Frame(root)
        mode_frame.grid(row=row, column=1, columnspan=2, sticky="w")
        ttk.Radiobutton(mode_frame, text="WiFi", variable=self.mode_var, value="wifi").pack(side="left")
        ttk.Radiobutton(mode_frame, text="Ethernet", variable=self.mode_var, value="firewall").pack(side="left")
        row += 1

        # Auto-reconnect checkbox + duration
        self.auto_reconnect_var = tk.BooleanVar(value=self.config.get("auto_reconnect", True))
        auto_frame = ttk.Frame(root)
        auto_frame.grid(row=row, column=0, columnspan=3, padx=10, pady=5, sticky="w")
        ttk.Checkbutton(auto_frame, text="Auto-reconnect after", variable=self.auto_reconnect_var, command=self.on_auto_reconnect_change).pack(side="left")
        self.duration_var = tk.StringVar(value=str(self.config.get("duration_ms", 5000)))
        self.duration_entry = ttk.Entry(auto_frame, textvariable=self.duration_var, width=8)
        self.duration_entry.pack(side="left", padx=5)
        ttk.Label(auto_frame, text="ms").pack(side="left")
        row += 1

        # Spam E checkbox
        self.spam_enabled_var = tk.BooleanVar(value=self.config.get("spam_enabled", True))
        self.spam_checkbox = ttk.Checkbutton(root, text="Spam E", variable=self.spam_enabled_var)
        self.spam_checkbox.grid(row=row, column=0, columnspan=2, padx=10, pady=5, sticky="w")
        row += 1

        # Spam interval
        ttk.Label(root, text="E interval (ms):").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        self.spam_interval_var = tk.StringVar(value=str(self.config.get("spam_interval_ms", 50)))
        ttk.Entry(root, textvariable=self.spam_interval_var, width=15).grid(row=row, column=1, padx=5, pady=5)
        row += 1

        # Spam before
        ttk.Label(root, text="E before (ms):").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        self.spam_before_var = tk.StringVar(value=str(self.config.get("spam_before_ms", 100)))
        self.spam_before_entry = ttk.Entry(root, textvariable=self.spam_before_var, width=15)
        self.spam_before_entry.grid(row=row, column=1, padx=5, pady=5)
        row += 1

        # Spam after
        ttk.Label(root, text="E after (ms):").grid(row=row, column=0, padx=10, pady=5, sticky="w")
        self.spam_after_var = tk.StringVar(value=str(self.config.get("spam_after_ms", 3000)))
        ttk.Entry(root, textvariable=self.spam_after_var, width=15).grid(row=row, column=1, padx=5, pady=5)
        row += 1

        # Status
        self.status_var = tk.StringVar(value="Ready")
        self.status_label = ttk.Label(root, textvariable=self.status_var, foreground="gray")
        self.status_label.grid(row=row, column=0, columnspan=3, pady=5)
        row += 1

        # Start button
        self.start_btn = ttk.Button(root, text="Start", command=self.toggle_listening)
        self.start_btn.grid(row=row, column=0, columnspan=3, pady=8)

        self.listening = False
        self.on_auto_reconnect_change()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_auto_reconnect_change(self):
        if self.auto_reconnect_var.get():
            self.duration_entry.config(state="normal")
            self.spam_before_entry.config(state="normal")
        else:
            self.duration_entry.config(state="disabled")
            self.spam_before_entry.config(state="disabled")

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
        if event.state & 0x4:
            parts.append("ctrl")
        if event.state & 0x20000 or event.state & 0x8:
            parts.append("alt")
        if event.state & 0x1:
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
        self.config = {
            "hotkey": self.hotkey_var.get(),
            "mode": self.mode_var.get(),
            "auto_reconnect": self.auto_reconnect_var.get(),
            "duration_ms": int(self.duration_var.get()) if self.duration_var.get().isdigit() else 5000,
            "spam_enabled": self.spam_enabled_var.get(),
            "spam_interval_ms": int(self.spam_interval_var.get()) if self.spam_interval_var.get().isdigit() else 50,
            "spam_before_ms": int(self.spam_before_var.get()) if self.spam_before_var.get().isdigit() else 100,
            "spam_after_ms": int(self.spam_after_var.get()) if self.spam_after_var.get().isdigit() else 3000
        }
        save_config(self.config)

    def cleanup_firewall_rules(self):
        subprocess.run(["netsh", "advfirewall", "firewall", "delete", "rule", "name=QuickDC_Block"],
                       creationflags=subprocess.CREATE_NO_WINDOW, capture_output=True)
        subprocess.run(["netsh", "advfirewall", "firewall", "delete", "rule", "name=QuickDC_Block_In"],
                       creationflags=subprocess.CREATE_NO_WINDOW, capture_output=True)

    def on_hotkey(self):
        if self.disconnected:
            self.do_reconnect()
        else:
            self.do_disconnect()

    def do_disconnect(self):
        self.status_var.set("Disconnected")
        self.status_label.config(foreground="orange")
        self.root.update()

        mode = self.mode_var.get()

        if mode == "wifi":
            # Get SSID first
            result = subprocess.run(
                ["netsh", "wlan", "show", "interfaces"],
                capture_output=True, text=True, creationflags=subprocess.CREATE_NO_WINDOW
            )
            for line in result.stdout.splitlines():
                if "SSID" in line and "BSSID" not in line:
                    self.ssid = line.split(":", 1)[1].strip()
                    break
            subprocess.run(["netsh", "wlan", "disconnect"], creationflags=subprocess.CREATE_NO_WINDOW)
        else:
            # Firewall block
            subprocess.run(["netsh", "advfirewall", "firewall", "add", "rule",
                           "name=QuickDC_Block", "dir=out", "action=block", "enable=yes"],
                          creationflags=subprocess.CREATE_NO_WINDOW)
            subprocess.run(["netsh", "advfirewall", "firewall", "add", "rule",
                           "name=QuickDC_Block_In", "dir=in", "action=block", "enable=yes"],
                          creationflags=subprocess.CREATE_NO_WINDOW)

        self.disconnected = True

        if self.auto_reconnect_var.get():
            threading.Thread(target=self.auto_reconnect_thread, daemon=True).start()

    def auto_reconnect_thread(self):
        duration_ms = int(self.duration_var.get()) if self.duration_var.get().isdigit() else 5000
        spam_enabled = self.spam_enabled_var.get()
        spam_interval = (int(self.spam_interval_var.get()) if self.spam_interval_var.get().isdigit() else 50) / 1000.0
        spam_before = (int(self.spam_before_var.get()) if self.spam_before_var.get().isdigit() else 100) / 1000.0
        spam_after = (int(self.spam_after_var.get()) if self.spam_after_var.get().isdigit() else 3000) / 1000.0

        if spam_enabled:
            wait_time = (duration_ms / 1000.0) - spam_before
            if wait_time > 0:
                time.sleep(wait_time)
            # Spam before
            end = time.perf_counter() + spam_before
            while time.perf_counter() < end:
                fast_press_e()
                precise_sleep(spam_interval)
            fast_press_e()
        else:
            time.sleep(duration_ms / 1000.0)

        self.do_reconnect_internal()

        if spam_enabled:
            fast_press_e()
            end = time.perf_counter() + spam_after
            while time.perf_counter() < end:
                fast_press_e()
                precise_sleep(spam_interval)

        self.status_var.set("Listening...")
        self.status_label.config(foreground="green")

    def do_reconnect(self):
        threading.Thread(target=self.manual_reconnect_thread, daemon=True).start()

    def manual_reconnect_thread(self):
        spam_enabled = self.spam_enabled_var.get()
        spam_interval = (int(self.spam_interval_var.get()) if self.spam_interval_var.get().isdigit() else 50) / 1000.0
        spam_after = (int(self.spam_after_var.get()) if self.spam_after_var.get().isdigit() else 3000) / 1000.0

        self.do_reconnect_internal()

        if spam_enabled:
            fast_press_e()
            end = time.perf_counter() + spam_after
            while time.perf_counter() < end:
                fast_press_e()
                precise_sleep(spam_interval)

        self.status_var.set("Listening...")
        self.status_label.config(foreground="green")

    def do_reconnect_internal(self):
        mode = self.mode_var.get()

        if mode == "wifi":
            if self.ssid:
                subprocess.run(["netsh", "wlan", "connect", f"name={self.ssid}"],
                              creationflags=subprocess.CREATE_NO_WINDOW)
        else:
            self.cleanup_firewall_rules()

        self.disconnected = False

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
            self.hotkey_registered = keyboard.add_hotkey(self.hotkey_var.get(), self.on_hotkey)
            self.listening = True
            self.start_btn.config(text="Stop")
            self.status_var.set("Listening...")
            self.status_label.config(foreground="green")

    def on_close(self):
        self.save_current_config()
        self.cleanup_firewall_rules()
        if self.hotkey_registered:
            keyboard.remove_hotkey(self.hotkey_registered)
        self.root.destroy()

if __name__ == "__main__":
    if not is_admin():
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
        sys.exit()

    root = tk.Tk()
    app = QuickDCApp(root)
    root.mainloop()
