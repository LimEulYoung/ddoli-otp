import ctypes
import json
import os
import sys
import threading
import time
import tkinter as tk
import winreg
from ctypes import wintypes
from tkinter import messagebox
from pathlib import Path

# Set DPI awareness (aligns ImageGrab coordinates with tkinter coordinates)
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    ctypes.windll.user32.SetProcessDPIAware()

from urllib.parse import urlparse, parse_qs, unquote

import pyotp
import pyperclip
import pystray
from pystray._util import win32
import cv2
import numpy as np
from PIL import Image, ImageDraw, ImageFont, ImageGrab

_APP_DIR = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "DdoliOTP"
_APP_DIR.mkdir(parents=True, exist_ok=True)
DATA_FILE = _APP_DIR / "otp_data.json"
AUTORUN_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"
AUTORUN_NAME = "Ddoli OTP"


def get_exe_path():
    """Return the current executable path (exe or pythonw + script)."""
    if getattr(sys, 'frozen', False):
        return f'"{sys.executable}"'
    return f'"{sys.executable}" "{Path(__file__).resolve()}"'


def is_autorun_enabled():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTORUN_KEY, 0, winreg.KEY_READ) as key:
            winreg.QueryValueEx(key, AUTORUN_NAME)
            return True
    except FileNotFoundError:
        return False


def set_autorun(enable):
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, AUTORUN_KEY, 0, winreg.KEY_SET_VALUE) as key:
        if enable:
            winreg.SetValueEx(key, AUTORUN_NAME, 0, winreg.REG_SZ, get_exe_path())
        else:
            try:
                winreg.DeleteValue(key, AUTORUN_NAME)
            except FileNotFoundError:
                pass


def load_data():
    if DATA_FILE.exists():
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def create_icon_image(size=64):
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    s = size / 512.0  # scale factor

    # Rounded rectangle background
    rx = int(112 * s)
    pad = int(16 * s)
    bg_color = (206, 100, 72)
    d.rounded_rectangle([pad, pad, size - pad, size - pad], radius=rx, fill=bg_color)

    cx = size / 2
    cy = size * 260 / 512
    cat_color = (252, 243, 234)
    ear_color = cat_color
    inner_ear = (232, 168, 124, 128)
    eye_color = (61, 43, 31)
    nose_color = (232, 168, 124)
    cheek = (232, 196, 170, 100)

    # Body
    bx, by, brx, bry = cx, cy + 20*s, 110*s, 95*s
    d.ellipse([bx-brx, by-bry, bx+brx, by+bry], fill=cat_color)

    # Head
    hr = 82*s
    d.ellipse([cx-hr, cy-65*s-hr, cx+hr, cy-65*s+hr], fill=cat_color)

    # Ears (left / right)
    d.polygon([(cx-65*s, cy-120*s), (cx-85*s, cy-195*s), (cx-25*s, cy-140*s)], fill=ear_color)
    d.polygon([(cx-60*s, cy-125*s), (cx-78*s, cy-180*s), (cx-32*s, cy-140*s)], fill=inner_ear)
    d.polygon([(cx+65*s, cy-120*s), (cx+85*s, cy-195*s), (cx+25*s, cy-140*s)], fill=ear_color)
    d.polygon([(cx+60*s, cy-125*s), (cx+78*s, cy-180*s), (cx+32*s, cy-140*s)], fill=inner_ear)

    # Cheeks
    cr = 18*s
    d.ellipse([cx-48*s-cr, cy-40*s-cr, cx-48*s+cr, cy-40*s+cr], fill=cheek)
    d.ellipse([cx+48*s-cr, cy-40*s-cr, cx+48*s+cr, cy-40*s+cr], fill=cheek)

    # Eyes
    er = 11*s
    d.ellipse([cx-28*s-er, cy-72*s-er, cx-28*s+er, cy-72*s+er], fill=eye_color)
    d.ellipse([cx+28*s-er, cy-72*s-er, cx+28*s+er, cy-72*s+er], fill=eye_color)

    # Eye highlights
    ehr = 3.5*s
    d.ellipse([cx-24*s-ehr, cy-76*s-ehr, cx-24*s+ehr, cy-76*s+ehr], fill=(255,255,255))
    d.ellipse([cx+32*s-ehr, cy-76*s-ehr, cx+32*s+ehr, cy-76*s+ehr], fill=(255,255,255))

    # Nose
    d.polygon([(cx-5*s, cy-54*s), (cx, cy-48*s), (cx+5*s, cy-54*s)], fill=nose_color)

    # OTP indicator dots (on body)
    for dx, alpha in [(-30, 128), (0, 178), (30, 230)]:
        dr = 9*s
        d.ellipse([cx+dx*s-dr, cy+25*s-dr, cx+dx*s+dr, cy+25*s+dr], fill=(217, 119, 87, alpha))

    # Paws
    d.ellipse([cx-30*s-18*s, cy+112*s-10*s, cx-30*s+18*s, cy+112*s+10*s], fill=cat_color)
    d.ellipse([cx+30*s-18*s, cy+112*s-10*s, cx+30*s+18*s, cy+112*s+10*s], fill=cat_color)

    return img


def generate_otp(secret):
    try:
        return pyotp.TOTP(secret).now()
    except Exception:
        return "??????"


def make_copy_action(name, secret):
    def action():
        pyperclip.copy(generate_otp(secret))
    return action


def parse_otpauth_uri(uri):
    """Parse otpauth://totp/Label?secret=...&issuer=... URI and return (name, secret)."""
    parsed = urlparse(uri)
    if parsed.scheme != "otpauth":
        return None, None
    params = parse_qs(parsed.query)
    secret = params.get("secret", [None])[0]
    if not secret:
        return None, None
    # Name: extract from issuer parameter or path
    issuer = params.get("issuer", [None])[0]
    label = unquote(parsed.path.lstrip("/"))
    if ":" in label:
        name = label.split(":")[0]
    elif issuer:
        name = issuer
    elif label:
        name = label
    else:
        name = "Unknown"
    return name, secret


def start_qr_capture(icon):
    """Drag a screen region to capture a QR code and register OTP."""
    def _capture():
        root = tk.Tk()
        root.attributes("-topmost", True)
        root.overrideredirect(True)

        # Full screen size
        screen_w = root.winfo_screenwidth()
        screen_h = root.winfo_screenheight()
        root.geometry(f"{screen_w}x{screen_h}+0+0")

        # Semi-transparent overlay
        root.attributes("-alpha", 0.3)
        root.configure(bg="black")

        canvas = tk.Canvas(root, bg="black", highlightthickness=0,
                           cursor="crosshair")
        canvas.pack(fill="both", expand=True)

        state = {"start_x": 0, "start_y": 0, "rect": None}

        def on_press(event):
            state["start_x"] = event.x_root
            state["start_y"] = event.y_root

        def on_drag(event):
            if state["rect"]:
                canvas.delete(state["rect"])
            x1 = state["start_x"]
            y1 = state["start_y"]
            x2 = event.x_root
            y2 = event.y_root
            state["rect"] = canvas.create_rectangle(
                x1, y1, x2, y2, outline="red", width=2)

        def on_release(event):
            x1 = min(state["start_x"], event.x_root)
            y1 = min(state["start_y"], event.y_root)
            x2 = max(state["start_x"], event.x_root)
            y2 = max(state["start_y"], event.y_root)

            if x2 - x1 < 10 or y2 - y1 < 10:
                root.destroy()
                return

            # Hide overlay before capture (prevents overlay from covering QR)
            root.withdraw()
            root.update()
            time.sleep(0.15)
            screenshot = ImageGrab.grab(bbox=(x1, y1, x2, y2))
            root.destroy()

            # QR decode (OpenCV)
            img_array = np.array(screenshot)
            img_bgr = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)
            detector = cv2.QRCodeDetector()
            qr_data, _, _ = detector.detectAndDecode(img_bgr)
            if not qr_data:
                msg_root = tk.Tk()
                msg_root.withdraw()
                msg_root.attributes("-topmost", True)
                messagebox.showerror("Error", "Could not recognize QR code.",
                                     parent=msg_root)
                msg_root.destroy()
                return

            name, secret = parse_otpauth_uri(qr_data)

            if not secret:
                msg_root = tk.Tk()
                msg_root.withdraw()
                msg_root.attributes("-topmost", True)
                messagebox.showerror("Error",
                                     "Not a valid OTP QR code.\n"
                                     f"Data: {qr_data}",
                                     parent=msg_root)
                msg_root.destroy()
                return

            # Validate OTP secret
            try:
                pyotp.TOTP(secret).now()
            except Exception:
                msg_root = tk.Tk()
                msg_root.withdraw()
                msg_root.attributes("-topmost", True)
                messagebox.showerror("Error",
                                     "The Secret Key from the QR code is invalid.",
                                     parent=msg_root)
                msg_root.destroy()
                return

            # Register
            data = load_data()
            data.append({"name": name, "secret": secret})
            save_data(data)
            rebuild_menu(icon)

            msg_root = tk.Tk()
            msg_root.withdraw()
            msg_root.attributes("-topmost", True)
            messagebox.showinfo("Done", f"'{name}' OTP has been registered.",
                                parent=msg_root)
            msg_root.destroy()

        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)
        root.bind("<Escape>", lambda e: root.destroy())
        root.mainloop()

    threading.Thread(target=_capture, daemon=True).start()


def show_register_dialog(icon):
    def _dialog():
        root = tk.Tk()
        root.title("Register OTP")
        root.resizable(False, False)
        root.attributes("-topmost", True)
        w, h = 380, 240
        root.geometry(f"{w}x{h}+{(root.winfo_screenwidth()-w)//2}+{(root.winfo_screenheight()-h)//2}")

        frame = tk.Frame(root, padx=20, pady=20)
        frame.pack(fill="both", expand=True)

        tk.Label(frame, text="Name (e.g. GitHub, AWS)").grid(row=0, column=0, sticky="w", pady=(0, 5))
        name_entry = tk.Entry(frame, width=35)
        name_entry.grid(row=1, column=0, sticky="ew", pady=(0, 15))

        tk.Label(frame, text="Secret Key (Base32)").grid(row=2, column=0, sticky="w", pady=(0, 5))
        secret_entry = tk.Entry(frame, width=35)
        secret_entry.grid(row=3, column=0, sticky="ew", pady=(0, 15))

        def on_submit():
            name = name_entry.get().strip()
            secret = secret_entry.get().strip().replace(" ", "")
            if not name:
                messagebox.showwarning("Input Error", "Please enter a name.", parent=root)
                return
            if not secret:
                messagebox.showwarning("Input Error", "Please enter the Secret Key.", parent=root)
                return
            try:
                pyotp.TOTP(secret).now()
            except Exception:
                messagebox.showerror("Error", "Invalid Secret Key.\nPlease check it is in Base32 format.", parent=root)
                return
            data = load_data()
            data.append({"name": name, "secret": secret})
            save_data(data)
            messagebox.showinfo("Done", f"'{name}' OTP has been registered.", parent=root)
            root.destroy()
            rebuild_menu(icon)

        btn_frame = tk.Frame(frame)
        btn_frame.grid(row=4, column=0, sticky="e")
        tk.Button(btn_frame, text="Register", width=10, command=on_submit).pack(side="right")
        tk.Button(btn_frame, text="Cancel", width=10, command=root.destroy).pack(side="right", padx=(0, 5))
        name_entry.focus_set()
        root.bind("<Return>", lambda e: on_submit())
        root.mainloop()

    threading.Thread(target=_dialog, daemon=True).start()


def show_manage_dialog(icon):
    def _dialog():
        data = load_data()
        if not data:
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            messagebox.showinfo("Info", "No OTP entries registered.", parent=root)
            root.destroy()
            return

        root = tk.Tk()
        root.title("Manage OTP")
        root.resizable(False, False)
        root.attributes("-topmost", True)
        w, h = 350, 450
        root.geometry(f"{w}x{h}+{(root.winfo_screenwidth()-w)//2}+{(root.winfo_screenheight()-h)//2}")

        frame = tk.Frame(root, padx=20, pady=20)
        frame.pack(fill="both", expand=True)
        tk.Label(frame, text="Select an item to manage:").pack(anchor="w", pady=(0, 10))
        listbox = tk.Listbox(frame, height=12)
        listbox.pack(fill="both", expand=True, pady=(0, 10))
        for item in data:
            listbox.insert(tk.END, item["name"])

        def get_selected():
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning("Selection Error", "Please select an item.", parent=root)
                return None
            return sel[0]

        def on_rename():
            idx = get_selected()
            if idx is None:
                return
            old_name = data[idx]["name"]

            rename_win = tk.Toplevel(root)
            rename_win.title("Rename")
            rename_win.resizable(False, False)
            rename_win.attributes("-topmost", True)
            rename_win.grab_set()
            rw, rh = 350, 170
            rename_win.geometry(f"{rw}x{rh}+{(root.winfo_screenwidth()-rw)//2}+{(root.winfo_screenheight()-rh)//2}")

            rf = tk.Frame(rename_win, padx=15, pady=15)
            rf.pack(fill="both", expand=True)
            tk.Label(rf, text="New name:").pack(anchor="w")
            name_entry = tk.Entry(rf, width=30)
            name_entry.pack(fill="x", pady=(5, 10))
            name_entry.insert(0, old_name)
            name_entry.select_range(0, tk.END)
            name_entry.focus_set()

            def do_rename():
                new_name = name_entry.get().strip()
                if not new_name:
                    messagebox.showwarning("Input Error", "Please enter a name.", parent=rename_win)
                    return
                data[idx]["name"] = new_name
                save_data(data)
                listbox.delete(idx)
                listbox.insert(idx, new_name)
                rename_win.destroy()
                rebuild_menu(icon)

            rbtn = tk.Frame(rf)
            rbtn.pack(anchor="e")
            tk.Button(rbtn, text="Rename", width=8, command=do_rename).pack(side="right")
            tk.Button(rbtn, text="Cancel", width=8, command=rename_win.destroy).pack(side="right", padx=(0, 5))
            rename_win.bind("<Return>", lambda e: do_rename())

        def on_delete():
            idx = get_selected()
            if idx is None:
                return
            name = data[idx]["name"]
            if messagebox.askyesno("Confirm", f"Delete '{name}' OTP?", parent=root):
                data.pop(idx)
                save_data(data)
                listbox.delete(idx)
                rebuild_menu(icon)

        btn_frame = tk.Frame(frame)
        btn_frame.pack(anchor="e")
        tk.Button(btn_frame, text="Delete", width=10, command=on_delete).pack(side="right")
        tk.Button(btn_frame, text="Rename", width=10, command=on_rename).pack(side="right", padx=(0, 5))
        tk.Button(btn_frame, text="Close", width=10, command=root.destroy).pack(side="right", padx=(0, 5))
        root.mainloop()

    threading.Thread(target=_dialog, daemon=True).start()


def build_menu_items(icon):
    """Build menu item list based on current OTP data."""
    data = load_data()
    items = []

    for entry in data:
        name = entry["name"]
        secret = entry["secret"]
        code = generate_otp(secret)
        remaining = 30 - int(time.time() % 30)
        items.append(pystray.MenuItem(f"{name}  -  {code}  ({remaining}s)", make_copy_action(name, secret)))

    if data:
        items.append(pystray.Menu.SEPARATOR)

    items.append(pystray.MenuItem("Register OTP", pystray.Menu(
        pystray.MenuItem("Scan QR Code", lambda: start_qr_capture(icon)),
        pystray.MenuItem("Enter Secret Key", lambda: show_register_dialog(icon)),
    )))
    items.append(pystray.MenuItem("Manage OTP", lambda: show_manage_dialog(icon)))
    items.append(pystray.Menu.SEPARATOR)
    items.append(pystray.MenuItem(
        "Run at Startup",
        lambda: set_autorun(not is_autorun_enabled()),
        checked=lambda _: is_autorun_enabled(),
    ))
    items.append(pystray.MenuItem("Exit", lambda: icon.stop()))

    return items


def rebuild_menu(icon):
    """Rebuild and replace the menu."""
    icon.menu = pystray.Menu(*build_menu_items(icon))
    icon._update_menu()


def show_popup_menu(icon):
    """Programmatically show the context menu."""
    # Rebuild menu each time to refresh OTP codes
    rebuild_menu(icon)

    if not icon._menu_handle:
        return

    win32.SetForegroundWindow(icon._hwnd)
    point = wintypes.POINT()
    win32.GetCursorPos(ctypes.byref(point))
    hmenu, descriptors = icon._menu_handle
    index = win32.TrackPopupMenuEx(
        hmenu,
        win32.TPM_RIGHTALIGN | win32.TPM_BOTTOMALIGN | win32.TPM_RETURNCMD,
        point.x, point.y,
        icon._menu_hwnd,
        None,
    )
    if index > 0:
        descriptors[index - 1](icon)


def main():
    icon_image = create_icon_image()
    icon = pystray.Icon("ddoli_otp", icon_image, "Ddoli OTP")

    # Initial menu setup
    icon.menu = pystray.Menu(*build_menu_items(icon))

    def setup(icon):
        icon.visible = True

        # Replace WM_NOTIFY handler in _message_handlers to
        # show menu popup on left-click and refresh OTP on right-click
        original_handler = icon._message_handlers[win32.WM_NOTIFY]

        def patched_handler(wparam, lparam):
            if lparam == win32.WM_LBUTTONUP:
                # Left-click: show menu popup
                show_popup_menu(icon)
            elif lparam == win32.WM_RBUTTONUP:
                # Right-click: refresh menu then run original logic
                rebuild_menu(icon)
                return original_handler(wparam, lparam)
            else:
                return original_handler(wparam, lparam)

        icon._message_handlers[win32.WM_NOTIFY] = patched_handler

    icon.run(setup=setup)


if __name__ == "__main__":
    main()
