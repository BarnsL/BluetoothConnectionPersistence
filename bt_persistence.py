"""
Bluetooth Connection Persistence
Automatically reconnects to a user-selected Bluetooth device on disconnection,
startup, and wake from sleep. Runs as a Windows system tray application.
"""

import sys
import os
import json
import time
import ctypes
import logging
import threading
import subprocess
import winreg
from pathlib import Path
from logging.handlers import RotatingFileHandler

import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw
import win32api
import win32con
import win32gui
import win32ts

# ── Constants ──────────────────────────────────────────────────────────────────
APP_NAME = "BluetoothConnectionPersistence"
APP_DISPLAY_NAME = "Bluetooth Connection Persistence"
CONFIG_DIR = Path(os.environ.get("APPDATA", "")) / APP_NAME
CONFIG_FILE = CONFIG_DIR / "config.json"
LOG_FILE = CONFIG_DIR / "service.log"
RECONNECT_INTERVAL = 5  # seconds between reconnect attempts
MAX_RECONNECT_INTERVAL = 60  # max back-off interval
STARTUP_REG_KEY = r"Software\Microsoft\Windows\CurrentVersion\Run"

# ── Logging ────────────────────────────────────────────────────────────────────
CONFIG_DIR.mkdir(parents=True, exist_ok=True)

logger = logging.getLogger(APP_NAME)
logger.setLevel(logging.INFO)
handler = RotatingFileHandler(LOG_FILE, maxBytes=1_000_000, backupCount=3)
handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logger.addHandler(handler)


# ── Configuration ──────────────────────────────────────────────────────────────
def load_config() -> dict:
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            logger.error("Failed to load config: %s", e)
    return {"devices": [], "auto_start": True}


def save_config(cfg: dict):
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)
    logger.info("Configuration saved.")


# ── Bluetooth helpers (uses Windows btcom / PowerShell) ────────────────────────
def get_paired_devices() -> list[dict]:
    """
    Enumerate paired Bluetooth devices via PowerShell.
    Returns deduplicated list of dicts with 'name', 'instance_id', 'status'.
    Filters out transport/service entries to show only actual devices.
    """
    ps_script = (
        "Get-PnpDevice -Class Bluetooth | "
        "Where-Object { $_.FriendlyName -and $_.InstanceId -match 'BTHENUM' } | "
        "Select-Object FriendlyName, InstanceId, Status | "
        "ConvertTo-Json -Compress"
    )
    # Patterns that indicate transport/service entries, not real devices
    SKIP_PATTERNS = {
        "avrcp transport", "avrcp", "handsfree", "a2dp",
        "phonebook access", "service discovery",
        "network nap", "personal area network",
        "serial port", "object push", "dial-up",
        "headset gateway", "audio sink", "audio source",
        "human interface", "hid device",
    }
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
            capture_output=True, text=True, timeout=30,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        if result.returncode != 0:
            logger.warning("PowerShell enumeration failed: %s", result.stderr.strip())
            return []
        raw = result.stdout.strip()
        if not raw:
            return []
        data = json.loads(raw)
        if isinstance(data, dict):
            data = [data]

        # Group by device name, prefer entries with status OK
        seen_names: dict[str, dict] = {}
        for d in data:
            name = d.get("FriendlyName", "Unknown")
            instance_id = d.get("InstanceId", "")
            status = d.get("Status", "Unknown")
            if not instance_id:
                continue
            # Skip transport/service profile entries
            name_lower = name.lower()
            if any(skip in name_lower for skip in SKIP_PATTERNS):
                continue
            # Keep the best entry per device name (prefer connected)
            if name not in seen_names or status == "OK":
                seen_names[name] = {
                    "name": name,
                    "instance_id": instance_id,
                    "status": status,
                }
        return list(seen_names.values())
    except Exception as e:
        logger.error("Error enumerating Bluetooth devices: %s", e)
        return []


def is_device_connected(instance_id: str) -> bool:
    """Check if a specific Bluetooth device is connected."""
    ps_script = (
        f"(Get-PnpDevice -InstanceId '{instance_id}' -ErrorAction SilentlyContinue).Status"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
            capture_output=True, text=True, timeout=15,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        status = result.stdout.strip()
        return status == "OK"
    except Exception as e:
        logger.error("Error checking device status: %s", e)
        return False


def reconnect_device(instance_id: str) -> bool:
    """
    Attempt to reconnect a Bluetooth device by disabling and re-enabling it.
    """
    # Validate instance_id format to prevent injection
    if not instance_id or not instance_id.startswith("BTHENUM\\"):
        logger.warning("Invalid instance ID format: %s", instance_id)
        return False

    disable_script = (
        f"Disable-PnpDevice -InstanceId '{instance_id}' -Confirm:$false -ErrorAction Stop"
    )
    enable_script = (
        f"Enable-PnpDevice -InstanceId '{instance_id}' -Confirm:$false -ErrorAction Stop"
    )
    try:
        subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", disable_script],
            capture_output=True, text=True, timeout=15,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        time.sleep(1)
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", enable_script],
            capture_output=True, text=True, timeout=15,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        if result.returncode == 0:
            logger.info("Reconnect command sent for %s", instance_id)
            return True
        else:
            logger.warning("Enable failed: %s", result.stderr.strip())
            return False
    except Exception as e:
        logger.error("Reconnect error for %s: %s", instance_id, e)
        return False


# ── Startup management ─────────────────────────────────────────────────────────
def get_exe_path() -> str:
    """Get the path of the running executable."""
    if getattr(sys, "frozen", False):
        return sys.executable
    return os.path.abspath(sys.argv[0])


def set_auto_start(enable: bool):
    """Add or remove this app from Windows startup registry (current user)."""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, STARTUP_REG_KEY, 0, winreg.KEY_SET_VALUE
        )
        if enable:
            exe = get_exe_path()
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, f'"{exe}" --background')
            logger.info("Auto-start enabled: %s", exe)
        else:
            try:
                winreg.DeleteValue(key, APP_NAME)
                logger.info("Auto-start disabled.")
            except FileNotFoundError:
                pass
        winreg.CloseKey(key)
    except OSError as e:
        logger.error("Registry error: %s", e)


def is_auto_start_enabled() -> bool:
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER, STARTUP_REG_KEY, 0, winreg.KEY_READ
        )
        winreg.QueryValueEx(key, APP_NAME)
        winreg.CloseKey(key)
        return True
    except (FileNotFoundError, OSError):
        return False


# ── System tray icon ───────────────────────────────────────────────────────────
def create_icon_image(connected: bool = False) -> Image.Image:
    """Generate a simple Bluetooth-style tray icon."""
    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    color = (0, 120, 215) if connected else (180, 50, 50)
    # Draw a simplified Bluetooth symbol
    cx, cy = size // 2, size // 2
    draw.ellipse([4, 4, size - 4, size - 4], fill=color)
    # B letter
    draw.text((cx - 7, cy - 12), "B", fill="white")
    return img


# ── Native Win32 device picker dialog ──────────────────────────────────────────
def _show_device_picker(devices: list[dict]) -> int | None:
    """
    Show a native Win32 dialog with a listbox of Bluetooth devices.
    Each entry shows the device name and its status (Connected / Disconnected).
    Returns the selected index, or None if cancelled.
    """
    import ctypes.wintypes

    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    gdi32 = ctypes.windll.gdi32

    # Win32 constants
    WS_OVERLAPPED = 0x00000000
    WS_CAPTION = 0x00C00000
    WS_SYSMENU = 0x00080000
    WS_VISIBLE = 0x10000000
    WS_CHILD = 0x40000000
    WS_VSCROLL = 0x00200000
    WS_BORDER = 0x00800000
    WS_TABSTOP = 0x00010000
    WS_EX_DLGMODALFRAME = 0x00000001
    LBS_NOTIFY = 0x0001
    LBS_NOINTEGRALHEIGHT = 0x0100
    LBS_OWNERDRAWFIXED = 0x0010
    LBS_HASSTRINGS = 0x0040
    BS_DEFPUSHBUTTON = 0x0001
    BS_PUSHBUTTON = 0x0000
    WM_CREATE = 0x0001
    WM_DESTROY = 0x0002
    WM_CLOSE = 0x0010
    WM_COMMAND = 0x0111
    WM_SETFONT = 0x0030
    WM_DRAWITEM = 0x002B
    WM_MEASUREITEM = 0x002C
    WM_CTLCOLORLISTBOX = 0x0134
    WM_CTLCOLORBTN = 0x0135
    WM_CTLCOLORSTATIC = 0x0138
    WM_CTLCOLORDLG = 0x0136
    WM_ERASEBKGND = 0x0014
    LB_ADDSTRING = 0x0180
    LB_GETCURSEL = 0x0188
    LB_SETCURSEL = 0x0186
    LB_SETITEMDATA = 0x019A
    LB_GETITEMDATA = 0x0199
    LBN_DBLCLK = 2
    BN_CLICKED = 0
    SW_SHOW = 5
    COLOR_WINDOW = 5
    ODA_DRAWENTIRE = 0x0001
    ODA_SELECT = 0x0002
    ODA_FOCUS = 0x0004
    ODT_LISTBOX = 2
    ODS_SELECTED = 0x0001

    WNDPROC = ctypes.WINFUNCTYPE(
        ctypes.c_long,
        ctypes.wintypes.HWND,
        ctypes.c_uint,
        ctypes.wintypes.WPARAM,
        ctypes.wintypes.LPARAM,
    )

    ID_LISTBOX = 100
    ID_OK = 101
    ID_CANCEL = 102

    result_holder = [None]
    hwnd_listbox = [None]

    # Dark theme colors
    BG_COLOR = 0x00282828        # #282828 (dark bg)
    TEXT_COLOR = 0x00FFFFFF       # white text
    ACCENT_COLOR = 0x00D77800    # #0078D7 (blue accent) — BGR
    ITEM_BG = 0x00323232         # #323232 (item bg)
    CONNECTED_COLOR = 0x0066CC66  # green — BGR
    DISCONNECTED_COLOR = 0x005050AA  # muted red — BGR
    SEPARATOR_COLOR = 0x00444444

    bg_brush = gdi32.CreateSolidBrush(BG_COLOR)
    item_brush = gdi32.CreateSolidBrush(ITEM_BG)
    accent_brush = gdi32.CreateSolidBrush(ACCENT_COLOR)

    class DRAWITEMSTRUCT(ctypes.Structure):
        _fields_ = [
            ("CtlType", ctypes.c_uint),
            ("CtlID", ctypes.c_uint),
            ("itemID", ctypes.c_uint),
            ("itemAction", ctypes.c_uint),
            ("itemState", ctypes.c_uint),
            ("hwndItem", ctypes.wintypes.HWND),
            ("hDC", ctypes.wintypes.HDC),
            ("rcItem", ctypes.wintypes.RECT),
            ("itemData", ctypes.POINTER(ctypes.c_ulong)),
        ]

    class MEASUREITEMSTRUCT(ctypes.Structure):
        _fields_ = [
            ("CtlType", ctypes.c_uint),
            ("CtlID", ctypes.c_uint),
            ("itemID", ctypes.c_uint),
            ("itemWidth", ctypes.c_uint),
            ("itemHeight", ctypes.c_uint),
            ("itemData", ctypes.POINTER(ctypes.c_ulong)),
        ]

    def wnd_proc(hwnd, msg, wparam, lparam):
        if msg == WM_CREATE:
            return 0

        elif msg == WM_ERASEBKGND:
            hdc = wparam
            rect = ctypes.wintypes.RECT()
            user32.GetClientRect(hwnd, ctypes.byref(rect))
            user32.FillRect(hdc, ctypes.byref(rect), bg_brush)
            return 1

        elif msg == WM_CTLCOLORLISTBOX:
            hdc = wparam
            gdi32.SetTextColor(hdc, TEXT_COLOR)
            gdi32.SetBkColor(hdc, ITEM_BG)
            return item_brush

        elif msg in (WM_CTLCOLORBTN, WM_CTLCOLORSTATIC, WM_CTLCOLORDLG):
            hdc = wparam
            gdi32.SetTextColor(hdc, TEXT_COLOR)
            gdi32.SetBkColor(hdc, BG_COLOR)
            return bg_brush

        elif msg == WM_MEASUREITEM:
            mis = ctypes.cast(lparam, ctypes.POINTER(MEASUREITEMSTRUCT)).contents
            mis.itemHeight = 48
            return 1

        elif msg == WM_DRAWITEM:
            dis = ctypes.cast(lparam, ctypes.POINTER(DRAWITEMSTRUCT)).contents
            if dis.CtlID != ID_LISTBOX:
                return 0

            hdc = dis.hDC
            rc = dis.rcItem
            selected = bool(dis.itemState & ODS_SELECTED)
            idx = dis.itemID

            if idx >= len(devices):
                return 0

            # Background
            fill_brush = accent_brush if selected else item_brush
            user32.FillRect(hdc, ctypes.byref(rc), fill_brush)

            # Draw separator line at bottom
            sep_rect = ctypes.wintypes.RECT(rc.left, rc.bottom - 1, rc.right, rc.bottom)
            sep_brush = gdi32.CreateSolidBrush(SEPARATOR_COLOR)
            user32.FillRect(hdc, ctypes.byref(sep_rect), sep_brush)
            gdi32.DeleteObject(sep_brush)

            gdi32.SetBkMode(hdc, 1)  # TRANSPARENT

            dev = devices[idx]
            name = dev["name"]
            is_connected = dev.get("status", "Unknown") == "OK"
            status_text = "Connected" if is_connected else "Disconnected"

            # Draw device name (bold, larger)
            name_font = gdi32.CreateFontW(
                -18, 0, 0, 0, 700, 0, 0, 0, 0, 0, 0, 0, 0, "Segoe UI"
            )
            old_font = gdi32.SelectObject(hdc, name_font)
            gdi32.SetTextColor(hdc, 0x00FFFFFF if selected else 0x00EEEEEE)
            name_rect = ctypes.wintypes.RECT(rc.left + 14, rc.top + 6, rc.right - 14, rc.top + 28)
            user32.DrawTextW(hdc, name, -1, ctypes.byref(name_rect), 0x0000)
            gdi32.SelectObject(hdc, old_font)
            gdi32.DeleteObject(name_font)

            # Draw status (smaller, colored)
            status_font = gdi32.CreateFontW(
                -13, 0, 0, 0, 400, 0, 0, 0, 0, 0, 0, 0, 0, "Segoe UI"
            )
            old_font2 = gdi32.SelectObject(hdc, status_font)
            status_color = CONNECTED_COLOR if is_connected else DISCONNECTED_COLOR
            gdi32.SetTextColor(hdc, status_color)
            status_rect = ctypes.wintypes.RECT(rc.left + 14, rc.top + 27, rc.right - 14, rc.bottom - 4)
            # Draw a small circle indicator
            indicator = "\u25CF "  # ● bullet
            full_status = indicator + status_text
            user32.DrawTextW(hdc, full_status, -1, ctypes.byref(status_rect), 0x0000)
            gdi32.SelectObject(hdc, old_font2)
            gdi32.DeleteObject(status_font)

            return 1

        elif msg == WM_COMMAND:
            control_id = wparam & 0xFFFF
            notify_code = (wparam >> 16) & 0xFFFF

            if control_id == ID_OK and notify_code == BN_CLICKED:
                sel = user32.SendMessageW(hwnd_listbox[0], LB_GETCURSEL, 0, 0)
                if sel >= 0:
                    result_holder[0] = sel
                user32.DestroyWindow(hwnd)
                return 0

            if control_id == ID_CANCEL and notify_code == BN_CLICKED:
                user32.DestroyWindow(hwnd)
                return 0

            if control_id == ID_LISTBOX and notify_code == LBN_DBLCLK:
                sel = user32.SendMessageW(hwnd_listbox[0], LB_GETCURSEL, 0, 0)
                if sel >= 0:
                    result_holder[0] = sel
                user32.DestroyWindow(hwnd)
                return 0

        elif msg == WM_CLOSE:
            user32.DestroyWindow(hwnd)
            return 0

        elif msg == WM_DESTROY:
            user32.PostQuitMessage(0)
            return 0

        return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    # prevent GC of the callback
    wnd_proc_cb = WNDPROC(wnd_proc)

    class_name = f"{APP_NAME}_DevicePicker"
    wc = ctypes.wintypes.WNDCLASSW()
    wc.lpfnWndProc = wnd_proc_cb
    wc.hInstance = kernel32.GetModuleHandleW(None)
    wc.lpszClassName = class_name
    wc.hCursor = user32.LoadCursorW(None, 32512)  # IDC_ARROW
    wc.hbrBackground = bg_brush

    user32.RegisterClassW(ctypes.byref(wc))

    # Center on screen
    dlg_w, dlg_h = 420, 480
    screen_w = user32.GetSystemMetrics(0)
    screen_h = user32.GetSystemMetrics(1)
    x = (screen_w - dlg_w) // 2
    y = (screen_h - dlg_h) // 2

    hwnd = user32.CreateWindowExW(
        WS_EX_DLGMODALFRAME,
        class_name,
        "Select Bluetooth Device",
        WS_OVERLAPPED | WS_CAPTION | WS_SYSMENU | WS_VISIBLE,
        x, y, dlg_w, dlg_h,
        None, None, wc.hInstance, None,
    )

    # UI font
    ui_font = gdi32.CreateFontW(-14, 0, 0, 0, 400, 0, 0, 0, 0, 0, 0, 0, 0, "Segoe UI")

    # Label
    h_label = user32.CreateWindowExW(
        0, "STATIC",
        "Select a paired device to monitor for auto-reconnect:",
        WS_CHILD | WS_VISIBLE,
        14, 12, 380, 22,
        hwnd, None, wc.hInstance, None,
    )
    user32.SendMessageW(h_label, WM_SETFONT, ui_font, 1)

    # Listbox (owner-drawn)
    lb_style = (
        WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_BORDER | WS_TABSTOP
        | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT | LBS_OWNERDRAWFIXED | LBS_HASSTRINGS
    )
    h_listbox = user32.CreateWindowExW(
        0, "LISTBOX", None,
        lb_style,
        14, 40, 378, 340,
        hwnd, ID_LISTBOX, wc.hInstance, None,
    )
    hwnd_listbox[0] = h_listbox

    for i, dev in enumerate(devices):
        label = dev["name"]
        user32.SendMessageW(h_listbox, LB_ADDSTRING, 0, label)

    # Select first item
    user32.SendMessageW(h_listbox, LB_SETCURSEL, 0, 0)

    # Buttons
    btn_font = gdi32.CreateFontW(-14, 0, 0, 0, 600, 0, 0, 0, 0, 0, 0, 0, 0, "Segoe UI")

    h_ok = user32.CreateWindowExW(
        0, "BUTTON", "Add Device",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
        142, 395, 120, 36,
        hwnd, ID_OK, wc.hInstance, None,
    )
    user32.SendMessageW(h_ok, WM_SETFONT, btn_font, 1)

    h_cancel = user32.CreateWindowExW(
        0, "BUTTON", "Cancel",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
        272, 395, 120, 36,
        hwnd, ID_CANCEL, wc.hInstance, None,
    )
    user32.SendMessageW(h_cancel, WM_SETFONT, btn_font, 1)

    # Message loop
    msg = ctypes.wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
        if not user32.IsDialogMessageW(hwnd, ctypes.byref(msg)):
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

    # Cleanup
    gdi32.DeleteObject(ui_font)
    gdi32.DeleteObject(btn_font)
    gdi32.DeleteObject(bg_brush)
    gdi32.DeleteObject(item_brush)
    gdi32.DeleteObject(accent_brush)
    user32.UnregisterClassW(class_name, wc.hInstance)

    return result_holder[0]


class BluetoothPersistenceApp:
    def __init__(self):
        self.config = load_config()
        self.running = True
        self.icon: pystray.Icon | None = None
        self.monitor_thread: threading.Thread | None = None
        self.power_thread: threading.Thread | None = None
        self._lock = threading.Lock()
        self._backoff = {}  # instance_id -> current interval

    def start(self):
        logger.info("Starting %s", APP_DISPLAY_NAME)

        # Set auto-start if configured
        if self.config.get("auto_start", True):
            set_auto_start(True)

        # Start monitoring thread
        self.monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self.monitor_thread.start()

        # Start power event listener thread
        self.power_thread = threading.Thread(target=self._power_event_listener, daemon=True)
        self.power_thread.start()

        # Build and run system tray icon (blocks on main thread)
        self._run_tray()

    def _run_tray(self):
        menu = pystray.Menu(
            item("Monitored Devices", pystray.Menu(self._build_device_menu)),
            item("Add Device...", self._on_add_device),
            pystray.Menu.SEPARATOR,
            item(
                "Start with Windows",
                self._on_toggle_autostart,
                checked=lambda _: is_auto_start_enabled(),
            ),
            item("View Log", self._on_view_log),
            pystray.Menu.SEPARATOR,
            item("Quit", self._on_quit),
        )
        self.icon = pystray.Icon(
            APP_NAME,
            create_icon_image(False),
            APP_DISPLAY_NAME,
            menu,
        )
        self.icon.run()

    def _build_device_menu(self) -> list:
        """Dynamically build submenu of monitored devices."""
        items = []
        for dev in self.config.get("devices", []):
            name = dev["name"]
            iid = dev["instance_id"]
            connected = is_device_connected(iid)
            status = "Connected" if connected else "Disconnected"
            label = f"{name} — {status}"
            items.append(
                item(label, self._make_remove_handler(iid, name))
            )
        if not items:
            items.append(item("(no devices)", lambda: None, enabled=False))
        return items

    def _make_remove_handler(self, instance_id: str, name: str):
        def handler(icon, menu_item):
            self._remove_device(instance_id, name)
        return handler

    def _on_add_device(self, icon, menu_item):
        """Show paired devices for user to select."""
        threading.Thread(target=self._add_device_dialog, daemon=True).start()

    def _add_device_dialog(self):
        devices = get_paired_devices()
        monitored_ids = {d["instance_id"] for d in self.config.get("devices", [])}
        # Only show devices that are currently connected AND not already monitored
        available = [
            d for d in devices
            if d["instance_id"] not in monitored_ids and d.get("status") == "OK"
        ]

        if not available:
            ctypes.windll.user32.MessageBoxW(
                0,
                "No connected Bluetooth devices available to add.\n\n"
                "Make sure your device is paired and connected in Windows Bluetooth settings.",
                APP_DISPLAY_NAME,
                0x40,
            )
            return

        selected_index = _show_device_picker(available)
        if selected_index is not None and 0 <= selected_index < len(available):
            dev = available[selected_index]
            with self._lock:
                self.config.setdefault("devices", []).append({
                    "name": dev["name"],
                    "instance_id": dev["instance_id"],
                })
                save_config(self.config)
            logger.info("Added device: %s (%s)", dev["name"], dev["instance_id"])
            if self.icon:
                self.icon.notify(f"Now monitoring: {dev['name']}", APP_DISPLAY_NAME)

    def _remove_device(self, instance_id: str, name: str):
        """Remove a device from monitoring."""
        confirm = ctypes.windll.user32.MessageBoxW(
            0,
            f"Stop monitoring '{name}'?",
            APP_DISPLAY_NAME,
            0x04 | 0x20,  # MB_YESNO | MB_ICONQUESTION
        )
        if confirm == 6:  # IDYES
            with self._lock:
                self.config["devices"] = [
                    d for d in self.config.get("devices", [])
                    if d["instance_id"] != instance_id
                ]
                save_config(self.config)
            logger.info("Removed device: %s", name)

    def _on_toggle_autostart(self, icon, menu_item):
        enabled = is_auto_start_enabled()
        set_auto_start(not enabled)
        self.config["auto_start"] = not enabled
        save_config(self.config)

    def _on_view_log(self, icon, menu_item):
        if LOG_FILE.exists():
            os.startfile(str(LOG_FILE))

    def _on_quit(self, icon, menu_item):
        logger.info("Shutting down.")
        self.running = False
        icon.stop()

    # ── Monitor loop ───────────────────────────────────────────────────────────
    def _monitor_loop(self):
        """Continuously check device connectivity and reconnect if needed."""
        # Short delay on startup to let Bluetooth stack initialize
        time.sleep(10)
        logger.info("Monitor loop started.")

        while self.running:
            with self._lock:
                devices = list(self.config.get("devices", []))

            for dev in devices:
                if not self.running:
                    break
                iid = dev["instance_id"]
                name = dev["name"]

                if is_device_connected(iid):
                    # Reset back-off on successful connection
                    self._backoff.pop(iid, None)
                    continue

                # Device disconnected — attempt reconnect
                interval = self._backoff.get(iid, RECONNECT_INTERVAL)
                logger.info("Device '%s' disconnected. Attempting reconnect...", name)

                success = reconnect_device(iid)
                if success:
                    # Wait and verify
                    time.sleep(3)
                    if is_device_connected(iid):
                        logger.info("Successfully reconnected '%s'.", name)
                        self._backoff.pop(iid, None)
                        if self.icon:
                            self.icon.notify(f"Reconnected: {name}", APP_DISPLAY_NAME)
                        continue

                # Back off
                self._backoff[iid] = min(interval * 2, MAX_RECONNECT_INTERVAL)
                logger.info(
                    "Reconnect failed for '%s'. Next attempt in %ds.",
                    name, self._backoff[iid],
                )

            # Sleep — use shorter intervals to stay responsive
            sleep_time = min(
                self._backoff.values(), default=RECONNECT_INTERVAL
            )
            for _ in range(int(sleep_time)):
                if not self.running:
                    return
                time.sleep(1)

    # ── Power event listener (sleep/wake) ──────────────────────────────────────
    def _power_event_listener(self):
        """
        Listen for Windows power events (suspend/resume) via a hidden window
        to trigger immediate reconnection after wake from sleep.
        """
        logger.info("Power event listener started.")

        GUID_MONITOR_POWER_ON = "{02731015-4510-4526-99e6-e5a17ebd1aea}"

        def wnd_proc(hwnd, msg, wparam, lparam):
            WM_POWERBROADCAST = 0x0218
            PBT_APMRESUMEAUTOMATIC = 0x0012
            PBT_APMRESUMESUSPEND = 0x0007
            PBT_APMSUSPEND = 0x0004

            if msg == WM_POWERBROADCAST:
                if wparam in (PBT_APMRESUMEAUTOMATIC, PBT_APMRESUMESUSPEND):
                    logger.info("System resumed from sleep/hibernation. Triggering reconnect...")
                    # Reset all back-offs and trigger immediate reconnect
                    self._backoff.clear()
                elif wparam == PBT_APMSUSPEND:
                    logger.info("System entering sleep/hibernation.")

            return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

        wc = win32gui.WNDCLASS()
        wc.lpfnWndProc = wnd_proc
        wc.lpszClassName = f"{APP_NAME}_PowerWatcher"
        wc.hInstance = win32api.GetModuleHandle(None)

        try:
            class_atom = win32gui.RegisterClass(wc)
            hwnd = win32gui.CreateWindow(
                class_atom, APP_NAME, 0, 0, 0, 0, 0, 0, 0, wc.hInstance, None
            )

            # Message pump
            while self.running:
                has_msg = win32gui.PeekMessage(hwnd, 0, 0, 0x0001)  # PM_REMOVE
                if has_msg and has_msg[1]:
                    msg = has_msg[1]
                    win32gui.TranslateMessage(msg)
                    win32gui.DispatchMessage(msg)
                else:
                    time.sleep(0.5)
        except Exception as e:
            logger.error("Power event listener error: %s", e)


# ── Entry point ────────────────────────────────────────────────────────────────
def main():
    # Prevent multiple instances using a named mutex
    mutex_name = f"Global\\{APP_NAME}_Mutex"
    mutex = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
    last_error = ctypes.windll.kernel32.GetLastError()
    ERROR_ALREADY_EXISTS = 183

    if last_error == ERROR_ALREADY_EXISTS:
        ctypes.windll.user32.MessageBoxW(
            0,
            f"{APP_DISPLAY_NAME} is already running.",
            APP_DISPLAY_NAME,
            0x40,
        )
        sys.exit(0)

    try:
        app = BluetoothPersistenceApp()
        app.start()
    except KeyboardInterrupt:
        logger.info("Interrupted by user.")
    except Exception as e:
        logger.critical("Fatal error: %s", e, exc_info=True)
        sys.exit(1)
    finally:
        if mutex:
            ctypes.windll.kernel32.ReleaseMutex(mutex)
            ctypes.windll.kernel32.CloseHandle(mutex)


if __name__ == "__main__":
    main()
