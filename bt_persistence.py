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

def _get_connected_bt_names() -> set[str]:
    """Use Windows Bluetooth API to get names of truly connected devices."""
    try:
        bt = ctypes.WinDLL("bthprops.cpl")
    except OSError:
        try:
            bt = ctypes.WinDLL("BluetoothAPIs.dll")
        except OSError:
            logger.warning("Could not load Bluetooth API DLL")
            return set()

    class SYSTEMTIME(ctypes.Structure):
        _fields_ = [
            ("wYear", ctypes.c_ushort), ("wMonth", ctypes.c_ushort),
            ("wDayOfWeek", ctypes.c_ushort), ("wDay", ctypes.c_ushort),
            ("wHour", ctypes.c_ushort), ("wMinute", ctypes.c_ushort),
            ("wSecond", ctypes.c_ushort), ("wMilliseconds", ctypes.c_ushort),
        ]

    class BLUETOOTH_DEVICE_INFO(ctypes.Structure):
        _fields_ = [
            ("dwSize", ctypes.c_ulong),
            ("Address", ctypes.c_ulonglong),
            ("ulClassofDevice", ctypes.c_ulong),
            ("fConnected", ctypes.c_int),
            ("fRemembered", ctypes.c_int),
            ("fAuthenticated", ctypes.c_int),
            ("stLastSeen", SYSTEMTIME),
            ("stLastUsed", SYSTEMTIME),
            ("szName", ctypes.c_wchar * 248),
        ]

    class BLUETOOTH_DEVICE_SEARCH_PARAMS(ctypes.Structure):
        _fields_ = [
            ("dwSize", ctypes.c_ulong),
            ("fReturnAuthenticated", ctypes.c_int),
            ("fReturnRemembered", ctypes.c_int),
            ("fReturnUnknown", ctypes.c_int),
            ("fReturnConnected", ctypes.c_int),
            ("fIssueInquiry", ctypes.c_int),
            ("cTimeoutMultiplier", ctypes.c_ubyte),
            ("hRadio", ctypes.c_void_p),
        ]

    try:
        # Set proper argtypes/restype for 64-bit handle compatibility
        bt.BluetoothFindFirstDevice.restype = ctypes.c_void_p
        bt.BluetoothFindFirstDevice.argtypes = [
            ctypes.POINTER(BLUETOOTH_DEVICE_SEARCH_PARAMS),
            ctypes.POINTER(BLUETOOTH_DEVICE_INFO),
        ]
        bt.BluetoothFindNextDevice.restype = ctypes.c_int
        bt.BluetoothFindNextDevice.argtypes = [
            ctypes.c_void_p, ctypes.POINTER(BLUETOOTH_DEVICE_INFO),
        ]
        bt.BluetoothFindDeviceClose.restype = ctypes.c_int
        bt.BluetoothFindDeviceClose.argtypes = [ctypes.c_void_p]

        params = BLUETOOTH_DEVICE_SEARCH_PARAMS()
        params.dwSize = ctypes.sizeof(BLUETOOTH_DEVICE_SEARCH_PARAMS)
        params.fReturnConnected = 1
        params.fReturnRemembered = 0
        params.fReturnAuthenticated = 0
        params.fReturnUnknown = 0
        params.fIssueInquiry = 0
        params.cTimeoutMultiplier = 0
        params.hRadio = None

        dev = BLUETOOTH_DEVICE_INFO()
        dev.dwSize = ctypes.sizeof(BLUETOOTH_DEVICE_INFO)

        connected = set()
        hFind = bt.BluetoothFindFirstDevice(ctypes.byref(params), ctypes.byref(dev))
        if hFind:
            if dev.fConnected:
                connected.add(dev.szName)
            while bt.BluetoothFindNextDevice(hFind, ctypes.byref(dev)):
                if dev.fConnected:
                    connected.add(dev.szName)
            bt.BluetoothFindDeviceClose(hFind)

        logger.info("Bluetooth API reports connected: %s", connected)
        return connected
    except Exception as e:
        logger.error("Error querying Bluetooth API: %s", e)
        return set()


def get_paired_devices() -> list[dict]:
    """
    Enumerate paired Bluetooth devices via PowerShell.
    Returns deduplicated list of dicts with 'name', 'instance_id', 'status'.
    Uses Windows Bluetooth API for actual connection status.
    """
    # Get truly connected device names from Windows Bluetooth API
    connected_names = _get_connected_bt_names()

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
            # Keep the best entry per device name (prefer OK PnP status for instance_id)
            if name not in seen_names or status == "OK":
                # Use Bluetooth API to determine actual connection, not PnP status
                actual_status = "OK" if name in connected_names else "Disconnected"
                seen_names[name] = {
                    "name": name,
                    "instance_id": instance_id,
                    "status": actual_status,
                }
        return list(seen_names.values())
    except Exception as e:
        logger.error("Error enumerating Bluetooth devices: %s", e)
        return []


def is_device_connected(instance_id: str) -> bool:
    """Check if a specific Bluetooth device is connected using Windows Bluetooth API."""
    # Get the device name from PnP, then check against BT API
    ps_script = (
        f"(Get-PnpDevice -InstanceId '{instance_id}' -ErrorAction SilentlyContinue).FriendlyName"
    )
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
            capture_output=True, text=True, timeout=15,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        name = result.stdout.strip()
        if not name:
            return False
        connected_names = _get_connected_bt_names()
        return name in connected_names
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
    Returns the selected index, or None if cancelled.
    """
    import ctypes.wintypes

    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    gdi32 = ctypes.windll.gdi32

    # ── 64-bit compatibility: set argtypes/restype for ALL Win32 calls ──
    kernel32.GetModuleHandleW.argtypes = [ctypes.wintypes.LPCWSTR]
    kernel32.GetModuleHandleW.restype = ctypes.wintypes.HINSTANCE

    user32.CreateWindowExW.argtypes = [
        ctypes.wintypes.DWORD, ctypes.wintypes.LPCWSTR, ctypes.wintypes.LPCWSTR,
        ctypes.wintypes.DWORD, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
        ctypes.wintypes.HWND, ctypes.wintypes.HMENU, ctypes.wintypes.HINSTANCE,
        ctypes.c_void_p,
    ]
    user32.CreateWindowExW.restype = ctypes.wintypes.HWND

    user32.DefWindowProcW.argtypes = [
        ctypes.wintypes.HWND, ctypes.c_uint,
        ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM,
    ]
    user32.DefWindowProcW.restype = ctypes.c_longlong

    user32.SendMessageW.argtypes = [
        ctypes.wintypes.HWND, ctypes.c_uint,
        ctypes.wintypes.WPARAM, ctypes.wintypes.LPARAM,
    ]
    user32.SendMessageW.restype = ctypes.c_longlong

    user32.RegisterClassW.argtypes = [ctypes.c_void_p]
    user32.RegisterClassW.restype = ctypes.wintypes.ATOM

    user32.UnregisterClassW.argtypes = [ctypes.wintypes.LPCWSTR, ctypes.wintypes.HINSTANCE]
    user32.UnregisterClassW.restype = ctypes.wintypes.BOOL

    user32.LoadCursorW.argtypes = [ctypes.wintypes.HINSTANCE, ctypes.wintypes.LPCWSTR]
    user32.LoadCursorW.restype = ctypes.wintypes.HANDLE

    user32.DestroyWindow.argtypes = [ctypes.wintypes.HWND]
    user32.DestroyWindow.restype = ctypes.wintypes.BOOL

    user32.PostQuitMessage.argtypes = [ctypes.c_int]
    user32.PostQuitMessage.restype = None

    user32.GetSystemMetrics.argtypes = [ctypes.c_int]
    user32.GetSystemMetrics.restype = ctypes.c_int

    user32.GetMessageW.argtypes = [
        ctypes.POINTER(ctypes.wintypes.MSG), ctypes.wintypes.HWND,
        ctypes.c_uint, ctypes.c_uint,
    ]
    user32.GetMessageW.restype = ctypes.wintypes.BOOL

    user32.IsDialogMessageW.argtypes = [ctypes.wintypes.HWND, ctypes.POINTER(ctypes.wintypes.MSG)]
    user32.IsDialogMessageW.restype = ctypes.wintypes.BOOL

    user32.TranslateMessage.argtypes = [ctypes.POINTER(ctypes.wintypes.MSG)]
    user32.TranslateMessage.restype = ctypes.wintypes.BOOL

    user32.DispatchMessageW.argtypes = [ctypes.POINTER(ctypes.wintypes.MSG)]
    user32.DispatchMessageW.restype = ctypes.c_longlong

    gdi32.CreateFontW.restype = ctypes.wintypes.HFONT
    gdi32.DeleteObject.argtypes = [ctypes.wintypes.HANDLE]
    gdi32.DeleteObject.restype = ctypes.wintypes.BOOL

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
    BS_DEFPUSHBUTTON = 0x0001
    BS_PUSHBUTTON = 0x0000
    WM_DESTROY = 0x0002
    WM_CLOSE = 0x0010
    WM_COMMAND = 0x0111
    WM_SETFONT = 0x0030
    LB_ADDSTRING = 0x0180
    LB_GETCURSEL = 0x0188
    LB_SETCURSEL = 0x0186
    LBN_DBLCLK = 2
    BN_CLICKED = 0

    WNDPROC = ctypes.WINFUNCTYPE(
        ctypes.c_longlong,
        ctypes.wintypes.HWND,
        ctypes.c_uint,
        ctypes.wintypes.WPARAM,
        ctypes.wintypes.LPARAM,
    )

    class WNDCLASSW(ctypes.Structure):
        _fields_ = [
            ("style", ctypes.c_uint),
            ("lpfnWndProc", WNDPROC),
            ("cbClsExtra", ctypes.c_int),
            ("cbWndExtra", ctypes.c_int),
            ("hInstance", ctypes.wintypes.HINSTANCE),
            ("hIcon", ctypes.wintypes.HICON),
            ("hCursor", ctypes.wintypes.HANDLE),
            ("hbrBackground", ctypes.wintypes.HBRUSH),
            ("lpszMenuName", ctypes.wintypes.LPCWSTR),
            ("lpszClassName", ctypes.wintypes.LPCWSTR),
        ]

    ID_LISTBOX = 100
    ID_OK = 101
    ID_CANCEL = 102

    result_holder = [None]
    hwnd_listbox = [None]

    def wnd_proc(hwnd, msg, wparam, lparam):
        if msg == WM_COMMAND:
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

    wnd_proc_cb = WNDPROC(wnd_proc)

    class_name = f"{APP_NAME}_DevicePicker"
    hInstance = kernel32.GetModuleHandleW(None)

    wc = WNDCLASSW()
    wc.lpfnWndProc = wnd_proc_cb
    wc.hInstance = hInstance
    wc.lpszClassName = class_name
    wc.hCursor = user32.LoadCursorW(None, ctypes.wintypes.LPCWSTR(32512))
    wc.hbrBackground = ctypes.wintypes.HBRUSH(6)  # COLOR_BTNFACE + 1

    user32.RegisterClassW(ctypes.byref(wc))

    # Center on screen
    dlg_w, dlg_h = 400, 380
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
        None, None, hInstance, None,
    )

    ui_font = gdi32.CreateFontW(-16, 0, 0, 0, 400, 0, 0, 0, 0, 0, 0, 0, 0, "Segoe UI")
    font_wparam = ctypes.cast(ui_font, ctypes.c_void_p).value or 0

    # Label
    h_label = user32.CreateWindowExW(
        0, "STATIC",
        "Select a device to monitor for auto-reconnect:",
        WS_CHILD | WS_VISIBLE,
        14, 10, 360, 22,
        hwnd, None, hInstance, None,
    )
    user32.SendMessageW(h_label, WM_SETFONT, font_wparam, 1)

    # Standard listbox (no owner-draw)
    lb_style = (
        WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_BORDER | WS_TABSTOP
        | LBS_NOTIFY | LBS_NOINTEGRALHEIGHT
    )
    h_listbox = user32.CreateWindowExW(
        0, "LISTBOX", None,
        lb_style,
        14, 38, 358, 240,
        hwnd, ctypes.wintypes.HMENU(ID_LISTBOX), hInstance, None,
    )
    hwnd_listbox[0] = h_listbox
    user32.SendMessageW(h_listbox, WM_SETFONT, font_wparam, 1)

    # Add items as formatted text strings
    _label_refs = []  # prevent GC of string buffers
    for dev in devices:
        status = "\u2713 Connected" if dev.get("status") == "OK" else "\u2717 Disconnected"
        label = f"{dev['name']}  —  {status}"
        buf = ctypes.c_wchar_p(label)
        _label_refs.append(buf)
        lp_str = ctypes.cast(buf, ctypes.c_void_p).value or 0
        user32.SendMessageW(h_listbox, LB_ADDSTRING, 0, lp_str)

    user32.SendMessageW(h_listbox, LB_SETCURSEL, 0, 0)

    # Buttons
    btn_font = gdi32.CreateFontW(-14, 0, 0, 0, 600, 0, 0, 0, 0, 0, 0, 0, 0, "Segoe UI")
    btn_font_wparam = ctypes.cast(btn_font, ctypes.c_void_p).value or 0

    h_ok = user32.CreateWindowExW(
        0, "BUTTON", "Add Device",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_DEFPUSHBUTTON,
        130, 290, 110, 34,
        hwnd, ctypes.wintypes.HMENU(ID_OK), hInstance, None,
    )
    user32.SendMessageW(h_ok, WM_SETFONT, btn_font_wparam, 1)

    h_cancel = user32.CreateWindowExW(
        0, "BUTTON", "Cancel",
        WS_CHILD | WS_VISIBLE | WS_TABSTOP | BS_PUSHBUTTON,
        250, 290, 110, 34,
        hwnd, ctypes.wintypes.HMENU(ID_CANCEL), hInstance, None,
    )
    user32.SendMessageW(h_cancel, WM_SETFONT, btn_font_wparam, 1)

    # Message loop
    msg = ctypes.wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
        if not user32.IsDialogMessageW(hwnd, ctypes.byref(msg)):
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

    # Cleanup
    gdi32.DeleteObject(ui_font)
    gdi32.DeleteObject(btn_font)
    user32.UnregisterClassW(class_name, hInstance)

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
        try:
            logger.info("Add Device dialog requested.")
            devices = get_paired_devices()
            logger.info("Found %d paired device(s).", len(devices))
            monitored_ids = {d["instance_id"] for d in self.config.get("devices", [])}
            # Only show devices that are currently connected AND not already monitored
            available = [
                d for d in devices
                if d["instance_id"] not in monitored_ids and d.get("status") == "OK"
            ]
            logger.info("Available (connected, not already monitored): %d", len(available))

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
        except Exception:
            logger.exception("Error in Add Device dialog")

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
