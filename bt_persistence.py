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
    Returns list of dicts with 'name' and 'address'.
    """
    ps_script = (
        "Get-PnpDevice -Class Bluetooth | "
        "Where-Object { $_.FriendlyName -and $_.InstanceId -match 'BTHENUM' } | "
        "Select-Object FriendlyName, InstanceId, Status | "
        "ConvertTo-Json -Compress"
    )
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
        devices = []
        for d in data:
            name = d.get("FriendlyName", "Unknown")
            instance_id = d.get("InstanceId", "")
            status = d.get("Status", "Unknown")
            if instance_id:
                devices.append({
                    "name": name,
                    "instance_id": instance_id,
                    "status": status,
                })
        return devices
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
        available = [d for d in devices if d["instance_id"] not in monitored_ids]

        if not available:
            ctypes.windll.user32.MessageBoxW(
                0,
                "No additional paired Bluetooth devices found.\n\n"
                "Make sure your device is paired in Windows Bluetooth settings.",
                APP_DISPLAY_NAME,
                0x40,  # MB_ICONINFORMATION
            )
            return

        # Build selection list
        lines = ["Select a device to monitor:\n"]
        for i, dev in enumerate(available, 1):
            lines.append(f"  {i}. {dev['name']} ({dev['status']})")
        lines.append(f"\nEnter a number (1-{len(available)}):")

        # Use a simple input box via PowerShell
        prompt_text = "\\n".join(lines).replace('"', '`"')
        ps_cmd = (
            f'Add-Type -AssemblyName Microsoft.VisualBasic; '
            f'[Microsoft.VisualBasic.Interaction]::InputBox('
            f'"{prompt_text}", "{APP_DISPLAY_NAME}", "1")'
        )
        try:
            result = subprocess.run(
                ["powershell", "-NoProfile", "-Command", ps_cmd],
                capture_output=True, text=True, timeout=120,
                creationflags=subprocess.CREATE_NO_WINDOW,
            )
            choice = result.stdout.strip()
            if not choice:
                return
            idx = int(choice) - 1
            if 0 <= idx < len(available):
                dev = available[idx]
                with self._lock:
                    self.config.setdefault("devices", []).append({
                        "name": dev["name"],
                        "instance_id": dev["instance_id"],
                    })
                    save_config(self.config)
                logger.info("Added device: %s (%s)", dev["name"], dev["instance_id"])
                self.icon.notify(f"Now monitoring: {dev['name']}", APP_DISPLAY_NAME)
            else:
                ctypes.windll.user32.MessageBoxW(
                    0, "Invalid selection.", APP_DISPLAY_NAME, 0x30
                )
        except (ValueError, subprocess.TimeoutExpired):
            pass

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
