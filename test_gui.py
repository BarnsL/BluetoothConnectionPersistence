"""Test the device picker dialog standalone."""
import subprocess, json, ctypes, ctypes.wintypes

APP_NAME = "BluetoothConnectionPersistence"

SKIP_PATTERNS = {
    "avrcp transport", "avrcp", "handsfree", "a2dp",
    "phonebook access", "service discovery",
    "network nap", "personal area network",
    "serial port", "object push", "dial-up",
    "headset gateway", "audio sink", "audio source",
    "human interface", "hid device",
}

ps_script = (
    'Get-PnpDevice -Class Bluetooth | '
    'Where-Object { $_.FriendlyName -and $_.InstanceId -match "BTHENUM" } | '
    'Select-Object FriendlyName, InstanceId, Status | '
    'ConvertTo-Json -Compress'
)
result = subprocess.run(
    ['powershell', '-NoProfile', '-NonInteractive', '-Command', ps_script],
    capture_output=True, text=True, timeout=30,
    creationflags=0x08000000,
)
raw = result.stdout.strip()
data = json.loads(raw)
if isinstance(data, dict):
    data = [data]

seen_names = {}
for d in data:
    name = d.get('FriendlyName', 'Unknown')
    instance_id = d.get('InstanceId', '')
    status = d.get('Status', 'Unknown')
    if not instance_id:
        continue
    name_lower = name.lower()
    if any(skip in name_lower for skip in SKIP_PATTERNS):
        continue
    if name not in seen_names or status == 'OK':
        seen_names[name] = {
            'name': name,
            'instance_id': instance_id,
            'status': status,
        }

devices = list(seen_names.values())
print(f"Found {len(devices)} devices, launching picker...")

# Inline test of _show_device_picker
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
gdi32 = ctypes.windll.gdi32

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
LBN_DBLCLK = 2
BN_CLICKED = 0
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

BG_COLOR = 0x00282828
TEXT_COLOR = 0x00FFFFFF
ACCENT_COLOR = 0x00D77800
ITEM_BG = 0x00323232
CONNECTED_COLOR = 0x0066CC66
DISCONNECTED_COLOR = 0x005050AA
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
    try:
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
            fill_brush = accent_brush if selected else item_brush
            user32.FillRect(hdc, ctypes.byref(rc), fill_brush)
            sep_rect = ctypes.wintypes.RECT(rc.left, rc.bottom - 1, rc.right, rc.bottom)
            sep_brush = gdi32.CreateSolidBrush(SEPARATOR_COLOR)
            user32.FillRect(hdc, ctypes.byref(sep_rect), sep_brush)
            gdi32.DeleteObject(sep_brush)
            gdi32.SetBkMode(hdc, 1)
            dev = devices[idx]
            name = dev["name"]
            is_connected = dev.get("status", "Unknown") == "OK"
            status_text = "Connected" if is_connected else "Disconnected"
            name_font = gdi32.CreateFontW(-18, 0, 0, 0, 700, 0, 0, 0, 0, 0, 0, 0, 0, "Segoe UI")
            old_font = gdi32.SelectObject(hdc, name_font)
            gdi32.SetTextColor(hdc, 0x00FFFFFF if selected else 0x00EEEEEE)
            name_rect = ctypes.wintypes.RECT(rc.left + 14, rc.top + 6, rc.right - 14, rc.top + 28)
            user32.DrawTextW(hdc, name, -1, ctypes.byref(name_rect), 0x0000)
            gdi32.SelectObject(hdc, old_font)
            gdi32.DeleteObject(name_font)
            status_font = gdi32.CreateFontW(-13, 0, 0, 0, 400, 0, 0, 0, 0, 0, 0, 0, 0, "Segoe UI")
            old_font2 = gdi32.SelectObject(hdc, status_font)
            status_color = CONNECTED_COLOR if is_connected else DISCONNECTED_COLOR
            gdi32.SetTextColor(hdc, status_color)
            status_rect = ctypes.wintypes.RECT(rc.left + 14, rc.top + 27, rc.right - 14, rc.bottom - 4)
            indicator = "\u25CF "
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
    except Exception as e:
        print(f"WndProc error: {e}")

    return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

wnd_proc_cb = WNDPROC(wnd_proc)

# WNDCLASSW - define properly
class WNDCLASSW(ctypes.Structure):
    _fields_ = [
        ("style", ctypes.c_uint),
        ("lpfnWndProc", WNDPROC),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", ctypes.wintypes.HINSTANCE),
        ("hIcon", ctypes.wintypes.HICON),
        ("hCursor", ctypes.wintypes.HICON),
        ("hbrBackground", ctypes.wintypes.HBRUSH),
        ("lpszMenuName", ctypes.wintypes.LPCWSTR),
        ("lpszClassName", ctypes.wintypes.LPCWSTR),
    ]

class_name = "TestDevicePicker"
wc = WNDCLASSW()
wc.lpfnWndProc = wnd_proc_cb
wc.hInstance = kernel32.GetModuleHandleW(None)
wc.lpszClassName = class_name
wc.hCursor = user32.LoadCursorW(None, 32512)
wc.hbrBackground = bg_brush

atom = user32.RegisterClassW(ctypes.byref(wc))
print(f"RegisterClassW atom: {atom}")

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
print(f"CreateWindowExW hwnd: {hwnd}")
if not hwnd:
    print(f"GetLastError: {kernel32.GetLastError()}")
    import sys; sys.exit(1)

ui_font = gdi32.CreateFontW(-14, 0, 0, 0, 400, 0, 0, 0, 0, 0, 0, 0, 0, "Segoe UI")

h_label = user32.CreateWindowExW(
    0, "STATIC",
    "Select a paired device to monitor for auto-reconnect:",
    WS_CHILD | WS_VISIBLE,
    14, 12, 380, 22,
    hwnd, None, wc.hInstance, None,
)
user32.SendMessageW(h_label, WM_SETFONT, ui_font, 1)

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
print(f"Listbox hwnd: {h_listbox}")

for i, dev in enumerate(devices):
    label = dev["name"]
    ret = user32.SendMessageW(h_listbox, LB_ADDSTRING, 0, label)
    print(f"  Added '{label}' -> index {ret}")

user32.SendMessageW(h_listbox, LB_SETCURSEL, 0, 0)

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

print("Running message loop...")
msg = ctypes.wintypes.MSG()
while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
    if not user32.IsDialogMessageW(hwnd, ctypes.byref(msg)):
        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageW(ctypes.byref(msg))

print(f"Result: {result_holder[0]}")
if result_holder[0] is not None:
    print(f"Selected device: {devices[result_holder[0]]['name']}")
