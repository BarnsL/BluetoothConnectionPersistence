# Bluetooth Connection Persistence

A lightweight Windows system tray application that automatically reconnects your Bluetooth devices when they disconnect — including after sleep, hibernation, or reboot.

## Features

- **Automatic Reconnection** — Monitors selected Bluetooth devices and reconnects them when disconnected
- **Survives Reboots** — Optionally starts with Windows via registry (current-user, no admin needed)
- **Sleep/Wake Aware** — Listens for Windows power events and immediately attempts reconnection after resume
- **System Tray App** — Runs quietly in the background with a tray icon for management
- **Exponential Back-off** — Avoids spamming reconnection attempts with intelligent retry intervals
- **Multi-Device Support** — Monitor as many paired Bluetooth devices as you want
- **Single Instance** — Prevents duplicate processes via a named mutex
- **Secure** — No network access, no telemetry, no elevated privileges required; config stored in user AppData

## Installation

### Option 1: Download the EXE (Recommended)

1. Go to the [Releases](../../releases) page
2. Download `BluetoothConnectionPersistence.exe`
3. Run it — the app will appear in your system tray
4. Right-click the tray icon → **Add Device...** to select a Bluetooth device to monitor

### Option 2: Run from Source

```powershell
# Clone the repo
git clone https://github.com/YOUR_USERNAME/BluetoothConnectionPersistence.git
cd BluetoothConnectionPersistence

# Create virtual environment
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# Install dependencies
pip install -r requirements.txt

# Run
python bt_persistence.py
```

### Option 3: Build the EXE Yourself

```powershell
pip install -r requirements.txt
pip install pyinstaller
pyinstaller bt_persistence.spec
```

The EXE will be in the `dist/` folder.

## Usage

| Action | How |
|---|---|
| **Add a device** | Right-click tray icon → Add Device... |
| **Remove a device** | Right-click tray icon → Monitored Devices → click the device |
| **Enable/disable auto-start** | Right-click tray icon → Start with Windows |
| **View logs** | Right-click tray icon → View Log |
| **Quit** | Right-click tray icon → Quit |

## How It Works

1. On startup, the app loads your saved device list from `%APPDATA%\BluetoothConnectionPersistence\config.json`
2. A background thread polls each monitored device's PnP status every 5 seconds
3. If a device is disconnected, it attempts reconnection by toggling the device's PnP state
4. A separate thread listens for `WM_POWERBROADCAST` messages to detect sleep/wake transitions
5. On wake, all back-off timers are reset for immediate reconnection

## Configuration

Config is stored at:
```
%APPDATA%\BluetoothConnectionPersistence\config.json
```

Example:
```json
{
  "devices": [
    {
      "name": "WH-1000XM5",
      "instance_id": "BTHENUM\\{0000111e-0000-1000-8000-00805f9b34fb}_LOCALMFG&0047\\7&1234..."
    }
  ],
  "auto_start": true
}
```

## Security

- **No network access** — The app never connects to the internet
- **No elevated privileges** — Runs as the current user; startup is via `HKCU` registry (not `HKLM`)
- **No telemetry or data collection**
- **Input validation** — Instance IDs are validated before use in commands
- **Single instance enforcement** — Named mutex prevents duplicate processes
- **Minimal dependencies** — Only well-known, widely-audited packages

## Requirements

- Windows 10 or later
- Bluetooth adapter
- Paired Bluetooth device(s)

## License

MIT License — see [LICENSE](LICENSE) for details.
