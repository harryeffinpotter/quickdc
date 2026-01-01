# QuickDC

Tool for the **ARC key glitch**.

## Features

- **Single hotkey** - Press once to disconnect, press again to reconnect
- **Two modes:**
  - **WiFi** - Disconnects/reconnects WiFi (like clicking the WiFi button)
  - **Ethernet** - Blocks all traffic via Windows Firewall (equivalent to pulling cable)
- **Auto-reconnect** - Optionally auto-reconnect after X milliseconds
- **Spam E** - Configurable key spam with millisecond precision

## How It Works

1. Set your hotkey
2. Choose WiFi or Ethernet mode
3. Enable/disable auto-reconnect
4. Click Start
5. Press hotkey to disconnect
6. Press hotkey again to reconnect (or wait for auto-reconnect)

## Technical Details

- Uses **Windows SendInput API** for fastest possible key presses
- Uses **high-precision timing** (`perf_counter` + busy-wait) for accurate millisecond intervals
- WiFi mode uses `netsh wlan disconnect/connect`
- Ethernet mode uses Windows Firewall rules to block all traffic instantly - no graceful shutdown, identical to unplugging cable
- Firewall rules are automatically cleaned up on exit

## Requirements

- Windows 10/11
- Run as Administrator

## Build from Source
Note: Must be run in WINDOWS
```bash
# Install dependencies
pip install -r requirements.txt
pip install pyinstaller

# Build exe
pyinstaller --onefile --windowed --name QuickDC --uac-admin quickdc.py
```

The exe will be in the `dist` folder.
