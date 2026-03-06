import subprocess, json

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

print(f"Filtered devices ({len(seen_names)}):")
for dev in seen_names.values():
    print(f"  {dev['name']:30s} | {dev['status']:10s} | {dev['instance_id'][:60]}")
