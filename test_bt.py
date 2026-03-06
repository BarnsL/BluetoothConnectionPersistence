import subprocess, json

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
print('RETURNCODE:', result.returncode)
print('STDERR:', result.stderr[:500] if result.stderr else 'none')
raw = result.stdout.strip()
if raw:
    data = json.loads(raw)
    if isinstance(data, dict):
        data = [data]
    for d in data:
        name = d.get('FriendlyName', '?')
        status = d.get('Status', '?')
        iid = d.get('InstanceId', '?')
        print(f"  {name} | {status} | {iid[:60]}")
    print(f'Total raw entries: {len(data)}')
else:
    print('NO OUTPUT')
