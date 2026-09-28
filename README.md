# How Tall Are You

A small Flask app for collecting height measurements and viewing participant rankings.

## Run locally

```powershell
cd web_app
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
$env:FLASK_SECRET_KEY = [guid]::NewGuid().ToString('N')
$env:OWNER_PASSWORD = 'choose-a-strong-private-password'
$env:WEB_HOST = '127.0.0.1'
python app.py
```

For VPN-only access, find the computer's VPN IPv4 address with `ipconfig`, then set `WEB_HOST` to that address instead of `127.0.0.1`. Start the app on the host computer and share only `http://<vpn-ip>:5000` with VPN users. Do not use `0.0.0.0` unless the Windows Firewall rule below restricts port 5000 to the VPN subnet.

Run PowerShell as Administrator once on the host, using the Wits subnet shown by `ipconfig`:

```powershell
New-NetFirewallRule -DisplayName 'How Tall Are You Wits only' -Direction Inbound -Action Allow -Protocol TCP -LocalPort 5000 -RemoteAddress 10.202.118.0/24
```

Keep the VPN provider's device authorization enabled, do not port-forward port 5000 from the internet, and keep the owner password private. Participant records and uploaded photos are intentionally kept local and are not committed to Git.