# Outlook Meeting Creator

Create Microsoft Outlook meetings directly from **WSL / Bash / Tilix** using PowerShell automation.

The use case is that can be used in while/for loop to create mutiple meeting meeting spam ;-)

Other use case is to create a meeing under some system alert, so firs time in the morning the IT guys have to go to the meeting room, sorry guys :-)

## Version

**v1.2 Limpia**

## Features

- Create Outlook calendar meetings from terminal
- Works from Ubuntu WSL / Tilix
- Uses PowerShell + Outlook COM
- Add multiple attendees
- Set date, time, duration and location
- Preview meeting before sending
- Direct send option
- Shows active Outlook account in console

## Requirements

- Windows 10 / 11
- Microsoft Outlook Classic Desktop installed
- WSL with Ubuntu
- `powershell.exe` available from WSL

## Files

- `meeting_v1.2.sh` → main script

## Usage

```bash
chmod +x meeting_v1.2.sh

./meeting_v1.2.sh \
--profile "Empresa" \
--subject "Reunión Comercial" \
--date 2026-04-25 \
--time 15:00 \
--duration 60 \
--attendees "a@x.com;b@x.com" \
--location "Teams" \
--body "Revisar pendientes"
```

## Direct Send

```bash
./meeting_v1.2.sh ... --send
```

## Notes

- Outlook usually uses the currently opened profile or the default profile.
- For best results, open Outlook first with the desired account.

## Example Workflow

WSL / Tilix → Bash Script → PowerShell → Outlook → Meeting Created

## License

MIT

## Author

Hubert Abasto
