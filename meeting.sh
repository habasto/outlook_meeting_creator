#!/usr/bin/env bash
set -euo pipefail

show_help() {
cat <<'EOF'
Outlook Meeting CLI Pro (WSL + Outlook clásico)

Uso:
  ./meeting.sh --profile "Empresa" --subject "Reunión" --date 2026-04-25 --time 15:00 \
    --duration 60 --attendees "a@x.com;b@x.com" --location "Teams"

Opciones:
  --profile     Nombre referencial del perfil/cuenta
  --subject     Motivo/asunto de la reunión
  --date        Fecha YYYY-MM-DD
  --time        Hora HH:MM (24h)
  --duration    Minutos (default 60)
  --attendees   Emails separados por ;
  --location    Lugar / Teams / Sala
  --body        Texto adicional
  --send        Envía directamente (sin vista previa)
  --help        Ayuda
EOF
}

PROFILE=""
SUBJECT=""
DATE=""
TIME=""
DURATION="60"
ATTENDEES=""
LOCATION="Teams"
BODY=""
SEND="0"

while [[ $# -gt 0 ]]; do
  case "$1" in
    --profile) PROFILE="$2"; shift 2;;
    --subject) SUBJECT="$2"; shift 2;;
    --date) DATE="$2"; shift 2;;
    --time) TIME="$2"; shift 2;;
    --duration) DURATION="$2"; shift 2;;
    --attendees) ATTENDEES="$2"; shift 2;;
    --location) LOCATION="$2"; shift 2;;
    --body) BODY="$2"; shift 2;;
    --send) SEND="1"; shift;;
    --help|-h) show_help; exit 0;;
    *) echo "Parámetro desconocido: $1"; exit 1;;
  esac
done

[[ -z "$SUBJECT" || -z "$DATE" || -z "$TIME" ]] && { show_help; exit 1; }

if ! [[ "$DATE" =~ ^[0-9]{4}-[0-9]{2}-[0-9]{2}$ ]]; then echo "Fecha inválida"; exit 1; fi
if ! [[ "$TIME" =~ ^[0-9]{2}:[0-9]{2}$ ]]; then echo "Hora inválida"; exit 1; fi
if ! [[ "$DURATION" =~ ^[0-9]+$ ]]; then echo "Duración inválida"; exit 1; fi

START="$DATE $TIME"

PS_SCRIPT=$(cat <<EOF
\$outlook = New-Object -ComObject Outlook.Application
\$ns = \$outlook.GetNamespace("MAPI")
\$appt = \$outlook.CreateItem(1)
\$appt.MeetingStatus = 1
\$appt.Subject = "$SUBJECT"
\$appt.Start = "$START"
\$appt.Duration = $DURATION
\$appt.Location = "$LOCATION"
\$appt.Body = "Perfil esperado: $PROFILE`r`n$BODY"
if ("$ATTENDEES" -ne "") {
  "$ATTENDEES".Split(";") | ForEach-Object {
    if (\$_.Trim() -ne "") { \$appt.Recipients.Add(\$_.Trim()) | Out-Null }
  }
}
if ($SEND -eq 1) { \$appt.Send() } else { \$appt.Display() }
EOF
)

powershell.exe -NoProfile -Command "$PS_SCRIPT"
