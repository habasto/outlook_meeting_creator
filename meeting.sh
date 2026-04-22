#!/usr/bin/env bash
set -euo pipefail

show_help() {
cat <<EOF
Uso:
./meeting.sh --profile "Empresa" --subject "Reunión" --date 2026-04-25 --time 15:00 \
--duration 60 --attendees "a@x.com;b@x.com" --location "Teams" --body "Texto"

Opciones:
--profile   Cuenta/perfil esperado
--subject   Asunto
--date      YYYY-MM-DD
--time      HH:MM
--duration  Minutos
--attendees Emails separados por ;
--location  Lugar
--body      Mensaje
--send      Enviar directo
--help      Ayuda
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

START="$DATE $TIME"

echo "Creando reunión..."
echo "Perfil esperado: $PROFILE"

powershell.exe -NoProfile -Command "
\$outlook = New-Object -ComObject Outlook.Application;
\$ns = \$outlook.GetNamespace('MAPI');
\$acct = \$ns.Accounts.Item(1);
Write-Host ('Cuenta activa Outlook: ' + \$acct.DisplayName);
\$appt = \$outlook.CreateItem(1);
\$appt.MeetingStatus = 1;
\$appt.Subject = '$SUBJECT';
\$appt.Start = '$START';
\$appt.Duration = $DURATION;
\$appt.Location = '$LOCATION';
\$appt.Body = '$BODY';
if ('$ATTENDEES' -ne '') {
 '$ATTENDEES'.Split(';') | ForEach-Object {
   if (\$_.Trim() -ne '') { \$appt.Recipients.Add(\$_.Trim()) | Out-Null }
 }
}
if ($SEND -eq 1) { \$appt.Send(); Write-Host 'Reunión enviada.' }
else { \$appt.Display(); Write-Host 'Reunión abierta para revisión.' }
"
