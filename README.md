# Outlook Meeting CLI Pro

Crear reuniones de Outlook desde **Tilix/WSL** usando Bash que llama a PowerShell + Outlook COM.

## Características

- Parámetros por línea de comandos
- Vista previa o envío directo
- Validación básica de fecha/hora
- Invitados múltiples
- Compatible con WSL

## Requisitos

- Windows con Outlook clásico instalado
- WSL Ubuntu
- `powershell.exe` disponible desde WSL

## Uso

```bash
chmod +x meeting.sh
./meeting.sh --profile "Empresa" --subject "Reunión semanal" --date 2026-04-25 --time 15:00 --duration 45 --attendees "a@x.com;b@x.com" --location "Teams"
```

## Envío directo

```bash
./meeting.sh ... --send
```

## Nota

Outlook COM normalmente usa el perfil actualmente abierto o el predeterminado.
