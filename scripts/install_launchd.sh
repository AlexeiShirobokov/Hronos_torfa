#!/usr/bin/env bash
# Установка/обновление launchd-задач для пайплайна и бота.
# Запуск:
#   bash scripts/install_launchd.sh
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
LD="$HOME/Library/LaunchAgents"
mkdir -p "$LD"

for L in com.alexei.hronos_torfa com.alexei.hronos_torfa_bot; do
  PLIST="$HERE/${L}.plist"
  DEST="$LD/${L}.plist"
  echo "→ копирую $PLIST → $DEST"
  cp -f "$PLIST" "$DEST"
  # bootout если уже загружен (mac OS 10.10+)
  launchctl bootout "gui/$(id -u)/${L}" 2>/dev/null || true
  launchctl bootstrap "gui/$(id -u)" "$DEST"
  launchctl enable "gui/$(id -u)/${L}"
  echo "✓ $L загружен"
done

echo
echo "Проверка:"
launchctl list | grep -i hronos || true
echo
echo "Логи будут в /Users/alexei/Claude_v1/logs/launchd_*.log"
