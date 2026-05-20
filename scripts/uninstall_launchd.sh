#!/usr/bin/env bash
# Удалить launchd-задачи. Не удаляет файлы проекта.
set -euo pipefail
LD="$HOME/Library/LaunchAgents"
for L in com.alexei.hronos_torfa com.alexei.hronos_torfa_bot; do
  launchctl bootout "gui/$(id -u)/${L}" 2>/dev/null || true
  rm -f "$LD/${L}.plist"
  echo "✓ удалён $L"
done
