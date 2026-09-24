#!/bin/zsh
set -euo pipefail
task_root="$(cd "$(dirname "$0")" && pwd)"
bundled_python="/Users/yanganru/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3"
if [[ -x "$bundled_python" ]]; then
  runtime_python="$bundled_python"
else
  runtime_python="$(command -v python3)"
fi
exec "$runtime_python" "$task_root/insurance_calculation_server.py" "$@"

