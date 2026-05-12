#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)"
cd "$ROOT_DIR"
export PYINSTALLER_CONFIG_DIR="$ROOT_DIR/.pyinstaller-cache"

if [ ! -x ".venv-packaging/bin/python" ]; then
  python3 -m venv .venv-packaging
fi

.venv-packaging/bin/python -m pip install -r requirements-packaging.txt
rm -rf .packaging-storage
mkdir -p .packaging-storage/uploads .packaging-storage/outputs
.venv-packaging/bin/python -m PyInstaller \
  --clean \
  --noconfirm \
  --windowed \
  --name "公文格式助手" \
  --add-data "public:public" \
  --add-data ".packaging-storage:storage" \
  desktop_app.py

echo "macOS app generated at: $ROOT_DIR/dist/公文格式助手.app"
