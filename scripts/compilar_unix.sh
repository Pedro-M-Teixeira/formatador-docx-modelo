#!/usr/bin/env bash
# ---------------------------------------------------------------------------
#  Gera o executável FormatadorDocx (macOS e Linux).
#  Utilização:  bash scripts/compilar_unix.sh
#  O resultado fica em dist/
# ---------------------------------------------------------------------------
set -euo pipefail
cd "$(dirname "$0")/.."

echo "=== A instalar dependências ==="
python3 -m pip install --upgrade pip
python3 -m pip install -r requirements.txt pyinstaller

echo "=== A executar o teste de fumo ==="
python3 scripts/teste_fumo.py

echo "=== A gerar o executável ==="
python3 -m PyInstaller build.spec --noconfirm --clean

echo
echo "Concluído. O resultado está em: $(pwd)/dist/"
