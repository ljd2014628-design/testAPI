#!/bin/bash
# 注意：PyInstaller 不支持交叉编译，在 Linux/macOS 上打出来的是对应平台的可执行文件，
# 想要 Windows 版 exe 请在 Windows 上运行 build_exe.bat
cd "$(dirname "$0")/.."
python3 -m pip install -r requirements.txt pyinstaller
python3 -m PyInstaller --onefile --windowed --noconfirm --clean --name ApiTestTool gui.py
echo "产物：dist/ApiTestTool.exe"
