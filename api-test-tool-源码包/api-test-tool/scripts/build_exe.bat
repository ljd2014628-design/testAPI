@echo off
chcp 65001 >nul
cd /d %~dp0\..
echo === 安装依赖 ===
python -m pip install -r requirements.txt pyinstaller
echo === 开始打包 ===
python -m PyInstaller --onefile --windowed --noconfirm --clean --name ApiTestTool gui.py
if exist dist\ApiTestTool.exe (
    echo.
    echo 打包完成：dist\ApiTestTool.exe
) else (
    echo.
    echo 打包失败，请检查上面的报错
)
pause
