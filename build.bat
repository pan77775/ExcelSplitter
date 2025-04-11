@echo off
echo 正在安裝必要的套件...
python -m pip install pyinstaller pandas openpyxl

echo 正在打包程式...
python -m PyInstaller --clean --noconfirm excel_splitter.spec

echo 打包完成！
echo 可執行檔位於 dist 資料夾中
pause 