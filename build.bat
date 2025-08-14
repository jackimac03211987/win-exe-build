@echo off
echo Building PDF Watermark Tool...

REM 安装依赖
pip install -r requirements.txt
pip install pyinstaller

REM 下载 Poppler
if not exist "poppler" (
    echo Downloading Poppler...
    curl -L https://github.com/oschwartz10612/poppler-windows/releases/download/v24.08.0-0/Release-24.08.0-0.zip -o poppler.zip
    tar -xf poppler.zip
    mkdir poppler
    xcopy /E /Y "poppler-24.08.0\Library\bin\*" "poppler\"
    rmdir /S /Q "poppler-24.08.0"
    del poppler.zip
)

REM 打包
pyinstaller --clean --noconfirm ^
    --name="PDFWatermark" ^
    --windowed ^
    --onedir ^
    --add-data="poppler;poppler" ^
    --hidden-import="openpyxl" ^
    --hidden-import="pandas.io.excel._openpyxl" ^
    --hidden-import="pdf2image" ^
    src/app_main.py

echo Build complete! Check dist/PDFWatermark folder.
pause
