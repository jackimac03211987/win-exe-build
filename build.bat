@echo off
echo Building PDF Watermark Tool...

REM 清理旧的构建
if exist "build" rmdir /S /Q "build"
if exist "dist" rmdir /S /Q "dist"

REM 安装依赖
echo Installing dependencies...
pip install --upgrade pip
pip install pillow pandas openpyxl pdf2image reportlab PyPDF2 numpy pyinstaller

REM 验证 openpyxl 安装
python -c "import openpyxl; print('openpyxl installed successfully')"

REM 下载 Poppler（如果不存在）
if not exist "poppler" (
    echo Downloading Poppler...
    curl -L https://github.com/oschwartz10612/poppler-windows/releases/download/v24.08.0-0/Release-24.08.0-0.zip -o poppler.zip
    tar -xf poppler.zip
    mkdir poppler
    xcopy /E /Y "poppler-24.08.0\Library\bin\*" "poppler\"
    rmdir /S /Q "poppler-24.08.0"
    del poppler.zip
)

REM 验证 Poppler 文件
echo Checking Poppler files...
dir poppler\*.exe

REM 打包（使用更详细的参数）
echo Building executable...
pyinstaller --clean --noconfirm ^
    --name="PDFWatermark" ^
    --windowed ^
    --onedir ^
    --add-binary="poppler\*.exe;." ^
    --add-binary="poppler\*.dll;." ^
    --hidden-import="openpyxl" ^
    --hidden-import="openpyxl.cell._writer" ^
    --hidden-import="pandas.io.excel._openpyxl" ^
    --hidden-import="pdf2image" ^
    --hidden-import="pdf2image.pdf2image" ^
    --collect-all="openpyxl" ^
    --collect-all="pandas" ^
    src\app_main.py

echo.
echo Build complete! Check dist\PDFWatermark folder.
echo.
echo Verifying build output...
dir dist\PDFWatermark\*.exe
dir dist\PDFWatermark\pdftoppm.exe

pause
