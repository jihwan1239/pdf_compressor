@echo off
chcp 65001 >nul
echo [1/3] 패키지 설치 중...
python -m pip install -r requirements.txt pyinstaller || goto :error
echo [2/3] exe 빌드 중...
python -m PyInstaller --noconfirm --onefile --windowed --name PDFCompressor --icon app.ico --version-file version_info.txt --add-data "app.ico;." pdf_compressor.py || goto :error
echo [3/3] 파일 해시(SHA-256) 계산 중...
certutil -hashfile dist\PDFCompressor.exe SHA256
echo.
echo 완료: dist\PDFCompressor.exe
pause
exit /b 0
:error
echo 빌드에 실패했습니다.
pause
exit /b 1
