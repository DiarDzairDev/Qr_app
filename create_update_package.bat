@echo off
echo Creating update package for QR Scanner...

rem 1. Nettoyer les anciens builds
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

rem 2. Activer l'environnement virtuel
call venv\Scripts\activate

rem 3. Compiler l'application
echo Building application...
pyinstaller qr_scanner.spec

rem 4. Vérifier que la compilation a réussi
if not exist "dist\qr_scanner" (
    echo Error: Build failed, dist\qr_scanner not found!
    pause
    exit /b 1
)

rem 5. Créer l'archive ZIP avec PowerShell
echo Creating ZIP package...
powershell -Command "Compress-Archive -Path 'dist\qr_scanner' -DestinationPath 'dist\qr_scanner.zip' -Force"

rem 6. Vérifier que l'archive a été créée
if exist "dist\qr_scanner.zip" (
    echo ✅ Package created successfully: dist\qr_scanner.zip
    for %%i in ("dist\qr_scanner.zip") do echo 📦 Package size: %%~zi bytes
    echo.
    echo Next steps:
    echo 1. Upload dist\qr_scanner.zip to your GitHub repository
    echo 2. Update the version number in version.txt
    echo 3. Make sure REMOTE_PACKAGE_URL points to the correct ZIP file
    pause
) else (
    echo ❌ Error: Failed to create package!
    pause
    exit /b 1
)
