@echo off
echo Creating update package for QR Scanner...

rem 1. Nettoyer les anciens builds
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

rem 2. Activer l'environnement virtuel
call venv\Scripts\activate

rem 3. Compiler l'updater stub d'abord
echo Building updater stub...
pyinstaller updater_stub.spec

rem 4. Vérifier que la compilation de l'updater a réussi
if not exist "dist\updater\updater_stub.exe" (
    echo Error: Updater stub build failed!
    pause
    exit /b 1
)

rem 5. Compiler l'application principale
echo Building main application...
pyinstaller qr_scanner.spec

rem 6. Vérifier que la compilation de l'application a réussi
if not exist "dist\Mouvement Stock" (
    echo Error: Main application build failed, dist\Mouvement Stock not found!
    pause
    exit /b 1
)

rem 7. Créer le dossier updater dans l'application
echo Creating updater directory...
if not exist "dist\Mouvement Stock\updater" mkdir "dist\Mouvement Stock\updater"

rem 8. Copier l'updater stub et ses dépendances
echo Copying updater stub to application directory...
copy "dist\updater\updater_stub.exe" "dist\Mouvement Stock\updater\"
if exist "dist\updater\_internal" (
    echo Copying updater _internal folder...
    xcopy /E /I /Y "dist\updater\_internal" "dist\Mouvement Stock\updater\_internal\"
)

rem 9. Vérifier que l'updater a été copié
if not exist "dist\Mouvement Stock\updater\updater_stub.exe" (
    echo Error: Failed to copy updater stub!
    pause
    exit /b 1
)

echo ✅ Updater stub integrated successfully!

rem 10. Créer l'archive ZIP avec PowerShell
echo Creating ZIP package...
powershell -Command "Compress-Archive -Path 'dist\Mouvement Stock' -DestinationPath 'dist\Mouvement Stock.zip' -Force"

rem 11. Vérifier que l'archive a été créée
if exist "dist\Mouvement Stock.zip" (
    echo ✅ Package created successfully: dist\Mouvement Stock.zip
    for %%i in ("dist\Mouvement Stock.zip") do echo 📦 Package size: %%~zi bytes
    echo.
    echo Package contents:
    echo - Main application: Mouvement Stock.exe
    echo - Application dependencies: _internal\
    echo - Updater: updater\updater_stub.exe
    echo - Updater dependencies: updater\_internal\
    echo.
    echo Next steps:
    echo 1. Upload dist\Mouvement Stock.zip to your GitHub repository
    echo 2. Update the version number in version.txt  
    echo 3. Make sure REMOTE_PACKAGE_URL points to the correct ZIP file
    pause
) else (
    echo ❌ Error: Failed to create package!
    pause
    exit /b 1
)
