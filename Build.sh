#!/bin/bash
# Script pour créer l'archive de mise à jour
# Usage: ./create_update_package.sh

echo "Creating update package for QR Scanner..."

# 1. Nettoyer les anciens builds
rm -rf build dist

# 2. Activer l'environnement virtuel
source venv/Scripts/activate

# 3. Compiler l'updater stub d'abord
echo "Building updater stub..."
pyinstaller updater_stub.spec

# 4. Vérifier que la compilation de l'updater a réussi
if [ ! -f "build/updater_stub/updater_stub.exe" ]; then
    echo "Error: Updater stub build failed!"
    exit 1
fi

# 5. Compiler l'application principale
echo "Building main application..."
pyinstaller qr_scanner.spec

# 6. Vérifier que la compilation de l'application a réussi
if [ ! -d "dist/Mouvement Stock" ]; then
    echo "Error: Main application build failed, dist/Mouvement Stock not found!"
    exit 1
fi

# 7. Créer le dossier updater dans l'application
echo "Creating updater directory..."
mkdir -p "dist/Mouvement Stock/updater"

# 8. Copier l'updater stub et ses dépendances
echo "Copying updater stub to application directory..."
cp "build/updater_stub/updater_stub.exe" "dist/Mouvement Stock/updater/"
if [ -d "build/updater_stub/_internal" ]; then
    echo "Copying updater _internal folder..."
    cp -r "build/updater_stub/_internal" "dist/Mouvement Stock/updater/"
fi

# 9. Vérifier que l'updater a été copié
if [ ! -f "dist/Mouvement Stock/updater/updater_stub.exe" ]; then
    echo "Error: Failed to copy updater stub!"
    exit 1
fi

echo "✅ Updater stub integrated successfully!"

# 10. Créer l'archive ZIP
echo "Creating ZIP package..."
cd dist
zip -r "Mouvement Stock.zip" "Mouvement Stock/"
cd ..

# 11. Vérifier que l'archive a été créée
if [ -f "dist/Mouvement Stock.zip" ]; then
    echo "✅ Package created successfully: dist/Mouvement Stock.zip"
    echo "📦 Package size: $(du -h "dist/Mouvement Stock.zip" | cut -f1)"
    echo ""
    echo "Package contents:"
    echo "- Main application: Mouvement Stock.exe"
    echo "- Application dependencies: _internal/"
    echo "- Updater: updater/updater_stub.exe"
    echo "- Updater dependencies: updater/_internal/"
    echo ""
    echo "Next steps:"
    echo "1. Upload dist/Mouvement Stock.zip to your GitHub repository"
    echo "2. Update the version number in version.txt"
    echo "3. Make sure REMOTE_PACKAGE_URL points to the correct ZIP file"
else
    echo "❌ Error: Failed to create package!"
    exit 1
fi
