#!/bin/bash
# Script pour créer l'archive de mise à jour
# Usage: ./create_update_package.sh

echo "Creating update package for QR Scanner..."

# 1. Nettoyer les anciens builds
rm -rf build dist

# 2. Activer l'environnement virtuel
source venv/Scripts/activate

# 3. Compiler l'application
echo "Building application..."
pyinstaller qr_scanner.spec

# 4. Vérifier que la compilation a réussi
if [ ! -d "dist/qr_scanner" ]; then
    echo "Error: Build failed, dist/qr_scanner not found!"
    exit 1
fi

# 5. Créer l'archive ZIP
echo "Creating ZIP package..."
cd dist
zip -r qr_scanner.zip qr_scanner/
cd ..

# 6. Vérifier que l'archive a été créée
if [ -f "dist/qr_scanner.zip" ]; then
    echo "✅ Package created successfully: dist/qr_scanner.zip"
    echo "📦 Package size: $(du -h dist/qr_scanner.zip | cut -f1)"
    echo ""
    echo "Next steps:"
    echo "1. Upload dist/qr_scanner.zip to your GitHub repository"
    echo "2. Update the version number in version.txt"
    echo "3. Make sure REMOTE_PACKAGE_URL points to the correct ZIP file"
else
    echo "❌ Error: Failed to create package!"
    exit 1
fi
