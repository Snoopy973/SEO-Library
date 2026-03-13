#!/bin/bash
# ═══════════════════════════════════════════════
# SEO Library — Script d'installation
#
# Usage :
#   1. Télécharge/copie le dossier SEO-Library sur le Bureau
#   2. Double-clic sur install.command (ou lance : bash install.sh)
#   3. C'est prêt !
# ═══════════════════════════════════════════════

set -e

echo ""
echo "╔══════════════════════════════════════════╗"
echo "║     🛠  Installation SEO Library          ║"
echo "╚══════════════════════════════════════════╝"
echo ""

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
export PATH="/usr/local/bin:/opt/homebrew/bin:/usr/bin:/bin:$PATH"

# ── 1. Vérifier Python 3 ──
echo "🔍 Vérification de Python 3..."
if command -v python3 &>/dev/null; then
    echo "   ✅ Python 3 trouvé : $(python3 --version)"
else
    echo "   ❌ Python 3 non trouvé. Installation requise."
    echo "   → Télécharge depuis https://www.python.org/downloads/"
    exit 1
fi

# ── 2. Installer les dépendances ──
echo ""
echo "📦 Installation des dépendances Python..."
DEPS=$(python3 -c "
import json
with open('$SCRIPT_DIR/config/apps.json') as f:
    data = json.load(f)
print(' '.join(data.get('dependencies', [])))
")

for dep in $DEPS; do
    if python3 -c "import $dep" 2>/dev/null; then
        echo "   ✅ $dep déjà installé"
    else
        echo "   📥 Installation de $dep..."
        pip3 install "$dep" --quiet
        echo "   ✅ $dep installé"
    fi
done

# ── 3. Rendre les scripts exécutables ──
echo ""
echo "🔧 Configuration des permissions..."
chmod +x "$SCRIPT_DIR/launcher.sh"
echo "   ✅ Permissions OK"

# ── 4. Créer l'app pour le Dock ──
echo ""
echo "🖥  Création de l'app SEO Library..."

APP_DIR="/Applications/SEO Library.app"
rm -rf "$APP_DIR"
CONTENTS="$APP_DIR/Contents"
mkdir -p "$CONTENTS/MacOS" "$CONTENTS/Resources"

cat > "$CONTENTS/Info.plist" << 'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>launcher</string>
    <key>CFBundleIdentifier</key>
    <string>com.seo-library.launcher</string>
    <key>CFBundleName</key>
    <string>SEO Library</string>
    <key>CFBundleDisplayName</key>
    <string>SEO Library</string>
    <key>CFBundleVersion</key>
    <string>1.0</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>NSHighResolutionCapable</key>
    <true/>
</dict>
</plist>
PLIST

cat > "$CONTENTS/MacOS/launcher" << EXEC
#!/bin/bash
export PATH="/usr/local/bin:/opt/homebrew/bin:/usr/bin:/bin:\$PATH"
LIBRARY_DIR="$SCRIPT_DIR"
if [ -f "\$LIBRARY_DIR/launcher.sh" ]; then
    bash "\$LIBRARY_DIR/launcher.sh"
else
    osascript -e "display alert \"SEO Library introuvable\" message \"Le dossier SEO-Library a été déplacé ou supprimé.\" as critical"
fi
EXEC

chmod +x "$CONTENTS/MacOS/launcher"
echo "   ✅ App créée dans /Applications/"

# ── 5. Ajouter au Dock ──
echo ""
echo "📌 Ajout au Dock..."

python3 << 'PYDOCK'
import plistlib, os

dock_plist = os.path.expanduser("~/Library/Preferences/com.apple.dock.plist")
with open(dock_plist, "rb") as f:
    dock = plistlib.load(f)

# Retirer les anciennes versions
dock["persistent-apps"] = [
    a for a in dock["persistent-apps"]
    if "SEO" not in str(a.get("tile-data", {}).get("file-label", ""))
]

# Ajouter
new_app = {
    "tile-data": {
        "file-data": {
            "_CFURLString": "file:///Applications/SEO%20Library.app/",
            "_CFURLStringType": 15
        },
        "file-label": "SEO Library"
    },
    "tile-type": "file-tile"
}
dock["persistent-apps"].append(new_app)

with open(dock_plist, "wb") as f:
    plistlib.dump(dock, f)
PYDOCK

killall Dock 2>/dev/null
echo "   ✅ SEO Library ajoutée au Dock"

# ── Terminé ──
echo ""
echo "╔══════════════════════════════════════════╗"
echo "║     ✅  Installation terminée !           ║"
echo "║                                          ║"
echo "║  → Clique sur 'SEO Library' dans le Dock ║"
echo "║  → Choisis l'outil à lancer              ║"
echo "╚══════════════════════════════════════════╝"
echo ""
