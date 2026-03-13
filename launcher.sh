#!/bin/bash
# ═══════════════════════════════════════════════
# SEO Library Launcher
# Lit apps.json et propose les outils disponibles
# ═══════════════════════════════════════════════

export PATH="/usr/local/bin:/opt/homebrew/bin:/usr/bin:/bin:$PATH"

# Trouver le dossier de la library (là où est ce script)
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CONFIG="$SCRIPT_DIR/config/apps.json"

if [ ! -f "$CONFIG" ]; then
    osascript -e 'display alert "Erreur" message "Fichier config/apps.json introuvable" as critical'
    exit 1
fi

# Lire les apps depuis apps.json via Python
APPS_LIST=$(python3 -c "
import json
with open('$CONFIG') as f:
    data = json.load(f)
for app in data['apps']:
    print(app['name'])
print('⚙️ Tout lancer')
")

# Convertir en liste AppleScript
AS_LIST=$(echo "$APPS_LIST" | python3 -c "
import sys
items = [l.strip() for l in sys.stdin if l.strip()]
print('{' + ', '.join(['\"' + i + '\"' for i in items]) + '}')
")

# Afficher le menu
CHOICE=$(osascript -e "
tell application \"System Events\"
    set choix to choose from list $AS_LIST with title \"SEO Library\" with prompt \"Quel outil lancer ?\" default items {item 1 of $AS_LIST}
    if choix is false then return \"cancel\"
    return item 1 of choix
end tell
")

[ "$CHOICE" = "cancel" ] && exit 0

# Lire la config de l'app choisie
APP_CONFIG=$(python3 -c "
import json
with open('$CONFIG') as f:
    data = json.load(f)
for app in data['apps']:
    if app['name'] == '''$CHOICE''':
        print(app['type'])
        print(app.get('path', ''))
        print(app.get('port', ''))
        print(app.get('prompt', ''))
        break
elif '''$CHOICE'''.startswith('⚙️'):
    print('all')
    print('')
    print('')
    print('')
")

APP_TYPE=$(echo "$APP_CONFIG" | sed -n '1p')
APP_PATH=$(echo "$APP_CONFIG" | sed -n '2p')
APP_PORT=$(echo "$APP_CONFIG" | sed -n '3p')
APP_PROMPT=$(echo "$APP_CONFIG" | sed -n '4p')

launch_streamlit() {
    local path="$1"
    local port="$2"
    cd "$SCRIPT_DIR/$(dirname "$path")"
    open "http://localhost:$port"
    streamlit run "$(basename "$path")" --server.port "$port" --server.headless true
}

launch_cli() {
    local path="$1"
    local prompt="$2"
    local full_path="$SCRIPT_DIR/$path"
    local dir=$(dirname "$full_path")
    local file=$(basename "$full_path")
    osascript -e "
    tell application \"Terminal\"
        activate
        do script \"cd '$dir' && echo '' && echo '🛒 SEO Library — $CHOICE' && echo '──────────────────────────────' && echo '' && read -p '$prompt: ' user_input && python3 '$file' \\\"\\\$user_input\\\"\"
    end tell
    "
}

case "$APP_TYPE" in
    streamlit)
        launch_streamlit "$APP_PATH" "$APP_PORT"
        ;;
    cli_interactive)
        launch_cli "$APP_PATH" "$APP_PROMPT"
        ;;
    all)
        # Lancer toutes les apps
        python3 -c "
import json
with open('$CONFIG') as f:
    data = json.load(f)
for app in data['apps']:
    print(app['type'] + '|' + app.get('path','') + '|' + app.get('port','') + '|' + app.get('prompt',''))
" | while IFS='|' read -r type path port prompt; do
            case "$type" in
                streamlit)
                    cd "$SCRIPT_DIR/$(dirname "$path")"
                    streamlit run "$(basename "$path")" --server.port "$port" --server.headless true &
                    sleep 2
                    open "http://localhost:$port"
                    ;;
                cli_interactive)
                    launch_cli "$path" "$prompt"
                    ;;
            esac
        done
        wait
        ;;
esac
