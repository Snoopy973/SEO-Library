import sys
import os

# Ajouter le dossier racine du projet au path pour importer app.py
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app import main

main()
