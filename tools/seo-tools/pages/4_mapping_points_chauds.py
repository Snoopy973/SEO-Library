"""
Page wrapper pour le Mapping des Points Chauds.
Importe et lance main() depuis l'app originale.
"""
import sys
from pathlib import Path

# Ajouter le dossier mapping-points-chauds au path pour les imports locaux
mapping_dir = Path(__file__).resolve().parent.parent.parent / "SEO-Library" / "tools" / "mapping-points-chauds"
if str(mapping_dir) not in sys.path:
    sys.path.insert(0, str(mapping_dir))

# Importer et lancer
from app import main as mapping_main
mapping_main()
