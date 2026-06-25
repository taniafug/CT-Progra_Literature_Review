import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]

STAGE_DIR = PROJECT_ROOT / "stage_4_full_pairing"

if str(STAGE_DIR) not in sys.path:
    sys.path.insert(0, str(STAGE_DIR))


from app_pairing_clean import main


main()