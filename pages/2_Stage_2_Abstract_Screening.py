import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]

STAGE_DIR = PROJECT_ROOT / "stage_2_abstract_screening"

if str(STAGE_DIR) not in sys.path:
    sys.path.insert(0, str(STAGE_DIR))

from stage_2_app import main

main()