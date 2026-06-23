"""Process Agilis .msg emails into the spreadsheet and Outlook drafts."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from frete_spot.pipeline import main

if __name__ == "__main__":
    main()
