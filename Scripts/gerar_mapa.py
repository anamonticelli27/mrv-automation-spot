"""Generate quote comparison map for a ticket."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from frete_spot.mapa import main

if __name__ == "__main__":
    main()
