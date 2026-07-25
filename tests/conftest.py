import sys
from pathlib import Path

# Make the src layout importable without an install step.
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
