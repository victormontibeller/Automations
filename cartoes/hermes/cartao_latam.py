"""Copie este arquivo (não um symlink) para ~/.hermes/scripts/."""
import os
from pathlib import Path
import sys

from cartoes import main

config = Path(os.environ.get("CARTOES_CONFIG", "~/Projetos/Automations/cartoes/config.json")).expanduser()
sys.exit(main(["--config", str(config), "--card", "latam", "--scheduled", "--send"]))
