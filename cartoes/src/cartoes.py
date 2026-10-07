"""Compatibility entry point for existing launchers importing cartoes.main."""
from resumos_cartoes.cli import main

__all__ = ["main"]

if __name__ == "__main__":
    raise SystemExit(main())
