from __future__ import annotations

from pathlib import Path

from .ui import run


def main() -> None:
    root = Path(__file__).resolve().parent.parent
    config = root / "config.json"
    example = root / "config.example.json"
    if not config.exists() and example.exists():
        config.write_text(example.read_text(encoding="utf-8"), encoding="utf-8")
    run()


if __name__ == "__main__":
    main()
