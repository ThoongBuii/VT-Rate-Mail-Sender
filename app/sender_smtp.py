from __future__ import annotations

import re
from typing import Optional


def split_emails(raw: str) -> list[str]:
    return [p.strip() for p in re.split(r"[;,]+", raw or "") if p.strip()]
