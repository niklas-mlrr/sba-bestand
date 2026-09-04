"""pytest-Fixtures um die geteilten Bausteine aus ``bestand.core.testing``."""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from bestand.core.testing import (  # noqa: E402,F401
    ISBN_DEUTSCH_5,
    ISBN_DEUTSCH_6,
    ISBN_DEUTSCH_7,
    ISBN_DEUTSCH_12,
    ISBN_ERDKUNDE_7,
    ISBN_ERDKUNDE_56,
    ISBN_ERDKUNDE_EA_12,
    SHEET_NAME,
    FakeClient,
    build_workbook,
)


@pytest.fixture()
def workbook_path(tmp_path: Path) -> Path:
    return build_workbook(tmp_path / "Bestand-Test.xlsx")


@pytest.fixture()
def fake_client() -> FakeClient:
    return FakeClient()
