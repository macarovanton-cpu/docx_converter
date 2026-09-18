import os

import pytest


def pytest_configure(config):
    config.addinivalue_line(
        "markers", "live: ходит в сеть (MinerU); без MINERU_API_KEY пропускается")


def pytest_collection_modifyitems(config, items):
    if os.environ.get("MINERU_API_KEY"):
        return
    skip = pytest.mark.skip(reason="нет MINERU_API_KEY")
    for item in items:
        if "live" in item.keywords:
            item.add_marker(skip)
