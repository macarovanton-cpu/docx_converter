import os

import pytest

# ПРАВКА #71: сеть включает только явный флаг. Ключ в окружении живёт постоянно,
# и на нём голый `pytest` молча уходил в облако и жёг квоту.
LIVE_ENV = "MINERU_LIVE"
SKIP_REASON = (f"живые тесты выключены: нет {LIVE_ENV}=1 "
               f"(одного MINERU_API_KEY мало — сеть включается флагом)")


def pytest_configure(config):
    config.addinivalue_line(
        "markers", f"live: ходит в сеть (MinerU); идёт только при {LIVE_ENV}=1")


def pytest_collection_modifyitems(config, items):
    if os.environ.get(LIVE_ENV) == "1":
        return
    skip = pytest.mark.skip(reason=SKIP_REASON)
    for item in items:
        if "live" in item.keywords:
            item.add_marker(skip)
