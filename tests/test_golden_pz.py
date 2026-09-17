"""ПРАВКА #55: эталон оформления ПЗ.

Профиль оформления переписывает каждую ветку рендера, поэтому регрессию в ПЗ
ловим не точечными ассертами, а посимвольным сравнением XML с эталоном,
снятым с коммита 396b40c (до #55). Эталон обновляется ТОЛЬКО осознанно, вместе
с правкой, которая меняет оформление ПЗ, и диффом в ревью.

Обновить эталон: python tests/test_golden_pz.py --update
"""

import sys
import zipfile
from pathlib import Path

import pytest
from lxml import etree

from convert import convert_md_to_docx

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
FIXTURES = Path(__file__).resolve().parent / "fixtures"
SOURCE = FIXTURES / "pz_full.md"

# Части пакета, на которые влияет оформление. docProps/core.xml не берём —
# там таймстампы, а thumbnail и theme профиль не трогает.
PARTS = {
    "word/document.xml":  "golden_pz_document.xml",
    "word/numbering.xml": "golden_pz_numbering.xml",
    "word/settings.xml":  "golden_pz_settings.xml",
}
NORMAL_STYLE = "golden_pz_normal_style.xml"


def _build(out_path):
    convert_md_to_docx(SOURCE.read_text(encoding="utf-8"), str(out_path))
    with zipfile.ZipFile(out_path) as z:
        parts = {name: z.read(name) for name in PARTS}
        styles = etree.fromstring(z.read("word/styles.xml"))
    normal = styles.find(f'.//{{{W}}}style[@{{{W}}}styleId="Normal"]')
    parts["Normal"] = etree.tostring(normal, pretty_print=True)
    return parts


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    return _build(tmp_path_factory.mktemp("golden") / "pz.docx")


@pytest.mark.parametrize("part,golden", sorted(PARTS.items()) + [("Normal", NORMAL_STYLE)])
def test_pz_output_matches_golden(built, part, golden):
    """Оформление ПЗ не изменилось ни на байт."""
    expected = (FIXTURES / golden).read_bytes()
    actual = built[part]
    if actual == expected:
        return
    # Читаемое сообщение: первая разошедшаяся строка, а не 25 КБ XML в консоли
    exp_lines = expected.decode("utf-8").splitlines()
    act_lines = actual.decode("utf-8").splitlines()
    for i, (e, a) in enumerate(zip(exp_lines, act_lines), 1):
        if e != a:
            pytest.fail(f"{part}: расхождение в строке {i}\n"
                        f"эталон:  ...{e[:400]}\n"
                        f"получено: ...{a[:400]}")
    pytest.fail(f"{part}: разная длина — эталон {len(exp_lines)} строк, "
                f"получено {len(act_lines)}")


if __name__ == "__main__":
    if "--update" not in sys.argv:
        raise SystemExit("нужен флаг --update, чтобы перезаписать эталон")
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        parts = _build(Path(tmp) / "pz.docx")
    for part, golden in list(PARTS.items()) + [("Normal", NORMAL_STYLE)]:
        (FIXTURES / golden).write_bytes(parts[part])
        print(f"updated {golden}")
