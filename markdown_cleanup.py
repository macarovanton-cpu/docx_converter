import re


_IP_ARTIFACTS = {
    "[Р68": "IP68",
    "[P68": "IP68",
    "[Р67": "IP67",
    "[P67": "IP67",
}

_TECHNICAL_HEADINGS = (
    "Навес:",
    "Помещение управления:",
    "Фундамент:",
    "Тензодатчики:",
    "Весовой терминал:",
    "Функции ПО:",
    "Интеграция:",
    "Комплект оборудования:",
    "Защита и питание:",
    "Оптика:",
)


def cleanup_ocr_markdown(text: str) -> str:
    """Apply conservative deterministic cleanup to raw OCR Markdown."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\f", "")

    text = re.sub(r";(?:e|•|·)\s+", "\n- ", text)
    text = re.sub(r"\.e\s+(?=[A-ZА-ЯЁа-яё])", "\n- ", text)

    text = re.sub(
        r"([^\n])\s+(?=(?:1 Основание|2 Заказчик|3 Стадия)\b)",
        r"\1\n\n",
        text,
    )
    text = re.sub(
        rf"([^\n\s])(?=(?:{'|'.join(map(re.escape, _TECHNICAL_HEADINGS))}))",
        r"\1\n",
        text,
    )

    text = re.sub(r"\bOOO\b", "ООО", text)
    for artifact, replacement in _IP_ARTIFACTS.items():
        text = text.replace(artifact, replacement)
    text = text.replace("He «слепла»", "не «слепла»")

    cleaned_lines = []
    for line in text.split("\n"):
        line = re.sub(r"[ \t]+", " ", line).strip()
        line = re.sub(r"^(?:e|•|·)\s+(.+)$", r"- \1", line)
        line = re.sub(r"^No\.?\s*(?=\d)", "№ ", line)
        cleaned_lines.append(line)

    text = "\n".join(cleaned_lines)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text
