import re


_IP_ARTIFACTS = {
    "[Р68": "IP68",
    "[P68": "IP68",
    "[Р67": "IP67",
    "[P67": "IP67",
}

def cleanup_ocr_markdown(text: str) -> str:
    """Apply conservative deterministic cleanup to raw OCR Markdown."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\f", "")

    text = re.sub(r";(?:e|•|·)\s+", "\n- ", text)
    text = re.sub(r"(?<=[0-9.,;:!?+\-–—)\]\}])e\s+(?=[A-ZА-ЯЁ])", "\n- ", text)

    text = re.sub(r"\bOOO\b", "ООО", text)
    for artifact, replacement in _IP_ARTIFACTS.items():
        text = text.replace(artifact, replacement)

    cleaned_lines = []
    for line in text.split("\n"):
        line = re.sub(r"[ \t]+", " ", line).strip()
        line = re.sub(r"^(?:e|•|·)\s+(.+)$", r"- \1", line)
        line = re.sub(r"\bNo\.?\s*(?=\d)", "№ ", line)
        cleaned_lines.append(line)

    text = "\n".join(cleaned_lines)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text
