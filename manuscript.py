import re


def is_chapter_heading(text: str) -> bool:
    stripped_text = text.strip()

    if stripped_text in {"Prologue", "Epilogue"}:
        return True

    return re.fullmatch(
        r"\d+\.\s+\S+",
        stripped_text,
    ) is not None