# Standard Library imports
from dataclasses import dataclass
from pathlib import Path
from docx import Document
from collections.abc import Callable

# Third-Party imports
from PySide6.QtWidgets import (
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QLabel,
    QLineEdit,
    QTextEdit,
    QVBoxLayout,
    QWidget,
    QMenu,
)

# Classes
@dataclass
class TitlePage:
    title: str = ""
    subtitle: str = ""
    author: str = ""
    start_index: int | None = None
    end_index: int | None = None


@dataclass
class TextSection:
    text: str = ""
    source_file: str | None = None
    start_index: int | None = None
    end_index: int | None = None


@dataclass
class FrontMatter:
    title_page: TitlePage | None = None
    dedication: TextSection | None = None
    copyright: TextSection | None = None
    trigger_warnings: TextSection | None = None
    map_file: str | None = None


class DedicationDialog(QDialog):
    def __init__(
        self,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)

        self.setWindowTitle("Dedication")

        layout = QVBoxLayout(self)

        layout.addWidget(
            QLabel(
                "Enter the dedication text to find "
                "in the manuscript:"
            )
        )

        self.text_edit = QTextEdit()
        layout.addWidget(
            self.text_edit
        )

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )

        buttons.accepted.connect(
            self.accept
        )
        buttons.rejected.connect(
            self.reject
        )

        layout.addWidget(buttons)

    def value(self) -> str:
        return self.text_edit.toPlainText().strip()


class TitlePageDialog(QDialog):
    def __init__(
        self,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)

        self.setWindowTitle("Identify Title Page")

        layout = QVBoxLayout(self)
        form_layout = QFormLayout()

        self.title_edit = QLineEdit()
        self.subtitle_edit = QLineEdit()
        self.author_edit = QLineEdit()

        form_layout.addRow(
            "Title:",
            self.title_edit,
        )
        form_layout.addRow(
            "Subtitle:",
            self.subtitle_edit,
        )
        form_layout.addRow(
            "Author:",
            self.author_edit,
        )

        layout.addLayout(form_layout)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok
            | QDialogButtonBox.StandardButton.Cancel
        )

        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout.addWidget(buttons)

    def values(self) -> tuple[str, str, str]:
        return (
            self.title_edit.text().strip(),
            self.subtitle_edit.text().strip(),
            self.author_edit.text().strip(),
        )

# Functions
def build_front_matter_menu(
    file_menu: QMenu,
    configure_title_page: Callable[[], None],
    configure_copyright: Callable[[], None],
    configure_dedication: Callable[[], None],
    configure_map: Callable[[], None],
    configure_trigger_warnings: Callable[[], None],
    remove_section: Callable[[str], None],
) -> QMenu:
    front_matter_menu = file_menu.addMenu(
        "Front Matter"
    )

    title_page_action = (
        front_matter_menu.addAction(
            "Title Page"
        )
    )

    title_page_action.triggered.connect(
        configure_title_page
    )

    copyright_action = (
        front_matter_menu.addAction(
            "Copyright"
        )
    )

    copyright_action.triggered.connect(
        configure_copyright
    )

    dedication_action = (
        front_matter_menu.addAction(
            "Dedication"
        )
    )

    dedication_action.triggered.connect(
        configure_dedication
    )

    map_action = (
        front_matter_menu.addAction(
            "Map"
        )
    )

    map_action.triggered.connect(
        configure_map
    )

    trigger_warnings_action = (
        front_matter_menu.addAction(
            "Trigger Warnings"
        )
    )

    trigger_warnings_action.triggered.connect(
        configure_trigger_warnings
    )
    # Remove menu
    remove_menu = file_menu.addMenu(
        "Remove"
    )

    remove_title_page_action = (
        remove_menu.addAction(
            "Title Page"
        )
    )

    remove_title_page_action.triggered.connect(
        lambda _checked=False: remove_section("title_page")
    )

    remove_copyright_action = (
        remove_menu.addAction(
            "Copyright"
        )
    )

    remove_copyright_action.triggered.connect(
        lambda _checked=False: remove_section("copyright")
    )

    remove_dedication_action = (
        remove_menu.addAction(
            "Dedication"
        )
    )

    remove_dedication_action.triggered.connect(
        lambda _checked=False: remove_section("dedication")
    )

    remove_map_action = (
        remove_menu.addAction(
            "Map"
        )
    )

    remove_map_action.triggered.connect(
        lambda _checked=False: remove_section("map_file")
    )

    remove_trigger_warnings_action = (
        remove_menu.addAction(
            "Trigger Warnings"
        )
    )

    remove_trigger_warnings_action.triggered.connect(
        lambda _checked=False: remove_section("trigger_warnings")
    )

    remove_menu.addSeparator()
    
    clear_all_action = (
        remove_menu.addAction(
            "Clear All"
        )
    )
    
    clear_all_action.triggered.connect(
            lambda: remove_section(
                "clear_all"
            )
        )

    return front_matter_menu

def normalize_text(text: str) -> str:
    return " ".join(
        text.casefold().split()
    )

def find_paragraph(
    paragraphs: list[str],
    search_text: str,
) -> int | None:
    target = normalize_text(search_text)

    for index, paragraph in enumerate(paragraphs):
        if normalize_text(paragraph) == target:
            return index

    return None

def find_title_page(
    paragraphs: list[str],
    title: str,
    subtitle: str,
    author: str,
) -> TitlePage | None:
    indexes = []

    title_index = find_paragraph(
        paragraphs,
        title,
    )

    if title_index is None:
        return None

    indexes.append(title_index)

    if subtitle.strip():
        subtitle_index = find_paragraph(
            paragraphs,
            subtitle,
        )

        if subtitle_index is None:
            return None

        indexes.append(subtitle_index)

    author_index = find_paragraph(
        paragraphs,
        author,
    )

    if author_index is None:
        return None

    indexes.append(author_index)

    return TitlePage(
        title=title,
        subtitle=subtitle,
        author=author,
        start_index=min(indexes),
        end_index=max(indexes),
    )

def load_text_file(file_path: str) -> str:
    path = Path(file_path)

    if path.suffix.lower() == ".docx":
        document = Document(file_path)

        return "\n".join(
            paragraph.text
            for paragraph in document.paragraphs
            if paragraph.text.strip()
        )

    return path.read_text(
        encoding="utf-8"
    )

def find_text_section(
        paragraphs: list[str],
        text: str,
    ) -> TextSection | None:
        search_paragraphs = [
            line.strip()
            for line in text.splitlines()
            if line.strip()
        ]

        if not search_paragraphs:
            return None

        normalized_search = [
            normalize_text(line)
            for line in search_paragraphs
        ]

        section_length = len(normalized_search)

        for start_index in range(
            len(paragraphs) - section_length + 1
        ):
            candidate = [
                normalize_text(paragraph)
                for paragraph in paragraphs[
                    start_index:
                    start_index + section_length
                ]
            ]

            if candidate == normalized_search:
                return TextSection(
                    text=text,
                    start_index=start_index,
                    end_index=(
                        start_index
                        + section_length
                        - 1
                    ),
                )

        return None

def remove_section(
    front_matter_data: FrontMatter,
    section_name: str,
) -> None:
    setattr(
        front_matter_data,
        section_name,
        None,
    )


def clear_all() -> FrontMatter:
    return FrontMatter()