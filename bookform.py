from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QHBoxLayout,
    QVBoxLayout,
    QLabel,
    QComboBox,
    QPushButton,
    QFileDialog,
    QTextEdit,
)

from PySide6.QtCore import Qt
from docx import Document


current_page = 0
pages = []

window = None
page_preview = None
page_number_label = None
font_size_combo = None


def show_page():
    if pages:
        page_preview.setPlainText(pages[current_page])
        page_number_label.setText(
            f"Page {current_page + 1} of {len(pages)}"
        )


def previous_page():
    global current_page

    if current_page > 0:
        current_page -= 1
        show_page()


def next_page():
    global current_page

    if current_page < len(pages) - 1:
        current_page += 1
        show_page()


def text_fits_page(text):
    page_preview.setPlainText(text)

    document_height = page_preview.document().size().height()
    visible_height = page_preview.viewport().height()

    return document_height <= visible_height


def split_paragraph_to_fit(paragraph):
    words = paragraph.split()
    fitting_words = []

    for index, word in enumerate(words):
        test_text = " ".join(fitting_words + [word])

        if text_fits_page(test_text):
            fitting_words.append(word)
        else:
            if not fitting_words:
                return word, " ".join(words[index + 1:])

            return (
                " ".join(fitting_words),
                " ".join(words[index:])
            )

    return " ".join(fitting_words), ""


def load_manuscript():
    global pages, current_page

    file_path, _ = QFileDialog.getOpenFileName(
        window,
        "Select Manuscript",
        "",
        "Word Documents (*.docx)"
    )

    if file_path:
        document = Document(file_path)

        paragraphs = []

        for paragraph in document.paragraphs:
            if paragraph.text.strip():
                paragraphs.append(paragraph.text)

        pages = []
        current_page_text = ""

        for paragraph in paragraphs:
            paragraph_remaining = paragraph

            while paragraph_remaining:
                if current_page_text:
                    test_text = (
                        current_page_text
                        + "\n\n"
                        + paragraph_remaining
                    )
                else:
                    test_text = paragraph_remaining

                if text_fits_page(test_text):
                    current_page_text = test_text
                    paragraph_remaining = ""
                else:
                    if current_page_text:
                        pages.append(current_page_text)
                        current_page_text = ""
                    else:
                        fitting_text, paragraph_remaining = (
                            split_paragraph_to_fit(
                                paragraph_remaining
                            )
                        )

                        pages.append(fitting_text)

        if current_page_text:
            pages.append(current_page_text)

        current_page = 0
        show_page()


def update_font_size():
    size_text = font_size_combo.currentText()
    size = int(size_text.split()[0])

    font = page_preview.font()
    font.setPointSize(size)
    page_preview.setFont(font)


def main():
    global window
    global page_preview
    global page_number_label
    global font_size_combo

    app = QApplication([])

    window = QWidget()
    window.setWindowTitle("BookForm")
    window.resize(800, 600)

    layout = QHBoxLayout()
    window.setLayout(layout)

    controls_panel = QWidget()
    preview_panel = QWidget()

    layout.addWidget(controls_panel)
    layout.addWidget(preview_panel)

    controls_layout = QVBoxLayout()
    controls_panel.setLayout(controls_layout)

    trim_label = QLabel("Trim Size")
    trim_combo = QComboBox()
    trim_combo.addItem("6 by 9 inches")

    controls_layout.addWidget(trim_label)
    controls_layout.addWidget(trim_combo)

    font_label = QLabel("Body Font")
    font_combo = QComboBox()
    font_combo.addItems([
        "Garamond",
        "Times New Roman",
        "Georgia",
    ])

    controls_layout.addWidget(font_label)
    controls_layout.addWidget(font_combo)

    font_size_label = QLabel("Font Size")

    font_size_combo = QComboBox()
    font_size_combo.addItems([
        "10 pt",
        "11 pt",
        "12 pt",
    ])

    font_size_combo.currentTextChanged.connect(
        update_font_size
    )

    controls_layout.addWidget(font_size_label)
    controls_layout.addWidget(font_size_combo)

    line_spacing_label = QLabel("Line Spacing")
    line_spacing_combo = QComboBox()
    line_spacing_combo.addItems([
        "1.0",
        "1.15",
        "1.5",
    ])

    controls_layout.addWidget(line_spacing_label)
    controls_layout.addWidget(line_spacing_combo)

    load_button = QPushButton("Load Manuscript")
    load_button.clicked.connect(load_manuscript)

    controls_layout.addWidget(load_button)
    controls_layout.addStretch()

    preview_layout = QVBoxLayout()
    preview_panel.setLayout(preview_layout)

    preview_title = QLabel("Book Preview")
    preview_layout.addWidget(preview_title)

    page_preview = QTextEdit()
    page_preview.setFixedSize(400, 600)
    page_preview.setReadOnly(True)
    page_preview.setFocusPolicy(Qt.NoFocus)

    page_preview.setVerticalScrollBarPolicy(
        Qt.ScrollBarAlwaysOff
    )

    page_preview.setHorizontalScrollBarPolicy(
        Qt.ScrollBarAlwaysOff
    )

    page_preview.setStyleSheet("""
        background-color: white;
        border: 1px solid gray;
        padding: 50px 40px;
    """)

    preview_layout.addWidget(page_preview)

    navigation_layout = QHBoxLayout()

    previous_button = QPushButton("Previous")
    next_button = QPushButton("Next")

    previous_button.clicked.connect(previous_page)
    next_button.clicked.connect(next_page)

    page_number_label = QLabel("Page 0 of 0")

    navigation_layout.addWidget(previous_button)
    navigation_layout.addWidget(page_number_label)
    navigation_layout.addWidget(next_button)

    preview_layout.addLayout(navigation_layout)

    window.show()

    app.exec()


if __name__ == "__main__":
    main()