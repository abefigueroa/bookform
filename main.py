"""book formatter to publish in KDP."""

# Standard library imports

# Third-party imports
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
from PySide6.QtGui import QTextCursor, QTextBlockFormat
from docx import Document

# Local imports
import constants

# Classes
class bookformwindow(QWidget):
    """Controls the BookForm application window."""

    def __init__(self) -> None:
        super().__init__()

        # application state
        self.current_page = 0
        self.pages = []
        self.paragraphs = []

        # window configuration
        self.setWindowTitle("BookForm")
        self.resize(800, 600)

        # layout setup
        self.layout = QHBoxLayout()
        self.setLayout(self.layout)

        # panel creation
        self.controls_panel = QWidget()
        self.preview_panel = QWidget()
        self.layout.addWidget(self.controls_panel)
        self.layout.addWidget(self.preview_panel)
        self.controls_layout = QVBoxLayout()
        self.controls_panel.setLayout(self.controls_layout)
        self.preview_layout = QVBoxLayout()
        self.preview_panel.setLayout(self.preview_layout)

        # formatting controls
        trim_label = QLabel("Trim Size")

        self.trim_combo = QComboBox()
        self.trim_combo.addItem("6 by 9 inches")

        self.controls_layout.addWidget(trim_label)
        self.controls_layout.addWidget(self.trim_combo)

        font_label = QLabel("Body Font")

        self.font_combo = QComboBox()
        self.font_combo.addItems([
            "Garamond",
            "Times New Roman",
            "Georgia",
        ])

        self.font_combo.currentTextChanged.connect(
            self.update_font
        )

        self.controls_layout.addWidget(font_label)
        self.controls_layout.addWidget(self.font_combo)

        font_size_label = QLabel("Font Size")

        self.font_size_combo = QComboBox()
        self.font_size_combo.addItems([
            "10 pt",
            "11 pt",
            "12 pt",
        ])

        self.font_size_combo.currentTextChanged.connect(
            self.update_font_size
        )

        self.controls_layout.addWidget(font_size_label)
        self.controls_layout.addWidget(self.font_size_combo)

        line_spacing_label = QLabel("Line Spacing")

        self.line_spacing_combo = QComboBox()
        self.line_spacing_combo.addItems([
            "1.0",
            "1.15",
            "1.5",
        ])

        self.line_spacing_combo.currentTextChanged.connect(
            self.line_spacing_changed
        )

        self.controls_layout.addWidget(line_spacing_label)
        self.controls_layout.addWidget(self.line_spacing_combo)

        # manuscript controls
        self.load_button = QPushButton("Load Manuscript")
        self.load_button = QPushButton("Load Manuscript")
        self.load_button.clicked.connect(self.load_manuscript)
        self.controls_layout.addWidget(self.load_button)

        self.controls_layout.addStretch()

        # preview setup
        preview_title = QLabel("Book Preview")
        self.preview_layout.addWidget(preview_title)

        self.page_preview = QTextEdit()
        self.page_preview.setFixedSize(
            constants.PREVIEW_WIDTH_PIXELS,
            constants.PREVIEW_HEIGHT_PIXELS,
        )
        self.page_preview.setReadOnly(True)
        self.page_preview.setFocusPolicy(Qt.NoFocus)

        self.page_preview.setVerticalScrollBarPolicy(
            Qt.ScrollBarAlwaysOff
        )

        self.page_preview.setHorizontalScrollBarPolicy(
            Qt.ScrollBarAlwaysOff
        )

        self.page_preview.setStyleSheet("""
            background-color: white;
            border: 1px solid gray;
            padding: 50px 40px;
        """)

        self.preview_layout.addWidget(self.page_preview)

        # page navigation
        navigation_layout = QHBoxLayout()

        self.previous_button = QPushButton("Previous")
        self.next_button = QPushButton("Next")
        self.page_number_label = QLabel("Page 0 of 0")

        self.previous_button.clicked.connect(self.previous_page)
        self.next_button.clicked.connect(self.next_page)
        
        navigation_layout.addWidget(self.previous_button)
        navigation_layout.addWidget(self.page_number_label)
        navigation_layout.addWidget(self.next_button)

        self.preview_layout.addLayout(navigation_layout)

    # Methods
    def show_page(self) -> None:
            if self.pages:
                self.page_preview.setPlainText(
                    self.pages[self.current_page]
                )
                self.update_line_spacing()
                
                self.page_number_label.setText(
                    f"Page {self.current_page + 1} of {len(self.pages)}"
                )

    def previous_page(self) -> None:
        if self.current_page > 0:
            self.current_page -= 1
            self.show_page()


    def next_page(self) -> None:
        if self.current_page < len(self.pages) - 1:
            self.current_page += 1
            self.show_page()

    def update_pages(self) -> None:
        self.pages = []
        current_page_text = ""

        for paragraph in self.paragraphs:
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
        
                if self.text_fits_page(test_text):
                    current_page_text = test_text
                    paragraph_remaining = ""
                else:
                    if current_page_text:
                        self.pages.append(current_page_text)
                        current_page_text = ""
                    else:
                        fitting_text, paragraph_remaining = (
                            self.split_paragraph_to_fit(
                                paragraph_remaining
                            )
                        )
        
                        self.pages.append(fitting_text)
        
        if current_page_text:
            self.pages.append(current_page_text)
        if self.pages:
            self.gutter_width = calculate_gutter_width(len(self.pages))

    def load_manuscript(self) -> None:
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Manuscript",
            "",
            "Word Documents (*.docx)"
        )

        if file_path:
            document = Document(file_path)

            self.paragraphs = []

            for paragraph in document.paragraphs:
                if paragraph.text.strip():
                    self.paragraphs.append(paragraph.text)

            self.update_pages()        

            self.current_page = 0
            self.show_page()

    def text_fits_page(self, text):
        self.page_preview.setPlainText(text)
        self.update_line_spacing()

        document_height = self.page_preview.document().size().height()
        visible_height = self.page_preview.viewport().height()

        return document_height <= visible_height


    def split_paragraph_to_fit(self, paragraph):
        words = paragraph.split()
        fitting_words = []

        for index, word in enumerate(words):
            test_text = " ".join(fitting_words + [word])

            if self.text_fits_page(test_text):
                fitting_words.append(word)
            else:
                if not fitting_words:
                    return word, " ".join(words[index + 1:])

                return (
                    " ".join(fitting_words),
                    " ".join(words[index:])
                )

        return " ".join(fitting_words), ""

    def update_font_size(self) -> None:
        size_text = self.font_size_combo.currentText()
        size = int(size_text.split()[0])

        font = self.page_preview.font()
        font.setPointSize(size)
        self.page_preview.setFont(font)

        self.update_pages()
        self.show_page()

    def update_font(self) -> None:
        text = self.font_combo.currentText()
        
        font = self.page_preview.font()
        font.setFamily(text)
        self.page_preview.setFont(font)

        self.update_pages()
        self.show_page()

    def update_line_spacing(self) -> None:
        spacing = float(self.line_spacing_combo.currentText())
        line_height = spacing * 100

        cursor = self.page_preview.textCursor()
        cursor.select(QTextCursor.Document)

        block_format = QTextBlockFormat()
        # Apply line spacing as a percentage of normal height
        block_format.setLineHeight(
            line_height,
            QTextBlockFormat.LineHeightTypes.ProportionalHeight.value
        )

        cursor.setBlockFormat(block_format)

    def line_spacing_changed(self) -> None:
        self.update_line_spacing()
        self.update_pages()
        self.show_page()

# Functions
def calculate_gutter_width(page_count: int) -> float:
    if page_count < 24:
        raise ValueError("KDP requires at least 24 pages.")

    if page_count <= 150:
        return 0.375
    elif page_count <= 300:
        return 0.500
    elif page_count <= 500:
        return 0.625
    elif page_count <= 700:
        return 0.750
    elif page_count <= 828:
        return 0.875

    raise ValueError("Page count exceeds the supported KDP range")

def inches_to_pixels(inches: float) -> int:
    return round(inches * constants.PIXELS_PER_INCH)
    
def main():
    app = QApplication([])

    window = bookformwindow()
    window.show()

    app.exec()

if __name__ == "__main__":
    main()