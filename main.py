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
        self.gutter_width = 0.375

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

        margin_profile_label = QLabel("Margin Profile")

        self.margin_profile_combo = QComboBox()
        self.margin_profile_combo.addItems(
            constants.MARGIN_PROFILES.keys()
        )

        self.margin_profile_combo.setCurrentText(
            constants.DEFAULT_MARGIN_PROFILE
        )

        self.margin_profile_combo.currentTextChanged.connect(
            self.margin_profile_changed
        )

        self.controls_layout.addWidget(margin_profile_label)
        self.controls_layout.addWidget(self.margin_profile_combo)

        # manuscript controls
        self.load_button = QPushButton("Load Manuscript")
        self.load_button = QPushButton("Load Manuscript")
        self.load_button.clicked.connect(self.load_manuscript)
        self.controls_layout.addWidget(self.load_button)

        self.controls_layout.addStretch()

        # preview setup
        preview_title = QLabel("Book Preview")
        self.preview_layout.addWidget(preview_title)

        # future two-page spread
        self.left_page_preview = self.create_page_preview()
        self.right_page_preview = self.create_page_preview()

        spread_layout = QHBoxLayout()
        spread_layout.addWidget(self.left_page_preview)
        spread_layout.addWidget(self.right_page_preview)

        # existing preview used by current pagination
        self.page_preview = self.create_page_preview()
        self.update_measurement_area()

        self.preview_layout.addLayout(spread_layout)

        self.kdp_warning_label = QLabel("")
        self.kdp_warning_label.setWordWrap(True)
        self.preview_layout.addWidget(self.kdp_warning_label)     
        
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
    def create_page_preview(self) -> QTextEdit:
        preview = QTextEdit()

        preview.setFixedSize(
            constants.PREVIEW_WIDTH_PIXELS,
            constants.PREVIEW_HEIGHT_PIXELS,
        )

        preview.setReadOnly(True)
        preview.setFocusPolicy(Qt.NoFocus)
        preview.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        preview.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)

        return preview

    
    def apply_page_margins(self) -> None:
        profile_name = self.margin_profile_combo.currentText()
        profile = constants.MARGIN_PROFILES[profile_name]

        top_margin = inches_to_pixels(profile["top"])
        bottom_margin = inches_to_pixels(profile["bottom"])
        outside_margin = inches_to_pixels(profile["outside"])
        gutter_margin = inches_to_pixels(self.gutter_width)

        page_number = self.current_page + 1

        if page_number % 2 == 1:
            left_margin = gutter_margin
            right_margin = outside_margin
        else:
            left_margin = outside_margin
            right_margin = gutter_margin

        self.left_page_preview.setStyleSheet(f"""
            background-color: white;
            border: 1px solid gray;
            padding-top: {top_margin}px;
            padding-bottom: {bottom_margin}px;
            padding-left: {outside_margin}px;
            padding-right: {gutter_margin}px;
        """)

        self.right_page_preview.setStyleSheet(f"""
            background-color: white;
            border: 1px solid gray;
            padding-top: {top_margin}px;
            padding-bottom: {bottom_margin}px;
            padding-left: {gutter_margin}px;
            padding-right: {outside_margin}px;
        """)

    def apply_font_settings(self, preview: QTextEdit) -> None:
        font_name = self.font_combo.currentText()
        size_text = self.font_size_combo.currentText()
        font_size = int(size_text.split()[0])

        font = preview.font()
        font.setFamily(font_name)
        font.setPointSize(font_size)
        preview.setFont(font)

    def show_page(self) -> None:
        if not self.pages:
            return

        self.apply_page_margins()

        right_index = self.current_page
        left_index = right_index - 1

        if left_index >= 0:
            self.left_page_preview.setPlainText(
                self.pages[left_index]
            )
        else:
            self.left_page_preview.clear()

        if right_index < len(self.pages):
            self.right_page_preview.setPlainText(
                self.pages[right_index]
            )
        else:
            self.right_page_preview.clear()

        if left_index >= 0 and right_index < len(self.pages):
            self.page_number_label.setText(
                f"Pages {left_index + 1}–{right_index + 1} "
                f"of {len(self.pages)}"
            )
        elif right_index < len(self.pages):
            self.page_number_label.setText(
            f"Page {right_index + 1} of {len(self.pages)}"
            )
        elif left_index >= 0:
            self.page_number_label.setText(
                f"Page {left_index + 1} of {len(self.pages)}"
            )

        self.apply_paragraph_formatting(self.left_page_preview)
        self.apply_paragraph_formatting(self.right_page_preview)

        self.apply_font_settings(self.left_page_preview)
        self.apply_font_settings(self.right_page_preview)


    def previous_page(self) -> None:
        if self.current_page >= 2:
            self.current_page -= 2
            self.show_page()


    def next_page(self) -> None:
        if self.current_page + 1 < len(self.pages):
            self.current_page += 2
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
                        + "\n"
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
            print(f"Page count: {len(self.pages)}")
            print(
                f"Word count: "
                f"{sum(len(paragraph.split()) for paragraph in self.paragraphs)}"
            )
            print(
                f"Measurement area: "
                f"{self.measurement_width} x "
                f"{self.measurement_height}"
            )
            page_count = len(self.pages)

            new_gutter_width = calculate_gutter_width(page_count)

            if new_gutter_width != self.gutter_width:
                self.gutter_width = new_gutter_width
                self.update_measurement_area()
                self.update_pages()
                return

            if page_count > constants.KDP_MAX_PAGE_COUNT:
                pages_over = page_count - constants.KDP_MAX_PAGE_COUNT

                self.kdp_warning_label.setText(
                    f"{page_count} pages — "
                    f"{pages_over} pages over the KDP 6 × 9 limit."
                )
            else:
                self.kdp_warning_label.setText("")

    def update_measurement_area(self) -> None:
        profile_name = self.margin_profile_combo.currentText()
        profile = constants.MARGIN_PROFILES[profile_name]

        top_margin = inches_to_pixels(profile["top"])
        bottom_margin = inches_to_pixels(profile["bottom"])
        outside_margin = inches_to_pixels(profile["outside"])
        gutter_margin = inches_to_pixels(self.gutter_width)

        usable_width = (
            constants.PREVIEW_WIDTH_PIXELS
            - gutter_margin
            - outside_margin
        )

        usable_height = (
            constants.PREVIEW_HEIGHT_PIXELS
            - top_margin
            - bottom_margin
        )

        self.measurement_width = usable_width
        self.measurement_height = usable_height

        self.page_preview.setFixedSize(
            usable_width,
            usable_height,
        )

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
            print(
                f"Measurement area: "
                f"{self.measurement_width} x "
                f"{self.measurement_height}"
            )
            print("Starting pagination...")
            self.update_pages()        

            self.current_page = 0
            self.show_page()


    def text_fits_page(self, text):
        self.page_preview.setPlainText(text)
        self.apply_paragraph_formatting(self.page_preview)

        self.page_preview.document().setTextWidth(
            self.measurement_width
        )

        document_height = self.page_preview.document().size().height()

        return document_height <= self.measurement_height


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

    def apply_paragraph_formatting(self, preview: QTextEdit) -> None:
        spacing = float(self.line_spacing_combo.currentText())
        line_height = spacing * 100

        first_line_indent = inches_to_pixels(
        constants.FIRST_LINE_INDENT_INCHES
        )

        cursor = preview.textCursor()
        cursor.select(QTextCursor.Document)

        block_format = QTextBlockFormat()
        # Apply line spacing as a percentage of normal height
        block_format.setLineHeight(
            line_height,
            QTextBlockFormat.LineHeightTypes.ProportionalHeight.value
        )

        block_format.setTextIndent(first_line_indent)

        cursor.setBlockFormat(block_format)

    def line_spacing_changed(self) -> None:
        self.apply_paragraph_formatting(self.page_preview)
        self.update_pages()
        self.show_page()

    def margin_profile_changed(self) -> None:
        self.update_measurement_area()
        self.update_pages()
        self.show_page()

# Functions
def calculate_gutter_width(page_count: int) -> float:
    if page_count <= 150:
        return 0.375
    elif page_count <= 300:
        return 0.500
    elif page_count <= 500:
        return 0.625
    elif page_count <= 700:
        return 0.750
    else:
        return 0.875

def inches_to_pixels(inches: float) -> int:
    return round(inches * constants.PIXELS_PER_INCH)


def main():
    app = QApplication([])

    window = bookformwindow()
    window.show()

    app.exec()

if __name__ == "__main__":
    main()