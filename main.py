"""book formatter to publish in KDP."""

# Standard library imports.
import manuscript

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
from PySide6.QtGui import (
    QImage,
    QTextBlockFormat,
    QTextCharFormat,
    QTextCursor,
)
from docx import Document

# Local imports
import constants
import book_layout

# Classes
class bookformwindow(QWidget):
    """Controls the BookForm application window."""

    def __init__(self) -> None:
        super().__init__()

        # application state
        self.current_page = 0
        self.pages = []
        self.paragraphs = []
        self.page_starts_with_continuation = []
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
            "Arial",
            "Palatino Linotype",
        ])

        self.font_combo.currentTextChanged.connect(
            self.update_font
        )

        self.controls_layout.addWidget(font_label)
        self.controls_layout.addWidget(self.font_combo)

        font_size_label = QLabel("Font Size")

        self.font_size_combo = QComboBox()
        self.font_size_combo.addItems([
            "8.5 pt",
            "9 pt",
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

        self.measurement_device = QImage(
            int(constants.PRINT_PAGE_WIDTH_POINTS),
            int(constants.PRINT_PAGE_HEIGHT_POINTS),
            QImage.Format.Format_ARGB32,
        )

        self.measurement_device.setDotsPerMeterX(
            constants.PRINT_DOTS_PER_METER
        )

        self.measurement_device.setDotsPerMeterY(
            constants.PRINT_DOTS_PER_METER
        )

        self.page_preview.document().documentLayout().setPaintDevice(
            self.measurement_device
        )
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

        preview.document().setDocumentMargin(0)

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

        top_margin = book_layout.inches_to_pixels(profile["top"])
        bottom_margin = book_layout.inches_to_pixels(profile["bottom"])
        outside_margin = book_layout.inches_to_pixels(profile["outside"])
        gutter_margin = book_layout.inches_to_pixels(self.gutter_width)

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

    def apply_font_settings(
        self,
        preview: QTextEdit,
    ) -> None:
        font_name = self.font_combo.currentText()
        size_text = self.font_size_combo.currentText()
        font_size = float(size_text.split()[0])

        font = preview.font()
        font.setFamily(font_name)

        if preview is self.page_preview:
            font.setPointSizeF(font_size)
        else:
            preview_font_size = (
                font_size
                * constants.PIXELS_PER_INCH
                / preview.logicalDpiY()
            )
            font.setPointSizeF(preview_font_size)

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

        left_continues = (
            left_index >= 0
            and self.page_starts_with_continuation[left_index]
        )

        right_continues = (
            right_index < len(self.pages)
            and self.page_starts_with_continuation[right_index]
        )

        self.apply_font_settings(
            self.left_page_preview
        )
        self.apply_font_settings(
            self.right_page_preview
        )

        self.apply_paragraph_formatting(
            self.left_page_preview,
            left_continues,
        )
        self.apply_paragraph_formatting(
            self.right_page_preview,
            right_continues,
        )

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
        left_continues = (
            left_index >= 0
            and self.page_starts_with_continuation[left_index]
        )

        right_continues = (
            right_index < len(self.pages)
            and self.page_starts_with_continuation[right_index]
        )


        print(
            "Left page:",
            f"viewport={self.left_page_preview.viewport().width()}x"
            f"{self.left_page_preview.viewport().height()}",
            f"document={self.left_page_preview.document().size().width():.1f}x"
            f"{self.left_page_preview.document().size().height():.1f}",
        )

        print(
            "Right page:",
            f"viewport={self.right_page_preview.viewport().width()}x"
            f"{self.right_page_preview.viewport().height()}",
            f"document={self.right_page_preview.document().size().width():.1f}x"
            f"{self.right_page_preview.document().size().height():.1f}",
        )


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
        self.page_starts_with_continuation = []

        current_page_text = ""
        current_page_starts_with_continuation = False

        for paragraph in self.paragraphs:
            if manuscript.is_chapter_heading(paragraph):
                if current_page_text:
                    self.add_page(
                        current_page_text,
                        current_page_starts_with_continuation,
                    )
                    current_page_text = ""
                    current_page_starts_with_continuation = False

            paragraph_remaining = paragraph
            paragraph_continues = False
        
            while paragraph_remaining:
                if current_page_text:
                    test_text = (
                        current_page_text
                        + "\n"
                        + paragraph_remaining
                    )
                    test_starts_with_continuation = (
                        current_page_starts_with_continuation
                    )
                else:
                    test_text = paragraph_remaining
                    test_starts_with_continuation = paragraph_continues
        
                if self.text_fits_page(
                    test_text,
                    test_starts_with_continuation,
                ):
                    if not current_page_text:
                        current_page_starts_with_continuation = (
                            paragraph_continues
                        )

                    current_page_text = test_text
                    paragraph_remaining = ""
                else:
                    if current_page_text:
                        fitting_text, paragraph_remaining = (
                            self.split_paragraph_to_fit(
                                paragraph_remaining,
                                current_page_text,
                                current_page_starts_with_continuation,
                            )
                        )

                        if fitting_text:
                            self.add_page(
                                current_page_text
                                + "\n"
                            + fitting_text,
                            current_page_starts_with_continuation,
                            )

                            if paragraph_remaining:
                                paragraph_continues = True

                        else:
                            self.add_page(
                                current_page_text,
                                current_page_starts_with_continuation,
                            )

                        current_page_text = ""
                        current_page_starts_with_continuation = False

                    else:
                        fitting_text, paragraph_remaining = (
                            self.split_paragraph_to_fit(
                                paragraph_remaining,
                                first_block_continues=paragraph_continues,
                            )
                        )

                        self.add_page(
                            fitting_text,
                            paragraph_continues,
                        )

                        if paragraph_remaining:
                            paragraph_continues = True
        
        if current_page_text:
            self.add_page(
                current_page_text,
                current_page_starts_with_continuation,
            )

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

            new_gutter_width = book_layout.calculate_gutter_width(page_count)

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

        top_margin = book_layout.inches_to_points(
            profile["top"]
        )
        bottom_margin = book_layout.inches_to_points(
            profile["bottom"]
        )
        outside_margin = book_layout.inches_to_points(
            profile["outside"]
        )
        gutter_margin = book_layout.inches_to_points(
            self.gutter_width
        )

        self.measurement_width = (
            constants.PRINT_PAGE_WIDTH_POINTS
            - gutter_margin
            - outside_margin
        )

        self.measurement_height = (
            constants.PRINT_PAGE_HEIGHT_POINTS
            - top_margin
            - bottom_margin
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
            styles_found = set()

            self.paragraphs = []

            for paragraph in document.paragraphs:
                styles_found.add(paragraph.style.name)

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


    def text_fits_page(
        self,
        text: str,
        first_block_continues: bool = False,
    ) -> bool:
        self.page_preview.setPlainText(text)

        self.apply_font_settings(self.page_preview)
        self.apply_paragraph_formatting(
            self.page_preview,
            first_block_continues,
        )

        self.page_preview.document().setTextWidth(
            self.measurement_width
        )

        document_height = self.page_preview.document().size().height()

        return document_height <= self.measurement_height


    def split_paragraph_to_fit(
        self,
        paragraph: str,
        current_page_text: str = "",
        first_block_continues: bool = False,
    ) -> tuple[str, str]:
        words = paragraph.split()
        fitting_words = []

        for index, word in enumerate(words):
            candidate = " ".join(fitting_words + [word])

            if current_page_text:
                test_text = (
                    current_page_text
                    + "\n"
                    + candidate
                )
            else:
                test_text = candidate

            if self.text_fits_page(
                test_text,
                first_block_continues,
            ):
                fitting_words.append(word)
            else:
                if not fitting_words:
                    if current_page_text:
                        return "", paragraph

                    return word, " ".join(words[index + 1:])

                return (
                    " ".join(fitting_words),
                    " ".join(words[index:])
                )

        return " ".join(fitting_words), ""

    def update_font_size(self) -> None:
        self.apply_font_settings(self.page_preview)

        self.update_pages()
        self.show_page()

    def update_font(self) -> None:
        self.apply_font_settings(self.page_preview)

        self.update_pages()
        self.show_page()

    def apply_paragraph_formatting(
        self,
        preview: QTextEdit,
        first_block_continues: bool = False,
    ) -> None:
        spacing = float(self.line_spacing_combo.currentText())
        line_height = spacing * 100

        if preview is self.page_preview:
            first_line_indent = book_layout.inches_to_points(
                constants.FIRST_LINE_INDENT_INCHES
            )
            heading_space_after = book_layout.inches_to_points(
                constants.CHAPTER_HEADING_SPACE_AFTER_INCHES
            )
        else:
            first_line_indent = book_layout.inches_to_pixels(
                constants.FIRST_LINE_INDENT_INCHES
            )
            heading_space_after = book_layout.inches_to_pixels(
                constants.CHAPTER_HEADING_SPACE_AFTER_INCHES
            )

        document = preview.document()
        block = document.begin()

        previous_was_heading = False
        block_number = 0

        while block.isValid():
            text = block.text().strip()
            is_heading = manuscript.is_chapter_heading(text)

            block_format = QTextBlockFormat()

            block_format.setLineHeight(
                line_height,
                QTextBlockFormat.LineHeightTypes.ProportionalHeight.value
            )

            if is_heading:
                block_format.setTextIndent(0)
                block_format.setAlignment(
                    Qt.AlignmentFlag.AlignCenter
                )
                block_format.setBottomMargin(
                    heading_space_after
                )

            elif previous_was_heading:
                block_format.setTextIndent(0)

            elif block_number == 0 and first_block_continues:
                block_format.setTextIndent(0)

            else:
                block_format.setTextIndent(
                    first_line_indent
                )

            cursor = QTextCursor(block)
            cursor.setBlockFormat(block_format)

            if is_heading:
                heading_format = QTextCharFormat()
                if preview is self.page_preview:
                    heading_font_size = constants.CHAPTER_HEADING_FONT_SIZE
                else:
                    heading_font_size = (
                        constants.CHAPTER_HEADING_FONT_SIZE
                        * constants.PIXELS_PER_INCH
                        / preview.logicalDpiY()
                    )

                heading_format.setFontPointSize(
                    heading_font_size
                )

                cursor.movePosition(
                    QTextCursor.MoveOperation.StartOfBlock
                )
                cursor.movePosition(
                    QTextCursor.MoveOperation.EndOfBlock,
                    QTextCursor.MoveMode.KeepAnchor,
                )

                cursor.setCharFormat(heading_format)

            previous_was_heading = is_heading
            block_number += 1
            block = block.next()

    def line_spacing_changed(self) -> None:
        self.apply_paragraph_formatting(self.page_preview)
        self.update_pages()
        self.show_page()

    def margin_profile_changed(self) -> None:
        self.update_measurement_area()
        self.update_pages()
        self.show_page()

    def add_page(
        self,
        text: str,
        starts_with_continuation: bool = False,
    ) -> None:
        self.pages.append(text)
        self.page_starts_with_continuation.append(
            starts_with_continuation
    )

# Functions

def is_chapter_heading(text: str) -> bool:
    return bool(re.match(r"^\d+\.\s+\S", text.strip()))

def has_numbering(paragraph) -> bool:
    paragraph_properties = paragraph._p.pPr

    return (
        paragraph_properties is not None
        and paragraph_properties.numPr is not None
    )


def main():
    app = QApplication([])

    window = bookformwindow()
    window.show()

    app.exec()

if __name__ == "__main__":
    main()