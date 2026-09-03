BOOKFORM_STYLESHEET = """
QMenuBar#mainMenuBar {
    background-color: #C8C8C8;
    color: #000000;
    border: none;
}

QMenuBar#mainMenuBar::item {
    background-color: #C8C8C8;
    color: #000000;
    padding: 4px 10px;
}

QMenuBar#mainMenuBar::item:selected {
    background-color: #AFAFAF;
    color: #000000;
}

QMenu {
    background-color: #E6E6E6;
    color: #000000;
    border: 1px solid #A0A0A0;
    padding: 0px;
}

QMenu::item {
    background-color: #E6E6E6;
    color: #000000;
    border: none;
    margin: 0px;
    padding: 5px 24px 5px 10px;
}

QMenu::item:selected {
    background-color: #BFBFBF;
    color: #000000;
    border: none;
}

QWidget {
    background-color: #20382B;
    color: #F2F5F2;
}

QLabel {
    color: #F2F5F2;
}

QComboBox {
    background-color: #F4F7F4;
    color: #1F2D24;
    border: 1px solid #6D8B76;
    border-radius: 5px;
    padding: 5px 8px;
}

QComboBox:hover {
    border: 1px solid #9AB7A2;
}

QComboBox QAbstractItemView {
    background-color: #F4F7F4;
    color: #1F2D24;
    selection-background-color: #6D8B76;
    selection-color: white;
}

QPushButton {
    background-color: #3E6B4E;
    color: white;
    border: 1px solid #527C60;
    border-radius: 6px;
    padding: 7px 12px;
}

QPushButton:hover {
    background-color: #4C805D;
}

QPushButton:pressed {
    background-color: #31563E;
}

QPushButton:disabled {
    background-color: #55645A;
    color: #BFC7C1;
}

QProgressBar {
    background-color: #F4F7F4;
    border: 1px solid #6D8B76;
    border-radius: 5px;
    text-align: center;
}

QProgressBar::chunk {
    background-color: #67A879;
    border-radius: 4px;
}

QTextEdit {
    background-color: white;
    color: black;
}
"""