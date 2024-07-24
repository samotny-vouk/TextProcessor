import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QTextEdit, QAction, QMenu, QFileDialog, QInputDialog, \
    QColorDialog, QFontDialog
from PyQt5.QtGui import QFont, QTextCharFormat


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Создание меню
        self.fileMenu = self.menuBar().addMenu("&Файл")
        self.openAction = QAction("&Открыть", self)
        self.openAction.triggered.connect(self.openFile)
        self.fileMenu.addAction(self.openAction)

        self.saveAction = QAction("&Сохранить", self)
        self.saveAction.triggered.connect(self.saveFile)
        self.fileMenu.addAction(self.saveAction)

        self.fileMenu.addSeparator()

        self.exitAction = QAction("&Выход", self)
        self.exitAction.triggered.connect(self.close)
        self.fileMenu.addAction(self.exitAction)

        self.editMenu = self.menuBar().addMenu("&Правка")
        self.undoAction = QAction("&Отменить", self)
        self.undoAction.triggered.connect(self.undo)
        self.editMenu.addAction(self.undoAction)

        self.redoAction = QAction("&Повторить", self)
        self.redoAction.triggered.connect(self.redo)
        self.editMenu.addAction(self.redoAction)

        self.formatMenu = self.menuBar().addMenu("&Формат")
        self.boldAction = QAction("&Жирный", self)
        self.boldAction.triggered.connect(self.setBold)
        self.formatMenu.addAction(self.boldAction)

        self.italicAction = QAction("&Курсив", self)
        self.italicAction.triggered.connect(self.setItalic)
        self.formatMenu.addAction(self.italicAction)

        self.underlineAction = QAction("&Подчеркивание", self)
        self.underlineAction.triggered.connect(self.setUnderline)
        self.formatMenu.addAction(self.underlineAction)

        self.fontAction = QAction("&Шрифт", self)
        self.fontAction.triggered.connect(self.changeFont)
        self.formatMenu.addAction(self.fontAction)

        self.fontSizeAction = QAction("&Размер шрифта", self)
        self.fontSizeAction.triggered.connect(self.changeFontSize)
        self.formatMenu.addAction(self.fontSizeAction)

        self.textColorAction = QAction("&Цвет текста", self)
        self.textColorAction.triggered.connect(self.changeTextColor)
        self.formatMenu.addAction(self.textColorAction)

        self.indentAction = QAction("&Отступ", self)
        self.indentAction.triggered.connect(self.applyIndent)
        self.formatMenu.addAction(self.indentAction)

        self.outdentAction = QAction("&Убрать отступ", self)
        self.outdentAction.triggered.connect(self.applyOutdent)
        self.formatMenu.addAction(self.outdentAction)

        self.lineSpacingAction = QAction("&Интервал между строками", self)
        self.lineSpacingAction.triggered.connect(self.changeLineSpacing)
        self.formatMenu.addAction(self.lineSpacingAction)

        # Создание текстового поля
        self.textEdit = QTextEdit(self)
        self.setCentralWidget(self.textEdit)

        # Базовые стили
        self.setStyleSheet("QMainWindow { background-color: #f0f0f0; }"
                          "QTextEdit { font-family: Arial; font-size: 12pt; }")

    def openFile(self):
        fileName = QFileDialog.getOpenFileName(self, "Открыть файл", "", "Текстовые файлы (*.txt);;Все файлы (*)")[0]
        if fileName:
            with open(fileName, 'r', encoding='utf-8') as file:
                self.textEdit.setText(file.read())

    def saveFile(self):
        fileName = QFileDialog.getSaveFileName(self, "Сохранить файл", "", "Текстовые файлы (*.txt);;Все файлы (*)")[0]
        if fileName:
            with open(fileName, 'w', encoding='utf-8') as file:
                file.write(self.textEdit.toPlainText())

    def undo(self):
        self.textEdit.undo()

    def redo(self):
        self.textEdit.redo()

    def setBold(self):
        cursor = self.textEdit.textCursor()
        fmt = QTextCharFormat()
        fmt.setFontWeight(QFont.Bold)
        cursor.mergeCharFormat(fmt)

    def setItalic(self):
        cursor = self.textEdit.textCursor()
        fmt = QTextCharFormat()
        fmt.setFontItalic(True)
        cursor.mergeCharFormat(fmt)

    def setUnderline(self):
        self.applyFormat(QTextCharFormat.UnderlineStyle, QTextCharFormat.UnderlineStyle.SingleUnderline)

    def changeFontSize(self):
        fontSize, ok = QInputDialog.getInt(self, "Размер шрифта", "Введите размер:", 10, 1, 100)
        if ok:
            cursor = self.textEdit.textCursor()
            format = cursor.charFormat()
            font = QFont(format.font())
            font.setPointSize(fontSize)
            format.setFont(font)
            cursor.mergeCharFormat(format)
            self.textEdit.setTextCursor(cursor)

    def changeFont(self):
        font, ok = QFontDialog.getFont()
        if ok:
            cursor = self.textEdit.textCursor()
            format = cursor.charFormat()
            format.setFont(font)
            cursor.setCharFormat(format)
            self.textEdit.setTextCursor(cursor)

    def changeTextColor(self):
        color = QColorDialog.getColor(self.textEdit.textColor())
        if color.isValid():
            self.textEdit.setTextColor(color)

    def applyIndent(self):
        self.applyFormat(QTextCharFormat.BlockIndent, self.textEdit.textCursor().blockFormat().indent() + 1)

    def applyOutdent(self):
        self.applyFormat(QTextCharFormat.BlockIndent, self.textEdit.textCursor().blockFormat().indent() - 1)

    def changeLineSpacing(self):
        spacing, ok = QInputDialog.getDouble(self, "Интервал между строками", "Введите интервал:", 1.0, 0.1, 3.0, 2)
        if ok:
            self.textEdit.setLineSpacing(spacing)



if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
