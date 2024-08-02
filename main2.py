import sqlite3
import sys
from PyQt5 import QtWidgets
from PyQt5.QtPrintSupport import QPrinter
from PyQt5.QtWidgets import QApplication, QMainWindow, QTextEdit, QAction, QFileDialog, QInputDialog, QColorDialog, \
    QFontDialog, QMessageBox, QVBoxLayout, QLabel, QLineEdit, QWidget, QPushButton, QComboBox
from PyQt5.QtGui import QFont, QTextCharFormat, QTextCursor, QDesktopServices, QColor, QBrush, QTextBlockFormat, \
    QTextImageFormat, QPixmap
from PyQt5.QtCore import Qt, QUrl
from docx import Document
import zipfile
import re
text = ""
DB_NAME = 'style.db'


def link_clicked(url):
    QDesktopServices.openUrl(QUrl(url))


class MyTextEdit(QTextEdit):
    def __init__(self):
        super().__init__()
        self.current_page = 1
        self.pages = []

    def addPageBreak(self):
        cursor = self.textCursor()
        block_format = QTextBlockFormat()
        block_format.setPageBreakPolicy(QTextBlockFormat.PageBreak_AlwaysBefore)

        cursor.insertBlock(block_format)  # Создание разрыва страницы
        self.pages.append(('Page {}'.format(self.current_page),
                           cursor.block().text()))  # Сохранение текста на странице
        self.current_page += 1

        # Пример добавления номера страницы
        cursor.insertText(f"\nPage {self.current_page}\n")

    def updateDocument(self):
        self.clear()
        for page_num, text in self.pages:
            self.append(text)
            self.append(page_num)  # Добавление номера страницы


class SearchReplaceApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Поиск и Замена')
        layout = QVBoxLayout()

        self.text_edit = QTextEdit(self)
        self.text_edit.setPlainText(text)
        layout.addWidget(QLabel("Документ:"))
        layout.addWidget(self.text_edit)

        self.search_input = QLineEdit(self)
        layout.addWidget(QLabel("Поиск:"))
        layout.addWidget(self.search_input)

        self.replace_input = QLineEdit(self)
        layout.addWidget(QLabel("Замена:"))
        layout.addWidget(self.replace_input)

        self.replace_button = QPushButton('Заменить', self)
        self.replace_button.clicked.connect(self.replace_text)
        layout.addWidget(self.replace_button)

        self.back = QPushButton('Назад', self)
        self.back.clicked.connect(lambda: toBack())
        layout.addWidget(self.back)

        self.setLayout(layout)

    def replace_text(self):
        document = self.text_edit.toPlainText()
        search_term = self.search_input.text()
        replace_term = self.replace_input.text()
        # use_regex = self.regex_checkbox.text().strip().lower() == 'да'
        use_regex = True

        if use_regex:
            try:
                pattern = re.compile(search_term)
                new_document = pattern.sub(replace_term, document)
                matches = pattern.findall(document)
            except re.error:
                QMessageBox.warning(self, "Ошибка", "Неверное регулярное выражение.")
                return
        else:
            matches = document.count(search_term)
            new_document = document.replace(search_term, replace_term)

        self.text_edit.setPlainText(new_document)
        QMessageBox.information(self, "Результаты", f"Найдено совпадений: {matches}")


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.textEdit = QTextEdit(self)
        self.setCentralWidget(self.textEdit)
        # self.setGeometry(100, 100, 600, 400)

        self.doc = Document()
        self.default_style = self.doc.styles['Normal']

        self.fileMenu = self.menuBar().addMenu("&Файл")
        self.openAction = QAction("&Открыть", self)
        self.openAction.triggered.connect(self.openFile)
        self.fileMenu.addAction(self.openAction)

        self.saveAction = QAction("&Сохранить", self)
        self.saveAction.triggered.connect(self.saveFile)
        self.fileMenu.addAction(self.saveAction)

        self.saveAsAction = QAction("Сохранить как...", self)
        self.saveAsAction.triggered.connect(self.saveFileAs)
        self.fileMenu.addAction(self.saveAsAction)

        self.exportMenu = self.fileMenu.addMenu("Экспорт")

        self.exportPdfAction = QAction("PDF", self)
        self.exportPdfAction.triggered.connect(self.exportPdf)
        self.exportMenu.addAction(self.exportPdfAction)

        self.exportTxtAction = QAction("TXT", self)
        self.exportTxtAction.triggered.connect(self.exportTxt)
        self.exportMenu.addAction(self.exportTxtAction)

        self.exportZipAction = QAction("ZIP", self)
        self.exportZipAction.triggered.connect(self.exportZip)
        self.exportMenu.addAction(self.exportZipAction)

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

        self.renameAction = QAction("&Найти и заменить", self)
        self.renameAction.triggered.connect(lambda: toNext(SearchReplaceApp))
        self.editMenu.addAction(self.renameAction)

        self.formatMenu = self.menuBar().addMenu("&Формат")

        self.insertImageAction = QAction("Вставить изображение", self)
        self.insertImageAction.triggered.connect(self.insertImage)
        self.formatMenu.addAction(self.insertImageAction)

        self.insertLinkAction = QAction("Вставить ссылку", self)
        self.insertLinkAction.triggered.connect(self.insertLink)
        self.formatMenu.addAction(self.insertLinkAction)

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

        self.backgroundTextAction = QAction('Цвет фона текста', self)
        self.backgroundTextAction.triggered.connect(self.chooseBackgroundColor)
        self.formatMenu.addAction(self.backgroundTextAction)

        self.pageColorAction = QAction('Цвет страницы', self)
        self.pageColorAction.triggered.connect(self.choosePageColor)
        self.formatMenu.addAction(self.pageColorAction)

        self.indentAction = QAction("&Отступ", self)
        self.indentAction.triggered.connect(self.applyIndent)
        self.formatMenu.addAction(self.indentAction)

        self.outdentAction = QAction("&Убрать отступ", self)
        self.outdentAction.triggered.connect(self.applyOutdent)
        self.formatMenu.addAction(self.outdentAction)

        self.lineSpacingAction = QAction("&Интервал между строками", self)
        self.lineSpacingAction.triggered.connect(self.changeLineSpacing)
        self.formatMenu.addAction(self.lineSpacingAction)

        self.solveMenu = self.menuBar().addMenu("&Параметры листа")

        self.a4Action = QAction("&А4", self)
        self.a4Action.triggered.connect(self.changePageSizeA4)
        self.solveMenu.addAction(self.a4Action)

        self.elseSolveAction = QAction("&Настроить", self)
        self.elseSolveAction.triggered.connect(self.changePageSizeSolve)
        self.solveMenu.addAction(self.elseSolveAction)

        self.pageNumber = QAction('Нумерация страниц')
        self.pageNumber.triggered.connect(self.addPageNumbers)
        self.solveMenu.addAction(self.pageNumber)

        self.pageBreak = QAction('Разрыв страницы')
        self.pageBreak.triggered.connect(self.addPageBreak)
        self.solveMenu.addAction(self.pageBreak)

        self.setStyleSheet("QMainWindow { background-color: #f0f0f0; }" "QTextEdit { font-family: Arial; font-size: 12pt; }")

        self.style_combo = QComboBox(self)
        self.load_styles()

        self.apply_style_button = QPushButton("Apply Style", self)
        self.add_style_button = QPushButton("Add Custom Style", self)

        self.apply_style_button.clicked.connect(self.applyStyle)
        self.add_style_button.clicked.connect(self.addCustomStyle)

        layout = QVBoxLayout()
        layout.addWidget(self.style_combo)
        layout.addWidget(self.apply_style_button)
        layout.addWidget(self.add_style_button)
        layout.addWidget(self.textEdit)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def load_styles(self):
        self.style_combo.clear()
        self.style_combo.addItems(["Normal", "Heading 1", "Heading 2"])

        con = sqlite3.connect(DB_NAME)
        cursor = con.cursor()
        cursor.execute("CREATE TABLE IF NOT EXISTS style (name TEXT PRIMARY KEY, font_size INTEGER, font_weight INTEGER, font_italic INTEGER, font_underline INTEGER, font_color TEXT, font_family TEXT, alignment TEXT)")
        con.commit()

        cursor.execute("SELECT name FROM style")
        styles = cursor.fetchall()

        for style in styles:
            self.style_combo.addItem(style[0])

        con.close()

    def applyStyle(self):
        selected_style = self.style_combo.currentText()

        if selected_style == "Normal":
            self.textEdit.setFontPointSize(12)
            self.textEdit.setFontWeight(50)
            self.textEdit.setFontItalic(False)
            self.textEdit.setFontUnderline(False)
            self.textEdit.setTextColor(QColor(0, 0, 0))
            self.textEdit.setFontFamily("Arial")
            self.textEdit.setAlignment(Qt.AlignLeft)
            # self.textEdit.setPageSize(QPrinter.A4)
        elif selected_style == "Heading 1":
            self.textEdit.setFontPointSize(24)
            self.textEdit.setFontWeight(100)
            self.textEdit.setFontItalic(False)
            self.textEdit.setFontUnderline(False)
            self.textEdit.setTextColor(QColor(0, 0, 0))
            self.textEdit.setFontFamily("Arial")
            self.textEdit.setAlignment(Qt.AlignCenter)
            # self.textEdit.printer.setPageSize(QPrinter.A4)
        elif selected_style == "Heading 2":
            self.textEdit.setFontPointSize(18)
            self.textEdit.setFontWeight(100)
            self.textEdit.setFontItalic(False)
            self.textEdit.setFontUnderline(False)
            self.textEdit.setTextColor(QColor(0, 0, 0))
            self.textEdit.setFontFamily("Arial")
            self.textEdit.setAlignment(Qt.AlignCenter)
            # self.textEdit.printer.setPageSize(QPrinter.A4)
        else:
            con = sqlite3.connect(DB_NAME)
            cursor = con.cursor()
            cursor.execute("SELECT font_size, font_weight, font_italic, font_underline, font_color, font_family, alignment FROM style WHERE name=?",(selected_style,))
            style = cursor.fetchone()

            if style:
                self.textEdit.setFontPointSize(style[0])
                self.textEdit.setFontWeight(style[1])
                self.textEdit.setFontItalic(bool(style[2]))
                self.textEdit.setFontUnderline(bool(style[3]))
                color = QColor(style[4]) if style[4] else QColor(0, 0, 0)

                self.textEdit.setTextColor(color)
                self.textEdit.setFontFamily(style[5])
                try:
                    self.textEdit.setAlignment(style[6])
                except:
                    self.textEdit.setAlignment(Qt.AlignLeft)
            else:
                QMessageBox.warning(self, "Error", "Style not found!")
            con.close()

    def addCustomStyle(self):
        style_name, ok = QInputDialog.getText(self, "Add Custom Style", "Enter style name:")
        if ok and style_name:
            font_size, ok = QInputDialog.getInt(self, "Set Font Size", "Enter font size:", 12, 1, 100)
            if not ok:
                return

            font_weight, ok = QInputDialog.getInt(self, "Set Font Weight", "Enter font weight (100: Normal, 400: Bold):", 0, 0, 100)
            if not ok:
                return

            font_italic, ok = QInputDialog.getItem(self, "Set Italic", "Is italic enabled?", ["True", "False"], 0, False)
            font_italic = font_italic == "True"
            if not ok:
                return

            font_underline, ok = QInputDialog.getItem(self, "Set Underline", "Is underline enabled?", ["True", "False"], 0, False)
            font_underline = font_underline == "True"
            if not ok:
                return

            color = QColorDialog.getColor()
            font_color = color.name() if color.isValid() else "#000000"

            font_family, ok = QInputDialog.getText(self, "Set Font Family", "Enter font family:")
            if not ok or not font_family:
                return

            alignment, ok = QInputDialog.getInt(self, "Set Alignment", "Enter alignment (0: Left, 1: Center, 2: Right):", 0, 0, 2)
            if ok:
                alignment_map = {
                    0: Qt.AlignLeft,
                    1: Qt.AlignCenter,
                    2: Qt.AlignRight
                }

                if alignment in alignment_map:
                    align = alignment_map.get(alignment)
                    print(f"Alignment selected: {alignment}, correspond to: {align}")  # Отладочный вывод
                else:
                    self.show_error("Invalid alignment value selected.")

            if ok:
                self.saveStyleToDB(style_name, font_size, font_weight, font_italic, font_underline, font_color,
                                   font_family, align)
                self.load_styles()

    def saveStyleToDB(self, style_name, font_size, font_weight, font_italic, font_underline, font_color, font_family, alignment):
        connection = sqlite3.connect(DB_NAME)
        cursor = connection.cursor()
        cursor.execute("INSERT OR REPLACE INTO style (name, font_size, font_weight, font_italic, font_underline, font_color, font_family, alignment) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                       (style_name, font_size, font_weight, font_italic, font_underline, font_color, font_family, alignment))
        connection.commit()
        connection.close()

    def changePageSizeA4(self):
        printer = QPrinter()
        printer.setPageSize(QPrinter.A4)

    def changePageSizeSolve(self):
        pass

    def addPageNumbers(self):
        text_document = self.textEdit.document()
        page_count = text_document.pageCount()
        for page in range(page_count):
            self.textEdit.append(f"Страница {page + 1} из {page_count}")

    def addPageBreak(self):
        text_edit = MyTextEdit()
        text_edit.setPlainText(self.textEdit.toPlainText())
        text_edit.addPageBreak()
        text_edit.updateDocument()
        text_edit.show()

    def updateDocument(self):
        self.clear()
        for page_num, text in self.pages:
            self.append(text)
            self.append(page_num)

    def insertImage(self):
        options = QFileDialog.Options()
        filePath, _ = QFileDialog.getOpenFileName(self, "Выбрать изображение", "",
                                                  "Images (*.png *.jpg *.jpeg *.bmp *.gif)", options=options)
        if filePath:
            self.textEdit.insertHtml(f'<img src="{filePath}" alt="Image" width="300"/>')

    def insertLink(self):
        cursor = self.textEdit.textCursor()
        selected_text = cursor.selectedText()

        if not selected_text:
            selected_text = cursor.block().text()
            cursor.setPosition(cursor.block().position())
            cursor.movePosition(QTextCursor.EndOfBlock, QTextCursor.KeepAnchor)

        link, ok = QInputDialog.getText(self, 'Вставить ссылку', 'Введите URL:')
        url_pattern = re.compile(r'http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+')

        if ok and url_pattern.match(link):
            char_format = cursor.charFormat()
            char_format.setForeground(Qt.blue)
            char_format.setFontUnderline(True)

            cursor.mergeCharFormat(char_format)
            cursor.insertHtml(f'<a href="{link}">{selected_text}</a>')
        else:
            QMessageBox.warning(self, "Ошибка", "Введенный URL некорректен.")

    def saveFileAs(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self, "Сохранить как...", "",
                                                  "Документ Word (*.docx);;Текстовый файл (*.txt);;Все файлы (*)",
                                                  options=options)
        if fileName:
            self.currentFile = fileName
            self.saveDocument(fileName)

    def saveDocument(self, fileName):
        text = self.textEdit.toHtml()
        self.doc.add_paragraph(text)
        self.doc.save(fileName)

    def exportPdf(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self, "Экспорт в PDF", "",
                                                  "PDF (*.pdf);;Все файлы (*)",
                                                  options=options)
        if fileName:
            docx_file = fileName.replace('.pdf', '.docx')
            self.doc.save(docx_file)
            self.convertDocxToPdf(docx_file, fileName)

    def exportTxt(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self, "Экспорт в TXT", "",
                                                  "Текстовый файл (*.txt);;Все файлы (*)",
                                                  options=options)
        if fileName:
            with open(fileName, 'w', encoding='utf-8') as f:
                f.write(self.textEdit.toPlainText())

    def exportZip(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, _ = QFileDialog.getSaveFileName(self, "Экспорт в ZIP", "",
                                                  "ZIP (*.zip);;Все файлы (*)",
                                                  options=options)
        if fileName:
            with zipfile.ZipFile(fileName, 'w', zipfile.ZIP_DEFLATED) as zipf:
                zipf.write(self.currentFile, arcname=self.currentFile.split('/')[-1])

    def convertDocxToPdf(self, docx_file, pdf_file):
        try:
            from docx2pdf import convert
            convert(docx_file, pdf_file)
        except ImportError:
            QMessageBox.warning(self, "Ошибка",
                                "Не установлен модуль docx2pdf. Установите его с помощью 'pip install docx2pdf'")
        except AssertionError as e:
            QMessageBox.warning(self, "Ошибка", f"Ошибка конвертации: {e}")

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
        try:
            self.textEdit.undo()
        except:
            QMessageBox.warning(self, "Ошибка", f"Нет элементов для удаления")

    def redo(self):
        try:
            self.textEdit.redo()
        except:
            QMessageBox.warning(self, "Ошибка", f"Нет элементов для возвращения")

    def setBold(self):
        cursor = self.textEdit.textCursor()
        fmt = cursor.charFormat()
        if self.textEdit.currentFont().weight() == QFont.Normal:
            fmt.setFontWeight(QFont.Black)
        elif self.textEdit.currentFont().weight() == QFont.Black:
            fmt.setFontWeight(QFont.Normal)
        else:
            fmt.setFontWeight(QFont.Normal)
        cursor.mergeCharFormat(fmt)

    def setItalic(self):
        cursor = self.textEdit.textCursor()
        fmt = cursor.charFormat()
        fmt.setFontItalic(not self.textEdit.currentFont().italic())
        cursor.mergeCharFormat(fmt)

    def setUnderline(self):
        cursor = self.textEdit.textCursor()
        fmt = cursor.charFormat()
        currentUnderline = fmt.underlineStyle()
        if currentUnderline == QTextCharFormat.UnderlineStyle.NoUnderline:
            fmt.setUnderlineStyle(QTextCharFormat.UnderlineStyle.SingleUnderline)
        elif currentUnderline == QTextCharFormat.UnderlineStyle.SingleUnderline:
            fmt.setUnderlineStyle(QTextCharFormat.UnderlineStyle.DashUnderline)
        elif currentUnderline == QTextCharFormat.UnderlineStyle.DashUnderline:
            fmt.setUnderlineStyle(QTextCharFormat.UnderlineStyle.DotLine)
        elif currentUnderline == QTextCharFormat.UnderlineStyle.DotLine:
            fmt.setUnderlineStyle(QTextCharFormat.UnderlineStyle.DashDotLine)
        elif currentUnderline == QTextCharFormat.UnderlineStyle.DashDotLine:
            fmt.setUnderlineStyle(QTextCharFormat.UnderlineStyle.DashDotDotLine)
        elif currentUnderline == QTextCharFormat.UnderlineStyle.DashDotDotLine:
            fmt.setUnderlineStyle(QTextCharFormat.UnderlineStyle.WaveUnderline)
        else:
            fmt.setUnderlineStyle(QTextCharFormat.UnderlineStyle.NoUnderline)

        cursor.mergeCharFormat(fmt)
        self.textEdit.setTextCursor(cursor)

    def changeFontSize(self):
        fontSize, ok = QInputDialog.getInt(self, "Размер шрифта", "Введите размер:", 10, 1, 100)
        if ok:
            cursor = self.textEdit.textCursor()
            fmt = cursor.charFormat()
            font = QFont(fmt.font())
            font.setPointSize(fontSize)
            fmt.setFont(font)
            cursor.mergeCharFormat(fmt)
            self.textEdit.setTextCursor(cursor)

    def changeFont(self):
        font, ok = QFontDialog.getFont()
        if ok:
            cursor = self.textEdit.textCursor()
            fmt = cursor.charFormat()
            fmt.setFont(font)
            cursor.setCharFormat(fmt)
            self.textEdit.setTextCursor(cursor)

    def changeTextColor(self):
        color = QColorDialog.getColor(self.textEdit.textColor())
        if color.isValid():
            self.textEdit.setTextColor(color)

    def applyIndent(self):
        cursor = self.textEdit.textCursor()
        block_format = cursor.blockFormat()
        block_format.setIndent(block_format.indent() + 1)
        cursor.setBlockFormat(block_format)

    def applyOutdent(self):
        cursor = self.textEdit.textCursor()
        block_format = cursor.blockFormat()
        current_indent = block_format.indent()

        if current_indent > 0:
            block_format.setIndent(current_indent - 1)
            cursor.setBlockFormat(block_format)

    def changeLineSpacing(self):
        spacing, ok = QInputDialog.getDouble(self, "Интервал между строками", "Введите интервал:", 1.0, 0.1, 3.0, 2)
        if ok:
            self.textEdit.setLineSpacing(spacing)

    def chooseBackgroundColor(self):
        color = QColorDialog.getColor()
        if color.isValid():
            fmt = self.textEdit.textCursor().charFormat()
            fmt.setBackground(QBrush(color))
            self.textEdit.textCursor().setCharFormat(fmt)

    def choosePageColor(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self.textEdit.setStyleSheet(f"background-color: {color.name()};")


def toNext(WindowNext):
    windowNext = WindowNext()
    widget.addWidget(windowNext)
    widget.setCurrentIndex(widget.currentIndex() + 1)


def toBack():
    if widget.count() > 1:
        widget.setCurrentIndex(widget.currentIndex() - 1)
        widget.removeWidget(widget.widget(widget.currentIndex() + 1))


app = QApplication(sys.argv)
widget = QtWidgets.QStackedWidget()
main_window = MainWindow()

widget.addWidget(main_window)
widget.setFixedWidth(800)
widget.setFixedHeight(600)
widget.show()

sys.exit(app.exec_())