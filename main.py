import sys
import pandas as pd
from PyQt5.QtWidgets import *
#from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QFileDialog, QTextEdit


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Text Duplicator')
        self.setGeometry(100, 100, 500, 400)

        self.text_a_path = None
        self.table_path = None

        self.text_a_label = QLabel('Text A:')
        self.text_a_text_edit = QTextEdit()
        self.text_a_text_edit.setReadOnly(True)
        self.text_a_button = QPushButton('Load Text A')
        self.text_a_button.clicked.connect(self.load_text_a)

        self.table_label = QLabel('Table:')
        self.table_text_edit = QTextEdit()
        self.table_text_edit.setReadOnly(True)
        self.table_button = QPushButton('Load Table')
        self.table_button.clicked.connect(self.load_table)

        self.output_button = QPushButton('Generate Text B')
        self.output_button.clicked.connect(self.generate_text_b)

        vbox = QVBoxLayout()
        vbox.addWidget(self.text_a_label)
        vbox.addWidget(self.text_a_text_edit)
        vbox.addWidget(self.text_a_button)
        vbox.addWidget(self.table_label)
        vbox.addWidget(self.table_text_edit)
        vbox.addWidget(self.table_button)
        vbox.addWidget(self.output_button)

        self.setLayout(vbox)

    def load_text_a(self):
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter('Text files (*.txt)')
        file_dialog.selectNameFilter('Text files (*.txt)')
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        if file_dialog.exec_():
            self.text_a_path = file_dialog.selectedFiles()[0]
            with open(self.text_a_path, 'r') as file:
                text_a = file.read()
            self.text_a_text_edit.setText(text_a)

    def load_table(self):
        file_dialog = QFileDialog(self)
        file_dialog.setNameFilter('Excel files (*.xls*)')
        file_dialog.selectNameFilter('Excel files (*.xls*)')
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        if file_dialog.exec_():
            self.table_path = file_dialog.selectedFiles()[0]
            table = pd.read_excel(self.table_path)
            self.table_text_edit.setText(table.to_string(index=False))

    def generate_text_b(self):
        if not self.text_a_path or not self.table_path:
            return

        with open(self.text_a_path, 'r') as f:
            text_a = f.read()

        table = pd.read_excel(self.table_path)

        with open('Text_B.txt', 'w') as f:
            for i in range(1, len(table)):
                text_b = text_a
                for j in range(len(table.columns)):
                    text_b = text_b.replace(table.iloc[0, j], table.iloc[i, j])
                f.write(text_b)

        self.output_button.setText('Text B generated')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())