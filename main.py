import sys  # Импорт системного модуля для работы с параметрами и функциями Python
import pandas as pd  # Импорт библиотеки pandas для работы с таблицами данных
from PyQt5.QtWidgets import *  # Импорт всех виджетов из библиотеки PyQt5

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()  # Вызов конструктора родительского класса
        self.setWindowTitle('Text Duplicator')  # Установка заголовка окна
        self.setGeometry(100, 100, 500, 400)  # Установка размера и позиции окна

        self.text_a_path = None  # Инициализация пути к файлу текст А
        self.table_path = None  # Инициализация пути к файлу таблицы

        # Создание виджетов для текстового файла A
        self.text_a_label = QLabel('Text A:')  # Метка для текстового файла A
        self.text_a_text_edit = QTextEdit()  # Поле для отображения содержимого текстового файла A
        self.text_a_text_edit.setReadOnly(True)  # Установка поля только для чтения
        self.text_a_button = QPushButton('Load Text A')  # Кнопка для загрузки текстового файла A
        self.text_a_button.clicked.connect(self.load_text_a)  # Привязка кнопки к функции загрузки текстового файла A

        # Создание виджетов для таблицы
        self.table_label = QLabel('Table:')  # Метка для таблицы
        self.table_text_edit = QTextEdit()  # Поле для отображения содержимого таблицы
        self.table_text_edit.setReadOnly(True)  # Установка поля только для чтения
        self.table_button = QPushButton('Load Table')  # Кнопка для загрузки таблицы
        self.table_button.clicked.connect(self.load_table)  # Привязка кнопки к функции загрузки таблицы

        # Создание кнопки для генерации текста B
        self.output_button = QPushButton('Generate Text B')  # Кнопка для генерации текста B
        self.output_button.clicked.connect(self.generate_text_b)  # Привязка кнопки к функции генерации текста B

        # Организация виджетов в окне с использованием вертикального макета
        vbox = QVBoxLayout()  # Создание вертикального макета
        vbox.addWidget(self.text_a_label)  # Добавление метки текстового файла A в макет
        vbox.addWidget(self.text_a_text_edit)  # Добавление поля текстового файла A в макет
        vbox.addWidget(self.text_a_button)  # Добавление кнопки текстового файла A в макет
        vbox.addWidget(self.table_label)  # Добавление метки таблицы в макет
        vbox.addWidget(self.table_text_edit)  # Добавление поля таблицы в макет
        vbox.addWidget(self.table_button)  # Добавление кнопки таблицы в макет
        vbox.addWidget(self.output_button)  # Добавление кнопки генерации текста B в макет

        self.setLayout(vbox)  # Установка вертикального макета для окна

    # Метод для загрузки текстового файла A
    def load_text_a(self):
        try:
            file_dialog = QFileDialog(self)  # Создание диалогового окна для выбора файлов
            file_dialog.setNameFilter('Text files (*.txt)')  # Установка фильтра для текстовых файлов
            file_dialog.selectNameFilter('Text files (*.txt)')  # Установка выбора фильтра
            file_dialog.setFileMode(QFileDialog.ExistingFile)  # Установка режима выбора существующего файла
            if file_dialog.exec_():  # Если файл выбран
                self.text_a_path = file_dialog.selectedFiles()[0]  # Получение пути к выбранному файлу
                with open(self.text_a_path, 'r') as file:  # Открытие файла для чтения
                    text_a = file.read()  # Чтение содержимого файла
                self.text_a_text_edit.setText(text_a)  # Отображение содержимого файла в текстовом поле
        except Exception as e:  # Обработка исключений
            QMessageBox.critical(self, "Error", f"Failed to load text file: {e}")  # Отображение сообщения об ошибке

    # Метод для загрузки таблицы из Excel-файла
    def load_table(self):
        try:
            file_dialog = QFileDialog(self)  # Создание диалогового окна для выбора файлов
            file_dialog.setNameFilter('Excel files (*.xls*);;Excel files (*.xlsx)')  # Установка фильтра для файлов Excel
            file_dialog.selectNameFilter('Excel files (*.xls*);;Excel files (*.xlsx)')  # Установка выбора фильтра
            file_dialog.setFileMode(QFileDialog.ExistingFile)  # Установка режима выбора существующего файла
            if file_dialog.exec_():  # Если файл выбран
                self.table_path = file_dialog.selectedFiles()[0]  # Получение пути к выбранному файлу
                if self.table_path.endswith('.xlsx'):  # Если файл имеет расширение .xlsx
                    table = pd.read_excel(self.table_path, engine='openpyxl')  # Чтение файла с использованием openpyxl
                else:
                    table = pd.read_excel(self.table_path, engine='xlrd')  # Чтение файла с использованием xlrd
                self.table_text_edit.setText(table.to_string(index=False))  # Отображение содержимого таблицы в текстовом поле
        except ImportError as e:  # Обработка исключений, связанных с импортом библиотек
            QMessageBox.critical(self, "Error", f"Failed to load the required library: {e}")  # Отображение сообщения об ошибке
        except Exception as e:  # Обработка всех остальных исключений
            QMessageBox.critical(self, "Error", f"Failed to load Excel file: {e}")  # Отображение сообщения об ошибке

    # Метод для генерации текста B
    def generate_text_b(self):
        if not self.text_a_path or not self.table_path:  # Проверка, загружены ли оба файла
            QMessageBox.warning(self, "Warning", "Both Text A and Table need to be loaded first.")  # Отображение предупреждения
            return
        try:
            with open(self.text_a_path, 'r') as f:  # Открытие текстового файла A для чтения
                text_a = f.read()  # Чтение содержимого текстового файла A

            if self.table_path.endswith('.xlsx'):  # Если таблица имеет расширение .xlsx
                table = pd.read_excel(self.table_path, engine='openpyxl')  # Чтение таблицы с использованием openpyxl
            else:
                table = pd.read_excel(self.table_path, engine='xlrd')  # Чтение таблицы с использованием xlrd

            with open('Text_B.txt', 'w') as f:  # Открытие файла для записи результата
                for i in range(1, len(table)):  # Проход по строкам таблицы, начиная со второй
                    text_b = text_a  # Копирование содержимого текстового файла A
                    for j in range(len(table.columns)):  # Проход по столбцам таблицы
                        try:
                            text_b = text_b.replace(str(table.iloc[0, j]), str(table.iloc[i, j]))  # Замена текста
                        except ValueError:  # Обработка исключений
                            print('Это не число. Выходим.' + str(table.iloc[0, j]))  # Вывод сообщения об ошибке
                    f.write(text_b)  # Запись измененного текста в файл
            QMessageBox.information(self, "Success", "Text B generated successfully.")  # Отображение сообщения об успешном завершении
        except Exception as e:  # Обработка всех исключений
            QMessageBox.critical(self, "Error", f"Failed to generate Text B: {e}")  # Отображение сообщения об ошибке


if __name__ == '__main__':
    app = QApplication(sys.argv)  # Создание экземпляра приложения
    window = MainWindow()  # Создание главного окна
    window.show()  # Отображение главного окна
    sys.exit(app.exec_())  # Запуск цикла обработки событий
