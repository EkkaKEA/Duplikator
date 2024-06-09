pip install pyinstaller
pyinstaller --onefile --noconsole --exclude-module tkinter --exclude-module PyQt5.QtSvg --exclude-module PyQt5.QtMultimedia --clean main.py
auto-py-to-exe