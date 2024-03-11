import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from PyQt5.uic import loadUi
import os
import pandas as pd

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        current_dir = os.path.dirname(os.path.abspath(__file__))
        ui_file = os.path.join(current_dir, "untitled.ui")
        loadUi(ui_file, self)  # Загружаем файл интерфейса

        # Привязываем обработчики событий к кнопкам
        self.pushButton_main_select.clicked.connect(self.open_file_dialog_main)
        self.pushButton_second_select.clicked.connect(self.open_file_dialog_second)
        self.pushButton_4.clicked.connect(self.compare_and_save)

    def open_file_dialog_main(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        file_name, _ = QFileDialog.getOpenFileName(self, "Выбрать файл", desktop_path, "Excel Files (*.xlsx *.xls)", options=options)
        if file_name:
            self.label_main_file.setText(os.path.basename(file_name))
            self.f1name = file_name

    def open_file_dialog_second(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
        file_name, _ = QFileDialog.getOpenFileName(self, "Выбрать файл", desktop_path, "Excel Files (*.xlsx *.xls)", options=options)
        if file_name:
            self.label_second_file.setText(os.path.basename(file_name))
            self.f2name = file_name

    def compare_and_save(self):
        try:
            df1 = pd.read_excel(self.f1name)
            df2 = pd.read_excel(self.f2name)

            df2.columns = df1.columns
            df2.index = df1.index

            diff_df = df1.compare(df2).droplevel(1, axis=1)

            output_file_name = self.lineEdit_output_name.text()
            if not output_file_name:
                output_file_name = os.path.join(os.path.expanduser('~'), 'Desktop', "Изменения.xlsx")
            else:
                output_file_name = os.path.join(os.path.expanduser('~'), 'Desktop', output_file_name)

            diff_df.to_excel(output_file_name, index=True)
            QMessageBox.information(self, "Готово", f"Сохранение завершено!\nФайл сохранен на рабочем столе как '{output_file_name}'.")
        except Exception as e:
            QMessageBox.warning(self, "Ошибка", f"Произошла ошибка: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
