"""
GUI运行主程序
"""
import sys

from PyQt5 import QtWidgets
from PyQt5.QtGui import QIcon

from src.common import read_config
from src.schedule import Schedule
from ui.main_ui import Ui_MainWindow
import ctypes

ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("myappid")


class MainUi(QtWidgets.QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MainUi, self).__init__(parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setFixedSize(self.width(), self.height())
        self.setWindowIcon(QIcon('../res/schedule.png'))
        self.btn_click()
        self.ui.lineEdit.setPlaceholderText(read_config("default").get("xlsx_template_name"))
        self.ui.lineEdit_2.setPlaceholderText(read_config("default").get("output_xlsx_name"))

    def btn_click(self):
        self.ui.pushButton.clicked.connect(self.main_func)

    def main_func(self):
        self.ui.label_4.setText("正在写入Excel，请稍后……")
        radiobutton_month_list = [self.ui.radioButton, self.ui.radioButton_2, self.ui.radioButton_3,
                                  self.ui.radioButton_4,
                                  self.ui.radioButton_5,
                                  self.ui.radioButton_6, self.ui.radioButton_7, self.ui.radioButton_8,
                                  self.ui.radioButton_9,
                                  self.ui.radioButton_10,
                                  self.ui.radioButton_11, self.ui.radioButton_12]
        month = \
            [radioButton.text().split("月")[0] for radioButton in radiobutton_month_list if radioButton.isChecked()][0]
        radiobutton_month_list = [self.ui.radioButton_14, self.ui.radioButton_13]
        skip_holidays = [radioButton.text() for radioButton in radiobutton_month_list if radioButton.isChecked()][0]
        file_template = self.ui.lineEdit.text()
        file_output = self.ui.lineEdit_2.text()

        try:
            main_function = Schedule(int(month), personal_file_name="personal", xlsx_template_name=file_template,
                                     skip=skip_holidays, xlsx_output_name=file_output)
            main_function.output_xlsx(main_function.creation_date(), main_function.read_personal_information(),
                                      main_function.read_personal_information("leader"))
        except PermissionError:
            self.ui.label_4.setText("Excel文件已被其他应用占用！请关闭文件！")
        else:
            self.ui.label_4.setText("排班表已填写完成!")



def generation_ui():
    app = QtWidgets.QApplication(sys.argv)  # 实例化一个应用对象
    main_ui = MainUi()
    main_ui.show()  # 显示窗口
    sys.exit(app.exec_())  # 程序循环,等待安全退出


if __name__ == '__main__':
    generation_ui()
