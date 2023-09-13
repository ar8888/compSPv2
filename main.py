import module_data as md
import logs

from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QApplication, QMainWindow
import sys
import os



class MyTread(QtCore.QThread):
    mysignal = QtCore.pyqtSignal(str)
    params = {}

    def __init__(self, parent=None):
        QtCore.QThread.__init__(self, parent)

    def run(self):
        try:
            macros_obj = self.params['macros']
            func = self.params['func']
            if func == 'get_sp_m':
                self.result = {'status': 'success', 'func': func}
                macros_obj.get_sp_monitoring(self.params['file_m'], self.mysignal)
            elif func == 'get_sp_1c':
                self.result = {'status': 'success', 'func': func}
                macros_obj.get_sp_1c(self.params['file_1c'], self.mysignal)
            elif func == 'compare_sp':
                self.result = {'status': 'success', 'func': self.params['func']}
                macros_obj.compare_sp(self.params['list_pharma'], self.mysignal)
        except Exception as err_all:
            self.mysignal.emit(f"Непредвиденная ошибка: {err_all}")


class Window(QMainWindow):
    def __init__(self):
        self.fl_file_1c = False
        self.fl_file_m = False
        #данные
        self.macros = md.Macros()
        # описываем окно
        super(Window, self).__init__()
        self.setWindowTitle("Обработчик")
        self.setGeometry(300, 100, 820, 620)
        # Добавлям кнопку выбора файла 1с
        self.btn_1c = QtWidgets.QPushButton(self)  # кнопка загрузить из файла
        self.btn_1c.move(10, 10)
        self.btn_1c.setText("Отчет 1с")
        self.btn_1c.setFixedWidth(150)
        self.btn_1c.clicked.connect(self.click_btn_1c)
        self.l_file_1c = QtWidgets.QLabel(self)
        self.l_file_1c.setGeometry(QtCore.QRect(170, 10, 200, 30))
        # Добавлям кнопку выбора файла мониторинга
        self.btn_m = QtWidgets.QPushButton(self)  # кнопка загрузить из файла
        self.btn_m.move(10, 50)
        self.btn_m.setText("Отчет мониторинг")
        self.btn_m.setFixedWidth(150)
        self.btn_m.clicked.connect(self.click_btn_m)
        self.l_file_m = QtWidgets.QLabel(self)
        self.l_file_m.setGeometry(QtCore.QRect(170, 50, 200, 30))
        #запуск сверки
        self.btn_run = QtWidgets.QPushButton(self)
        self.btn_run.setGeometry(10, 100, 180, 50)
        self.btn_run.setText('Запустить сравнение')
        self.btn_run.clicked.connect(self.click_btn_run)
        self.btn_run.setEnabled(False)

        # список сетей
        self.btn_filter = QtWidgets.QPushButton(self)
        self.btn_filter.setGeometry(QtCore.QRect(450, 10, 220, 30))
        self.btn_filter.setText('Получить список АС')
        self.btn_filter.clicked.connect(self.click_filter)
        self.btn_filter.setEnabled(False)

        self.btn_check_all = QtWidgets.QPushButton(self)
        self.btn_check_all.setGeometry(QtCore.QRect(450, 50, 80, 25))
        self.btn_check_all.setText('Выбрать все')
        self.btn_check_all.clicked.connect(self.click_btn_check_all)
        self.btn_check_all.setEnabled(False)

        self.btn_uncheck_all = QtWidgets.QPushButton(self)
        self.btn_uncheck_all.setGeometry(QtCore.QRect(535, 50, 70, 25))
        self.btn_uncheck_all.setText('Снять все')
        self.btn_uncheck_all.clicked.connect(self.click_btn_uncheck_all)
        self.btn_uncheck_all.setEnabled(False)

        self.txt_pharma_search = QtWidgets.QLineEdit(self)
        self.txt_pharma_search.setGeometry(QtCore.QRect(450, 80, 170, 25))
        self.btn_pharma_search = QtWidgets.QPushButton(self)
        self.btn_pharma_search.setGeometry(QtCore.QRect(630, 80, 50, 25))
        self.btn_pharma_search.setText('найти')
        self.btn_pharma_search.clicked.connect(self.search_pharma)
        self.btn_pharma_search.setEnabled(False)

        self.lb_pharma = QtWidgets.QListWidget(self)
        self.lb_pharma.setGeometry(QtCore.QRect(450, 120, 320, 480))
        self.lb_pharma.setObjectName('Выберите АС')



        # добавляем элемент для логов
        self.lbl_log = QtWidgets.QLabel(self)
        self.lbl_log.setText('вывод информации о работе программы')
        self.lbl_log.setGeometry(QtCore.QRect(10, 380, 300, 20))
        self.txt_logs = QtWidgets.QTextEdit(self)
        self.txt_logs.setGeometry(QtCore.QRect(10, 400, 400, 200))
        self.txt_logs.setReadOnly(True)
        self.txt_logs.setBackgroundRole(QtGui.QPalette.Base)
        p = self.txt_logs.palette()
        p.setColor(self.txt_logs.backgroundRole(), QtGui.QColor(225, 230, 229))
        self.txt_logs.setPalette(p)
        # поток для работы
        self.mythread = MyTread()
        self.mythread.started.connect(self.lock_btn)
        self.mythread.finished.connect(self.mythread_finish)
        self.mythread.mysignal.connect(self.mythread_change, QtCore.Qt.QueuedConnection)



    def click_btn_1c(self):
        tmp = QtWidgets.QFileDialog.getOpenFileName(self, "Выберите файл", "", "Excel files (*.xlsx)")
        file_1c = tmp[0]
        file_name = os.path.basename(file_1c)
        self.l_file_1c.setText(file_name)
        self.mythread.params = {'func': 'get_sp_1c', 'file_1c': file_1c, 'macros': self.macros}
        self.mythread.start()
        self.btn_run.setEnabled(False)
        self.btn_filter.setEnabled(False)
        self.fl_file_1c = True



    def click_btn_m(self):
        tmp = QtWidgets.QFileDialog.getOpenFileName(self, "Выберите файл", "", "Excel files (*.xls)")
        file_m = tmp[0]
        file_name = os.path.basename(file_m)
        self.l_file_m.setText(file_name)
        # предварительно проверить параметр
        self.mythread.params = {'func': 'get_sp_m', 'file_m': file_m, 'macros': self.macros}
        self.mythread.start()
        self.fl_file_m = True


    def lock_btn(self):
        self.btn_1c.setEnabled(False)
        self.btn_m.setEnabled(False)
        self.btn_run.setEnabled(False)
        self.btn_filter.setEnabled(False)

    def unlock_btn(self):
        self.btn_1c.setEnabled(True)
        self.btn_m.setEnabled(True)
        if self.fl_file_1c is True and self.fl_file_m is True:
            self.btn_filter.setEnabled(True)
            self.btn_run.setEnabled(True)


    def mythread_finish(self):
        self.unlock_btn()


    def mythread_change(self, s):
        self.txt_logs.append(s)

    def search_pharma(self):
        txt_search = self.txt_pharma_search.text().lower().strip()
        if txt_search is None or txt_search == '':
            return False
        for i in range(self.lb_pharma.count()):
            pharma = self.lb_pharma.item(i).text().lower()
            if pharma.find(txt_search) > -1:
                self.lb_pharma.setCurrentRow(i)
                break
        return True

    def click_btn_check_all(self):
        for i in range(self.lb_pharma.count()):
            self.lb_pharma.item(i).setCheckState(QtCore.Qt.Checked)

    def click_btn_uncheck_all(self):
        for i in range(self.lb_pharma.count()):
            self.lb_pharma.item(i).setCheckState(QtCore.Qt.Unchecked)

    def click_filter(self):
        list_pharma = self.macros.get_list_pharma()
        if len(list_pharma) > 0:
            for i, option in enumerate(list_pharma):
                item_ch = QtWidgets.QListWidgetItem()
                item_ch.setText(option)
                item_ch.setCheckState(QtCore.Qt.Checked)
                self.lb_pharma.addItem(item_ch)
        self.btn_check_all.setEnabled(True)
        self.btn_uncheck_all.setEnabled(True)
        self.btn_pharma_search.setEnabled(True)


    def click_btn_run(self):
        #проверки добавить
        list_pharma = []
        for i in range(self.lb_pharma.count()):
            if self.lb_pharma.item(i).checkState() == QtCore.Qt.Checked:
                list_pharma.append(self.lb_pharma.item(i).text().strip())
        self.mythread.params = {'func': 'compare_sp', 'list_pharma': list_pharma, 'macros': self.macros}
        self.mythread.start()


def application():
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    application()
