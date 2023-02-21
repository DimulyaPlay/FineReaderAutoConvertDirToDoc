# -*- coding: utf-8 -*-
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QObject, QThread, pyqtSignal
from threading import Thread
from main_functions import *
import os


class Worker(QThread):
    def __init__(self, agregator, isText, isImage):
        super().__init__()
        self.tagregator = agregator
        self.tistext = isText
        self.tisimage = isImage
        self._Running = True

    add_string_to_log = pyqtSignal(str)

    def run(self):
        current_files = os.listdir(self.tagregator.input_path)
        self.add_string_to_log.emit(f'{ctime()} - Сканирование запущено')
        while self._Running:
            new_files = os.listdir(self.tagregator.input_path)
            new_current_files = [i for i in new_files if i not in current_files and i != os.path.basename(
                self.tagregator.output_path_text) and i != os.path.basename(self.tagregator.output_path_img)]
            if new_current_files:
                self.add_string_to_log.emit(f'{ctime()} - Обнаружены файлы')
                for new_file in new_current_files:
                    filepath_in = self.tagregator.input_path + r'\\' + new_file
                    if os.path.isfile(filepath_in):
                        if self.tistext:
                            self.add_string_to_log.emit(f'{ctime()} - Начата процедура P2T для {filepath_in}')
                            try:
                                self.tagregator.save_to_doc_as_text(filepath_in)
                            except Exception as e:
                                pass
                            self.add_string_to_log.emit(f'{ctime()} - Завершена процедура P2T для {filepath_in}')
                        if self.tisimage:
                            self.add_string_to_log.emit(f'{ctime()} - Начата процедура P2I для {filepath_in}')
                            try:
                                self.tagregator.save_to_docx_as_img(filepath_in)

                            except Exception as e:
                                pass
                        self.add_string_to_log.emit(f'{ctime()} - Завершена процедура P2I для {filepath_in}')
            current_files = new_files
            sleep(self.tagregator.delay)
        self.exit(0)


class Ui_MainWindow(object):
    def setupUi(self, MainWindow, agregator):
        self.mw = MainWindow
        MainWindow.setFixedSize(330, 480)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.tabWidget = QtWidgets.QTabWidget(self.mw)
        self.tabWidget.setGeometry(QtCore.QRect(5, 5, 320, 230))
        self.textTab = QtWidgets.QWidget()
        self.tabWidget.addTab(self.textTab,"В текст")
        self.imgTab = QtWidgets.QWidget()
        self.checkBox_to_text = QtWidgets.QCheckBox(self.textTab)
        self.checkBox_to_text.setGeometry(QtCore.QRect(10, 5, 221, 20))
        self.checkBox_to_text.setFont(font)
        self.checkBox_to_text.setChecked(agregator.toText)
        self.label_import_txt = QtWidgets.QLabel(self.textTab)
        self.label_import_txt.setGeometry(QtCore.QRect(10, 25, 211, 20))
        self.label_import_txt.setFont(font)
        self.lineEdit_import_path_text = QtWidgets.QLineEdit(self.textTab)
        self.lineEdit_import_path_text.setGeometry(QtCore.QRect(10, 45, 275, 20))
        self.lineEdit_import_path_text.setText(agregator.input_path)
        self.toolButton_choose_import_path_text = QtWidgets.QToolButton(self.textTab)
        self.toolButton_choose_import_path_text.setGeometry(QtCore.QRect(290, 45, 20, 20))
        self.toolButton_choose_import_path_text.clicked.connect(lambda: self.openFolderNameDialog(self.lineEdit_import_path_text))
        self.label_export_text = QtWidgets.QLabel(self.textTab)
        self.label_export_text.setGeometry(QtCore.QRect(10, 70, 280, 20))
        self.label_export_text.setFont(font)
        self.lineEdit_export_text_path = QtWidgets.QLineEdit(self.textTab)
        self.lineEdit_export_text_path.setGeometry(QtCore.QRect(10, 90, 275, 20))
        self.lineEdit_export_text_path.setText(agregator.output_path_text)
        self.toolButton_choose_export_text_path = QtWidgets.QToolButton(self.textTab)
        self.toolButton_choose_export_text_path.setGeometry(QtCore.QRect(290, 90, 20, 20))
        self.toolButton_choose_export_text_path.clicked.connect(lambda: self.openFolderNameDialog(self.lineEdit_export_text_path))
        self.checkBox_to_text_use_prefix = QtWidgets.QCheckBox(self.textTab)
        self.checkBox_to_text_use_prefix.setGeometry(QtCore.QRect(10, 120, 221, 20))
        self.checkBox_to_text_use_prefix.setFont(font)
        self.checkBox_to_text_use_prefix.setChecked(agregator.toTextUsePrefix)
        self.lineEdit_text_prefix = QtWidgets.QLineEdit(self.textTab)
        self.lineEdit_text_prefix.setGeometry(QtCore.QRect(10, 140, 275, 20))
        self.lineEdit_text_prefix.setText(agregator.toTextPrefix)

        self.tabWidget.addTab(self.imgTab, 'В изображения')
        # LINES AND LABELS

        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(10, 240, 211, 21))
        self.label_3.setFont(font)
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(10, 190, 281, 21))
        self.label_4.setFont(font)
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(10, 120, 281, 21))
        self.label_5.setFont(font)

        # FOLDER FOR SCAN

        # FOLDER FOR FINEREADER OUTPUT

        # FILECMD PATH
        self.lineEdit_finecmd_path = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_finecmd_path.setGeometry(QtCore.QRect(10, 140, 281, 20))
        self.lineEdit_finecmd_path.setText(agregator.finecmd_path)
        self.toolButton_choose_finecmd_path = QtWidgets.QToolButton(self.centralwidget)
        self.toolButton_choose_finecmd_path.setGeometry(QtCore.QRect(300, 140, 20, 20))
        self.toolButton_choose_finecmd_path.clicked.connect(lambda: self.openFileNameDialog(self.lineEdit_finecmd_path))

        # FOLDER FOR IMG OUTPUT
        self.checkBox_to_wordImages = QtWidgets.QCheckBox(self.centralwidget)
        self.checkBox_to_wordImages.setGeometry(QtCore.QRect(10, 170, 301, 21))
        self.checkBox_to_wordImages.setFont(font)
        self.lineEdit_wordimg_path = QtWidgets.QLineEdit(self.centralwidget)
        self.lineEdit_wordimg_path.setGeometry(QtCore.QRect(10, 210, 281, 20))
        self.lineEdit_wordimg_path.setText(agregator.output_path_img)
        self.toolButton_choose_export_wordimg_path = QtWidgets.QToolButton(self.centralwidget)
        self.toolButton_choose_export_wordimg_path.setGeometry(QtCore.QRect(300, 210, 20, 20))
        self.toolButton_choose_export_wordimg_path.clicked.connect(lambda: self.openFolderNameDialog(self.lineEdit_wordimg_path))

        # DELAY SPINBOX
        self.spinBox_period = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox_period.setGeometry(QtCore.QRect(230, 240, 91, 20))
        self.spinBox_period.setMinimum(1)
        self.spinBox_period.setValue(agregator.delay)

        self.pushButton_start = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_start.setGeometry(QtCore.QRect(10, 270, 71, 31))
        self.pushButton_start.setFont(font)
        self.pushButton_start.clicked.connect(lambda: self.start_worker_scan())

        self.pushButton_stop = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_stop.setGeometry(QtCore.QRect(90, 270, 81, 31))
        self.pushButton_stop.setFont(font)
        self.pushButton_stop.clicked.connect(lambda: self.stop_worker_scan())
        self.pushButton_stop.setDisabled(True)

        self.pushButton_scan_once = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_scan_once.setGeometry(QtCore.QRect(180, 270, 141, 31))
        self.pushButton_scan_once.setFont(font)
        self.pushButton_scan_once.clicked.connect(lambda: self.apply_settings())

        self.plainTextEdit_logger = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.plainTextEdit_logger.setGeometry(QtCore.QRect(5, 310, 320, 148))
        self.plainTextEdit_logger.setReadOnly(True)
        self.plainTextEdit_logger.appendPlainText('Утилита для мониторинга папки и конвертации появляющихся в ней файлов через FineReader в текст, а также просто сохранения в виде изображений на страницах Word (для прикрепления сканов касс определений в базу)')
        self.plainTextEdit_logger.appendPlainText('Dmitry Sosnin, Krasnokamsky gs, github.com/dumulyaplay')

        self.statusBar = QtWidgets.QStatusBar(MainWindow)
        self.statusBar.setFont(font)
        self.statusBar.setSizeGripEnabled(False)
        self.statusBar.showMessage('Сканирование не запущено.')
        MainWindow.setCentralWidget(self.centralwidget)
        MainWindow.setStatusBar(self.statusBar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "PDFtoWORD"))
        self.label_import_txt.setText(_translate("MainWindow", "Введите путь для импорта сканов"))
        self.label_export_text.setText(_translate("MainWindow", "Введите путь для экспорта в .doc (FineReader)"))
        self.label_3.setText(_translate("MainWindow", "Периодичность сканирования, сек:"))
        self.toolButton_choose_import_path_text.setText(_translate("MainWindow", "..."))
        self.toolButton_choose_export_text_path.setText(_translate("MainWindow", "..."))
        self.checkBox_to_text.setText(_translate("MainWindow", "Сохранять как текст (FineReader)"))
        self.checkBox_to_text_use_prefix.setText(_translate("MainWindow", "Реагировать только на префикс"))
        self.checkBox_to_wordImages.setText(_translate("MainWindow", "Сохранять как изображения (Префикс toimg-)"))
        self.toolButton_choose_export_wordimg_path.setText(_translate("MainWindow", "..."))
        self.label_4.setText(_translate("MainWindow", "Введите путь для экспорта в .docx"))
        self.toolButton_choose_finecmd_path.setText(_translate("MainWindow", "..."))
        self.label_5.setText(_translate("MainWindow", "Введите путь к FineCmd.exe"))
        self.pushButton_start.setText(_translate("MainWindow", "Запустить"))
        self.pushButton_stop.setText(_translate("MainWindow", "Остановить"))
        self.pushButton_scan_once.setText(_translate("MainWindow", "Сохранить параметры"))

    def openFileNameDialog(self, lineEdit):
        filepath = QtWidgets.QFileDialog.getOpenFileName(self.mw, 'Направь на путь абсолютный до FineCmd.exe')
        if filepath:
            lineEdit.setText(fr'{filepath[0]}')

    def openFolderNameDialog(self, lineEdit):
        folderpath = QtWidgets.QFileDialog.getExistingDirectory(self.mw, 'Укажи путь')
        if folderpath:
            lineEdit.setText(fr'{folderpath}')

    def start_worker_scan(self):
        self.pushButton_stop.setDisabled(False)
        self.pushButton_start.setDisabled(True)
        self.apply_settings()
        self.thread = QThread()
        self.worker = Worker(agregator, self.checkBox_to_text.isChecked(), self.checkBox_to_wordImages.isChecked())
        self.worker.moveToThread(self.thread)
        self.worker.add_string_to_log.connect(self.addLogRow)
        self.thread.started.connect(self.worker.run)
        self.thread.start()
        self.statusBar.showMessage('Сканирование запущено')

    def addLogRow(self, string):
        self.plainTextEdit_logger.appendPlainText(string)

    def stop_worker_scan(self):
        self.worker._Running = False
        self.worker = None
        self.thread.terminate()
        self.addLogRow('Сканирование было остановлено пользователем.')
        self.statusBar.showMessage('Сканирование не запущено')
        self.pushButton_start.setDisabled(False)
        self.pushButton_stop.setDisabled(True)

    def apply_settings(self):
        agregator.cfg['finecmd_path'] = fr'{self.lineEdit_finecmd_path.text()}'
        agregator.cfg['input_folder'] = fr'{self.lineEdit_import_path_text.text()}'
        agregator.cfg['export_wordtext_folder'] = fr'{self.lineEdit_export_text_path.text()}'
        agregator.cfg['export_wordimages_folder'] = fr'{self.lineEdit_wordimg_path.text()}'
        agregator.cfg['delay'] = self.spinBox_period.value()
        # agregator.cfg['toText']
        # agregator.cfg['toImg']
        # agregator.cfg['toTextUsePrefix']
        # agregator.cfg['toImgUsePrefix']
        # agregator.cfg['toTextPrefix']
        # agregator.cfg['toImgPrefix']
        agregator.write_config_to_file()
        agregator.read_create_config()
        self.addLogRow('Настройки были сохранены')


def my_exception_hook(exctype, value, traceback):
    print(exctype, value, traceback)
    sys._excepthook(exctype, value, traceback)
    sys.exit(1)


sys._excepthook = sys.excepthook
sys.excepthook = my_exception_hook


if __name__ == "__main__":
    import sys
    agregator = main_agregator()
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow, agregator)
    MainWindow.show()
    sys.exit(app.exec_())
