# -*- coding: utf-8 -*-
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QObject, QThread, pyqtSignal
from threading import Thread
from main_functions import *
import os
import traceback

# pyinstaller --noconfirm --onefile --windowed --icon "C:/Users/CourtUser/Desktop/release/FineReaderAutoConvertDirToDoc/document-convert.png" --add-data "C:/Users/CourtUser/Desktop/release/FineReaderAutoConvertDirToDoc/document-convert.png;."  "C:/Users/CourtUser/Desktop/release/FineReaderAutoConvertDirToDoc/main_window.py"

class Worker(QThread):
    def __init__(self, agregator, workerType):
        super().__init__()
        self.tagregator = agregator
        self._Running = True
        self.workerType = workerType

    add_string_to_log = pyqtSignal(str)

    def run(self):
        if self.workerType == 'both':
            input_path = self.tagregator.input_path_txt if self.tagregator.toText else self.tagregator.input_path_img
        if self.workerType == 'img':
            input_path = self.tagregator.input_path_img
        if self.workerType == 'txt':
            input_path = self.tagregator.input_path_txt
        current_files = os.listdir(input_path)
        self.add_string_to_log.emit(f'{ctime()} - Сканирование запущено: {self.workerType}' )

        while self._Running:
            work_counter = 1
            new_files = os.listdir(input_path)
            new_current_files = [i for i in new_files if i not in current_files and i != os.path.basename(
                self.tagregator.output_path_text) and i != os.path.basename(self.tagregator.output_path_img)]
            if new_current_files:
                for new_file in new_current_files:
                    prefixImg = True if new_file.startswith(self.tagregator.toImgPrefix) else False
                    prefixTxt = True if new_file.startswith(self.tagregator.toTextPrefix) else False
                    filepath_in = input_path + r'\\' + new_file
                    if os.path.isfile(filepath_in):
                        if self.tagregator.toText and (self.workerType == 'both' or self.workerType == 'txt'):
                            if (self.tagregator.toTextUsePrefix and prefixTxt) or not self.tagregator.toTextUsePrefix:
                                self.add_string_to_log.emit(f'{ctime()} - Начата процедура P2T для {filepath_in}')
                                try:
                                    self.tagregator.save_to_doc_as_text(fr'{filepath_in}')
                                    self.add_string_to_log.emit(
                                        f'№{work_counter}. {ctime()} - Завершена процедура P2T для {filepath_in}')
                                    work_counter += 1
                                except Exception as e:
                                    self.add_string_to_log.emit(
                                        f'{ctime()} - Ошибка {e}')
                                    pass
                        if self.tagregator.toImg and (self.workerType == 'both' or self.workerType == 'img'):
                            if (self.tagregator.toImgUsePrefix and prefixImg) or not self.tagregator.toImgUsePrefix:
                                self.add_string_to_log.emit(f'{self.tagregator.toImgUsePrefix} {prefixImg} {prefixImg}')
                                self.add_string_to_log.emit(f'{ctime()} - Начата процедура P2I для {filepath_in}')
                                try:
                                    self.tagregator.save_to_docx_as_img(filepath_in, self.tagregator.dpi)
                                    self.add_string_to_log.emit(f'№{work_counter}. {ctime()} - Завершена процедура P2I для {filepath_in}')
                                    work_counter += 1
                                except Exception as e:
                                    self.add_string_to_log.emit(
                                        f'{ctime()} - Ошибка {e}')
                                    pass
            current_files = new_files
            sleep(self.tagregator.delay)
        self.exit(0)


class Ui_MainWindow(object):
    def setupUi(self, MainWindow, agregator):
        self.mw = MainWindow
        MainWindow.setFixedSize(335, 480)
        MainWindow.setWindowIcon(QtGui.QIcon(agregator.icon))
        self.tray = QtWidgets.QSystemTrayIcon(MainWindow)
        self.tray.setIcon(QtGui.QIcon(agregator.icon))
        self.tray.setVisible(True)
        show_action = QtWidgets.QAction("Показать", MainWindow)
        quit_action = QtWidgets.QAction("Выйти", MainWindow)
        hide_action = QtWidgets.QAction("Скрыть", MainWindow)
        show_action.triggered.connect(MainWindow.show)
        hide_action.triggered.connect(MainWindow.hide)
        quit_action.triggered.connect(app.quit)
        tray_menu = QtWidgets.QMenu()
        tray_menu.addAction(show_action)
        tray_menu.addAction(hide_action)
        tray_menu.addAction(quit_action)
        self.tray.setContextMenu(tray_menu)
        self.mw.setWindowFlags(QtCore.Qt.Tool)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        font = QtGui.QFont()
        font.setPointSize(10)
        self.tabWidget = QtWidgets.QTabWidget(self.mw)
        self.tabWidget.setGeometry(QtCore.QRect(5, 5, 325, 235))
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
        self.lineEdit_import_path_text.setText(agregator.input_path_txt)
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
        self.label_finecmd_path = QtWidgets.QLabel(self.textTab)
        self.label_finecmd_path.setGeometry(QtCore.QRect(10, 165, 281, 21))
        self.label_finecmd_path.setFont(font)
        self.lineEdit_finecmd_path = QtWidgets.QLineEdit(self.textTab)
        self.lineEdit_finecmd_path.setGeometry(QtCore.QRect(10, 185, 275, 20))
        self.lineEdit_finecmd_path.setText(agregator.finecmd_path)
        self.toolButton_choose_finecmd_path = QtWidgets.QToolButton(self.textTab)
        self.toolButton_choose_finecmd_path.setGeometry(QtCore.QRect(290, 185, 20, 20))
        self.toolButton_choose_finecmd_path.clicked.connect(lambda: self.openFileNameDialog(self.lineEdit_finecmd_path))

        self.tabWidget.addTab(self.imgTab, 'В изображения')
        self.checkBox_to_wordImages = QtWidgets.QCheckBox(self.imgTab)
        self.checkBox_to_wordImages.setGeometry(QtCore.QRect(10, 5, 241, 20))
        self.checkBox_to_wordImages.setFont(font)
        self.checkBox_to_wordImages.setChecked(agregator.toImg)
        self.label_import_img = QtWidgets.QLabel(self.imgTab)
        self.label_import_img.setGeometry(QtCore.QRect(10, 25, 211, 20))
        self.label_import_img.setFont(font)
        self.lineEdit_import_path_img = QtWidgets.QLineEdit(self.imgTab)
        self.lineEdit_import_path_img.setGeometry(QtCore.QRect(10, 45, 275, 20))
        self.lineEdit_import_path_img.setText(agregator.input_path_img)
        self.toolButton_choose_import_path_img = QtWidgets.QToolButton(self.imgTab)
        self.toolButton_choose_import_path_img.setGeometry(QtCore.QRect(290, 45, 20, 20))
        self.toolButton_choose_import_path_img.clicked.connect(lambda: self.openFolderNameDialog(self.lineEdit_import_path_img))
        self.label_export_img = QtWidgets.QLabel(self.imgTab)
        self.label_export_img.setGeometry(QtCore.QRect(10, 70, 281, 21))
        self.label_export_img.setFont(font)
        self.lineEdit_export_img_path = QtWidgets.QLineEdit(self.imgTab)
        self.lineEdit_export_img_path.setGeometry(QtCore.QRect(10, 90, 275, 20))
        self.lineEdit_export_img_path.setText(agregator.output_path_img)
        self.toolButton_choose_export_img_path = QtWidgets.QToolButton(self.imgTab)
        self.toolButton_choose_export_img_path.setGeometry(QtCore.QRect(290, 90, 20, 20))
        self.toolButton_choose_export_img_path.clicked.connect(
            lambda: self.openFolderNameDialog(self.lineEdit_export_img_path))
        self.checkBox_to_img_use_prefix = QtWidgets.QCheckBox(self.imgTab)
        self.checkBox_to_img_use_prefix.setGeometry(QtCore.QRect(10, 120, 221, 20))
        self.checkBox_to_img_use_prefix.setFont(font)
        self.checkBox_to_img_use_prefix.setChecked(agregator.toImgUsePrefix)
        self.lineEdit_img_prefix = QtWidgets.QLineEdit(self.imgTab)
        self.lineEdit_img_prefix.setGeometry(QtCore.QRect(10, 140, 275, 20))
        self.lineEdit_img_prefix.setText(agregator.toImgPrefix)
        self.label_dpi = QtWidgets.QLabel(self.imgTab)
        self.label_dpi.setGeometry(QtCore.QRect(10, 180, 281, 21))
        self.label_dpi.setFont(font)
        self.spinBox_dpi = QtWidgets.QSpinBox(self.imgTab)
        self.spinBox_dpi.setGeometry(QtCore.QRect(215, 180, 101, 20))
        self.spinBox_dpi.setMinimum(75)
        self.spinBox_dpi.setMaximum(600)
        self.spinBox_dpi.setSingleStep(75)
        self.spinBox_dpi.setValue(agregator.dpi)
        # DELAY SPINBOX
        self.label_delay = QtWidgets.QLabel(self.centralwidget)
        self.label_delay.setGeometry(QtCore.QRect(10, 245, 211, 21))
        self.label_delay.setFont(font)
        self.spinBox_period = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBox_period.setGeometry(QtCore.QRect(230, 245, 91, 20))
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

        self.pushButton_save = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_save.setGeometry(QtCore.QRect(180, 270, 141, 31))
        self.pushButton_save.setFont(font)
        self.pushButton_save.clicked.connect(lambda: self.apply_settings())

        self.plainTextEdit_logger = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.plainTextEdit_logger.setGeometry(QtCore.QRect(5, 310, 320, 148))
        self.plainTextEdit_logger.setReadOnly(True)
        self.plainTextEdit_logger.appendPlainText('Утилита для мониторинга папки и конвертации появляющихся в ней файлов через FineReader в текст, а также просто сохранения в виде изображений на страницах Word (для прикрепления сканов касс определений в базу)')
        self.plainTextEdit_logger.appendPlainText('Dmitry Sosnin, Krasnokamsky gs, github.com/dimulyaplay')

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
        self.label_import_img.setText(_translate("MainWindow", "Введите путь для импорта сканов"))
        self.label_export_text.setText(_translate("MainWindow", "Введите путь для экспорта в .doc (FineReader)"))
        self.label_delay.setText(_translate("MainWindow", "Периодичность сканирования, сек:"))
        self.toolButton_choose_import_path_text.setText(_translate("MainWindow", "..."))
        self.toolButton_choose_export_text_path.setText(_translate("MainWindow", "..."))
        self.toolButton_choose_import_path_img.setText(_translate("MainWindow", "..."))
        self.toolButton_choose_export_img_path.setText(_translate("MainWindow", "..."))
        self.checkBox_to_text.setText(_translate("MainWindow", "Сохранять как текст (FineReader)"))
        self.checkBox_to_text_use_prefix.setText(_translate("MainWindow", "Реагировать только на префикс"))
        self.checkBox_to_img_use_prefix.setText(_translate("MainWindow", "Реагировать только на префикс"))
        self.checkBox_to_wordImages.setText(_translate("MainWindow", "Сохранять как изображения в word"))
        self.label_export_img.setText(_translate("MainWindow", "Введите путь для экспорта в .docx"))
        self.toolButton_choose_finecmd_path.setText(_translate("MainWindow", "..."))
        self.label_finecmd_path.setText(_translate("MainWindow", "Введите путь к FineCmd.exe"))
        self.label_dpi.setText(_translate("MainWindow", "DPI растеризации изображений:"))
        self.pushButton_start.setText(_translate("MainWindow", "Запустить"))
        self.pushButton_stop.setText(_translate("MainWindow", "Остановить"))
        self.pushButton_save.setText(_translate("MainWindow", "Сохранить параметры"))

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
        res = self.apply_settings()
        if not res:
            return
        if agregator.same_input_paths:
            self.thread = QThread()
            self.worker = Worker(agregator, 'both')
            self.worker.moveToThread(self.thread)
            self.worker.add_string_to_log.connect(self.addLogRow)
            self.thread.started.connect(self.worker.run)
            self.thread.start()
        else:
            if agregator.toText:
                self.thread1 = QThread()
                self.worker1 = Worker(agregator, 'txt')
                self.worker1.moveToThread(self.thread1)
                self.worker1.add_string_to_log.connect(self.addLogRow)
                self.thread1.started.connect(self.worker1.run)
                self.thread1.start()
            if agregator.toImg:
                self.thread2 = QThread()
                self.worker2 = Worker(agregator, 'img')
                self.worker2.moveToThread(self.thread2)
                self.worker2.add_string_to_log.connect(self.addLogRow)
                self.thread2.started.connect(self.worker2.run)
                self.thread2.start()
        self.statusBar.showMessage('Сканирование запущено')

    def addLogRow(self, string):
        self.plainTextEdit_logger.appendPlainText(string)

    def stop_worker_scan(self):
        try:
            self.worker._Running = False
            self.thread.terminate()
        except:
            pass
        try:
            self.worker1._Running = False
            self.thread1.terminate()
        except:
            pass
        try:
            self.worker2._Running = False
            self.thread2.terminate()
        except:
            pass
        self.addLogRow('Сканирование было остановлено пользователем.')
        self.statusBar.showMessage('Сканирование не запущено')
        self.pushButton_start.setDisabled(False)
        self.pushButton_stop.setDisabled(True)

    def apply_settings(self):
        restricted_paths = {self.lineEdit_export_text_path.text(), self.lineEdit_export_img_path.text()}
        if self.lineEdit_import_path_text.text() in restricted_paths or self.lineEdit_import_path_img.text() in restricted_paths:
            self.addLogRow('Пути входов и выходов не могут совпадать, настройки не были сохранены.')
            return False
        agregator.cfg['finecmd_path'] = fr'{self.lineEdit_finecmd_path.text()}'
        agregator.cfg['input_folder_txt'] = fr'{self.lineEdit_import_path_text.text()}'
        agregator.cfg['input_folder_img'] = fr'{self.lineEdit_import_path_img.text()}'
        agregator.cfg['export_wordtext_folder'] = fr'{self.lineEdit_export_text_path.text()}'
        agregator.cfg['export_wordimages_folder'] = fr'{self.lineEdit_export_img_path.text()}'
        agregator.cfg['delay'] = self.spinBox_period.value()
        agregator.cfg['toText'] = self.checkBox_to_text.isChecked()
        agregator.cfg['toImg'] = self.checkBox_to_wordImages.isChecked()
        agregator.cfg['toTextUsePrefix'] = self.checkBox_to_text_use_prefix.isChecked()
        agregator.cfg['toImgUsePrefix'] = self.checkBox_to_img_use_prefix.isChecked()
        agregator.cfg['toTextPrefix'] = self.lineEdit_text_prefix.text()
        agregator.cfg['toImgPrefix'] = self.lineEdit_img_prefix.text()
        agregator.cfg['dpi'] = self.spinBox_dpi.value()
        agregator.write_config_to_file()
        agregator.read_create_config()
        self.addLogRow('Настройки были сохранены')
        return True


if __name__ == "__main__":
    import sys
    try:
        agregator = main_agregator()
        app = QtWidgets.QApplication(sys.argv)
        MainWindow = QtWidgets.QMainWindow()
        ui = Ui_MainWindow()
        ui.setupUi(MainWindow, agregator)
        if '-start_scan' in sys.argv:
            ui.start_worker_scan()
            ui.tray.showMessage(
                "PDF2WORD",
                "Сканирование запущено",
                QtWidgets.QSystemTrayIcon.Information,
                100)
        else:
            ui.tray.showMessage(
                "PDF2WORD",
                "Приложение запущено",
                QtWidgets.QSystemTrayIcon.Information,
                100)
    except:
        traceback.print_exc()
    sys.exit(app.exec_())
