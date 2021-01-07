import sys, smtplib, xlrd
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QTableWidgetItem
from PyQt5 import QtCore, QtGui, QtWidgets
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import email.mime.application

max_count_rows_in_table = 10000
start_row_with_data = 1


class Ui_Settings_Dialog(object):
    def setupUi(self, Settings_Dialog):
        Settings_Dialog.setObjectName("Settings_Dialog")
        Settings_Dialog.setWindowModality(QtCore.Qt.ApplicationModal)
        Settings_Dialog.resize(377, 168)
        self.buttonBox = QtWidgets.QDialogButtonBox(Settings_Dialog)
        self.buttonBox.setGeometry(QtCore.QRect(20, 130, 341, 32))
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel | QtWidgets.QDialogButtonBox.Ok)
        self.buttonBox.setObjectName("buttonBox")
        self.groupBox = QtWidgets.QGroupBox(Settings_Dialog)
        self.groupBox.setGeometry(QtCore.QRect(10, 10, 361, 121))
        self.groupBox.setObjectName("groupBox")
        self.label = QtWidgets.QLabel(self.groupBox)
        self.label.setGeometry(QtCore.QRect(10, 30, 47, 13))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.groupBox)
        self.label_2.setGeometry(QtCore.QRect(10, 60, 47, 13))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.groupBox)
        self.label_3.setGeometry(QtCore.QRect(10, 90, 47, 13))
        self.label_3.setObjectName("label_3")
        self.lineEdit = QtWidgets.QLineEdit(self.groupBox)
        self.lineEdit.setGeometry(QtCore.QRect(50, 30, 301, 20))
        self.lineEdit.setObjectName("lineEdit")
        self.lineEdit_2 = QtWidgets.QLineEdit(self.groupBox)
        self.lineEdit_2.setGeometry(QtCore.QRect(50, 60, 301, 20))
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.lineEdit_3 = QtWidgets.QLineEdit(self.groupBox)
        self.lineEdit_3.setGeometry(QtCore.QRect(50, 90, 301, 20))
        self.lineEdit_3.setEchoMode(QtWidgets.QLineEdit.Password)
        self.lineEdit_3.setObjectName("lineEdit_3")

        self.retranslateUi(Settings_Dialog)
        self.buttonBox.accepted.connect(Settings_Dialog.accept)
        self.buttonBox.rejected.connect(Settings_Dialog.reject)
        QtCore.QMetaObject.connectSlotsByName(Settings_Dialog)

    def retranslateUi(self, Settings_Dialog):
        _translate = QtCore.QCoreApplication.translate
        Settings_Dialog.setWindowTitle(_translate("Settings_Dialog", "Настройка почтового ящика"))
        self.groupBox.setTitle(_translate("Settings_Dialog", "Настройте ящик, с которого будут отправляться письма"))
        self.label.setText(_translate("Settings_Dialog", "Почта"))
        self.label_2.setText(_translate("Settings_Dialog", "Логин"))
        self.label_3.setText(_translate("Settings_Dialog", "Пароль"))
        self.lineEdit.setText(_translate("Settings_Dialog", "smtp.mail.ru"))


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 600)
        MainWindow.setMinimumSize(QtCore.QSize(800, 600))
        MainWindow.setMaximumSize(QtCore.QSize(800, 600))
        icon = QtGui.QIcon.fromTheme("mail")
        MainWindow.setWindowIcon(icon)
        MainWindow.setDocumentMode(False)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.groupBox = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox.setGeometry(QtCore.QRect(10, 10, 781, 291))
        self.groupBox.setObjectName("groupBox")
        self.tabWidget = QtWidgets.QTabWidget(self.groupBox)
        self.tabWidget.setGeometry(QtCore.QRect(5, 20, 772, 266))
        self.tabWidget.setAutoFillBackground(True)
        self.tabWidget.setStyleSheet("QTextEdit { \n"
                                     "      background-color: #fffede; \n"
                                     "}")
        self.tabWidget.setObjectName("tabWidget")
        self.tab1 = QtWidgets.QWidget()
        self.tab1.setObjectName("tab1")
        self.textEdit1 = QtWidgets.QTextEdit(self.tab1)
        self.textEdit1.setGeometry(QtCore.QRect(0, 0, 765, 240))
        self.textEdit1.setStyleSheet("QTextEdit { \n"
                                     "      background-color: #e6ffff; \n"
                                     "}")
        self.textEdit1.setObjectName("textEdit1")
        self.tabWidget.addTab(self.tab1, "")
        self.tab2 = QtWidgets.QWidget()
        self.tab2.setObjectName("tab2")
        self.textEdit2 = QtWidgets.QTextEdit(self.tab2)
        self.textEdit2.setGeometry(QtCore.QRect(0, 0, 765, 240))
        self.textEdit2.setObjectName("textEdit2")
        self.tabWidget.addTab(self.tab2, "")
        self.tab3 = QtWidgets.QWidget()
        self.tab3.setObjectName("tab3")
        self.textEdit3 = QtWidgets.QTextEdit(self.tab3)
        self.textEdit3.setGeometry(QtCore.QRect(0, 0, 765, 240))
        self.textEdit3.setStyleSheet("QTextEdit {      \n"
                                     " background-color: #deffd4; \n"
                                     "}")
        self.textEdit3.setObjectName("textEdit3")
        self.tabWidget.addTab(self.tab3, "")
        self.groupBox_2 = QtWidgets.QGroupBox(self.centralwidget)
        self.groupBox_2.setGeometry(QtCore.QRect(10, 300, 781, 247))
        self.groupBox_2.setObjectName("groupBox_2")
        self.tableWidget = QtWidgets.QTableWidget(self.groupBox_2)
        self.tableWidget.setGeometry(QtCore.QRect(8, 15, 764, 226))
        self.tableWidget.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.tableWidget.setFrameShadow(QtWidgets.QFrame.Plain)
        self.tableWidget.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAlwaysOn)
        self.tableWidget.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
        self.tableWidget.setAlternatingRowColors(False)
        self.tableWidget.setGridStyle(QtCore.Qt.SolidLine)
        self.tableWidget.setWordWrap(True)
        self.tableWidget.setCornerButtonEnabled(True)
        self.tableWidget.setRowCount(7)
        self.tableWidget.setColumnCount(7)
        self.tableWidget.setObjectName("tableWidget")
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(5, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(6, item)
        item = QtWidgets.QTableWidgetItem()
        brush = QtGui.QBrush(QtGui.QColor(240, 240, 240))
        brush.setStyle(QtCore.Qt.NoBrush)
        item.setBackground(brush)
        self.tableWidget.setItem(0, 0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setItem(0, 1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setItem(0, 2, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setItem(0, 3, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setItem(0, 4, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setItem(0, 5, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setItem(0, 6, item)
        self.tableWidget.horizontalHeader().setVisible(True)
        self.tableWidget.horizontalHeader().setCascadingSectionResizes(True)
        self.tableWidget.horizontalHeader().setDefaultSectionSize(104)
        self.tableWidget.horizontalHeader().setHighlightSections(True)
        self.tableWidget.horizontalHeader().setStretchLastSection(True)
        self.tableWidget.verticalHeader().setVisible(False)
        self.tableWidget.verticalHeader().setHighlightSections(True)
        self.tableWidget.verticalHeader().setStretchLastSection(False)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 21))
        self.menubar.setObjectName("menubar")
        self.menu = QtWidgets.QMenu(self.menubar)
        self.menu.setObjectName("menu")
        self.menu_2 = QtWidgets.QMenu(self.menubar)
        self.menu_2.setObjectName("menu_2")
        self.menu_He = QtWidgets.QMenu(self.menubar)
        self.menu_He.setObjectName("menu_He")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.action_Get_from_Excel = QtWidgets.QAction(MainWindow)
        self.action_Get_from_Excel.setObjectName("action_Get_from_Excel")
        self.action_Export_To_Excel = QtWidgets.QAction(MainWindow)
        self.action_Export_To_Excel.setObjectName("action_Export_To_Excel")
        self.action_Settings = QtWidgets.QAction(MainWindow)
        self.action_Settings.setObjectName("action_Settings")
        self.action_Exit = QtWidgets.QAction(MainWindow)
        self.action_Exit.setObjectName("action_Exit")
        self.action_Send_Mail = QtWidgets.QAction(MainWindow)
        self.action_Send_Mail.setObjectName("action_Send_Mail")
        self.action_Help = QtWidgets.QAction(MainWindow)
        self.action_Help.setObjectName("action_Help")
        self.action_About = QtWidgets.QAction(MainWindow)
        self.action_About.setObjectName("action_About")
        self.menu.addAction(self.action_Get_from_Excel)
        self.menu.addAction(self.action_Export_To_Excel)
        self.menu.addAction(self.action_Settings)
        self.menu.addSeparator()
        self.menu.addAction(self.action_Exit)
        self.menu_2.addAction(self.action_Send_Mail)
        self.menu_He.addAction(self.action_Help)
        self.menu_He.addAction(self.action_About)
        self.menubar.addAction(self.menu.menuAction())
        self.menubar.addAction(self.menu_2.menuAction())
        self.menubar.addAction(self.menu_He.menuAction())

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(0)
        # self.action_Exit.triggered.connect(MainWindow.close)
        # QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Рассылка электронных писем на основе шаблонов"))
        self.groupBox.setTitle(_translate("MainWindow",
                                          "Вы можете формировать письма на основе этих шаблонов. Наберите здесь текст письма. Вы можете использовать до трех различных текстов."))
        self.textEdit1.setHtml(_translate("MainWindow",
                                          "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
                                          "<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
                                          "p, li { white-space: pre-wrap; }\n"
                                          "</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:8.25pt; font-weight:400; font-style:normal;\">\n"
                                          "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\';\">Первая строка шаблона используется в качестве темы для письма.</span></p>\n"
                                          "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\';\">Далее здесь нужно набрать сам текст письма, в котором можно использовать мета-подстановки. Например:</p>\n"
                                          "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\';\">Уважаемый %meta1%!</span></p>\n"
                                          "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\';\">Приглашаю Вас пройти онлайн тренировку к зачету для экспертов региональных предметных комиссий ЕГЭ по информатике.  </span></p>\n"
                                          "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\';\">Сайт для прохождения тренировки здесь: </span></p>\n"
                                          "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\';\">Ваш логин:%meta4%</span></p>\n"
                                          "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\';\">Ваш пароль: %meta5%</span></p>\n"
                                          "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-family:\'MS Shell Dlg 2\';\"><br /></p>\n"
                                          "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\';\">С уважением, председатель комиссии,</span></p>\n"
                                          "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\';\">Г.А.Г.</p>\n"
                                          "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\';\"></p><br>\n"
                                          "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\';\">Для добавления в таблицу адресатов новой строки в ручном режиме используйте клавишу F4</span></p></body></html>"))

        "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\';\"> \n"


        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab1), _translate("MainWindow", "Шаблон №1"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab2), _translate("MainWindow", "Шаблон №2"))
        self.textEdit3.setHtml(_translate("MainWindow",
                                          "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
                                          "<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
                                          "p, li { white-space: pre-wrap; }\n"
                                          "</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:8.25pt; font-weight:400; font-style:normal;\">\n"
                                          "<p style=\"-qt-paragraph-type:empty; margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px; font-family:\'MS Shell Dlg 2\';\"><br /></p></body></html>"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab3), _translate("MainWindow", "Шаблон №3"))
        self.groupBox_2.setTitle(
            _translate("MainWindow", "Список рассылки и мета-данные для подстановки в шаблоны писем"))
        stylesheet = "::section{Background-color:rgb(230,230,230)}"
        self.tableWidget.horizontalHeader().setStyleSheet(stylesheet)
        self.tableWidget.setColumnWidth(0, 155)
        self.tableWidget.setColumnWidth(1, 55)
        item = self.tableWidget.horizontalHeaderItem(0)
        item.setText(_translate("MainWindow", "Адресат"))
        item.setToolTip("электронная почта адресата")
        item = self.tableWidget.horizontalHeaderItem(1)
        item.setText(_translate("MainWindow", "Шаблон"))
        item.setToolTip("Номер шаблона 1,2 или 3")
        # item.setTextAlignment(0)
        item = self.tableWidget.horizontalHeaderItem(2)
        item.setText(_translate("MainWindow", "Meta 1"))
        item.setToolTip("для подстановки этого значения используйте в тексте шаблона %meta1%")
        item = self.tableWidget.horizontalHeaderItem(3)
        item.setText(_translate("MainWindow", "Meta 2"))
        item.setToolTip("для подстановки этого значения используйте в тексте шаблона %meta2%")
        item = self.tableWidget.horizontalHeaderItem(4)
        item.setText(_translate("MainWindow", "Meta 3"))
        item.setToolTip("для подстановки этого значения используйте в тексте шаблона %meta3%")
        item = self.tableWidget.horizontalHeaderItem(5)
        item.setText(_translate("MainWindow", "Meta 4"))
        item.setToolTip("для подстановки этого значения используйте в тексте шаблона %meta4%")
        item = self.tableWidget.horizontalHeaderItem(6)
        item.setText(_translate("MainWindow", "Meta 5"))
        item.setToolTip("для подстановки этого значения используйте в тексте шаблона %meta5%")
        __sortingEnabled = self.tableWidget.isSortingEnabled()
        self.tableWidget.setSortingEnabled(False)
        item = self.tableWidget.item(0, 0)
        item.setText(_translate("MainWindow", "sample@mail.ру"))
        item = self.tableWidget.item(0, 1)
        item.setText(_translate("MainWindow", "1"))
        item = self.tableWidget.item(0, 2)
        item.setText(_translate("MainWindow", "Иван Иванович"))
        item = self.tableWidget.item(0, 3)
        item.setText(_translate("MainWindow", "Иванов"))
        item = self.tableWidget.item(0, 5)
        item.setText(_translate("MainWindow", "EGE0264578"))
        item = self.tableWidget.item(0, 6)
        item.setText(_translate("MainWindow", "54684321"))
        self.tableWidget.setSortingEnabled(__sortingEnabled)
        self.menu.setTitle(_translate("MainWindow", "Файл"))
        self.menu_2.setTitle(_translate("MainWindow", "Рассылка писем"))
        self.menu_He.setTitle(_translate("MainWindow", "Справка"))
        self.action_Get_from_Excel.setText(_translate("MainWindow", "Загрузить данные из Excel"))
        self.action_Export_To_Excel.setText(_translate("MainWindow", "Выгрузить шаблон Excel"))
        self.action_Settings.setText(_translate("MainWindow", "Настроить почтовый сервер"))
        self.action_Exit.setText(_translate("MainWindow", "Выход"))
        self.action_Send_Mail.setText(_translate("MainWindow", "Отправить письма всем адресатам"))
        self.action_Help.setText(_translate("MainWindow", "Справка"))
        self.action_About.setText(_translate("MainWindow", "О программе"))

    def keyPressEvent(self, event):
        if event.key() == 16777267: #нажата F4
            self.tableWidget.insertRow(self.tableWidget.rowCount())


class Base_form(Ui_MainWindow, QMainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowIcon(QtGui.QIcon('res/mail_ico.png'))
        self.action_Exit.triggered.connect(self.base_form_close)
        self.action_Settings.triggered.connect(self.start_settings_dialog)
        self.action_Get_from_Excel.triggered.connect(self.import_from_excel)
        self.action_Send_Mail.triggered.connect(self.send_mail)
        self.tableWidget.cellChanged.connect(self.row_column_clicked)
        # self.tableWidget.cellPressed.connect(self.cell_press)

    # def cell_press(self):


    def row_column_clicked(self):
        try:
            cell = self.tableWidget.currentItem().text()
            row = self.tableWidget.currentRow()
            col = self.tableWidget.currentColumn()
            if col == 1:
                value = self.tableWidget.item(row, col)
                if str(cell) == "1":
                    self.tableWidget.item(row, 0).setBackground(QtGui.QColor(230, 255, 255))
                elif str(cell) == "2":
                    self.tableWidget.item(row, 0).setBackground(QtGui.QColor(255, 254, 222))
                elif str(cell) == "3":
                    self.tableWidget.item(row, 0).setBackground(QtGui.QColor(222, 255, 212))
                else:
                    self.tableWidget.item(row, 0).setBackground(QtGui.QColor(255, 255, 255))
        except Exception:
            pass

    def base_form_close(self):
        # здесь можно сохранить данные
        self.close()

    def send_mail(self):
        if not (settings_dialog.autorisation):
            self.start_settings_dialog()
            if not (settings_dialog.autorisation):
                self.statusbar.showMessage("Ошибка авторизации. Письма не отправлены", 2000)
                return
        send_count = 0
        try:
            for i in range(self.tableWidget.rowCount()):
                shablon = self.tableWidget.item(i, 1).text()
                if shablon == '1' or shablon == '2' or shablon == '3':
                    if shablon == '1':
                        text = self.textEdit1.toPlainText().split('\n')
                    elif shablon == '2':
                        text = self.textEdit2.toPlainText().split('\n')
                    else:
                        text = self.textEdit3.toPlainText().split('\n')
                    subject = ""
                    try:
                        subject = text[0]
                    except Exception:
                        pass
                    try:
                        letter_text = "\n".join(text[1:])
                        letter_text = letter_text.replace("%meta1%", self.tableWidget.item(i, 2).text())
                        letter_text = letter_text.replace("%meta2%", self.tableWidget.item(i, 3).text())
                        letter_text = letter_text.replace("%meta3%", self.tableWidget.item(i, 4).text())
                        letter_text = letter_text.replace("%meta4%", self.tableWidget.item(i, 5).text())
                        letter_text = letter_text.replace("%meta5%", self.tableWidget.item(i, 6).text())
                    except Exception:
                        pass
                    try:
                        msg = MIMEMultipart()
                        msg['From'] = settings_dialog.lineEdit_2.text()
                        me = settings_dialog.lineEdit_2.text()
                        msg['To'] = self.tableWidget.item(i, 0).text()
                        adr = self.tableWidget.item(i, 0).text()
                        msg['Subject'] = subject
                        msg.attach(MIMEText(letter_text, 'plain'))
                        settings_dialog.server.sendmail(me, [adr], msg.as_string())
                        send_count += 1
                    except Exception:
                        self.statusbar.showMessage(f'письмо {adr} не отправлено', 500)
        except Exception:
            pass
        self.statusbar.showMessage(f'Отправлено писем: {send_count}', 2000)


    def start_settings_dialog(self):
        if settings_dialog.exec_():
            try:
                settings_dialog.server = smtplib.SMTP(settings_dialog.lineEdit.displayText())
                settings_dialog.server.starttls()
                settings_dialog.server.login(settings_dialog.lineEdit_2.text(), settings_dialog.lineEdit_3.text())
                settings_dialog.autorisation = True
                self.statusbar.showMessage("Cоединение с сервером установлено.", 2000)
            except Exception:
                settings_dialog.autorisation = False
                self.statusbar.showMessage("Ошибка соединения с сервером. Проверьте логин и пароль.", 2000)

    def import_from_excel(self):
        fname = ""
        fname = QFileDialog.getOpenFileName(self, "Открыть список рассылки", "", "Excel (*.xlsx *.xls)")[0]
        try:
            workbook = xlrd.open_workbook(fname, on_demand=True)
            sheet = workbook.get_sheet(0)
            while self.tableWidget.rowCount() > 0:
                self.tableWidget.removeRow(0)
            for i in range(start_row_with_data, max_count_rows_in_table):
                if sheet.cell(i, 0).value != "":
                    if self.tableWidget.rowCount() < i:
                        self.tableWidget.insertRow(i - 1)
                self.tableWidget.setItem(i - 1, 0, QTableWidgetItem(str(sheet.cell(i, 0).value)))
                try:
                    self.tableWidget.setItem(i - 1, 1, QTableWidgetItem(str(int(sheet.cell(i, 1).value))))
                    if str(int(sheet.cell(i, 1).value)) == "1":
                        self.tableWidget.item(i - 1, 0).setBackground(QtGui.QColor(230, 255, 255))
                    elif str(int(sheet.cell(i, 1).value)) == "2":
                        self.tableWidget.item(i - 1, 0).setBackground(QtGui.QColor(255, 254, 222))
                    elif str(int(sheet.cell(i, 1).value)) == "3":
                        self.tableWidget.item(i - 1, 0).setBackground(QtGui.QColor(222, 255, 212))
                    else:
                        self.tableWidget.item(i - 1, 0).setBackground(QtGui.QColor(255, 255, 255))
                except Exception:
                    pass
                self.tableWidget.setItem(i - 1, 2, QTableWidgetItem(str(sheet.cell(i, 2).value)))
                self.tableWidget.setItem(i - 1, 3, QTableWidgetItem(str(sheet.cell(i, 3).value)))
                self.tableWidget.setItem(i - 1, 4, QTableWidgetItem(str(sheet.cell(i, 4).value)))
                self.tableWidget.setItem(i - 1, 5, QTableWidgetItem(str(sheet.cell(i, 5).value)))
                self.tableWidget.setItem(i - 1, 6, QTableWidgetItem(str(sheet.cell(i, 6).value)))

        except Exception:
            self.statusbar.showMessage("Ошибка чтения из файла", 2000)


class Settings_dialog(Ui_Settings_Dialog, QtWidgets.QDialog):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowIcon(QtGui.QIcon('res/mail_ico.png'))
        self.autorisation = False
        self.server_name = ""
        self.server_login = ""
        self.server_password = ""


if __name__ == '__main__':
    app = QApplication(sys.argv)
    base_form = Base_form()
    settings_dialog = Settings_dialog()
    base_form.show()
    sys.exit(app.exec_())
