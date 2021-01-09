import sys, smtplib, pyexcel
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QTableWidgetItem
from PyQt5 import QtCore, QtGui, QtWidgets
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


class Ui_Form_help(object):
    def setupUi(self, Form_help):
        Form_help.setObjectName("Form_help")
        Form_help.resize(781, 557)
        Form_help.setMinimumSize(QtCore.QSize(781, 557))
        Form_help.setMaximumSize(QtCore.QSize(781, 557))
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(Form_help.sizePolicy().hasHeightForWidth())
        Form_help.setSizePolicy(sizePolicy)
        Form_help.setToolTipDuration(1)
        self.textBrowser = QtWidgets.QTextBrowser(Form_help)
        self.textBrowser.setGeometry(QtCore.QRect(10, 0, 761, 511))
        self.textBrowser.setObjectName("textBrowser")
        self.pushButton = QtWidgets.QPushButton(Form_help)
        self.pushButton.setGeometry(QtCore.QRect(10, 520, 761, 31))
        self.pushButton.setObjectName("pushButton")

        self.retranslateUi(Form_help)
        self.pushButton.clicked.connect(Form_help.close)
        QtCore.QMetaObject.connectSlotsByName(Form_help)

    def retranslateUi(self, Form_help):
        _translate = QtCore.QCoreApplication.translate
        Form_help.setWindowTitle(_translate("Form_help", "Инструкция по работе с программой"))
        self.textBrowser.setHtml(_translate("Form_help",
                                            "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
                                            "<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
                                            "p, li { white-space: pre-wrap; }\n"
                                            "</style></head><body style=\" font-family:\'MS Shell Dlg 2\'; font-size:8.25pt; font-weight:400; font-style:normal;\">\n"
                                            "<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\'; font-size:12pt;\">Эта программа предназначена для автоматической рассылки писем схожего содержания нескольким адресатам, данные которых заполнены в электронной таблице Excel. Она поддерживает до трех различных текстов писем и позволяет подставлять в текст письма информацию из пяти столбцов электронной таблицы. </span></p>\n"
                                            "<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\'; font-size:12pt;\">Всего 4 простых шага для отправки всех писем:</span></p>\n"
                                            "<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\'; font-size:12pt; color:#0000ff;\">1.</span><span style=\" font-family:\'MS Shell Dlg 2\'; font-size:12pt;\"> Подготовьте в Excel список рассылки с одной строкой заголовка. Первый столбец должен содержать адреса электронной почты адресатов, которым требуется отправить письма. "
                                            "Второй столбец должен содержать номер шаблона с текстом письма (1, 2, или 3). Для отправки одинаковых писем всем адресатам, в этот столбец следует вписывать значение 1. Если оставить второй столбец незаполненным, то письмо данному адресату отправляться не будет. Столбцы с третьего по седьмой Вы можете заполнить по своему усмотрению. Данные из них могут быть использованы в письме. </span></p>\n"
                                            "<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\'; font-size:12pt;\">(Таблица, расположенная в нижней части окна, предоставляет Вам базовые возможности для непосредственного редактирования списка рассылки без использования электронных таблиц. Для добавления новой строки следует нажать F4. Для удаления текста из ячейки, его предварительно нужно выделить.) </span></p>\n"
                                            "<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\'; font-size:12pt; color:#0000ff;\">2. </span><span style=\" font-family:\'MS Shell Dlg 2\'; font-size:12pt;\">Запустите программу. В цветных вкладках «Шаблон №__» наберите письма. Если желаете, чтобы в текст письма подставлялись те или иные данные из таблицы, то, набирая текст письма, впишите, например, %meta1%. Используйте цифры от 1 до 5. Внимание: первая строка цветного поля, в котором набирается письмо, содержит тему письма, а начиная со второй строки – сам текст. Три цветные вкладки соответствуют трем различным текстам писем, которые Вы можете устанавливать индивидуально для каждого адресата. Поле адреса в таблице будет подсвечиваться тем же цветом, что и вкладка с текстом письма. </span></p>\n"
                                            "<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\'; font-size:12pt; color:#0000ff;\">3.</span><span style=\" font-family:\'MS Shell Dlg 2\'; font-size:12pt;\"> В меню «Файл» настройте почтовый smtp-сервер ящика, с которого будут отправлены все письма. Например, «smtp.mail.ru» или «smtp.yandex.ru». Впишите логин и пароль ящика, с которого будут отправлены письма, в соответствующих полях. Сообщение о статусе соединения с сервером появится в статус-баре внизу.</span></p>\n"
                                            "<p align=\"justify\" style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"><span style=\" font-family:\'MS Shell Dlg 2\'; font-size:12pt; color:#0000ff;\">4.</span><span style=\" font-family:\'MS Shell Dlg 2\'; font-size:12pt;\"> Щелкните в таблице по любому из адресов, перейдите в меню «Рассылка писем» и выполните рассылку. Цвет адресов, по которым рассылка прошла успешно, будет изменен на зеленый.</span></p></body></html>"))
        self.pushButton.setText(_translate("Form_help", "Закрыть"))


class Ui_Settings_Dialog(object):
    def setupUi(self, Settings_Dialog):
        Settings_Dialog.setObjectName("Settings_Dialog")
        Settings_Dialog.setWindowModality(QtCore.Qt.ApplicationModal)
        Settings_Dialog.resize(377, 168)
        Settings_Dialog.setMinimumSize(QtCore.QSize(377, 168))
        Settings_Dialog.setMaximumSize(QtCore.QSize(377, 168))
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
        item.setText(_translate("MainWindow", "meta 1"))
        item.setToolTip("для подстановки этого значения используйте в тексте шаблона %meta1%")
        item = self.tableWidget.horizontalHeaderItem(3)
        item.setText(_translate("MainWindow", "meta 2"))
        item.setToolTip("для подстановки этого значения используйте в тексте шаблона %meta2%")
        item = self.tableWidget.horizontalHeaderItem(4)
        item.setText(_translate("MainWindow", "meta 3"))
        item.setToolTip("для подстановки этого значения используйте в тексте шаблона %meta3%")
        item = self.tableWidget.horizontalHeaderItem(5)
        item.setText(_translate("MainWindow", "meta 4"))
        item.setToolTip("для подстановки этого значения используйте в тексте шаблона %meta4%")
        item = self.tableWidget.horizontalHeaderItem(6)
        item.setText(_translate("MainWindow", "meta 5"))
        item.setToolTip("для подстановки этого значения используйте в тексте шаблона %meta5%")
        __sortingEnabled = self.tableWidget.isSortingEnabled()
        self.tableWidget.setSortingEnabled(False)
        item = self.tableWidget.item(0, 0)
        item.setText(_translate("MainWindow", "sample@mail.ру"))
        item.setBackground(QtGui.QColor(230, 255, 255))
        item = self.tableWidget.item(0, 1)
        item.setText(_translate("MainWindow", "1"))
        item.setTextAlignment(QtCore.Qt.AlignCenter)
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
        if event.key() == 16777267:  # нажата F4
            self.tableWidget.insertRow(self.tableWidget.rowCount())


class Base_form(Ui_MainWindow, QMainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowIcon(QtGui.QIcon('res/mail_ico.png'))
        self.action_Exit.triggered.connect(self.base_form_close)
        self.action_Settings.triggered.connect(self.start_settings_dialog)
        self.action_Get_from_Excel.triggered.connect(self.import_from_excel)
        self.action_Export_To_Excel.triggered.connect(self.export_to_excel)
        self.action_Send_Mail.triggered.connect(self.send_mail)
        self.action_Help.triggered.connect(self.start_help)
        self.tableWidget.cellChanged.connect(self.row_column_clicked)

    def start_help(self):
        help_form.show()

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

    def export_to_excel(self):
        a_list = [["Адресат", "Шаблон", "Имя Отчетво", "Фамилия", "Данные №3", "Данные №4", "Данные №5"],
                  ['sample@mail.ру', '1', 'Иван Иванович', 'Иванов', '', '', ''],
                  ['Названия столбцов можно менять,', '', '', '', '', '', ''],
                  ['но первая строка должна содержать названия столбцов.', '', '', '', '', '', ''],
                  ['Не забудьте переименовать файл и удалить эти комментарии.', '', '', '', '', '', '']
                  ]
        pyexcel.save_as(array=a_list, dest_file_name="example.xls")
        self.statusbar.showMessage("Шаблон списка рассылки example.xls создан в текущем каталоге", 3000)

    def send_mail(self):
        self.statusbar.showMessage("Процесс отправки писем запущен. Отправленные письма будут помечены зеленым цветом.", 3000)
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
                        subject = subject.replace("%meta1%", self.tableWidget.item(i, 2).text())
                        subject = subject.replace("%meta2%", self.tableWidget.item(i, 3).text())
                        subject = subject.replace("%meta3%", self.tableWidget.item(i, 4).text())
                        subject = subject.replace("%meta4%", self.tableWidget.item(i, 5).text())
                        subject = subject.replace("%meta5%", self.tableWidget.item(i, 6).text())
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
                        self.tableWidget.item(i, 0).setBackground(QtGui.QColor(0, 255, 0))

                    except Exception:
                        self.statusbar.showMessage(f'письмо {adr} не отправлено', 500)
                        settings_dialog.autorisation = False
                        try:
                            settings_dialog.server = smtplib.SMTP(settings_dialog.lineEdit.displayText())
                            settings_dialog.server.starttls()
                            settings_dialog.server.login(settings_dialog.lineEdit_2.text(), settings_dialog.lineEdit_3.text())
                            settings_dialog.autorisation = True
                            self.statusbar.showMessage("Cоединение с сервером восстановлено.", 500)
                            continue
                        except Exception:
                            settings_dialog.autorisation = False
                            self.statusbar.showMessage("Потеряно соединение с сервером.", 3000)
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
                self.statusbar.showMessage("Cоединение с сервером установлено.", 3000)
            except Exception:
                settings_dialog.autorisation = False
                self.statusbar.showMessage("Ошибка соединения с сервером. Проверьте логин и пароль.", 3000)

    def import_from_excel(self):
        fname = ""
        fname = QFileDialog.getOpenFileName(self, "Открыть список рассылки", "", "Excel (*.xlsx *.xls)")[0]
        try:
            work_sheet = pyexcel.get_array(file_name=fname)
            while self.tableWidget.rowCount() > 0:
                self.tableWidget.removeRow(0)
            i = -1
            for sheet_row in work_sheet:
                i += 1
                if i == 0:
                    continue
                if sheet_row[0] != "":
                    if self.tableWidget.rowCount() < i:
                        self.tableWidget.insertRow(i - 1)
                self.tableWidget.setItem(i - 1, 0, QTableWidgetItem(str(sheet_row[0])))
                try:
                    self.tableWidget.setItem(i - 1, 1, QTableWidgetItem(str(int(sheet_row[1]))))
                    self.tableWidget.item(i - 1, 1).setTextAlignment(QtCore.Qt.AlignCenter)
                    if str(int(sheet_row[1])) == "1":
                        self.tableWidget.item(i - 1, 0).setBackground(QtGui.QColor(230, 255, 255))
                    elif str(int(sheet_row[1])) == "2":
                        self.tableWidget.item(i - 1, 0).setBackground(QtGui.QColor(255, 254, 222))
                    elif str(int(sheet_row[1])) == "3":
                        self.tableWidget.item(i - 1, 0).setBackground(QtGui.QColor(222, 255, 212))
                    else:
                        self.tableWidget.item(i - 1, 0).setBackground(QtGui.QColor(255, 255, 255))
                except Exception:
                    pass
                self.tableWidget.setItem(i - 1, 2, QTableWidgetItem(str(sheet_row[2])))
                self.tableWidget.setItem(i - 1, 3, QTableWidgetItem(str(sheet_row[3])))
                self.tableWidget.setItem(i - 1, 4, QTableWidgetItem(str(sheet_row[4])))
                self.tableWidget.setItem(i - 1, 5, QTableWidgetItem(str(sheet_row[5])))
                self.tableWidget.setItem(i - 1, 6, QTableWidgetItem(str(sheet_row[6])))

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


class Help_form(Ui_Form_help, QtWidgets.QDialog):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowIcon(QtGui.QIcon('res/mail_ico.png'))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    base_form = Base_form()
    settings_dialog = Settings_dialog()
    help_form = Help_form()
    base_form.show()
    sys.exit(app.exec_())
