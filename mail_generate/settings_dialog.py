# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'settings_dialog.ui'
#
# Created by: PyQt5 UI code generator 5.13.2
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets


class Ui_Settings_Dialog(object):
    def setupUi(self, Settings_Dialog):
        Settings_Dialog.setObjectName("Settings_Dialog")
        Settings_Dialog.setWindowModality(QtCore.Qt.ApplicationModal)
        Settings_Dialog.resize(377, 168)
        self.buttonBox = QtWidgets.QDialogButtonBox(Settings_Dialog)
        self.buttonBox.setGeometry(QtCore.QRect(20, 130, 341, 32))
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel|QtWidgets.QDialogButtonBox.Ok)
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
