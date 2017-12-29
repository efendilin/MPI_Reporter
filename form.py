# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'form.ui'
#
# Created by: PyQt5 UI code generator 5.9
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_formDialog(object):
    def setupUi(self, formDialog):
        formDialog.setObjectName("formDialog")
        formDialog.resize(491, 409)
        self.verticalLayout = QtWidgets.QVBoxLayout(formDialog)
        self.verticalLayout.setObjectName("verticalLayout")
        self.plainTextEdit = QtWidgets.QPlainTextEdit(formDialog)
        self.plainTextEdit.setObjectName("plainTextEdit")
        self.verticalLayout.addWidget(self.plainTextEdit)
        self.buttonCancel = QtWidgets.QDialogButtonBox(formDialog)
        self.buttonCancel.setOrientation(QtCore.Qt.Horizontal)
        self.buttonCancel.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel|QtWidgets.QDialogButtonBox.Ok)
        self.buttonCancel.setObjectName("buttonCancel")
        self.verticalLayout.addWidget(self.buttonCancel)

        self.retranslateUi(formDialog)
        self.buttonCancel.accepted.connect(formDialog.accept)
        self.buttonCancel.rejected.connect(formDialog.reject)
        QtCore.QMetaObject.connectSlotsByName(formDialog)

    def retranslateUi(self, formDialog):
        _translate = QtCore.QCoreApplication.translate
        formDialog.setWindowTitle(_translate("formDialog", "Dialog"))

