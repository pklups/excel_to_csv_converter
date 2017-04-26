from PyQt5 import QtCore, QtGui, QtWidgets
import sys
import os
import interface  # import the file that holds the copy of the ui in py

# interface file was prepared by running pyuic5 on .ui file in terminal
# pyuic5 interface.ui -o interface.py


class XlsToCsvApp(QtWidgets.QMainWindow, interface.Ui_MainWindow):
    def __init__(self, parent=None):
        # Explaining super is out of the scope of this article
        # So please google it if you're not familiar with it
        # Simple reason why we use it here is that it allows us to
        # access variables, methods etc in the design.py file
        super(XlsToCsvApp, self).__init__(parent)
        self.setupUi(self)  # This is defined in design.py file automatically
        # It sets up layout and widgets that are defined
        self.btnBrowse.clicked.connect(self.browse_folder)  # When the button is pressed
        # Execute browse_folder function

    def browse_folder(self):
        self.listWidget.clear()  # In case there are any existing elements in the list
        directory = QtWidgets.QFileDialog.getExistingDirectory(self, "Pick a folder")
        # execute getExistingDirectory dialog and set the directory variable to be equal
        # to the user selected directory

        if directory:  # if user didn't pick a directory don't continue
            for file_name in os.listdir(directory):  # for all files, if any, in the directory
                self.listWidget.addItem(file_name)  # add file to the listWidget


def main():
    app = QtWidgets.QApplication(sys.argv)  # A new instance of QApplication
    form = XlsToCsvApp()  # We set the form to be our ExampleApp (design)
    form.show()  # Show the form
    app.exec_()  # and execute the app

if __name__ == '__main__':  # if we're running file directly and not importing it
    main()  # run the main function


