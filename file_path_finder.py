import sys
import os
import openpyxl
import pyodbc
import btm_up_by_exl
from PyQt5 import uic, QtWidgets, QtGui


def resource_path(relative_path):

    if hasattr(sys, '_MEIPASS2'):

        return os.path.join(sys._MEIPASS2, relative_path)

    return os.path.join(os.path.abspath("."), relative_path)


# Load UI
print(resource_path("file_path_finder.ui"))
form_class = uic.loadUiType(resource_path("file_path_finder.ui"))[0]
print(form_class)


class WindowClass(QtWidgets.QMainWindow, form_class) :

    def __init__(self) :
        super().__init__()
        self.setupUi(self)

        global cursor
        global conn

        # Load and connect OBDC
        conn_String = r'DRIVER={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};' \
                      r'DBQ=' + resource_path("MasterData.xlsx") + ';' \
                      r'ReadOnly=1'

        conn = pyodbc.connect(conn_String, autocommit=True)
        cursor = conn.cursor()

        sQuery = "Select * from Vol_Ver"
        cursor.execute(sQuery)
        volume_list = list(cursor.fetchall())
        # print(volume_list)

        self.exit=QtWidgets.QAction("Exit Application",shortcut=QtGui.QKeySequence("Ctrl+q"),triggered=lambda:self.exit_app(conn,cursor))
        self.addAction(self.exit)

        for row in volume_list:
            self.volume_ver.addItem(str(row[0]))

        self.vol_label.setText(str(volume_list[self.volume_ver.currentIndex()][1]))

        #버튼에 기능을 연결하는 코드
        self.file_path_find_btn.clicked.connect(self.open_file)
        self.save_path_find_btn.clicked.connect(self.save_file)
        self.exe_btn.clicked.connect(self.execution_proc)
        self.volume_ver.currentIndexChanged.connect(lambda : self.on_change_vol(volume_list))
        #self.btn_2.clicked.connect(self.button2Function)

    def closeEvent(self, event):

        print ("User has clicked the red x on the main window")
        cursor.close()
        conn.close()
        event.accept()

    def exit_app(self, conn, cursor):

        print("Shortcut pressed") #verification of shortcut press
        self.close()

    def execution_proc(self):

        sht_name = self.sht_name_list.currentText()

        if self.chk_allocation.isChecked() is True:

            self.job_name.setText("Allocating")
            #btm_up_by_exl.btm_up_chg(read_file_name, conn_String, sht_name, UI_set)
            self.job_name.setText("Done")

        if self.chk_datacube.isChecked() is True:
            self.job_name.setText("Make a data cube")
            btm_up_by_exl.btm_up_datacube(read_file_name, write_file_name, sht_name, self,cursor)
            self.job_name.setText("Done")

    def on_change_vol(self,volume_list):
        self.vol_label.setText(str(volume_list[self.volume_ver.currentIndex()][1]))

    def save_file(self):

        global write_file_name

        write_file_name = QtWidgets.QFileDialog.getOpenFileName(self, "파일 저장","","*.xlsm")[0]

        if write_file_name is not None:
            self.save_path_find_tBox.setText(write_file_name)

    def open_file(self):

        global read_file_name

        # Find file
        read_file_name = QtWidgets.QFileDialog.getOpenFileName(self, "파일 열기 . . .","","*.xlsx")[0]

        if read_file_name is not None:
            # Set combobox
            self.file_path_find_tBox.setText(read_file_name)
            tempBk = openpyxl.load_workbook(read_file_name, data_only=True)
            sht_list = tempBk.sheetnames
            self.sht_name_list.addItems(sht_list)
            tempBk.close()

if __name__ == "__main__" :

    app = QtWidgets.QApplication(sys.argv)
    myWindow = WindowClass()
    myWindow.show()
    app.exec_()