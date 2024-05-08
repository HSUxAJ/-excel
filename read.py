import sys
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QFileDialog
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import pyqtSlot
import pandas as pd
import openpyxl

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setGeometry(50, 50, 320, 200)
        self.setWindowTitle("拆解Excel內文")

        file_label_font = QFont()
        file_label_font.setBold(True)

        self.file_label = QLabel(self)
        self.file_label.setFont(file_label_font)
        self.file_label.setGeometry(20, 20, 280, 50)

        import_button = QPushButton("匯入Excel", self)
        import_button.clicked.connect(self.import_file)
        import_button.setGeometry(110, 80, 100, 30)

        download_button = QPushButton("下載處理後的Excel", self)
        download_button.clicked.connect(self.download_processed_excel)
        download_button.setGeometry(70, 120, 180, 30)

    def import_file(self):
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName()

        if file_path:
            # print("Imported file:", file_path)

            # Load the Excel file
            wb = openpyxl.load_workbook(file_path, data_only=True)
            sheet = wb.active

            # Get the name from sheet['B']
            name = sheet['B2'].value[:10]
            if name:
                name = str(name).strip()
            else:
                name = 'Untitled'

            # Process Excel data
            cols = [
                'Company',
                'Transfer Type',
                'Pick Up Date',
                'Pick Up Time',
                'Flight Number',
                'Passenger Name',
                'Phone Number',
                'Vehicle Model',
                'No. of Passenger',
                'From', 'To',
                'Additional Services',
                'Special Requirements',
                'Other Contact Information',
                'Order Number',
                'Total Price',
                'E-mail',
                '駕駛',
                '電話',
                '車號',
                '車型'
                # 'Day/Night Time',
                # 'Flight Departure/Arrival Time',
                # 'Package Name',
            ]
            data = pd.DataFrame(columns=cols)

            for i in range(1, len(sheet['H'])):
                price = str(sheet['L'][i].value)
                order_numbere = str(sheet['C'][i].value)
                passenger_name = str(sheet['G'][i].value)
                info = str(sheet['H'][i].value)
                lines = info.split('\n')
                dic = {
                    'Company': 'KLOOK',
                    'Total Price': price,
                    'Passenger Name': passenger_name,
                    'Order Number': order_numbere,
                    '駕駛': '',
                    '電話': '',
                    '車號': '',
                    '車型': '',
                    'Pick Up Date': '',
                    'Pick Up Time': ''
                    }
                for line in lines:
                    spot_pos = line.find(':')
                    key = line[:spot_pos]
                    val = line[spot_pos+2:]
                    if (key == 'Pick Up Time'):
                        dic['Pick Up Date'], dic['Pick Up Time'] = val.split(' ', 1)
                        continue
                    elif (key == 'Transfer Type'):
                        if (val == 'Pick-up'):
                            val = '接機'
                        else:
                            val = '送機'
                    elif (key == 'Phone Number'):
                        val = '.+' + val
                    if (key == 'Flight Departure/Arrival Time' and 'Pick Up Time' in val):
                        _, tmp_val = val.split(': ', 1)
                        dic['Pick Up Date'], dic['Pick Up Time'] = tmp_val.split(' ', 1)
                    dic[key] = val

                li = []
                for key in cols:
                    li.append(dic[key])

                temp_data = pd.DataFrame([li], columns=cols)
                data = pd.concat([data, temp_data], ignore_index=True)

            self.processed_data = data
            self.file_label.setText("已匯入檔案: " + file_path)

            self.output_name = name

    def download_processed_excel(self):
        if hasattr(self, 'processed_data'):
            if hasattr(self, 'output_name'):
                output_path, _ = QFileDialog.getSaveFileName(self, "儲存處理後的Excel", self.output_name + ".xlsx",
                                                             "Excel Files (*.xlsx)")
            else:
                output_path, _ = QFileDialog.getSaveFileName(self, "儲存處理後的Excel", "", "Excel Files (*.xlsx)")

            if output_path:
                self.processed_data.to_excel(output_path, sheet_name='Sheet1', index=False)
                self.file_label.setText("已儲存處理後的檔案: " + output_path)
        # else:
        #     print("無處理資料可用.")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
