import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QVBoxLayout, QWidget, QLabel, QComboBox, QPushButton, QMessageBox
)
from openpyxl import Workbook

class DHIS2PeriodSelector(QMainWindow):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle('DHIS2 Period Selector')
        self.setGeometry(100, 100, 400, 200)

        main_layout = QVBoxLayout()

        # Period Type Combobox
        self.period_type_label = QLabel('Period Type:', self)
        self.period_type_combo = QComboBox(self)
        self.period_type_combo.addItems(['Monthly', 'Quarterly', 'Semi-annual', 'Annual'])
        main_layout.addWidget(self.period_type_label)
        main_layout.addWidget(self.period_type_combo)

        # Year Combobox
        self.year_label = QLabel('Year:', self)
        self.year_combo = QComboBox(self)
        self.year_combo.addItems([str(year) for year in range(2019, 2027)])
        main_layout.addWidget(self.year_label)
        main_layout.addWidget(self.year_combo)

        # Period Combobox
        self.period_label = QLabel('Period:', self)
        self.period_combo = QComboBox(self)
        self.update_periods()
        main_layout.addWidget(self.period_label)
        main_layout.addWidget(self.period_combo)

        # Update periods when period type or year changes
        self.period_type_combo.currentIndexChanged.connect(self.update_periods)
        self.year_combo.currentIndexChanged.connect(self.update_periods)

        # Submit Button
        self.submit_button = QPushButton('Submit', self)
        self.submit_button.clicked.connect(self.submit)
        main_layout.addWidget(self.submit_button)

        container = QWidget()
        container.setLayout(main_layout)
        self.setCentralWidget(container)

    def update_periods(self):
        period_type = self.period_type_combo.currentText()
        year = int(self.year_combo.currentText())
        self.period_combo.clear()

        if period_type == 'Monthly':
            months = [
                'January', 'February', 'March', 'April', 'May', 'June',
                'July', 'August', 'September', 'October', 'November', 'December'
            ]
            self.period_combo.addItems([f'{month}-{year}' for month in months])
        elif period_type == 'Quarterly':
            self.period_combo.addItems([
                f'January to March {year}', 
                f'April to June {year}', 
                f'July to October {year}', 
                f'September to December {year}'
            ])
        elif period_type == 'Semi-annual':
            self.period_combo.addItems([
                f'April - September {year}',
                f'October - March {year + 1}'
            ])
        elif period_type == 'Annual':
            self.period_combo.addItems([
                f'October {year -1 } - September {year}'
            ])

    def submit(self):
        period_type = self.period_type_combo.currentText()
        year = int(self.year_combo.currentText())
        period = self.period_combo.currentText()

        # Determine the period code
        if period_type == 'Monthly':
            month = period.split('-')[0]
            month_number = {
                'January': '01', 'February': '02', 'March': '03', 'April': '04',
                'May': '05', 'June': '06', 'July': '07', 'August': '08',
                'September': '09', 'October': '10', 'November': '11', 'December': '12'
            }[month]
            period_code = f'{year}{month_number}'
        elif period_type == 'Quarterly':
            quarter = period.split(' ')[0]
            quarter_number = {
                'January': 'Q1', 'April': 'Q2', 'July': 'Q3', 'September': 'Q4'
            }[quarter]
            period_code = f'{year}{quarter_number}'
        elif period_type == 'Semi-annual':
            semi_annual = period.split(' ')[0]
            semi_annual_code = {
                'April': 'AprilS1', 'October': 'AprilS2'
            }[semi_annual]
            period_code = f'{year}{semi_annual_code}'
        elif period_type == 'Annual':
            period_code = f'{year-1}Oct'

        # Print the period code in the terminal
        print(f'Period Code: {period_code}')

        QMessageBox.information(self, 'Success', f'Period code: {period_code}')

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = DHIS2PeriodSelector()
    window.show()
    sys.exit(app.exec_())