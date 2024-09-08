import sys
import requests
import os
from requests.auth import HTTPBasicAuth
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QMessageBox, QProgressBar  # Import QProgressBar

class DHIS2Importer(QWidget):
    def __init__(self):
        super().__init__()
        self.progress_bar = QProgressBar(self)  # Initialize progress bar
        self.percentage_label = QLabel("0%", self)  # Initialize percentage label
        self.initUI()  # Call initUI after initializing progress bar

    def initUI(self):
        layout = QVBoxLayout()

        # Input fields for URL, username, and password
        self.url_input = QLineEdit(self)
        self.url_input.setPlaceholderText("Enter DHIS2 URL")
        layout.addWidget(QLabel("DHIS2 URL:"))
        layout.addWidget(self.url_input)

        self.username_input = QLineEdit(self)
        self.username_input.setPlaceholderText("Enter Username")
        layout.addWidget(QLabel("Username:"))
        layout.addWidget(self.username_input)

        self.password_input = QLineEdit(self)
        self.password_input.setPlaceholderText("Enter Password")
        self.password_input.setEchoMode(QLineEdit.Password)
        layout.addWidget(QLabel("Password:"))
        layout.addWidget(self.password_input)

        # Upload button
        upload_button = QPushButton("Upload Data", self)
        upload_button.clicked.connect(self.upload_data)
        layout.addWidget(upload_button)

        # Close button
        close_button = QPushButton("Close", self)
        close_button.clicked.connect(self.close)
        layout.addWidget(close_button)

        # Add progress bar and percentage label to layout
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.percentage_label)

        self.setLayout(layout)
        self.setWindowTitle('DHIS2 Data Importer')
        self.show()

    def upload_data(self):
        dhis2_url = self.url_input.text() + '/api/dataValueSets'
        username = self.username_input.text()
        password = self.password_input.text()

        # Path to the CSV file
        file_path = os.path.join(os.path.dirname(__file__), 'Merged', 'file_to_import.csv')

        # Headers for the request
        headers = {
            'Content-Type': 'application/csv'
        }

        # Set progress bar range
        self.progress_bar.setRange(0, 100)  # Set range for percentage
        self.progress_bar.setVisible(True)  # Show progress bar

        # Read the CSV file and prepare it for upload
        with open(file_path, 'rb') as file:
            total_size = os.path.getsize(file_path)  # Get total file size
            uploaded_size = 0

            # Make the POST request to upload the file
            response = requests.post(
                dhis2_url,
                headers=headers,
                data=file,
                auth=HTTPBasicAuth(username, password)
            )

        # Hide progress bar after request
        self.progress_bar.setVisible(False)  # Hide progress bar
        self.percentage_label.setText("100%")  # Set label to 100% after completion

        # Check the response status
        if response.status_code == 200:
            response_json = response.json()
            message = {
                "status": "SUCCESS",
                "importCount": response_json.get("importCount", {})
            }
            QMessageBox.information(self, "Success", f"Data imported successfully!\n{message}")
        else:
            QMessageBox.critical(self, "Error", f"Failed to import data. Status code: {response.status_code}\nResponse: {response.text}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = DHIS2Importer()
    sys.exit(app.exec_())