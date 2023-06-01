import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QPushButton, QVBoxLayout, QTextEdit,
                             QDialog, QLabel, QLineEdit, QDialogButtonBox, QFileDialog)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt, QThreadPool, QRunnable, QDateTime


class App(QWidget):

    def __init__(self):
        super().__init__()
        self.title = 'Reports By Lazaro Gonzalez'
        self.left = 100
        self.top = 100
        self.width = 320
        self.height = 300
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        # Set the window icon
        self.setWindowIcon(QIcon(':/icon.png'))

        # Create a layout
        layout = QVBoxLayout()

        # Create a button and style it
        button1 = QPushButton('Generate Daily, Weekly, Monthly, Yearly Stats', self)
        button1.setToolTip('Generate statistics')
        button1.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')

        # Set the cursor to a pointer when hovering the button
        button1.setCursor(Qt.PointingHandCursor)

        # Connect the button to the function that generates the statistics
        button1.clicked.connect(self.on_click)

        # Create a button and style it
        button2 = QPushButton('None Work Codes / Schedule Adherence Reports', self)
        button2.setToolTip('Break Down Non Work Codes By Agent')
        button2.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')

        # Set the cursor to a pointer when hovering the button
        button2.setCursor(Qt.PointingHandCursor)

        # Connect the button to the function that generates the statistics
        button2.clicked.connect(self.non_work_report_click)


         # Create a button and style it
        button3 = QPushButton('Proponisi Stats', self)
        button3.setToolTip('Creates Proponisi csv file')
        button3.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')

        # Set the cursor to a pointer when hovering the button
        button3.setCursor(Qt.PointingHandCursor)

        # Connect the button to the function that generates the statistics
        button3.clicked.connect(self.proponisi_report_click)

        # Create a button and style it
        button4 = QPushButton('Proponisi QA Stats', self)
        button4.setToolTip('Adds QA Stats To Proponisi File')
        button4.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')

        # Set the cursor to a pointer when hovering the button
        button4.setCursor(Qt.PointingHandCursor)

        # Connect the button to the function that generates the statistics
        button4.clicked.connect(self.qa_stats_proponisi_click)

        button5 = QPushButton('Send Quality Assurance', self)
        button5.setToolTip('Import quality assurance')
        button5.setStyleSheet('''
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background-color: #66BB6A;
            }
            QPushButton:pressed {
                background-color: #2E7D32;
            }
        ''')

        # Set the cursor to a pointer when hovering the button
        button5.setCursor(Qt.PointingHandCursor)

        # Connect the button to the function that sends quality assurance
        button5.clicked.connect(self.send_quality_assurance_click)


        # Create a QTextEdit for output
        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)

        # Add the current date and time to the output text
        current_datetime = QDateTime.currentDateTime()
        self.output_text.append(
            f"Date & Time: {current_datetime.toString('yyyy-MM-dd hh:mm:ss')}")
        self.output_text.append(
        "The newly generated stats downloaded from ICBM must be saved "
        "in the same folder as the applications. The filename must always "
        "be kept as 'User Productivity Summary' and should not be changed. \n"
        )
        self.output_text.append(
        "Non Work Codes Report will create report.txt file highlighting "
        "with ***** if an agent went over 15 minutes on a Non Work code. "
        "It will will track all Non Work Codes even 1 second. This must be exported "
        "from ICBM in csv format \n"
        )
        self.output_text.append(
        "Proponisi Stats will create a csv file with the stats of all the agents "
        "this must be exported from ICBM to the same directory as the application in "
        "the format of csv \n"
        )


        # Add buttons and QTextEdit to the layout
        layout.addWidget(button1)
        layout.addWidget(button2)
        layout.addWidget(button3) 
        layout.addWidget(button4) 
        layout.addWidget(button5) 
        layout.addWidget(self.output_text)

        self.setLayout(layout)
        self.show()

    def get_date(self):
        dialog = QDialog()
        dialog.setWindowTitle("Enter Date")

        label = QLabel("Enter the date in mm.dd.yyyy format:", dialog)
        line_edit = QLineEdit(dialog)
        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, dialog)

        layout = QVBoxLayout(dialog)
        layout.addWidget(label)
        layout.addWidget(line_edit)
        layout.addWidget(button_box)

        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)

        result = dialog.exec_()

        if result == QDialog.Accepted:
            return line_edit.text()
        else:
            return None

    def on_click(self):
        # Set the cursor to WaitCursor
        QApplication.setOverrideCursor(Qt.WaitCursor)

        # Ask user to select the stats file
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        file_path, _ = file_dialog.getOpenFileName(self, "Select file", "", "CSV Files (*.csv);;Excel Files (*.xls *.xlsx)")
        if not file_path:
            QApplication.restoreOverrideCursor()
            return
        self.append_output("Selected file: " + file_path)

        # Create a runnable object and run it in a separate thread
        stats_runnable = StatsRunnable(self.append_output, file_path)
        QThreadPool.globalInstance().start(stats_runnable)

        # Set the cursor back to ArrowCursor
        QApplication.restoreOverrideCursor()

    def append_output(self, text):
        self.output_text.append(text)
    
    def non_work_report_click(self):
        from NonWorkCodes import run_report
        try:
            file_dialog = QFileDialog()
            file_dialog.setFileMode(QFileDialog.ExistingFile)
            file_path, _ = file_dialog.getOpenFileName(self, "Select file", "", "CSV Files (*.csv)")
            if not file_path:
                return
            self.append_output("Selected file: " + file_path)
            output_text = run_report(file_path=file_path)
            self.append_output(output_text)
        except Exception as e:
            self.append_output(str(e))

    def proponisi_report_click(self):
        import proponisi
        
        
        date = self.get_date()
        if not date:
            return
        # date_object = datetime.strptime(date, "%m.%d.%Y")
        # Select first file
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        file_path1, _ = file_dialog.getOpenFileName(self, "Select Stats File CSV Format", "", "CSV Files (*.csv)")
        if not file_path1:
            return
        self.append_output("Selected file: " + file_path1)
        
        # Run your proponisi function with only one file
        output_text = proponisi.run_report(date_string=date, file_path1=file_path1)
        self.append_output(output_text)

    def qa_stats_proponisi_click(self):
        import proponisiQA

        self.append_output("To add the QA Stats 4 files must be selected, 1st: Proponisi File Created By 'Proponisi Stats' Button \n")
        self.append_output("2nd select the file agents_list_do_not_delete which is in the same directory as the program \n")
        self.append_output("3rd select the file Boca QA list, exported by ICBM, must be CSV format \n")
        self.append_output("4th select the file Ocoee QA list, exported by ICBM, must be CSV format \n")

        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)

        file_path1, _ = file_dialog.getOpenFileName(self, "Select Proponisi File Created By 'Proponisi Stats' Button", "", "CSV Files (*.csv)")
        if not file_path1:
            return
        self.append_output('1 / 4')

        file_path2, _ = file_dialog.getOpenFileName(self, "Select agents_list_do_not_delete Created By 'Proponisi Stats' Button", "", "CSV Files (*.csv)")
        if not file_path2:
            return
        self.append_output('2 / 4')

        file_path3, _ = file_dialog.getOpenFileName(self, "Select the file Ocoee/Boca QA list, exported by ICBM, must be CSV format", "", "CSV Files (*.csv)")
        if not file_path3:
            return
        self.append_output('3 / 4')

        file_path4, _ = file_dialog.getOpenFileName(self, "Select the file Ocoee/Boca QA list, exported by ICBM, must be CSV format", "", "CSV Files (*.csv)")
        if not file_path4:
            self.append_output('You only selected 1 QA file, make sure to select Boca QA list and Ocoee QA list')
            output_text = proponisiQA.qa_stats_proponisi_click(file_path1=file_path1, file_path2=file_path2, file_path3=file_path3)
        else: 
            self.append_output('4 / 4')
            output_text = proponisiQA.qa_stats_proponisi_click(file_path1=file_path1, file_path2=file_path2, file_path3=file_path3, file_path4=file_path4)

        return 
    
    def send_quality_assurance_click(self):
        import QA

        date_range = self.get_date()
        if not date_range:
            return

        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)

        file_path_1, _ = file_dialog.getOpenFileName(self, "File 1 is the QA in CSV format from ICBM", "", "CSV Files (*.csv)")
        if not file_path_1:
            return
        self.append_output('Selected QA file in CSV')

        file_path_2, _ = file_dialog.getOpenFileName(self, "Select agents_list_do_not_delete Created By 'Proponisi Stats' Button", "", "CSV Files (*.csv)")
        if not file_path_2:
            return
        self.append_output('Selected agents_list_do_not_delete file in CSV')

        try:
            QA.send_qa_to_work_email(file_path_1, file_path_2, date_range)
            self.append_output("Quality assurance sent successfully.")
        except Exception as e:
            self.append_output("Error sending quality assurance:")
            self.append_output(str(e))


class StatsRunnable(QRunnable):
    def __init__(self, append_output_func, file_path):
        super().__init__()
        self.append_output_func = append_output_func
        self.file_path = file_path

    def run(self):
        from stats import run_daily_stats
        try:
            errors = run_daily_stats(file_path=self.file_path)
            if errors:
                self.append_output_func(str(errors))
            else:
                self.append_output_func('Stats created')
        except Exception as e:
            self.append_output_func(str(e))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())

