import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QPushButton,
    QLineEdit, QLabel, QComboBox, QFileDialog, QProgressBar, QTextEdit, QMessageBox,
    QGroupBox, QDialog, QListWidget, QListWidgetItem, QAbstractItemView, QInputDialog
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
import configparser
import os
import pandas as pd
from tableauserverclient import Server, PersonalAccessTokenAuth, PDFRequestOptions
from io import BytesIO
from threading import Thread

class PDFExportApp(QWidget):
    def __init__(self):
        super().__init__()
        self.config_file = 'settings.ini'
        self.config = configparser.ConfigParser()
        self.conditions = []  # List to store condition widgets
        self.excel_path = ""  # Store the current selected Excel file path
        self.tableau_views = []  # Store views from Tableau workbook
        self.country_column = None  # To store the detected country column

        self.setWindowTitle('PDF Export Tool')
        self.setGeometry(100, 100, 800, 600)

        self.initUI()

    def initUI(self):
        main_layout = QVBoxLayout()

        self.setup_server_configuration_panel(main_layout)
        self.setup_file_locations_panel(main_layout)
        self.setup_conditions_panel(main_layout)
        self.setup_progress_panel(main_layout)
        self.setup_control_buttons(main_layout)

        self.setLayout(main_layout)
        
        self.load_configuration()  # Load the configuration after setting up the UI

    def setup_server_configuration_panel(self, layout):
        config_box = QGroupBox("Server Configuration")
        config_box.setFont(QFont('Arial', weight=QFont.Normal))  # Set section title to Normal

        config_layout = QGridLayout()
        config_box.setLayout(config_layout)

        self.server_url = QLineEdit()
        self.token_name = QLineEdit()
        self.token_secret = QLineEdit()
        self.site_id = QLineEdit()
        self.workbook_name = QLineEdit()

        load_views_btn = QPushButton("🔍 Load Views")
        load_views_btn.clicked.connect(self.load_tableau_views)

        reset_btn = QPushButton("🔄 Reset")
        reset_btn.clicked.connect(self.reset_server_configuration)
        
        # Set size policy for smaller button
        reset_btn.setFixedSize(70, 30)  # Width, Height

        config_layout.addWidget(QLabel("Server URL"), 0, 0)
        config_layout.addWidget(self.server_url, 0, 1, 1, 2)
        config_layout.addWidget(QLabel("Token Name"), 1, 0)
        config_layout.addWidget(self.token_name, 1, 1, 1, 2)
        config_layout.addWidget(QLabel("Token Secret"), 2, 0)
        config_layout.addWidget(self.token_secret, 2, 1, 1, 2)
        config_layout.addWidget(QLabel("Site ID"), 3, 0)
        config_layout.addWidget(self.site_id, 3, 1, 1, 2)
        config_layout.addWidget(QLabel("Workbook Name"), 4, 0)
        config_layout.addWidget(self.workbook_name, 4, 1)
        config_layout.addWidget(load_views_btn, 4, 2)
        config_layout.addWidget(reset_btn, 5, 0, 1, 3)

        layout.addWidget(config_box)

    def reset_server_configuration(self):
        self.server_url.clear()
        self.token_name.clear()
        self.token_secret.clear()
        self.site_id.clear()
        self.workbook_name.clear()
        self.tableau_views.clear()  # Clear the loaded views

    def setup_file_locations_panel(self, layout):
        file_box = QGroupBox("File and Folder Locations")
        file_box.setFont(QFont('Arial', weight=QFont.Normal))  # Set section title to Normal

        file_layout = QGridLayout()
        file_box.setLayout(file_layout)

        self.excel_file = QLineEdit()
        self.sheet_name = QComboBox()
        self.output_folder = QLineEdit()

        self.sheet_name.setEnabled(False)  # Disable initially

        browse_excel = QPushButton("📂 Browse")
        browse_excel.clicked.connect(self.browse_excel_file)
        browse_output = QPushButton("📂 Browse")
        browse_output.clicked.connect(self.browse_output_folder)

        reset_btn = QPushButton("🔄 Reset")
        reset_btn.setFixedSize(70, 30)  # Width, Height
        reset_btn.clicked.connect(self.reset_file_locations)

        file_layout.addWidget(QLabel("Excel File"), 0, 0)
        file_layout.addWidget(self.excel_file, 0, 1)
        file_layout.addWidget(browse_excel, 0, 2)
        file_layout.addWidget(QLabel("Sheet Name"), 1, 0)
        file_layout.addWidget(self.sheet_name, 1, 1, 1, 2)
        file_layout.addWidget(QLabel("Output Folder"), 2, 0)
        file_layout.addWidget(self.output_folder, 2, 1)
        file_layout.addWidget(browse_output, 2, 2)
        file_layout.addWidget(reset_btn, 3, 0, 1, 3)

        layout.addWidget(file_box)

    def reset_file_locations(self):
        self.excel_file.clear()
        self.sheet_name.clear()
        self.sheet_name.setEnabled(False)
        self.output_folder.clear()
        self.excel_path = ""

    def browse_excel_file(self):
        file_name, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)")
        if file_name:
            self.excel_file.setText(file_name)
            self.excel_path = file_name  # Store the selected file path
            self.load_sheets()

    def load_sheets(self):
        if self.excel_path:
            try:
                xls = pd.ExcelFile(self.excel_path)
                self.sheet_name.clear()
                self.sheet_name.addItems(xls.sheet_names)
                self.sheet_name.setEnabled(True)  # Enable sheet name selection

                # Set the saved sheet name if it exists in the config
                saved_sheet_name = self.config.get('Paths', 'sheet_name', fallback='')
                if saved_sheet_name and saved_sheet_name in xls.sheet_names:
                    self.sheet_name.setCurrentText(saved_sheet_name)

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not load sheets: {e}")

    def browse_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.output_folder.setText(folder)

    def setup_conditions_panel(self, layout):
        self.conditions_box = QGroupBox("Conditions")
        self.conditions_box.setFont(QFont('Arial', weight=QFont.Normal))  # Set section title to Normal
        self.conditions_layout = QVBoxLayout()
        self.conditions_box.setLayout(self.conditions_layout)

        add_condition_btn = QPushButton("➕ Add Condition")
        add_condition_btn.clicked.connect(self.add_condition_line)

        # Set size policy for smaller button
        add_condition_btn.setFixedSize(120, 30)  # Width, Height

        # Center the button within the layout
        button_layout = QHBoxLayout()
        button_layout.addWidget(add_condition_btn, alignment=Qt.AlignCenter)
        
        self.conditions_layout.addLayout(button_layout)
        self.conditions_layout.addStretch()  # Add stretch to push the button to the top

        layout.addWidget(self.conditions_box)

    def add_condition_line(self, condition=None):
        hbox = QHBoxLayout()

        column_txt = QComboBox()
        column_txt.setEnabled(False)  # Disable until columns are loaded
        type_choice = QComboBox()
        type_choice.addItems(['Equals', 'Higher', 'Lower'])
        value_txt = QLineEdit()
        views_btn = QPushButton("Select Views")
        del_btn = QPushButton("🗑️ Delete")

        exclude_views = []

        if condition:
            column_txt.setCurrentText(condition.get('column', ''))
            type_choice.setCurrentText(condition.get('type', 'Equals'))
            value_txt.setText(condition.get('value', ''))
            exclude_views = condition.get('views', '')

        views_btn.clicked.connect(lambda: self.select_views(hbox, exclude_views))

        hbox.addWidget(QLabel("Field"))
        hbox.addWidget(column_txt)
        hbox.addWidget(QLabel("Type"))
        hbox.addWidget(type_choice)
        hbox.addWidget(QLabel("Value"))
        hbox.addWidget(value_txt)
        hbox.addWidget(views_btn)
        hbox.addWidget(del_btn)

        del_btn.clicked.connect(lambda: self.remove_condition_line(hbox))

        self.conditions_layout.addLayout(hbox)
        self.conditions.append({
            'hbox': hbox,
            'column_txt': column_txt,
            'type_choice': type_choice,
            'value_txt': value_txt,
            'views_btn': views_btn,
            'views': exclude_views
        })

        # Load columns if the Excel file and sheet are already selected
        if self.excel_path and self.sheet_name.currentText():
            self.load_columns_for_conditions()

        # Set the loaded condition's column
        if condition and self.excel_path and self.sheet_name.currentText():
            df = pd.read_excel(self.excel_path, sheet_name=self.sheet_name.currentText())
            column_names = list(df.columns)
            column_txt.addItems(column_names)
            column_txt.setEnabled(True)

            if condition['column'] in column_names:
                column_txt.setCurrentText(condition['column'])

    def remove_condition_line(self, hbox):
        for i in reversed(range(hbox.count())):
            widget = hbox.itemAt(i).widget()
            if widget is not None:
                widget.setParent(None)
        self.conditions = [cond for cond in self.conditions if cond['hbox'] != hbox]

    def select_views(self, hbox, initial_selection):
        dialog = QDialog(self)
        dialog.setWindowTitle("Select Views")

        layout = QVBoxLayout()
        list_widget = QListWidget()

        for view in self.tableau_views:
            item = QListWidgetItem(view)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if view in initial_selection else Qt.Unchecked)
            list_widget.addItem(item)

        list_widget.setSelectionMode(QAbstractItemView.MultiSelection)

        select_btn = QPushButton("Select")
        select_btn.clicked.connect(lambda: self.apply_view_selection(dialog, list_widget, hbox))

        layout.addWidget(list_widget)
        layout.addWidget(select_btn)
        dialog.setLayout(layout)
        dialog.exec_()

    def apply_view_selection(self, dialog, list_widget, hbox):
        selected_views = []
        for index in range(list_widget.count()):
            item = list_widget.item(index)
            if item.checkState() == Qt.Checked:
                selected_views.append(item.text())
        for cond in self.conditions:
            if cond['hbox'] == hbox:
                cond['views'] = selected_views
        dialog.accept()

    def setup_progress_panel(self, layout):
        progress_box = QGroupBox("") # Title here
        progress_box.setFont(QFont('Arial', weight=QFont.Normal))  # Set section title to Normal
        progress_layout = QVBoxLayout()
        progress_box.setLayout(progress_layout)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)

        progress_layout.addWidget(self.progress_bar)
        progress_layout.addWidget(self.log_text)

        layout.addWidget(progress_box)

    def setup_control_buttons(self, layout):
        button_layout = QHBoxLayout()

        start_btn = QPushButton("▶️ Start")
        cancel_btn = QPushButton("❌ Cancel")
        save_btn = QPushButton("💾 Save Configuration")

        button_layout.addWidget(start_btn)
        button_layout.addWidget(cancel_btn)
        button_layout.addWidget(save_btn)

        start_btn.clicked.connect(self.OnStart)
        cancel_btn.clicked.connect(self.OnCancel)
        save_btn.clicked.connect(self.OnSave)

        layout.addLayout(button_layout)

    def load_configuration(self):
        if os.path.exists(self.config_file):
            self.config.read(self.config_file)

            self.server_url.setText(self.config.get('Server', 'url', fallback=''))
            self.token_name.setText(self.config.get('Server', 'token_name', fallback=''))
            self.token_secret.setText(self.config.get('Server', 'token_secret', fallback=''))
            self.site_id.setText(self.config.get('Server', 'site_id', fallback=''))
            self.workbook_name.setText(self.config.get('Server', 'workbook_name', fallback=''))
            self.excel_file.setText(self.config.get('Paths', 'excel_file', fallback=''))
            self.output_folder.setText(self.config.get('Paths', 'output_folder', fallback=''))

            self.excel_path = self.excel_file.text()

            if self.excel_path and os.path.exists(self.excel_path):
                # Defer the loading of sheets to the first user interaction
                self.load_sheets()  # Load sheets from the specified Excel file
                
                saved_sheet_name = self.config.get('Paths', 'sheet_name', fallback='')
                if saved_sheet_name:
                    self.sheet_name.setCurrentText(saved_sheet_name)

            conditions_count = self.config.getint('General', 'conditions_count', fallback=0)
            for idx in range(conditions_count):
                section = f'Condition_{idx}'
                if self.config.has_section(section):
                    condition = {
                        'column': self.config.get(section, 'column', fallback=''),
                        'type': self.config.get(section, 'type', fallback='Equals'),
                        'value': self.config.get(section, 'value', fallback=''),
                        'views': self.config.get(section, 'views', fallback='').split(',')
                    }
                    self.add_condition_line(condition)

    def load_columns_for_conditions(self):
        """ Load columns for the selected sheet to populate condition fields. """
        sheet = self.sheet_name.currentText()
        if self.excel_path and sheet:
            try:
                df = pd.read_excel(self.excel_path, sheet_name=sheet)
                column_names = list(df.columns)

                for cond in self.conditions:
                    # Store the current selected value
                    current_text = cond['column_txt'].currentText()
                    
                    cond['column_txt'].clear()
                    cond['column_txt'].addItems(column_names)
                    cond['column_txt'].setEnabled(True)
                    
                    # Restore the previous selected value if it exists in the new list
                    if current_text in column_names:
                        cond['column_txt'].setCurrentText(current_text)
                    else:
                        # If the previously selected column is not in the new list, reset to the first column
                        cond['column_txt'].setCurrentIndex(0)

            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not load columns: {e}")

    def load_tableau_views(self):
        tableau_server = self.server_url.text()
        token_name = self.token_name.text()
        token_secret = self.token_secret.text()
        tableau_site = self.site_id.text()
        workbook_name = self.workbook_name.text()

        auth = PersonalAccessTokenAuth(token_name, token_secret, site_id=tableau_site)
        server = Server(tableau_server, use_server_version=True)

        try:
            server.auth.sign_in_with_personal_access_token(auth)

            all_workbooks, _ = server.workbooks.get()
            workbook_id = next((workbook.id for workbook in all_workbooks if workbook.name == workbook_name), None)

            if workbook_id is None:
                raise Exception(f"Workbook {workbook_name} not found")

            workbook = server.workbooks.get_by_id(workbook_id)
            server.workbooks.populate_views(workbook)
            self.tableau_views = [view.name for view in workbook.views]

            QMessageBox.information(self, "Info", f"Loaded {len(self.tableau_views)} views from {workbook_name}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to load views: {str(e)}")
        finally:
            server.auth.sign_out()

    def OnStart(self):
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("0%") #Progress: 
        self.log_text.clear()

        self.worker = Thread(target=self.run_task)
        self.worker.start()

    def run_task(self):
        try:
            tableau_server = self.server_url.text()
            token_name = self.token_name.text()
            token_secret = self.token_secret.text()
            tableau_site = self.site_id.text()
            excel_file = self.excel_file.text()
            sheet_name = self.sheet_name.currentText()
            output_folder = self.output_folder.text()
            workbook_name = self.workbook_name.text()

            auth = PersonalAccessTokenAuth(token_name, token_secret, site_id=tableau_site)
            server = Server(tableau_server, use_server_version=True)

            # Sign in to Tableau Server
            server.auth.sign_in_with_personal_access_token(auth)

            df = pd.read_excel(excel_file, sheet_name=sheet_name)

            # Check for the country column here
            self.detect_country_column(df)

            if not self.country_column:
                self.update_log("Country column not found. Please ensure a country column is available.")
                return

            all_workbooks, _ = server.workbooks.get()
            workbook_id = next((workbook.id for workbook in all_workbooks if workbook.name == workbook_name), None)

            if workbook_id is None:
                raise Exception(f"Workbook {workbook_name} not found")

            workbook = server.workbooks.get_by_id(workbook_id)
            server.workbooks.populate_views(workbook)
            all_views = workbook.views

            total_countries = len(df)
            for index, row in df.iterrows():
                country = row[self.country_column]
                iso = row['ISO3']
                region = row['Region']

                excluded_views = set()

                for condition in self.conditions:
                    column = condition['column_txt'].currentText()
                    value = condition['value_txt'].text()
                    views = condition['views']

                    if condition['type_choice'].currentText() == 'Equals' and str(row[column]) == value:
                        excluded_views.update(views)
                    elif condition['type_choice'].currentText() == 'Higher' and row[column] > float(value):
                        excluded_views.update(views)
                    elif condition['type_choice'].currentText() == 'Lower' and row[column] < float(value):
                        excluded_views.update(views)

                included_views = [view for view in all_views if view.name not in excluded_views]
                view_numbering = {view.name: i + 1 for i, view in enumerate(included_views)}

                region_folder = os.path.join(output_folder, region)
                os.makedirs(region_folder, exist_ok=True)
                country_folder = os.path.join(region_folder, country)
                os.makedirs(country_folder, exist_ok=True)

                for view in included_views:
                    view_index = view_numbering[view.name]
                    retries = 0
                    max_retries = 3
                    while retries < max_retries:
                        try:
                            pdf_req_option = PDFRequestOptions(page_type=PDFRequestOptions.PageType.Unspecified)
                            pdf_req_option.vf('ISO3', iso)
                            server.views.populate_pdf(view, pdf_req_option)

                            time.sleep(2)  # Wait for PDF processing

                            pdf_stream = BytesIO(view.pdf)
                            pdf_file = os.path.join(country_folder, f'{view_index}. {view.name}.pdf')
                            with open(pdf_file, 'wb') as f:
                                f.write(pdf_stream.read())
                            self.update_log(f"Saved PDF for {country} - {view_index}. {view.name}")
                            break
                        except Exception as e:
                            self.update_log(f"Attempt {retries + 1} failed for {country} - {view.name}. Error: {str(e)}")
                            retries += 1
                            time.sleep(5)

                self.update_log(f"Completed processing for {country}")
                self.update_progress(int((index + 1) / total_countries * 100))

        except Exception as e:
            self.update_log(f"An error occurred: {str(e)}")
        finally:
            server.auth.sign_out()

    def detect_country_column(self, df):
        """Detect the country column from the given dataframe."""
        self.country_column = None
        possible_country_columns = ['Country', 'country', 'Country Name', 'country name']
        for col in possible_country_columns:
            if col in df.columns:
                self.country_column = col
                self.update_log(f"Detected country column: {col}")
                return

        if not self.country_column:
            # Allow user to select manually if no match is found
            self.country_column, ok = QInputDialog.getItem(
                self, "Select Country Column", 
                "Country column not detected. Please select manually:", list(df.columns), 0, False
            )
            if ok:
                self.update_log(f"User selected country column: {self.country_column}")
            else:
                self.country_column = None  # Reset if user cancels
                self.update_log("Country column selection was cancelled by the user.")

    def update_log(self, message):
        self.log_text.append(message + "\n")

    def update_progress(self, value):
        self.progress_bar.setValue(value)
        self.progress_bar.setFormat(f"{value}%") #Progress: 

    def OnCancel(self):
        self.close()

    def OnSave(self):
        self.config['Server'] = {
            'url': self.server_url.text(),
            'token_name': self.token_name.text(),
            'token_secret': self.token_secret.text(),
            'site_id': self.site_id.text(),
            'workbook_name': self.workbook_name.text()
        }
        self.config['Paths'] = {
            'excel_file': self.excel_file.text(),
            'sheet_name': self.sheet_name.currentText(),  # Save selected sheet name
            'output_folder': self.output_folder.text()
        }

        self.config['General'] = {
            'conditions_count': str(len(self.conditions))
        }

        for idx, cond in enumerate(self.conditions):
            section = f'Condition_{idx}'
            if not self.config.has_section(section):
                self.config.add_section(section)
            self.config[section] = {
                'column': cond['column_txt'].currentText(),
                'type': cond['type_choice'].currentText(),
                'value': cond['value_txt'].text(),
                'views': ','.join(cond['views'])
            }

        with open(self.config_file, 'w') as configfile:
            self.config.write(configfile)
        
        QMessageBox.information(self, "Info", "Configuration saved successfully")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PDFExportApp()
    ex.show()
    sys.exit(app.exec_())
