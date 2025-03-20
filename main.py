import subprocess
import sys
import os.path
from datetime import datetime

import numpy as np
from PyQt6 import QtGui
from PyQt6.QtCore import QCoreApplication, Qt #, QSettings
from PyQt6.QtGui import QStandardItem, QStandardItemModel
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
import pandas as pd

from mainwindow import Ui_MainWindow
from tablemodel import OutputTableModel
from savefile import SaveFileThread
from loadfile import LoadFileThread

class MainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        self.setWindowTitle('Excel Analyzer')
        self.setWindowIcon(QtGui.QIcon(QtGui.QPixmap(icon_path)))

        # Load Settings
        #self.settings = QSettings('Manyullyn17', 'Excel_Analyzer')

        # Variables
        self.fileOne = None
        self.fileTwo = None
        self.fileOneXls = None
        self.fileTwoXls = None
        self.fileOneData = None
        self.fileTwoData = None
        self.twoSheetMode = False
        self.twoFileMode = False
        self.fileOneFinishedLoading = False
        self.fileTwoFinishedLoading = False
        self.sheetOne = None
        self.sheetTwo = None
        self.operation = 0  # 0 -> Find Duplicates, 1 -> Find Differences,
                            # 2 -> Count Occurences, 3 -> Change Date Format
        self.mode = 0
        self.differenceModes = ['--  Column Modes  --', 'Strict', 'Unordered', '--  Row Modes  --', 'Strict', 'Strict-Column', 'Loose']
        #                                  0                1          2                3               4           5             6
        self.dateFormats = ['%Y-%m-%d', '%d.%m.%Y', '%d-%m-%Y', '%m/%d/%Y', '%m-%d-%Y', 'Excel']
        self.table_model = OutputTableModel(pd.DataFrame())

        self.load_thread_one = None
        self.load_thread_two = None
        self.save_thread = None

        # Settings
        # self.setting1 = False

        # self.update_settings()

        # Set initial state
        self.outputTable.setVisible(False)
        self.operation_change()

        self.outputTable.setModel(self.table_model)

        # Connect Widgets to methods
        self.oneFileRadioButton.clicked.connect(self.switch_1_file_2_files)
        self.twoFileRadioButton.clicked.connect(self.switch_1_file_2_files)
        self.selectFileOneButton.clicked.connect(self.load_file_one)
        self.selectFileTwoButton.clicked.connect(self.load_file_two)
        self.fileOneList.clicked.connect(self.select_sheet_one)
        self.fileTwoList.clicked.connect(self.select_sheet_two)
        self.operationSelectBox.currentIndexChanged.connect(self.operation_change)
        self.modeSelectBox.currentIndexChanged.connect(self.mode_change)
        self.executeButton.clicked.connect(self.execute)
        self.saveButton.clicked.connect(self.execute_with_save)
        self.showTableButton.clicked.connect(self.toggle_table)

    def switch_1_file_2_files(self):
        if self.oneFileRadioButton.isChecked():
            self.twoFileMode = False
            self.selectFileTwoButton.setVisible(False)
            self.selectFileOneButton.setText('Select File')
            self.fileTwoPath.setVisible(False)
            if self.fileTwoList.count() == 0 and self.fileOneXls:
                for sheet in self.fileOneXls.sheet_names:
                    self.fileTwoList.addItem(sheet)
        else:
            self.twoFileMode = True
            self.selectFileTwoButton.setVisible(True)
            self.selectFileOneButton.setText('Select File 1')
            self.fileTwoPath.setVisible(True)
            self.fileTwoList.clear()

    def load_file_one(self):
        file, _ = QFileDialog.getOpenFileName(self, 'Select Input Excel File', '', 'Excel Files (*.xlsx;*.xls)')

        if file:
            self.fileOneFinishedLoading = False
            self.fileOneList.clear()
            self.fileOneList.addItem('Loading File...')
            QCoreApplication.processEvents()
            self.fileOne = file
            self.fileOnePath.clear()
            self.fileOnePath.setText(self.fileOne)

            try:
                self.fileOneXls = pd.ExcelFile(self.fileOne)
                sheet_names = self.fileOneXls.sheet_names
                self.fileOneList.addItem(f'Loading {len(sheet_names)} sheets...')

                self.load_thread_one = LoadFileThread(self.fileOne, 1)
                self.load_thread_one.finished.connect(self.on_file_loaded)
                self.load_thread_one.error.connect(self.on_file_load_error)
                self.load_thread_one.start()

            except Exception as e:
                QMessageBox.critical(self, 'Error', f'Failed to load file 1: {e}')
                self.fileOneList.clear()
                self.fileOneList.addItem('Failed to load file 1')

    def load_file_two(self):
        file, _ = QFileDialog.getOpenFileName(self, 'Select Input Excel File', '', 'Excel Files (*.xlsx;*.xls)')

        if file:
            self.fileTwoFinishedLoading = False
            self.fileTwoList.clear()
            self.fileTwoList.addItem('Loading File...')
            QCoreApplication.processEvents()
            self.fileTwo = file
            self.fileTwoPath.clear()
            self.fileTwoPath.setText(self.fileTwo)

            try:
                self.fileTwoXls = pd.ExcelFile(self.fileTwo)
                sheet_names = self.fileTwoXls.sheet_names
                self.fileTwoList.addItem(f'Loading {len(sheet_names)} sheets...')

                self.load_thread_one = LoadFileThread(self.fileTwo, 2)
                self.load_thread_one.finished.connect(self.on_file_loaded)
                self.load_thread_one.error.connect(self.on_file_load_error)
                self.load_thread_one.start()

            except Exception as e:
                QMessageBox.critical(self, 'Error', f'Failed to load file 2: {e}')
                self.fileTwoList.clear()
                self.fileTwoList.addItem('Failed to load file 2')

    def on_file_loaded(self, loaded_data: pd.DataFrame, filenum: int):
        if filenum == 1:
            self.fileOneData = loaded_data
            self.fileOneList.clear()
            if not self.twoFileMode:
                self.fileTwoList.clear()
                for sheet in self.fileOneXls.sheet_names:
                    self.fileOneList.addItem(sheet)
                    self.fileTwoList.addItem(sheet)
            else:
                for sheet in self.fileOneXls.sheet_names:
                    self.fileOneList.addItem(sheet)
            if self.operation in (2, 3):
                self.fileTwoList.clear()
                self.fileTwoList.addItem('Select sheet to see columns...')
            self.fileOneFinishedLoading = True
        elif filenum == 2:
            self.fileTwoData = loaded_data
            self.fileTwoList.clear()
            for sheet in self.fileTwoXls.sheet_names:
                self.fileTwoList.addItem(sheet)
            self.fileTwoFinishedLoading = True
        else:
            QMessageBox.critical(self, 'Error', 'No filenum provided but loading finished')
    
    def on_file_load_error(self, e: Exception, filenum: int):
        QMessageBox.critical(self, 'Error', f'Failed to load file {filenum}: {e}')
        if filenum == 1:
            self.fileOneList.clear()
            self.fileOneList.addItem(f'Failed to load file')
            if not self.twoFileMode:
                self.fileTwoList.clear()
                self.fileTwoList.addItem(f'Failed to load file')
        elif filenum == 2:
            self.fileTwoList.clear()
            self.fileTwoList.addItem(f'Failed to load file')
        else:
            QMessageBox.critical(self, 'Error', 'No filenum provided and loading failed')

    def select_sheet_one(self):
        if self.fileOneFinishedLoading:
            try: # Get selected sheet name
                self.sheetOne = self.fileOneXls.sheet_names[self.fileOneList.currentIndex().row()]
                if self.operation in (2, 3):
                    self.fileTwoList.clear()
                    self.fileTwoList.addItem('All')

                    if self.sheetOne in self.fileOneData:
                        columns = self.fileOneData[self.sheetOne].columns
                        for column in columns:
                            self.fileTwoList.addItem(column)
            except (IndexError, KeyError, AttributeError):
                return
        else:
            QMessageBox.information(self, 'File not loaded', 'Please wait for File One to finish loading')

    def select_sheet_two(self):
        if self.twoFileMode: # Two File Mode
            if self.fileTwoFinishedLoading:
                try: # Get selected sheet name
                    self.sheetTwo = self.fileTwoXls.sheet_names[self.fileTwoList.currentIndex().row()]
                except (IndexError, KeyError, AttributeError):
                    return
            else:
                QMessageBox.information(self, 'File not loaded', 'Please wait for File Two to finish loading')

        elif self.operation in (2, 3):
            self.sheetTwo = self.fileTwoList.itemFromIndex(self.fileTwoList.currentIndex()).text()

        else: # One File Mode
            if self.fileOneFinishedLoading:
                try:  # Get selected sheet name
                    self.sheetTwo = self.fileOneXls.sheet_names[self.fileTwoList.currentIndex().row()]
                except (IndexError, KeyError, AttributeError):
                    return
            else:
                QMessageBox.information(self, 'File not loaded', 'Please wait for File One to finish loading')

    def operation_change(self):
        self.operation = self.operationSelectBox.currentIndex()

        def set_two_sheet_mode(state: bool):
            self.twoSheetMode = state
            self.fileTwoList.setVisible(state)
            self.sheetTwoLabel.setVisible(state)
            if not state:
                self.sheetOneLabel.setText('Sheet:')

        def set_two_file_mode(state: bool):
            if state:
                self.twoFileRadioButton.setToolTip(None)
            else:
                self.twoFileRadioButton.setToolTip('To enable Two File Mode select a different Operation Mode')
                self.oneFileRadioButton.setChecked(True)
                QCoreApplication.processEvents()
                self.switch_1_file_2_files()
            self.twoFileRadioButton.setEnabled(state)

        def set_column_list():
            self.sheetOneLabel.setText('Sheet:')
            self.sheetTwoLabel.setText('Column:')
            self.fileTwoList.clear()
            self.fileTwoList.addItem('Select sheet to see columns...')

        def set_sheet_list():
            self.sheetOneLabel.setText('Sheet 1:')
            self.sheetTwoLabel.setText('Sheet 2:')
            self.fileTwoList.clear()
            if self.twoFileMode and self.fileTwoXls:
                for sheet in self.fileTwoXls.sheet_names:
                    self.fileTwoList.addItem(sheet)
            elif self.fileOneXls:
                for sheet in self.fileOneXls.sheet_names:
                    self.fileTwoList.addItem(sheet)

        def set_mode_select(state: bool):
            self.modeLabel.setVisible(state)
            self.modeSelectBox.setVisible(state)
            self.modeSelectBox.clear()

        if self.operation == 0: # Find Duplicates
            set_two_sheet_mode(False)
            set_two_file_mode(False)
            set_mode_select(False)

        elif self.operation == 1: # Find Differences
            set_two_sheet_mode(True)
            set_sheet_list()
            set_two_file_mode(True)
            set_mode_select(True)
            model = QStandardItemModel()
            for mode in self.differenceModes:
                if '--' in mode:
                    separator = QStandardItem(mode)
                    separator.setFlags(Qt.ItemFlag.NoItemFlags)
                    model.appendRow(separator)
                else:
                    item = QStandardItem(mode)
                    model.appendRow(item)

            self.modeSelectBox.setModel(model)


            #self.modeSelectBox.addItems(self.differenceModes)

        elif self.operation == 2: # Count Occurences
            set_two_sheet_mode(True)
            set_column_list()
            set_two_file_mode(False)
            set_mode_select(False)

        elif self.operation == 3: # Change Date Format
            set_two_sheet_mode(True)
            set_column_list()
            set_two_file_mode(False)
            set_mode_select(True)
            for dateFormat in self.dateFormats:
                self.modeSelectBox.addItem(datetime.now().strftime(format=dateFormat))

        else: # No operation selected
            QMessageBox.information(self, 'Huh?', f'How did you get here?\nOperation Number: {self.operation}')

    def mode_change(self):
        self.mode = self.modeSelectBox.currentIndex()

    def execute_with_save(self):
        self.execute(True)

    def execute(self, excel: bool=False):
        if not self.fileOneFinishedLoading or (not self.fileTwoFinishedLoading and self.twoFileMode):
            QMessageBox.information(self, 'Files not loaded', 'Please wait for files to finish loading')
            return

        if not self.sheetOne or (not self.sheetTwo and self.twoSheetMode):
            QMessageBox.information(self, 'Sheets not selected',
                                    'Please select the two sheets you want to use for the operation')
            return

        if self.operation == 0: # Find Duplicates
            self.statusbar.showMessage(f'Finding Duplicates in Sheet {self.sheetOne}')
            QCoreApplication.processEvents()
            duplicates = self.find_duplicates(excel)
            self.statusbar.showMessage(f'Found {duplicates} duplicates in Sheet {self.sheetOne}')
            return

        elif self.operation == 1: # Find Differences
            self.statusbar.showMessage(f'Finding Differences between Sheet {self.sheetOne} and Sheet {self.sheetTwo}')
            QCoreApplication.processEvents()
            differences = self.find_differences(excel)
            self.statusbar.showMessage(f'Found {differences} Differences between Sheet {self.sheetOne}'
                                       f'and Sheet {self.sheetTwo}')
            return

        elif self.operation == 2: # Count Occurences
            self.statusbar.showMessage(f'Counting Occurences in Sheet {self.sheetOne}')
            QCoreApplication.processEvents()
            self.count_occurences(excel)
            self.statusbar.showMessage(f'Finished counting Occurences in Sheet {self.sheetOne}')
            return

        elif self.operation == 3: # Change Date Format
            if self.mode in list(range(self.modeSelectBox.count())):
                to_format = datetime.now().strftime(format=self.dateFormats[self.mode])
            else:
                QMessageBox.information(self, 'No Mode selected', 'Please select a Mode')
                return

            self.statusbar.showMessage(f'Changing Date Format to {to_format} for Sheet {self.sheetOne}')
            QCoreApplication.processEvents()
            self.change_date_format(excel)
            self.statusbar.showMessage(f'Finished changing Date Format to {to_format} for Sheet {self.sheetOne}')
            return

        else:
            QMessageBox.information(self, 'No Operation selected', 'Please select an Operation')
            return

    def find_duplicates(self, excel: bool=False):
        # Create an empty DataFrame to store duplicates
        duplicates_df = pd.DataFrame()
        total_duplicates = -1

        try:
            sheet_data = self.fileOneData[self.sheetOne]
            # Check duplicates for each column
            for column in sheet_data.columns:
                # Get the duplicated values in each column, excluding the first occurrence
                duplicated_values = sheet_data[column][sheet_data[column].duplicated()].unique()
                if duplicated_values.size > 0:
                    duplicates_df[column] = pd.Series(duplicated_values)

            duplicates_df = duplicates_df.reset_index(drop=True).fillna('')

            # call function to save result / display in table
            self.save_output(duplicates_df, self.operation, -1, excel)

            # Get the total number of duplicate rows in the whole DataFrame
            total_duplicates = sheet_data.duplicated().sum()
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Error finding duplicates\n{e}')
        return total_duplicates

    def find_differences(self, excel: bool=False):
        diff_sheet1 = pd.DataFrame()
        diff_sheet2 = pd.DataFrame()
        tmp_df = pd.DataFrame()

        if self.twoFileMode: # Sheets from 2 different files
            sheet1 = self.fileOneData[self.sheetOne]
            sheet2 = self.fileTwoData[self.sheetTwo]
        else: # Sheets from the same file
            sheet1 = self.fileOneData[self.sheetOne]
            sheet2 = self.fileOneData[self.sheetTwo]

        try:
            # Column Modes
            if self.mode in (1, 2):
                # Create copies of the sheets to hold the differences
                diff_sheet1 = sheet1.copy()
                diff_sheet2 = sheet2.copy()

                # Identify columns unique to each sheet
                unique_to_sheet1 = set(sheet1.columns) - set(sheet2.columns)
                unique_to_sheet2 = set(sheet2.columns) - set(sheet1.columns)

                # Loop through common columns and compare values
                for column in sheet1.columns.intersection(sheet2.columns):
                    diff_sheet1[column] = sheet1[column].where(sheet1[column] != sheet2[column], np.nan)
                    diff_sheet2[column] = sheet2[column].where(sheet1[column] != sheet2[column], np.nan)

                # Optionally, check if a value exists anywhere in the column (ignoring row order)
                if self.mode == 2: # Unordered Mode
                    for column in sheet1.columns.intersection(sheet2.columns):
                        unique_values_sheet2 = set(sheet2[column].dropna().values)  # Unique non-null values
                        unique_values_sheet1 = set(sheet1[column].dropna().values)

                        diff_sheet1[column] = sheet1[column].where(
                            sheet1[column].apply(lambda x: x not in unique_values_sheet2), np.nan)
                        diff_sheet2[column] = sheet2[column].where(
                            sheet2[column].apply(lambda x: x not in unique_values_sheet1), np.nan)

                # Remove empty cells from sheet 1
                for column in diff_sheet1.columns:
                    # Remove empty & NaN values
                    non_empty_values = diff_sheet1[column].replace('', np.nan).dropna().reset_index(drop=True)
                    tmp_df[column] = pd.Series(non_empty_values)  # Add the cleaned column to tmp DataFrame

                #tmp_df.drop(columns=unique_to_sheet2, inplace=True, errors='ignore')
                diff_sheet1 = tmp_df.fillna('') # Set sheet 1 to tmp DataFrame
                tmp_df = pd.DataFrame() # Reset tmp_df

                # Remove empty cells from sheet 2
                for column in diff_sheet2.columns:
                    # Remove empty & NaN values
                    non_empty_values = diff_sheet2[column].replace('', np.nan).dropna().reset_index(drop=True)
                    tmp_df[column] = pd.Series(non_empty_values)  # Add the cleaned column to tmp DataFrame

                #tmp_df.drop(columns=unique_to_sheet1, inplace=True, errors='ignore')
                diff_sheet2 = tmp_df.fillna('') # Set sheet 2 to tmp DataFrame

                # Remove columns that are only in one sheet
                diff_sheet2.drop(columns=list(unique_to_sheet1), inplace=True, errors='ignore')
                diff_sheet1.drop(columns=list(unique_to_sheet2), inplace=True, errors='ignore')

                # Remove empty columns
                diff_sheet1.dropna(how='all', axis='columns', inplace=True)
                diff_sheet2.dropna(how='all', axis='columns', inplace=True)

            # Row Modes
            elif self.mode in (4, 5, 6):
                diff_sheet1 = pd.DataFrame(columns=sheet1.columns)
                diff_sheet2 = pd.DataFrame(columns=sheet2.columns)

                # Strict Mode
                if self.mode == 4:
                    # For strict mode, compare the rows exactly
                    # Determine the min and max number of rows
                    min_rows = min(len(sheet1), len(sheet2))
                    max_rows = max(len(sheet1), len(sheet2))

                    # Compare rows that exist in both sheets
                    for idx in range(min_rows):
                        row1 = sheet1.iloc[idx]
                        row2 = sheet2.iloc[idx]

                        # Normalize empty strings to NaN for comparison
                        row1 = row1.replace('', np.nan)
                        row2 = row2.replace('', np.nan)

                        # If rows are different, add them to the difference DataFrames
                        if not row1.equals(row2):
                            diff_sheet1 = pd.concat([diff_sheet1, pd.DataFrame([row1],
                                                    columns=sheet1.columns)], ignore_index=True)
                            diff_sheet2 = pd.concat([diff_sheet2, pd.DataFrame([row2],
                                                    columns=sheet2.columns)], ignore_index=True)

                    # If sheet1 has extra rows, add them to diff_sheet1
                    if len(sheet1) > min_rows:
                        diff_sheet1 = pd.concat([diff_sheet1, sheet1.iloc[min_rows:max_rows]], ignore_index=True)

                    # If sheet2 has extra rows, add them to diff_sheet2
                    if len(sheet2) > min_rows:
                        diff_sheet2 = pd.concat([diff_sheet2, sheet2.iloc[min_rows:max_rows]], ignore_index=True)

                # Strict-Column Mode
                if self.mode == 5:
                    if not sheet1.columns.equals(sheet2.columns):
                        diff_sheet1 = sheet1.copy()
                        diff_sheet2 = sheet2.copy()
                    else:
                        # Find rows that are in sheet1 but not in sheet2
                        diff_sheet1 = sheet1.loc[~sheet1.apply(tuple, axis=1).isin(sheet2.apply(tuple, axis=1))].copy()

                        # Find rows that are in sheet2 but not in sheet1
                        diff_sheet2 = sheet2.loc[~sheet2.apply(tuple, axis=1).isin(sheet1.apply(tuple, axis=1))].copy()

                # Loose Mode
                if self.mode == 6:
                    # Step 1: Compare columns (ignoring order)
                    if set(sheet1.columns) != set(sheet2.columns):
                        diff_sheet1 = sheet1.copy()
                        diff_sheet2 = sheet2.copy()
                    else:
                        # Step 2: Reorder columns to be in the same order for both DataFrames
                        sheet1 = sheet1[sheet2.columns]
                        sheet2 = sheet2[sheet1.columns]
    
                        # Step 3: Compare rows using the tuple logic
                        diff_sheet1 = sheet1.loc[~sheet1.apply(tuple, axis=1).isin(sheet2.apply(tuple, axis=1))].copy()
                        diff_sheet2 = sheet2.loc[~sheet2.apply(tuple, axis=1).isin(sheet1.apply(tuple, axis=1))].copy()

        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Error finding differences: {e}')
            return

        # Count differences
        if self.mode in (1, 2):
            differences = diff_sheet1.ne('').to_numpy().sum() + diff_sheet2.ne('').to_numpy().sum()
        else:
            differences = len(diff_sheet1) + len(diff_sheet2)

        # call function to save result / display in table
        self.save_output(diff_sheet1, self.operation, self.mode, excel, True, diff_sheet2,
                         self.sheetOne + 'Diff', self.sheetTwo + 'Diff')

        return differences

    def count_occurences(self, excel: bool=False):
        df = self.fileOneData[self.sheetOne]
        columns = [self.sheetTwo]
        result = []
        if columns[0] == 'All':
            columns = self.fileOneData[self.sheetOne].columns

        try:
            # Loop through each column name in the provided list
            for column in columns:
                # Count occurrences of each value in the column
                value_counts = df[column].value_counts()

                # For each unique value in the column, add it along with its count to the result
                for value, count in value_counts.items():
                    result.append({
                        'Column': column,
                        'Value': value,
                        'Count': count
                    })
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Error while counting occurences: {e}')

        # Create a DataFrame from the result list
        counts_df = pd.DataFrame(result)

        # call function to save result / display in table
        self.save_output(counts_df, self.operation, -1, excel)

    def change_date_format(self, excel: bool=False):
        df = self.fileOneData[self.sheetOne]
        columns = [self.sheetTwo]
        date_format = self.dateFormats[self.mode]
        if columns[0] == 'All':
            columns = self.fileOneData[self.sheetOne].columns

        def convert_to_excel(_df: pd.DataFrame, _column: str):
            """Converts a column of plain text dates (dd.mm.yyyy or yyyy.mm.dd) to Excel date format (serial number)"""
            # Try converting DD.MM.YYYY first
            _df[_column] = pd.to_datetime(_df[_column], errors='coerce', dayfirst=True)

            # If there are any NaT (Not a Time) values, try converting to YYYY.MM.DD format
            if _df[_column].isna().sum() > 0:
                _df[_column] = pd.to_datetime(_df[_column], errors='coerce', dayfirst=False)

            # Convert to Excel date format (serial number)
            _df[_column] = (_df[_column] - pd.Timestamp('1899-12-30')).dt.days

            return _df

        def convert_to_plain_text(_df: pd.DataFrame, _column: str, _date_format: str):
            """Converts a column of Excel serial date format to plain text format (e.g., yyyy-mm-dd)."""
            # Check if the column contains numeric (Excel date format) values
            if pd.api.types.is_numeric_dtype(_df[_column]):
                # Convert Excel serial date (numeric) to datetime
                _df[_column] = pd.to_datetime(_df[_column], origin='1899-12-30', unit='D')

            # Convert to plain text with the specified format
            _df[_column] = _df[_column].dt.strftime(_date_format)

            return _df

        try:
            for column in columns:
                if date_format == 'Excel':
                    # Convert plain text date to Excel date format
                    df = convert_to_excel(df, column)
                elif not pd.api.types.is_object_dtype(df[column]):
                    # Convert Excel date to plain text format
                    df = convert_to_plain_text(df, column, date_format)
                else:
                    # Convert plain text date to Excel date format
                    df = convert_to_excel(df, column)
                    # Convert Excel date to plain text format
                    df = convert_to_plain_text(df, column, date_format)
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Error while converting Date Format: {e}')

        # call function to save result / display in table
        self.save_output(df, self.operation, self.mode, excel)

    def save_output(self, df: pd.DataFrame, operation: int, mode: int=-1, save_to_excel: bool=False,
                    dual_file: bool = False, df2: pd.DataFrame = None, sheet1_name: str = None, sheet2_name: str = None):
        if not save_to_excel:
            if dual_file:
                # Add sheet identifier column
                # df['SheetName'] = self.sheetOne
                # df2['SheetName'] = self.sheetTwo
                df.loc[:, 'SheetName'] = self.sheetOne
                df2.loc[:, 'SheetName'] = self.sheetTwo

                # Merge sheets into one DataFrame
                df = pd.concat([df, df2], ignore_index=True).fillna('')
                columns_order = [col for col in df.columns if col != 'SheetName'] + ['SheetName']
                df = df[columns_order]  # Reorder columns

            self.table_model.beginResetModel()
            self.table_model._data = df
            self.table_model.endResetModel()
            if not self.outputTable.isVisible():
                self.toggle_table()
            return
        else:
            # Get output file
            operation_name = str.replace(self.operationSelectBox.itemText(operation), ' ', '')
            if mode != -1 and operation != 3:  # special case for operation with different modes
                operation_name += '-' + str.replace(self.modeSelectBox.itemText(mode), ' ', '')
            input_dir = os.path.dirname(self.fileOne)
            input_name = os.path.splitext(os.path.basename(self.fileOne))[0]
            output_file = os.path.normpath(os.path.join(input_dir, f'{input_name}_{operation_name}.xlsx'))

            self.save_thread = SaveFileThread(df, output_file, operation_name, dual_file, df2, sheet1_name, sheet2_name)
            self.save_thread.error.connect(self.save_error)
            self.save_thread.finished.connect(self.save_finished)
            self.save_thread.start()

    def save_error(self, e: str):
        QMessageBox.critical(self, 'Error', f'Error saving to Excel file: {e}')

    def save_finished(self, output_file: str):
        result = QMessageBox.question(self, 'Saved', f'Results saved to Excel file\n{output_file}\nOpen file?')

        if result == QMessageBox.StandardButton.Yes:
            try:
                if os.name == 'nt':  # Windows
                    os.startfile(output_file)
                elif os.name == 'posix':  # Mac/Linux
                    subprocess.run(['open', output_file])
            except Exception as e:
                QMessageBox.critical(self, 'Error', f'Could not open the file: {e}')

    def toggle_table(self):
        if self.outputTable.isVisible():
            self.outputTable.setVisible(False)
            self.resize(self.width() // 2, self.height())
            self.showTableButton.setText('>')
        else:
            self.outputTable.setVisible(True)
            self.resize(self.width() * 2, self.height())
            self.showTableButton.setText('<')

if __name__ == '__main__':
    if getattr(sys, 'frozen', False):
        # Running as a bundled PyInstaller executable
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    else:
        # Running as a normal Python script
        base_path = os.path.dirname(os.path.abspath(__file__))

    icon_path = os.path.join(base_path, 'Excel_Analyzer_Icon.ico')
    icon_png_path = os.path.join(base_path, 'Excel_Analyzer_Icon.png')

    app = QApplication(sys.argv)
    app.setWindowIcon(QtGui.QIcon(QtGui.QPixmap(icon_path)))
    window = MainWindow()
    window.show()  # Show the window
    sys.exit(app.exec())
