from PyQt6.QtCore import QThread, pyqtSignal
import pandas as pd

class LoadFileThread(QThread):
    """Loads the file in a separate thread to avoid freezing the UI"""
    finished = pyqtSignal(dict, int)
    error = pyqtSignal(str, int)

    def __init__(self, input_file: str, filenum: int):
        super().__init__()
        self.input_file = input_file
        self.filenum = filenum

    def run(self):
        try:
            xls = pd.ExcelFile(self.input_file)
            loaded_data = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}
            self.finished.emit(loaded_data, self.filenum)
        except Exception as e:
            self.error.emit(str(e), self.filenum)
