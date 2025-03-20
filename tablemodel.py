from PyQt6.QtCore import QAbstractTableModel, Qt
import pandas as pd

class OutputTableModel(QAbstractTableModel):
    def __init__(self, data: pd.DataFrame):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None

        if role == Qt.ItemDataRole.DisplayRole:
            return str(self._data.iloc[index.row(), index.column()])

        return None

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if role != Qt.ItemDataRole.DisplayRole:
            return None
        if orientation == Qt.Orientation.Horizontal:
            return str(self._data.columns[section]) # Column headers
        elif orientation == Qt.Orientation.Vertical:
            return str(self._data.index[section]+1) # Row headers
        return None

    def sort(self, column: int, order: Qt.SortOrder=Qt.SortOrder.AscendingOrder):
        """
        Sort the data in the model by a given column and order.
        This method is triggered when the user clicks on the column header.
        """
        # Sort the DataFrame based on the column
        self.beginResetModel()  # Tell the model that it will be reset
        self._data = self._data.sort_values(self._data.columns[column],
                                            ascending=(order != Qt.SortOrder.AscendingOrder))
        self.endResetModel()  # Notify the model that the reset is complete
