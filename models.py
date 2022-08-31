from PyQt5.QtSql import QSqlDatabase, QSqlQuery, QSqlTableModel
from PyQt5.QtCore import Qt
import numpy as np

class LoginInfo():
    def __init__(self, login_site:str, id:str, password:str, no_of_tabs:int):
        self.login_site = login_site
        self.id = id
        self.password = password

class ReservationInfo():
    def __init__(self, wanted_date: int, wanted_time1: int, wanted_time2: int, wanted_time3: int):
        self.wanted_date = wanted_date
        self.wanted_time1 = wanted_time1
        self.wanted_time2 = wanted_time2
        self.wanted_time3 = wanted_time3
        
class MyTableModel(QSqlTableModel):
    # 210513
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parnet=None):
        return self._data.shape[1]

    def data(self, index, role):
        if index.isValid():
            if role == Qt.DisplayRole:
                if type(self._data.iloc[index.row(), index.column()]) == str:
                    value = f"{str(self._data.iloc[index.row(), index.column()]):>10s}"
                    return value
                elif type(self._data.iloc[index.row(), index.column()]) == np.float64:
                    # print(type(self._data.iloc[index.row(), index.column()]))
                    # print(self._data.iloc[index.row(), index.column()])
                    value = f"{str(f'{self._data.iloc[index.row(), index.column()]:>20,.2f}')}"
                    return value
                else:
                    # print(type(self._data.iloc[index.row(), index.column()]))
                    # print(self._data.iloc[index.row(), index.column()])
                    value = self._data.iloc[index.row(), index.column()]
                    return value
        return None