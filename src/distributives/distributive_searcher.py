from .distributive import Distributive
from names.names import *

from datetime import datetime, timedelta
from typing import Optional
from pandas import ExcelFile
from loguru import logger

import pandas as pd


class DistributiveSearcher:
    def __init__(self, excel_file: ExcelFile) -> None:
        self.excel_file = excel_file

    def get_main_distributive(self, distr_name: str) -> Optional[Distributive]:
        distr_name = distr_name.upper()
        if distr_name not in self.excel_file.sheet_names:
            logger.warning(f"Дистрибутив с именем \"{distr_name}\" не найден в списке Листов!")
            return self.get_main_distributive_from_mz(distr_name)

        df = pd.read_excel(self.excel_file, sheet_name=distr_name)

        oldest_date = df[MainDistrColumns.DATE_COLUMN_NAME].min()

        current_time = datetime.now()
        
        if (current_time - oldest_date) < timedelta(hours=96):
            return self.get_main_distributive_from_mz(distr_name)

        suitable_rows = df[df[MainDistrColumns.DATE_COLUMN_NAME] == oldest_date]
        row = suitable_rows.iloc[0]

        return Distributive(row[MainDistrColumns.NUMBER_COLUMN_NAME], row[MainDistrColumns.DATE_COLUMN_NAME], distr_name, distr_name)

    def get_main_distributive_from_mz(self, distr_name: str) -> Optional[Distributive]:
        """
        Raises:
            - ValueError if sheet МЗ not found
        """
        whole_dataframe = pd.read_excel(self.excel_file, sheet_name=Sheets.MZ_NAME)
        for mz_column in MzColumns.MZ_COLUMNS:
            distr_rows = whole_dataframe[whole_dataframe[mz_column] == distr_name]
            
            oldest_date = distr_rows[MzColumns.DATE_COLUMN_NAME].min()
            current_time = datetime.now()
            if (current_time - oldest_date) < timedelta(hours=96):
                continue
            suitable_rows = distr_rows[distr_rows[MzColumns.DATE_COLUMN_NAME] == oldest_date]
            row = suitable_rows.iloc[0]
            return Distributive(round(row[MzColumns.NUMBER_COLUMN_NAME]), row[MzColumns.DATE_COLUMN_NAME], Sheets.MZ_NAME, distr_name)

    def get_dop_distributive(self, distr_name: str) -> Optional[Distributive]:
        whole_dataframe = pd.read_excel(self.excel_file, sheet_name=Sheets.DOP_NAME)
        if distr_name not in whole_dataframe.columns:
            logger.warning(f"Дистрибутив с именем \"{distr_name}\" не найден в дополнительных дистрибутивах!")
            return None
        distr_column_index = whole_dataframe.columns.values.tolist().index(distr_name)
        date_column_index = distr_column_index + 1
        date_column_name: str = whole_dataframe.columns.values.tolist()[date_column_index]
        distr_dataframe = whole_dataframe[[distr_name, date_column_name]]

        oldest_date = distr_dataframe[date_column_name].min()
        current_time = datetime.now()
        if (current_time - oldest_date) < timedelta(hours=96):
            return None
        suitable_rows = distr_dataframe[distr_dataframe[date_column_name] == oldest_date]
        row = suitable_rows.iloc[0]
        return Distributive(round(row[distr_name]), row[date_column_name], Sheets.DOP_NAME, distr_name)
