
from datetime import datetime
import pandas as pd

from pandas import DataFrame, ExcelFile, ExcelWriter

from distributives.distributive import Distributive
from names.names import *


class Saver:
    def __init__(self, source_excel_file: ExcelFile) -> None:
        self.source_excel_file = source_excel_file
        self.sheet_names = self.source_excel_file.sheet_names

    def save_to_excel_file(self, output_file: str, main_distr: Distributive, dop_distrs: list[Distributive]):
        sheet_names = self.source_excel_file.sheet_names
        dfs = []
        for sheet in self.source_excel_file.sheet_names:
            dfs.append(pd.read_excel(self.source_excel_file, sheet_name=sheet))
        self.source_excel_file.close()
        
        with ExcelWriter(output_file) as writter:
            for i, sheet in enumerate(sheet_names):
                if sheet in [Sheets.MZ_NAME, Sheets.DOP_NAME]:
                    continue
                elif sheet == main_distr.sheet_name:
                    self.__write_main_distr(dfs[i], writter, main_distr, dop_distrs)
                else:
                    df = dfs[i]
                    df.to_excel(writter, sheet_name=sheet, index=False)
            # TODO: process Sheets.MZ_NAME, Sheets.DOP_NAME
            mz_dataframe = dfs[sheet_names.index(Sheets.MZ_NAME)]
            self.__write_mz_sheet(mz_dataframe, writter, main_distr, dop_distrs)
            dop_dataframe = dfs[sheet_names.index(Sheets.DOP_NAME)]
            self.__write_dops_sheet(dop_dataframe, writter, dop_distrs)

    def __write_main_distr(self, df: DataFrame, writter: ExcelWriter, main_distr: Distributive, dop_distrs: list[Distributive]):
        row_index = df[df[MainDistrColumns.NUMBER_COLUMN_NAME] == main_distr.number].index[0]

        # Remove "Value 'asdasd' has dtype incompatible with float64" warnings
        df[MainDistrColumns.COMPLECT_COLUMN_NAME] = df[MainDistrColumns.COMPLECT_COLUMN_NAME].fillna('').astype(object)
        df[MainDistrColumns.DOPS_COLUMN_NAME] = df[MainDistrColumns.DOPS_COLUMN_NAME].fillna('').astype(object)
        df[MainDistrColumns.USER_COLUMN_NAME] = df[MainDistrColumns.USER_COLUMN_NAME].fillna('').astype(object)

        df.loc[row_index, MainDistrColumns.COMPLECT_COLUMN_NAME] = "; ".join([str(main_distr.number), *[str(x.number) for x in dop_distrs]])
        df.loc[row_index, MainDistrColumns.DOPS_COLUMN_NAME] = ", ".join([str(x.name) for x in dop_distrs])
        df.loc[row_index, MainDistrColumns.USER_COLUMN_NAME] = main_distr.user_name
        df.loc[row_index, MainDistrColumns.DATE_COLUMN_NAME] = datetime.now()
        df.to_excel(writter, sheet_name=main_distr.name, index=False)

    def __write_mz_sheet(self, df: DataFrame, writter: ExcelWriter, main_distr: Distributive, dop_distrs: list[Distributive]):
        if main_distr.sheet_name != Sheets.MZ_NAME:
            df.to_excel(writter, sheet_name=Sheets.MZ_NAME, index=False)
            return
        row_index = df[df[MzColumns.NUMBER_COLUMN_NAME] == main_distr.number].index[0]

        # Remove "Value 'asdasd' has dtype incompatible with float64" warnings
        df[MzColumns.COMPLECT_COLUMN_NAME] = df[MzColumns.COMPLECT_COLUMN_NAME].fillna('').astype(object)
        df[MzColumns.DOPS_COLUMN_NAME] = df[MzColumns.DOPS_COLUMN_NAME].fillna('').astype(object)
        df[MzColumns.USER_COLUMN_NAME] = df[MzColumns.USER_COLUMN_NAME].fillna('').astype(object)

        df.loc[row_index, MzColumns.COMPLECT_COLUMN_NAME] = "; ".join([str(main_distr.number), *[str(x.number) for x in dop_distrs]])
        df.loc[row_index, MzColumns.DOPS_COLUMN_NAME] = ", ".join([str(x.name) for x in dop_distrs])
        df.loc[row_index, MzColumns.USER_COLUMN_NAME] = main_distr.user_name
        df.loc[row_index, MzColumns.DATE_COLUMN_NAME] = datetime.now()
        df.to_excel(writter, sheet_name=Sheets.MZ_NAME, index=False)

    def __write_dops_sheet(self, df: DataFrame, writter: ExcelWriter, dop_distrs: list[Distributive]):
        for dop_distr in dop_distrs:
            distr_column_index = df.columns.values.tolist().index(dop_distr.name)
            date_column_index = distr_column_index + 1
            date_column_name: str = df.columns.values.tolist()[date_column_index]
            distr_dataframe = df.loc[:, [dop_distr.name, date_column_name]]

            row_index = distr_dataframe[distr_dataframe[dop_distr.name] == dop_distr.number].index[0]
            df.loc[row_index, date_column_name] = datetime.now()
        df.to_excel(writter, sheet_name=Sheets.DOP_NAME, index=False)
