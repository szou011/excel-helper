import win32com.client as win32
import xlsxwriter
import pandas as pd

import logging

logging.basicConfig(
    format="%(levelname)s: %(asctime)s - %(message)s",
    datefmt="%m/%d/%Y %I:%M:%S %p",
    level=logging.INFO,
)


def refresh_workbook(xl_file):

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    try:
        xl_workbook = excel.Workbooks(xl_file)

    except Exception as e:
        try:
            xl_workbook = excel.Workbooks.Open(xl_file)
        except Exception as e:
            print(e)
            xl_workbook = None

    xl_workbook.RefreshAll()

    xl_workbook.Close(True)
    xl_workbook = None
    excel = None

    return None


def split_csv(full_csv, header_csv, details_csv, max_lines_to_check=30):
    with open(full_csv, "r", encoding="utf-8-sig") as full:
        for i in range(max_lines_to_check):
            if not full.readline().strip():
                empty_line = i
                break

    with open(full_csv, "r", encoding="utf-8-sig") as full, open(
        header_csv, "w", encoding="utf-8"
    ) as header, open(details_csv, "w", encoding="utf-8") as details:
        for i in range(empty_line):
            header.write(full.readline())

        # skip the blank line
        next(full)

        for line in full.readlines():
            details.write(line)


class ExcelWriter:
    def __init__(self, filename):
        self.workbook = xlsxwriter.Workbook(filename)

        # set up format
        self.string_format = self.workbook.add_format(
            {"font_name": "Arial", "font_size": 8}
        )
        self.number_format = self.workbook.add_format(
            {"num_format": "#,##0", "font_name": "Arial", "font_size": 8}
        )
        self.date_format = self.workbook.add_format(
            {"num_format": "dd/mm/yyyy", "font_name": "Arial", "font_size": 8}
        )
        self.header_format = self.workbook.add_format(
            {"bold": True, "font_name": "Arial", "font_size": 8}
        )

    def add_dataframe(
        self,
        dataframe,
        sheetname="Sheet1",
        offset_row=0,
        offset_col=0,
        col_format=None,
    ):

        # retrive the number of index level, column level (in case there are multiindexes) and the dataframe shape
        index_level = dataframe.index.nlevels
        column_level = dataframe.columns.nlevels
        row_num, col_num = dataframe.shape

        # remove NaN values from dataframe
        self._remove_nan(dataframe)

        # set up worksheet
        try:
            worksheet = self.workbook.add_worksheet(sheetname)
        except xlsxwriter.exceptions.DuplicateWorksheetName:
            print(f'Sheet name {sheetname} is duplicated.')
            #worksheet = self.workbook.get_worksheet_by_name(sheetname)

        logging.info(f'Worksheet created {worksheet.get_name()}')

        # write index
        for level in range(index_level):
            for index_num, value in enumerate(dataframe.index.get_level_values(level)):
                # index starting position will be offset by the 'column level'
                worksheet.write(
                    index_num + column_level + offset_row,
                    level + offset_col,
                    value,
                    self.header_format,
                )

        # write column/header
        for level in range(column_level):
            for index_num, value in enumerate(
                dataframe.columns.get_level_values(level)
            ):
                # index starting position will be offset by the 'column level'
                worksheet.write(
                    level + offset_row,
                    index_num + index_level + offset_col,
                    value,
                    self.header_format,
                )

        # write data
        for col in range(col_num):

            # check if there is NaN in column
            if dataframe.iloc[:, col].hasnans:

                dataframe.iloc[:, col] = dataframe.iloc[:, col].fillna(0)

            if col_format and (col in col_format):
                data_format = col_format[col]
            elif (
                (dataframe.iloc[:, col].dtypes == "Int64")
                or (dataframe.iloc[:, col].dtypes == "int64")
                or (dataframe.iloc[:, col].dtypes == "float")
            ):
                data_format = self.number_format
            elif dataframe.iloc[:, col].dtypes == "datetime64[ns]":
                data_format = self.date_format
            else:
                data_format = self.string_format

            for row in range(row_num):
                worksheet.write(
                    row + column_level + offset_row,
                    col + index_level + offset_col,
                    dataframe.iloc[row, col],
                    data_format,
                )

        return worksheet

    def _remove_nan(self, dataframe):

        for col in dataframe.columns:
            if dataframe[col].hasnans:
                if (
                    (dataframe[col].dtypes == "Int64")
                    or (dataframe[col].dtypes == "int64")
                    or (dataframe[col].dtypes == "float")
                ):
                    logging.info(
                        f"Column {col} has NaN number values, these values have been filled with zeros"
                    )
                    dataframe[col].fillna(0, inplace=True)

                elif dataframe[col].dtypes == "object":
                    logging.info(
                        f"Column {col} has NaN string values, these values have been filled with blank string"
                    )
                    dataframe[col].fillna("", inplace=True)

                elif dataframe[col].dtypes == "datetime64[ns]":
                    logging.info(
                        f"Column {col} has NaN datetime values, these values have been filled with 1899-12-31"
                    )
                    dataframe[col].fillna(
                        pd.Timestamp(year=1899, month=12, day=31),
                        inplace=True
                    )

                elif dataframe[col].dtypes == "category":
                    logging.info(
                        f"Column {col} has NaN category values, these values have been filled with None"
                    )
                    dataframe[col].add_categories("None").fillna("None", inplace=True)
                    

        return None

    def save(self):
        self.workbook.close()
        return None
