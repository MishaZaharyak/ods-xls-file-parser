from pathlib import Path
import re
import pyexcel as pc
from typing import List, Union, Optional, Tuple
from exceptions import IsNoneException

SheetsTotalHours = Tuple[List[Tuple[str, float]], float]


class BaseColumn:
    """ base sheet column class """

    def __init__(self, data: List[str]):
        self.data = data

    @property
    def value(self):
        """ :returns column value """
        return self.data


class TimeColumn(BaseColumn):
    """ provides functionality for time column """
    DIGITS_REGEX = r"\d+"
    HOUR_REGEX = r"(\d+\s?h)"
    MINUTES_REGEX = r"(\d+\s?min)"

    def get_total_hours(self, time_list: List[str]) -> float:
        """ sum time column values, values should be in following
            formats: "1 h 30 min", "1h 30min", "1h30min"

         :param time_list: time column values """

        time_string = ' '.join(time_list)
        hours = self.parse_and_sum_hours(time_string)
        minutes = self.parse_and_sum_minutes(time_string)
        return hours + minutes / 60

    def parse_and_sum_hours(self, time_string: str) -> int:
        """ sum hours from time string
         :param time_string: time string """
        hours_string_list = re.findall(self.HOUR_REGEX, time_string)

        if hours_string_list:
            hours_string = " ".join(hours_string_list)
            return sum(map(int, re.findall(self.DIGITS_REGEX, hours_string)))

        return 0

    def parse_and_sum_minutes(self, time_string: str) -> int:
        """ sum minutes from time string
         :param time_string: time string """
        minutes_string_list = re.findall(self.MINUTES_REGEX, time_string)

        if minutes_string_list:
            minutes_string = " ".join(minutes_string_list)
            return sum(map(int, re.findall(self.DIGITS_REGEX, minutes_string)))

        return 0

    @property
    def value(self):
        return self.get_total_hours(self.data)


class BaseSheet:
    """ provides basic functionality to interact with pyexcel.Sheet class """

    def __init__(self, sheet: pc.Sheet, name_columns_by_row: Optional[int]):
        """ :param sheet: pyexcel.Sheet instance
            :param name_columns_by_row: sheet row index, optional """
        self.sheet = sheet
        self._set_columns_names(name_columns_by_row)

    def __repr__(self):
        """ shows table representation of a sheet """
        return self.sheet.__repr__()

    def get_column_data(self, column_name: str) -> List[str]:
        # fixme: when sheet is used second time, first line where colnames
        #  are defined magically (or not) disappear and we get ValueError
        return self.sheet.named_column_at(column_name)

    def _set_columns_names(self, row_index: int) -> None:
        """ use row values to name all columns
            :param row_index: row index, if not provided then first row will be used """

        if not row_index:
            row_index = 0

        try:
            self.sheet.name_columns_by_row(row_index)
        except TypeError:
            raise IsNoneException(f"\"name_columns_by_row\" expect type 'int' got '{type(row_index).__name__}'")


class TimeSheet(BaseSheet):
    """ provides functionality to parse and summarize total time """

    def get_total_hours(self, column_name: str) -> float:
        """
        :param column_name: name of a column that provides time data
        :returns total hours
        """
        data = self.get_column_data(column_name)
        return TimeColumn(data).value


class MainSheet(TimeSheet):
    """ main sheet class that inherits all others sheets """
    pass


class OdsBookBase:
    """ provides basic functionality to interact with .ods file """

    def __init__(self, file_path: str):
        """ :param file_path: path to .ods file """
        self.file_path = file_path
        self._sheet: Union[MainSheet, None] = None
        self._load()

    def _load(self) -> None:
        """ load .ods document data """
        self.book: pc.Book = pc.get_book(file_name=self.file_path)
        self.sheet_names = self.book.sheet_names()

    def load_sheet(self, sheet_name, name_columns_by_row=None):
        """ to work with a specific sheet we need to load its data first
            :param sheet_name: name of a sheet that we want to work with
            :param name_columns_by_row: optional """

        sheet = self.book.sheet_by_name(sheet_name)
        self._sheet = MainSheet(sheet, name_columns_by_row)

    @property
    def sheet(self):
        """ :returns loaded sheet """
        assert self._sheet is not None, "You should call \"load_sheet()\" method first"
        return self._sheet


class OdsBookTimeSheets(OdsBookBase):
    """ additional functionality for time sheets """

    def get_total_hours(self, column_name: str, name_columns_by_row: Optional[int] = None) -> SheetsTotalHours:
        """
        :param column_name: name of a column that provides time data
        :param name_columns_by_row: sheet row index, optional
        :returns total hours from all sheets
        """
        sheets_total_hours = []
        total_hours = 0

        for name in self.sheet_names:
            sheet_data = self.book.sheet_by_name(name)
            sheet = MainSheet(sheet_data, name_columns_by_row)
            sheet_total_hours = sheet.get_total_hours(column_name)
            total_hours += sheet_total_hours
            sheets_total_hours.append((name, sheet_total_hours))

        return sheets_total_hours, total_hours

    def print_total_hours(self, column_name: str, name_columns_by_row: Optional[int] = None) -> None:
        """ same as self.get_total_hours(), but the result is printed """
        sheets_total_hours, total_hours = self.get_total_hours(column_name, name_columns_by_row)

        for sheet_name, total in sheets_total_hours:
            print(f"{sheet_name} | {total}")
            print("-" * 40)

        print(f"Total: {total_hours}")


class OdsBook(OdsBookTimeSheets):
    """ main book class that inherits all others books """
    pass


SHEET_NAME = "January"
FILE_NAME = "mykhailo_zakhariak_work_hours.ods"
DIR_NAME = f"{str(Path.home())}/Documents"


def get_file_path():
    """ returns the hours tracking file """
    file = Path(f"{DIR_NAME}/{FILE_NAME}")

    if file.exists() and file.is_file():
        return str(file.absolute())

    raise FileExistsError(f"file {FILE_NAME} doesn't exists in {DIR_NAME}")


if __name__ == '__main__':
    file_path = get_file_path()
    ods = OdsBook(file_path)
    ods.load_sheet(SHEET_NAME)
    total_hours = ods.sheet.get_total_hours("time")
    # print(total_hours)
    # sheets_total_hours, total_hours = ods.get_total_hours("time")
    # print(sheets_total_hours)
    # ods.print_total_hours("time")