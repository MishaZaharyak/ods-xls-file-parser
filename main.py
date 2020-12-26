from pathlib import Path
import re
import pyexcel as pc

SHEET_NAME = "December"
SHEET_COLUMN_NAMES = ["date", "ticket", "time", "day total hours"]
DIGITS_REGEX = r"\d+"
FILE_NAME = "mykhailo_zakhariak_work_hours.ods"
DIR_NAME = f"{str(Path.home())}/Documents"


def get_file_path():
    """ returns the hours tracking file """
    file = Path(f"{DIR_NAME}/{FILE_NAME}")

    if file.exists() and file.is_file():
        return str(file.absolute())

    raise FileExistsError(f"file {FILE_NAME} doesn't exists in {DIR_NAME}")


def get_sheet_data_by_name(sheet_name, file_path):
    return pc.get_sheet(name=sheet_name, file_name=file_path)


def get_column_data(sheet, column_name):
    return sheet.named_column_at(column_name)


def parse_and_sum_hours(time_string):
    hour_regex = r"(\d+\s?h)"
    hours_string_list = re.findall(hour_regex, time_string)

    if hours_string_list:
        hours_string = " ".join(hours_string_list)
        return sum(map(int, re.findall(DIGITS_REGEX, hours_string)))

    return 0


def parse_and_sum_minutes(time_string):
    minutes_regex = r"(\d+\s?min)"
    minutes_string_list = re.findall(minutes_regex, time_string)

    if minutes_string_list:
        minutes_string = " ".join(minutes_string_list)
        return sum(map(int, re.findall(DIGITS_REGEX, minutes_string)))

    return 0


def get_total_hours(time_list):
    time_string = ' '.join(time_list)
    hours = parse_and_sum_hours(time_string)
    minutes = parse_and_sum_minutes(time_string)

    return hours + minutes / 60


def main():
    file_path = get_file_path()
    sheet = get_sheet_data_by_name(SHEET_NAME, file_path)
    sheet.colnames = SHEET_COLUMN_NAMES
    column_data = get_column_data(sheet, "time")
    total_hours = get_total_hours(column_data)
    print(total_hours)


if __name__ == '__main__':
    main()
