import os.path
import sys
import time
import yaml
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter
from validators import (
    RequiredValidator,
    TypeValidator,
    LengthValidator,
    RegexValidator,
    EmailValidator,
    ChoiceValidator,
    DateTimeValidator,
    ExcelDateValidator,
    NonNegativeValidator,
    ComparatorValidator,
)


def is_valid_cell(valdn_type, value, coordinate, errors):
    classmap = {
        "Required": RequiredValidator,
        "Type": TypeValidator,
        "Length": LengthValidator,
        "Regex": RegexValidator,
        "Email": EmailValidator,
        "Choice": ChoiceValidator,
        "Date": DateTimeValidator,
        "ExcelDate": ExcelDateValidator,
        "ComparatorValidator": ComparatorValidator,
        "NonNegativeValidator": NonNegativeValidator,
    }

    violations = []
    # TODO: try to avoid [0]
    name = list(valdn_type.keys())[0]
    data = list(valdn_type.values())[0]
    validator = classmap[name](data)  # creating the object with the parameters
    try:
        validator.validate(value)
    except Exception as ex:
        violations.append(ex)
    # TODO: try to avoid violations

    if violations:
        errors.append((coordinate, violations))


def set_config(yaml_config, yaml_validator_cls):
    """
    Takes the config yaml file and converts it into a dictionary.
    """
    with open(yaml_config, "r") as yml:
        config = yaml.safe_load(yml).get(yaml_validator_cls)
        config["default"] = config.get("validations").get("default")[0] or None
        return config


def validate(config, worksheet):
    errors = []
    columns_to_validate = config.get("validations").get("columns")

    # ********************** COLUMN_CASES ******************* #
    column_letter_to_header = {}
    max_active_column = 0
    for row in worksheet.iter_rows(max_row=1):
        for cell in row:
            if cell.value:
                column_letter_to_header[cell.column_letter] = cell.value
    # ********************** COLUMN_CASES ******************* #
    start_row = 1
    if config.get("iterate_by_header_name"):
        start_row = 2
    else:
        for key, _ in column_letter_to_header.items():
            column_letter_to_header[key] = key
    for row in worksheet.iter_rows(
        min_row=start_row, max_col=len(column_letter_to_header)
    ):
        for cell in row:
            try:
                value = cell.value
            except ValueError:
                errors.append((cell.coordinate, ValueError))
            if column_letter_to_header[cell.column_letter] in config.get(
                "excludes", []
            ):
                continue
            if column_letter_to_header[cell.column_letter] in columns_to_validate:
                for valdn_type in columns_to_validate[
                    column_letter_to_header[cell.column_letter]
                ]:
                    is_valid_cell(valdn_type, value, cell.coordinate, errors)

            elif config.get("default"):
                is_valid_cell(config["default"], value, cell.coordinate, errors)

    print(errors)  # TODO: do something about error logging


def validate_excel(xlsx_filepath, yaml_filepath):
    # import ipdb; ipdb.set_trace()
    # file_size = os.path.getsize(xlsx_filepath) > 10485760

    try:
        workbook = load_workbook(xlsx_filepath)
        sheets = workbook.sheetnames
        for sheet in sheets:
            config = set_config(yaml_filepath, sheet)
            worksheet = workbook[sheet]
            validate(config, worksheet)
    except Exception as e:
        sys.exit("Error occured: " + str(e))


if __name__ == "__main__":
    xlsx_filepath = "/Users/I1597/Documents/repositories/excel_validator/sample.xlsx"
    yaml_filepath = "/Users/I1597/Documents/repositories/excel_validator/ucc_thn.yml"
    validate_excel(xlsx_filepath, yaml_filepath)
