from email import header
import sys
import yaml
import copy
from openpyxl.reader.excel import load_workbook
from validators import *

ERRORS= []

def is_valid_cell(valdn_type, cell, sheet, column_header):
    for typ, data in valdn_type.items():
        metadata = {
            "sheet": sheet,
            "header": column_header,
            "cell": cell.coordinate,
        }
        validating_class = getattr(sys.modules[__name__], typ)
        validator = validating_class(data)  # creating the object with the parameters
        try:
            validator.validate(cell.value)
        except Exception as ex:
            metadata["error"] = ex.args[0]
            ERRORS.append(metadata)


def set_config(yaml_config):
    """
    Takes the config yaml file and converts it into a dictionary.
    """
    with open(yaml_config, "r") as yml:
        config = yaml.safe_load(yml)
        # config["default"] = config.get("validations").get("default")[0] or None
        return config


def validate(config, worksheet):
    columns_to_validate = config.get("validations").get("columns")
    must_have_columns = config.get("must_have_columns")
    # ********************** COLUMN_CASES ******************* #
    column_letter_to_header = {}
    for row in worksheet.iter_rows(max_row=1):
        for cell in row:
            if cell.value:
                column_letter_to_header[cell.column_letter] = cell.value
    # ********************** COLUMN_CASES ******************* #
    if not set(must_have_columns).issubset(set(column_letter_to_header.values())):
        raise Exception(
            {
                "sheet": worksheet.title,
                "error": "Sheet {sheet} must have column(s) {missing_columns}".format(
                    sheet=worksheet.title,
                    missing_columns=", ".join(
                        list(
                            set(must_have_columns)
                            - set(column_letter_to_header.values())
                        )
                    ),
                ),
            }
        )
    start_row = 1
    if config.get("iterate_by_header_name"):
        start_row = 2
    else:
        for key, _ in column_letter_to_header.items():
            column_letter_to_header[key] = key
    for row in worksheet.iter_rows(min_row=start_row, max_col=len(column_letter_to_header.keys())):
        for cell in row:
            column_header = column_letter_to_header[cell.column_letter]
            if column_header in config.get("exclude", []):
                continue
            if column_header in columns_to_validate:
                for valdn_type in columns_to_validate[column_header]:
                    is_valid_cell(
                        valdn_type=valdn_type,
                        cell=cell,
                        sheet=worksheet.title,
                        column_header=column_header
                    )

            elif config.get("validations").get("default"):
                is_valid_cell(
                    valdn_type=config.get("validations").get("default")[0],
                    cell=cell,
                    sheet=worksheet.title,
                    column_header=column_header
                )

    for error in ERRORS:
        print(error)  # TODO: do something about error logging


def validate_excel(xlsx_filepath, yaml_filepath):
    try:
        workbook = load_workbook(xlsx_filepath)
        sheets = workbook.sheetnames
        config = set_config(yaml_filepath)
        config_sheets = config.get("sheets")
        if not set(config_sheets).issubset(set(sheets)):
            raise Exception(
            {
                "error": "The uploaded file must have sheet(s) {missing_sheets}".format(
                    missing_sheets=", ".join(
                        list(
                            set(config_sheets)
                            - set(sheets)
                        )
                    ),
                ),
            }
        )
        for sheet in sheets:
            if config.get(sheet):
                worksheet = workbook[sheet]
                validate(config.get(sheet), worksheet)
    except Exception as e:
        sys.exit(e.args[0])


if __name__ == "__main__":
    import time

    t1 = time.time()
    xlsx_filepath = "/Users/I1597/Documents/repositories/excel_validator/sample_1.xlsx"
    yaml_filepath = "/Users/I1597/Documents/repositories/excel_validator/ucc_thn.yml"
    validate_excel(xlsx_filepath, yaml_filepath)
    t2 = time.time()
    print("Total Time", (t2 - t1))
