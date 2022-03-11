import inspect
import logging
import traceback
from collections import defaultdict

import pandas as pd
import yaml
import validators_pd as _validators
from validators_pd import *  # noqa

logger = logging.getLogger(__name__)

SHEET_ERRORS = defaultdict(list)

validating_objects = {
    v_cls[0]: eval(v_cls[0])()
    for v_cls in inspect.getmembers(_validators, inspect.isclass)
    if v_cls[1].__module__ == _validators.__name__
}


def is_valid_cell(values, validations, sheet, header, column_letter):
    index = 1
    for value in values:
        index += 1
        for valdn_type in validations:
            for typ, data in valdn_type.items():
                metadata = {
                    "type": "cell",
                    "header": header,
                    "rowNumber": index,
                    "cellLocation": column_letter+str(index)
                }
                validator = validating_objects.get(typ)
                try:
                    validator.validate(value, data)
                except Exception as ex:
                    metadata["errorMessage"] = ex.args[0]
                    SHEET_ERRORS[sheet].append(metadata)


def set_config(yaml_config):
    """
    Takes the config yaml file and converts it into a dictionary.
    """
    with open(yaml_config, "r") as yml:
        config = yaml.safe_load(yml)
        return config


def validate(config, worksheet, sheet):
    columns_to_validate = config.get("validations").get("columns") or []
    must_have_columns = config.get("must_have_columns")
    read_as_string = config.get("read_as_string") or []
    headers = worksheet.columns.to_list()
    if not set(must_have_columns).issubset(set(headers)):
        SHEET_ERRORS[sheet].append(
            {
                "type": "column",
                "errorMessage": "Sheet {sheet} must have column(s) {missing_columns}".format(
                    sheet=sheet,
                    missing_columns=", ".join(
                        list(set(must_have_columns) - set(headers))
                    ),
                ),
            }
        )
    for column in headers:
        column_letter = chr(worksheet.columns.get_loc(column) + 65)
        if column in read_as_string:
            worksheet[column] = worksheet[column].apply(lambda x: str(x))
        if column in config.get("exclude", []):
            continue
        if column in columns_to_validate:
            is_valid_cell(
                worksheet[column],
                validations=columns_to_validate[column],
                sheet=sheet,
                header=column,
                column_letter=column_letter
            )
        elif config.get("validations").get("default"):
            is_valid_cell(
                worksheet[column],
                validations=config.get("validations").get("default"),
                sheet=sheet,
                header=column,
                column_letter=column_letter
            )


def run_validations(xlsx_file, yaml_file, return_sheet_data=False):
    errors = {}
    sheet_data = {}
    try:
        xlsx = pd.ExcelFile(xlsx_file, engine="openpyxl")
        sheets = xlsx.sheet_names
        config = set_config(yaml_file)
        config_sheets = config.get("sheets")
        if not set(config_sheets).issubset(set(sheets)):
            errors["fileError"] = {
                "errorMessage": "The uploaded file must have sheet(s) {missing_sheets}".format(
                    missing_sheets=", ".join(list(set(config_sheets) - set(sheets))),
                ),
            }
        for sheet in sheets:
            if config.get(sheet):
                worksheet = pd.read_excel(xlsx, sheet_name=sheet)
                worksheet = worksheet.astype(object).where(
                    worksheet.notna(), None
                )  # turning NaN to None
                validate(config.get(sheet), worksheet, sheet)
                sheet_data[sheet] = worksheet.to_dict(orient="records")
        if SHEET_ERRORS:
            errors["sheetErrors"] = []
            for key, value in SHEET_ERRORS.items():
                err = {}
                err["sheetName"] = key
                err["errors"] = value
                errors["sheetErrors"].append(err)

        if errors:
            return (False, errors, {})  # is_valid, errors, sheet_data
        else:
            if return_sheet_data:
                return (True, {}, sheet_data)
            else:
                return (True, {}, {})

    except Exception as e:
        traceback.print_exc()
        logger.exception(e)


if __name__ == "__main__":
    import time
    t1 = time.time()
    xlsx_filepath = "/Users/I1597/Downloads/ucc_data_columns_final_2.xlsx"
    yaml_filepath = "/Users/I1597/Documents/repositories/excel_validator/thn_final_validator.yml"
    # xlsx_filepath = "/Users/I1597/Documents/repositories/excel_validator/one_lakh_record.xlsx"
    # yaml_filepath = "/Users/I1597/Documents/repositories/excel_validator/sales_record.yml"

    is_valid, errors, data = run_validations(xlsx_filepath, yaml_filepath)
    print(is_valid)
    print(errors)
    t2 = time.time()
    print("Total Time", (t2 - t1))
