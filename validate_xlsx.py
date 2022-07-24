import argparse
import inspect
import logging
import sys
import traceback
from collections import Counter, defaultdict

import pandas as pd
import yaml

import validators as _validators
from validators import *  # noqa

logger = logging.getLogger(__name__)

SHEET_ERRORS = defaultdict(list)

validating_objects = {
    v_cls[0]: getattr(sys.modules[__name__], v_cls[0])()
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
                    "cellLocation": column_letter + str(index),
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


def col_index_to_col_letter(column_index):
    column_letter = ""
    while column_index > 0:
        column_index, remainder = divmod(column_index - 1, 26)
        column_letter = chr(65 + remainder) + column_letter
    return column_letter


def validate(config, worksheet, sheet):
    columns_to_validate = config.get("validations").get("columns") or []
    must_have_columns = config.get("must_have_columns")
    read_as_string = config.get("read_as_string") or []
    unique_columns = config.get("unique_columns") or []
    headers = worksheet.columns.to_list()
    duplicates = {}
    if not set(must_have_columns).issubset(set(headers)):
        SHEET_ERRORS[sheet].append(
            {
                "type": "column",
                "errorMessage": "Sheet '{sheet}' must have column(s): {missing_columns}".format(
                    sheet=sheet,
                    missing_columns=", ".join(
                        list(set(must_have_columns) - set(headers))
                    ),
                ),
            }
        )

    for column in headers:
        column_letter = col_index_to_col_letter(
            worksheet.columns.get_loc(column) + 1
        )  # pandas column locations are 0 based.
        if column in read_as_string:
            worksheet[column] = worksheet[column].apply(lambda x: str(x))
        if column in unique_columns:
            if not worksheet[column].is_unique:
                counter = Counter(worksheet[column])
                duplicates[column] = {k: v for k, v in counter.items() if v > 1}
        if column in config.get("exclude", []):
            continue
        if column in columns_to_validate:
            is_valid_cell(
                worksheet[column],
                validations=columns_to_validate[column],
                sheet=sheet,
                header=column,
                column_letter=column_letter,
            )
        elif config.get("validations").get("default"):
            is_valid_cell(
                worksheet[column],
                validations=config.get("validations").get("default"),
                sheet=sheet,
                header=column,
                column_letter=column_letter,
            )

    if duplicates:
        SHEET_ERRORS[sheet].append(
            {
                "type": "duplicates",
                "errorMessage": "There are duplicate values in the following columns:",
                "columns": duplicates,
            }
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
            errors["fileErrors"] = {
                "errorMessage": "The uploaded file must have sheet(s): {missing_sheets}".format(
                    missing_sheets=", ".join(list(set(config_sheets) - set(sheets))),
                ),
            }
        for sheet in config_sheets:
            if config.get(sheet) and sheet in sheets:
                worksheet = pd.read_excel(xlsx, sheet_name=sheet, keep_default_na=False)
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
    parser = argparse.ArgumentParser(description="Add path to the yaml and xlsx files.")
    parser.add_argument("--yaml_file_path", required=True, nargs=1, type=str)
    parser.add_argument("--xlsx_file_path", required=True, nargs=1, type=str)
    args = parser.parse_args()
    yaml_filepath = args.yaml_file_path[0]
    xlsx_filepath = args.xlsx_file_path[0]
    is_valid, errors, data = run_validations(xlsx_filepath, yaml_filepath)
    if is_valid:
        print(
            "The excel file has no validation issues as per the provided configurations in the yaml file."
        )
    else:
        print(
            "The excel file has validation issues. Please refer the ValitionErrors.log file to find the issues."
        )
        with open("ValidationErrors.log", "w+") as error_logs:
            error_logs.write("Please fix the errors mentioned in the error file.\n\n\n")
            if errors.get("fileErrors"):
                error_logs.write(errors.get("fileErrors").get("errorMessage") + "\n\n")
            df = pd.json_normalize(
                data=errors.get("sheetErrors"), record_path="errors", meta=["sheetName"]
            ).fillna("NA")
            df.insert(0, "sheetName", df.pop("sheetName"))
            error_logs.write(df.to_string())
