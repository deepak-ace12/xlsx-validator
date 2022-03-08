import yaml
import sys
import inspect
from collections import defaultdict
from openpyxl.reader.excel import load_workbook
import validators_pd as _validators
import pandas as pd
from validators_pd import *


ERRORS = defaultdict(list)

validating_objects = {v_cls[0]: eval(v_cls[0])() for v_cls in inspect.getmembers(_validators, inspect.isclass) if v_cls[1].__module__ == _validators.__name__}
def is_valid_cell(values, validations, sheet, header):
    index = 1
    for value in values:
        index += 1
        for valdn_type in validations:
            for typ, data in valdn_type.items():
                metadata = {
                    "header": header,
                    "rowNumber": index
                }
                # validating_class = eval(typ) #getattr(sys.modules[__name__], typ)
                validator = validating_objects.get(typ)
                try:
                    validator.validate(value, data)
                except Exception as ex:
                    metadata["error"] = ex.args[0]
                    ERRORS[sheet].append(metadata)


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
    # ********************** COLUMN_CASES ******************* #
    headers = worksheet.columns.to_list()
    
    # ********************** COLUMN_CASES ******************* #
    if not set(must_have_columns).issubset(set(headers)):
        raise Exception(
            {
                "sheet": worksheet.title,
                "error": "Sheet {sheet} must have column(s) {missing_columns}".format(
                    sheet=worksheet.title,
                    missing_columns=", ".join(
                        list(
                            set(must_have_columns)
                            - set(headers)
                        )
                    ),
                ),
            }
        )
    import time
    t11 = time.time()
    for column in headers:
        if column in read_as_string:
             worksheet[column] = worksheet[column].apply(lambda x: str(x))
        if column in config.get("exclude", []):
            continue
        t1 = time.time()
        if column in columns_to_validate:
            is_valid_cell(worksheet[column], validations=columns_to_validate[column], sheet=sheet, header=column)          
        elif config.get("validations").get("default"):
            is_valid_cell(worksheet[column], validations=config.get("validations").get("default"), sheet=sheet, header=column)
        t2 = time.time()
        print(column, (t2-t1))
    t22 = time.time()
    print("valdnt", (t22-t11))

def validate_excel(xlsx_filepath, yaml_filepath):

    try:
        import time
        t1 = time.time()
        xlsx = pd.ExcelFile(xlsx_filepath , engine="openpyxl")
        sheets = xlsx.sheet_names
        config = set_config(yaml_filepath)
        config_sheets = config.get("sheets")
        t2 = time.time()
        print("loading",(t2-t1))
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
                t11 = time.time()
                worksheet = pd.read_excel(xlsx, sheet_name=sheet)
                worksheet = worksheet.astype(object).where(worksheet.notna(), None) # turning NaN to None
                t22 = time.time()
                print("sheet", (t22-t11))
                validate(config.get(sheet), worksheet, sheet)
                t33 = time.time()
        
        print()
    except Exception as e:
        sys.exit(e.args[0])
    
    print(dict(ERRORS))



if __name__ == "__main__":
    import time
    t1 = time.time()
    xlsx_filepath = "/Users/I1597/Downloads/ucc_data_columns_final.xlsx"
    yaml_filepath = "/Users/I1597/Documents/repositories/excel_validator/thn_final_validator.yml"
    # xlsx_filepath = "/Users/I1597/Documents/repositories/excel_validator/one_lakh_record.xlsx"
    # yaml_filepath = "/Users/I1597/Documents/repositories/excel_validator/sales_record.yml"
    validate_excel(xlsx_filepath, yaml_filepath)
    t2 = time.time()
    print("Total Time", (t2 - t1))