import os.path
import sys
import time
import yaml
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string, get_column_letter
from validators import RequiredValidator, TypeValidator, LengthValidator, RegexValidator, EmailValidator, ChoiceValidator, DateTimeValidator, ExcelDateValidator, NonNegativeValidator, ComparatorValidator


def is_valid_cell(valdn_type, value, coordinate, errors, value2 = None):
    classmap = {
        'Required': RequiredValidator,
        'Type': TypeValidator,
        'Length': LengthValidator,
        'Regex': RegexValidator,
        'Email': EmailValidator,
        'Choice': ChoiceValidator,
        'Date': DateTimeValidator,
        'ExcelDate': ExcelDateValidator,
        "ComparatorValidator": ComparatorValidator,
        "NonNegativeValidator": NonNegativeValidator,
    }

    violations = []
    #TODO: try to avoid [0]
    name = list(valdn_type.keys())[0]
    data =list(valdn_type.values())[0]
    validator = classmap[name](data) #creating the object with the parameters
    result = validator.validate(value)

    #TODO: try to avoid violations
    if not result:
        violations.append(validator.get_error_message())

    if violations:
        errors.append((coordinate, violations))

    return result


def set_rules(yaml_config, yaml_validator_cls):
    """
    Takes the config yaml file and converts it into a dictionary.
    """
    rules = {}
    with open(yaml_config, "r") as yml:
        config = yaml.safe_load(yml).get(yaml_validator_cls)
        rules["default"] = config.get("validations").get("default") or None
        excluded_indexes = [column_index_from_string(column) for column in config.get("excludes")]
        return rules


def validate(rules, xlsx_file, sheet_name):
    errors = []
    workbook = load_workbook(xlsx_file, data_only=True, read_only=True)
    worksheet = workbook[sheet_name]
    columns_to_validate = rules.get("validations").get("columns")
    
    # ********************** COLUMN_CASES ******************* #
    column_letter_to_header = {}
    for row in worksheet.iter_rows(max_row=1):
        for cell in row:
            column_letter_to_header[cell.column_letter] = cell.value
    # ********************** COLUMN_CASES ******************* #
    start_row = 1
    if rules.get("iterate_by_header_name"):
        start_row = 2
    else:
        for key, _ in column_letter_to_header.items():
            column_letter_to_header[key] = key
    
    for row in worksheet.iter_rows(min_row=start_row):
        for cell in row:
            try:
                value = cell.value
            except ValueError:
                errors.append((cell.coordinate, ValueError))
            if column_letter_to_header[cell.column_letter] in rules.get("excludes", []):
                continue
            if column_letter_to_header[cell.column_letter] in columns_to_validate:
                for valdn_type in columns_to_validate[column_letter_to_header[cell.column_letter]]:
                    res = is_valid_cell(valdn_type, value, cell.coordinates, errors)
                    if not res:
                        break

            elif rules.get("default"):
                is_valid_cell(rules['default'], value, cell.coordinates, errors)

    print(errors) #TODO: do something about error logging


def validate_excel(yaml_filepath):
    

    rules = set_rules(yaml_filepath)
    sheet_name = "patients"
    xlsx_filepath = ""
    file_size = os.path.getsize(xlsx_filepath) > 10485760

    try:
        validate(rules, xlsx_filepath, sheet_name)
    except Exception as e:
        sys.exit("Error occured: " + str(e))

    