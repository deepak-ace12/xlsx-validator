import os
import yaml
import argparse


class YamlValidator:
    def __init__(self, dev_file, ref_file=None):
        ref_file = os.path.abspath("reference_yaml.yml")
        with open(dev_file, "r") as d_yml:
            self.dev_yaml = yaml.safe_load(d_yml)

        with open(ref_file, "r") as r_yml:
            self.ref_yaml = yaml.safe_load(r_yml)

    def validate_yaml(self):
        validation_errors = []
        sheets = self.dev_yaml.get("sheets")
        if not sheets:
            validation.append("Yaml must include the list of sheets.")

        for sheet in sheets:
            sheet_data = self.dev_yaml.get(sheet)
            high_level_keys = sheet_data.keys()
            is_valid, missing_keys = self.has_all_keys(
                "high_level_keys", high_level_keys
            )
            if not is_valid:
                validation_errors.append(
                    "{sheet} sheet file must contain: {missing_keys}".format(
                        sheet=sheet, missing_keys=", ".join(missing_keys)
                    )
                )
            columns = sheet_data.get("validations").get("columns")
            if not columns:
                validation_errors.append(
                    "validation class must have the list of columns"
                )

            for column_name, validations in columns.items():
                for valdn_type in validations:
                    for class_name, validation in valdn_type.items():
                        is_valid, missing_keys = self.has_all_keys(
                            class_name, validation.keys()
                        )
                        if not is_valid:
                            validation_errors.append(
                                "{column_name} column's {class_name} class is missing {missing_keys}".format(
                                    column_name=column_name,
                                    class_name=class_name,
                                    missing_keys=", ".join(missing_keys),
                                )
                            )
            if validation_errors:
                raise Exception(validation_errors)
            else:
                print(
                    "Yaml file is validated successfully. There's no validation error."
                )

    def has_all_keys(self, yml_class, keys):
        if set(self.ref_yaml.get(yml_class)) == set(keys):
            return (True, None)
        else:
            return (False, list(set(self.ref_yaml.get(yml_class)) - set(keys)))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Add path of your yaml file to validate it.")
    parser.add_argument("--dev_yaml_file_path", required=True, nargs=1, type=str)
    args = parser.parse_args()
    dev_file = args.dev_yaml_file_path[0]
    YamlValidator(dev_file).validate_yaml()
