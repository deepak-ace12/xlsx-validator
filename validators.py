import re
from datetime import datetime
from abc import ABCMeta, abstractmethod, abstractproperty
from openpyxl.utils.datetime import from_excel


class BaseValidator:
    __metaclass__ = ABCMeta

    @abstractmethod
    def validate(self, value):
        if hasattr(self, "trim"):
            return value.strip()

        return value

    # When we search for an attribute in a class that is involved
    # in python multiple inheritance, an order is followed.
    # First, it is searched in the current class. If not found,
    # the search moves to parent classes. This is left-to-right,
    # depth-first. This order is called linearization of class Child,
    # and the set of rules applied are called MRO (Method Resolution Order).
    # To get the MRO of a class, you can use either the __mro__ attribute
    # or the mro() method.
    @classmethod
    def __subclasshook__(cls, C):
        if cls is BaseValidator:
            if any("validate" in B.__dict__ for B in C.__mro__):
                return True
        return NotImplemented

    def __init__(self, params):
        if params:
            for key, value in params.items():
                setattr(self, key, value)


class OptionValidator(BaseValidator):
    def validate(self, value):
        if value:
            value = super(OptionValidator, self).validate(value)
            if not self.case_sensitive:
                value = value.lower()
                if value not in [option.lower() for option in self.options]:
                    raise Exception(self.error_msg)
            elif value not in self.options:
                raise Exception(self.error_msg)


class DateTimeValidator(BaseValidator):
    def validate(self, value):
        if value:
            value = super(DateTimeValidator, self).validate(value)
            if type(value) is datetime: # For Date Only
                try:
                    value = value.strftime(self.format)
                except Exception as ex:
                    raise Exception(self.error_msg)
            try:
                datetime.strptime(value, self.format)
            except ValueError:
                raise Exception(self.error_msg)


class EmailValidator(BaseValidator):
    def validate(self, value):
        if value:
            value = super(EmailValidator, self).validate(value)
            regex = r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b"
            if not re.fullmatch(regex, value):
                raise Exception(self.error_msg)


class ExcelDateValidator(BaseValidator):
    def validate(self, value):
        if value:
            value = super(ExcelDateValidator, self).validate(value)
            try:
                if str(value).isdigit():
                    from_excel(int(value))
                else:
                    from_excel(float(value))
            except Exception as ex:
                raise Exception(self.error_msg)


class LengthValidator(BaseValidator):
    def validate(self, value):
        if value:
            value = super(LengthValidator, self).validate(value)
            if self.operation == "min" and len(value) < self.threshold:
                raise Exception(self.error_msg)
            if self.operation == "max" and len(value) > self.threshold:
                raise Exception(self.error_msg)


class RequiredValidator(BaseValidator):
    def validate(self, value):
        value = super(RequiredValidator, self).validate(value)
        if value in ["", None]:
            raise Exception(self.error_msg)


class RegexValidator(BaseValidator):
    def validate(self, value):
        if value:
            value = super(RegexValidator, self).validate(value)
            if not re.match(self.pattern, value):
                raise Exception(self.error_msg)


class TypeValidator(BaseValidator):
    def validate(self, value):
        string_to_func = {"int": int, "float": float, "bool": bool}
        if value:
            value = super(TypeValidator, self).validate(value)
            try:
                string_to_func[self.type](value)
            except Exception as ex:
                raise Exception(self.error_msg)


class NonNegativeValidator(TypeValidator):
    def validate(self, value):
        if value:
            super(NonNegativeValidator, self).validate(value)
        if value < 0:
            raise Exception(self.error_msg)


class ComparatorValidator(BaseValidator):
    def validate(self, value):

        if value:
            value = str(value)
            if not str(value).replace(".", "").isdigit():
                raise Exception("Cell value must be a number.")
            if self.operation == "gt" and eval(value) >= self.threshold:
                raise Exception(self.error_msg)
            if self.operation == "lt" and eval(value) <= self.threshold:
                raise Exception(self.error_msg)
