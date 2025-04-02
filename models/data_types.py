from enum import Enum

class DataType(Enum):
    SINGLE_VALUE = "single_value"
    LIST = "list"
    TABLE = "table"
    DICTIONARY = "dictionary"
    SERIES = "series"
    HIERARCHICAL = "hierarchical"
