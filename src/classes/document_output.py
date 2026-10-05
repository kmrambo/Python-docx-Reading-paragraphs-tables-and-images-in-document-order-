from pandas import DataFrame
from enum import Enum


# >> Enum class to define the output format of the document and its respective metadata
class DocumentOutput(Enum):
    JSON = {
        'filetype': '.json',
        'args': {},
        'function': DataFrame.to_json,
    }
    CSV = {
        'filetype': '.csv',
        'args': {'index': False},
        'function': DataFrame.to_csv,
    }
