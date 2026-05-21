from dataclasses import dataclass
from typing import List, Any, Optional

@dataclass
class Cell:
    value: Any
    font_name: Optional[str] = None
    font_size: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    font_color: Optional[str] = None
    fill_color: Optional[str] = None
    horizontal_alignment: Optional[str] = None
    vertical_alignment: Optional[str] = None

@dataclass
class Row:
    cells: List[Cell]

@dataclass
class Table:
    name: str
    rows: List[Row]

@dataclass
class Sheet:
    name: str
    tables: List[Table]
