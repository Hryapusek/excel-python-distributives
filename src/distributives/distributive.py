from dataclasses import dataclass
from datetime import datetime


@dataclass
class Distributive:
    number: str = None
    date: datetime = None
    sheet_name: str = None
    name: str = None
    user_name: str = None
