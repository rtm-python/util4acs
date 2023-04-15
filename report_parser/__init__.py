"""
This package parse batch reports received from ACS into JSON-file next to operate with structured data.

"""

import logging
from abc import ABC
from dataclasses import dataclass
from typing import List
from datetime import datetime

logging.basicConfig(
    format="[%(asctime)s] %(name)s: %(levelname)s (%(module)s.%(funcName)s.%(lineno)d) - %(message)s",
    level=logging.ERROR,
)
logger = logging.getLogger(__name__)


@dataclass
class AccessData:
    restricted_area: str
    enter_in: datetime
    enter_in_turnstyle: str
    exit_out: datetime
    exit_out_turnstyle: str


@dataclass
class EmployeeAccess:
    name: str
    unit: str
    access_data_list: List[AccessData]


@dataclass
class Parser(ABC):
    config: dict

    def parse(self, filename: str) -> List[EmployeeAccess]:
        pass
