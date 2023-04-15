"""
This package parse batch reports received from ACS into JSON-file next to operate with structured data.

"""

import logging

logging.basicConfig(
    format="[%(asctime)s] %(name)s: %(levelname)s (%(module)s.%(funcName)s.%(lineno)d) - %(message)s",
    level=logging.DEBUG,
)
logger = logging.getLogger(__name__)
