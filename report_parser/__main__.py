"""
This is an additional module to demonstrate package functionality (parses reports from folder and creates JSON data file).

Example:
    $ python -m report_parser <reports_folder>
"""

from typing import List
from report_parser import logger


def main(reports: List[str]) -> None:
    """This function parses reports from defined folder and creates JSON file with structured data.

    Args:
        reports_folder  folder path where reports from ACS stored

    Returns:
        None
    """
    logger.info(f'Package "{__package__}" demonstration')


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        logger.error("No report filenames or folders provided.")
        exit(1)
    main(sys.argv[1:])
