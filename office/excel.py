"""
excel.py - utilities to help with Microsoft Excel (TM)
"""

import logging
from pathlib import Path
import string
import sys
from typing import Dict
import win32com.client as win32
import pywintypes as pwt


class Excel:

    def __init__(self, visible=True):
        try:
            app = win32.GetActiveObject("Excel.Application")
            logging.debug("Running Excel instance found, returning object")
        except pwt.com_error: # pylint: disable=E1101
            app = win32.gencache.EnsureDispatch("Excel.Application")
            app.Visible = visible
            logging.debug("No running Excel instances, returning new instance")
        self.app = app
        self.wbs = {} # type: Dict[str, bool]

    def Workbooks(self, path: Path):
        if hasattr(path, 'name') is False:
            print("Must pass in a 'Path' object. Can't continue.")
            sys.exit(1)
        try:
            wb = self.app.Workbooks(str(path.name))
            self.wbs[wb.Name] = False
        except Exception as x:
            wb = self.app.Workbooks.Open(str(path))
            self.wbs[wb.Name] = True
        return wb

    Workbook = Workbooks

    def __del__(self):
        for wbname, should_close in self.wbs.items():
            if should_close:
                self.app.Workbooks(wbname).Close()


def num2col(num):
    """Convert given column letter to an Excel column number."""
    result = []
    while num:
        num, rem = divmod(num-1, 26)
        result[:0] = string.ascii_uppercase[rem]
    return ''.join(result)

def col2num(ltr):
    num = 0
    for c in ltr:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num


if __name__ == '__main__':
    for nxt in range(17, 120, 3):
        print(f'= SUM({num2col(nxt)}5:{num2col(nxt+2)}5)', end='\t')
