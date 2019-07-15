"""
office.py - utilities to help with Microsoft Office (TM)
"""

import logging
import win32com.client as win32
import pywintypes as pwt


def open_office_app(which, visible=True):
    """
    Get running Office app instance if possible, else return new instance.

    App can be Word, Excel, Outlook, etc.
    """
    try:
        app = win32.GetActiveObject("{}.Application".format(which))
        logging.debug("Running %s instance found, returning object", which)
    except pwt.com_error: # pylint: disable=E1101
        app = win32.gencache.EnsureDispatch("{}.Application".format(which))
        app.Visible = visible
        logging.debug("No running %s instances, returning new instance", which)
    return app
