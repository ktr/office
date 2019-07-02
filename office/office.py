"""
office.py - utilities to help with Microsoft Office (TM)

Example:

    outlook = Outlook()
    tbl = outlook.list_to_tbl([[1, 2, 3], [4, 5, 6]])
    with open(r'H:\test1.png', 'rb') as io:
        img = outlook.inline_img(io)
    outlook.create_mail('test@example.com', 'Hello!', f'Hello!<br><br>{tbl}<br><br>{img}<br><br>Goodbye!', show=True)

"""

import base64
import datetime 
from win32com.client import Dispatch
import win32com.client


class Outlook:

    def __init__(self):
        self.ns = self.outlook = self.inbox = None

    def _connect(self):
        self.outlook = Dispatch("Outlook.Application")
        self.ns = self.outlook.GetNamespace("MAPI")
        self.inbox = self.ns.GetDefaultFolder("6")

    def create_mail(self, to: str, subject: str, body: str, cc: str='', attachments: list=[], show: bool=False, send: bool=False):
        if not self.outlook:
            self._connect()
        msg = self.outlook.CreateItem(0x0)
        msg.Subject = subject
        for path in attachments:
            msg.Attachments.Add(path)
        msg.To = to
        msg.CC = cc
        msg.HTMLBody = body
        if show:
            msg.Display()
        if send:
            msg.Send()

    def inline_img(self, io) -> str:
        """
        Returns html snippet to embed an image in an email.

        `io` should be an open file (e.g., open(<path>, 'rb')).
        """
        encoded_image = base64.b64encode(io.read()).decode("utf-8")
        return '<img src="data:image/png;base64,%s"/>' % encoded_image

    def tbl_style(self, styles={}):
        styles.setdefault('header-bg', '#1F77B4')
        styles.setdefault('header-fg', '#FFFFFF')
        styles.setdefault('th-border-color', '#222222')
        styles.setdefault('td-border-color', '#222222')
        return '''\
<style type="text/css">
  table {
    border-collapse:collapse;
    border-spacing:0;
  }
  table td {
    font-family:Arial, sans-serif;
    font-size:14px;
    padding:5px 10px;
    border-style:solid;
    border-width:1px;
    overflow:hidden;
    word-break:normal;
    border-color: %(td-border-color)s;
    text-align: right;
    width: 100px;
  }
  table th {
    font-family:Arial, Helvetica, sans-serif !important;
    font-size:14px;
    font-weight:bold;
    padding:5px 10px;
    border-style:solid;
    border-width:1px;
    overflow:hidden;
    word-break:normal;
    background-color:%(header-bg)s;
    color:%(header-fg)s;
    vertical-align:top;
    border-color: %(th-border-color)s;
    width: 100px;
  }
</style>''' % styles

    def list_to_tbl(self, lst: list, first_is_hdr: bool=True) -> str:
        """
        Returns html table of `lst`.
        """
        head = '<table class="tg">\n'
        rowg = lambda row, mk='td': '<tr>\n' + '\n'.join([f'<{mk}>{_}</{mk}>' for _ in row]) + '\n</tr>'
        row1 = rowg(lst[0], 'th') + '\n'
        rows = row1 + '\n'.join([rowg(_) for _ in lst[1:]])
        foot = '\n</table>'
        return self.tbl_style() + head + rows + foot


if __name__ == "__main__":
    outlook = Outlook()
    tbl = outlook.list_to_tbl([[1, 2, 3], [4, 5, 6]])
    with open(r'C:\test1.png', 'rb') as io:
        img = outlook.inline_img(io)
    outlook.create_mail('test@example.com', 'Hello!', f'Hello!<br><br>{tbl}<br><br>{img}<br><br>Goodbye!', show=True)
