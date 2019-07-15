"""
outlook.py - utilities to help with Microsoft Outlook (TM)

Example:

    outlook = Outlook()
    tbl = outlook.list_to_tbl([[1, 2, 3], [4, 5, 6]])
    with open(r'H:\test1.png', 'rb') as io:
        img = outlook.inline_img(io)
    outlook.create_mail('test@example.com', 'Hello!', f'Hello!<br><br>{tbl}<br><br>{img}<br><br>Goodbye!', show=True)

"""

import base64
import datetime 
import logging
import pytz
from win32com.client import Dispatch
import win32com.client


class Outlook:

    def __init__(self):
        self.ns = self.outlook = self.inbox = None

    def _connect(self):
        # https://docs.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
        if not self.outlook:
            self.outlook = Dispatch("Outlook.Application")
            self.ns = self.outlook.GetNamespace("MAPI")
            self.calendar = self.ns.GetDefaultFolder(9)
            self.inbox = self.ns.GetDefaultFolder(6)

    def create_mail(self, to: str, subject: str, body: str, cc: str='', attachments: list=[], show: bool=False, send: bool=False):
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

    def find_open_slots(self, appts, duration=None):
        # appts should be list of start/end times
        start = appts[0][0]
        end = appts[-1][1]
        # find open slots between 9am–6pm
        utc = pytz.UTC
        hours = (utc.localize(datetime.datetime(start.year, start.month, start.day, 9)),
                 utc.localize(datetime.datetime(end.year, end.month, end.day, 18)))

        if duration is None:
            duration = datetime.timedelta(minutes=30)

        slots = sorted([(hours[0], hours[0])] + appts + [(hours[1], hours[1])])
        open_slots = []
        for start, end in ((slots[i][1], slots[i+1][0]) for i in range(len(slots)-1)):
            while start + duration <= end:
                open_slots.append([start, start + duration])
                start += duration
        this = open_slots[0]
        print(f'\n{this[0]:%Y-%m-%d}')
        for slot in open_slots[1:]:
            # only offer up times after default start time
            if slot[0].hour < hours[0].hour or slot[0].hour > hours[1].hour:
                continue
            # or before the default end time
            if slot[1].hour < hours[0].hour or slot[1].hour > hours[1].hour:
                continue
            # otherwise, check if we should combine consecutive time sequences
            if slot[0] <= this[1]:
                this[1] = slot[1]
            # if not, print them out and move on to the next one
            else:
                print(f'  {this[0]:%I:%M %p} to {this[1]:%I:%M %p}')
                if this[0].day != slot[0].day:
                    print(f'\n{this[0]:%Y-%m-%d}')
                this = slot

    def show_appts(self, begin=None, end=None):
        self._connect()
        if begin is None:
            begin = datetime.date.today() + datetime.timedelta(days=1) # tomorrow
        if end is None:
            end = begin + datetime.timedelta(days=2) # duration of 1 day

        # http://msdn.microsoft.com/en-us/library/office/aa210899(v=office.11).aspx
        appts = self.calendar.Items 

        appts.IncludeRecurrences = "True"
        # Need the following call to 'Sort', otherwise will include all
        # recurrences (whether they are in the list or not!)
        appts.Sort("[Start]")
        where = f"[Start] >= '{begin.strftime('%m/%d/%Y')}' AND [End] <= '{end.strftime('%m/%d/%Y')}'"

        msg = "{1:%H:%M}–{2:%H:%M}, {0} (Organizer: {3})"
        appt_lst = []
        for item in appts.Restrict(where):
            print(msg.format(item.Subject, item.Start, item.End, item.Organizer))
            appt_lst.append((item.Start, item.End,))
        print("\nOpen Slots:")
        self.find_open_slots(appt_lst)


if __name__ == "__main__":
    outlook = Outlook()
    create_html_sample = 0
    show_appts = 1
    if create_html_sample:
        tbl = outlook.list_to_tbl([[1, 2, 3], [4, 5, 6]])
        with open(r'C:\test1.png', 'rb') as io:
            img = outlook.inline_img(io)
        outlook.create_mail('test@example.com', 'Hello!', f'Hello!<br><br>{tbl}<br><br>{img}<br><br>Goodbye!', show=True)
    if show_appts:
        outlook.show_appts()
