#!/usr/bin/env python2.7
# encoding: utf-8

"""
send_payslip.py

Created by 4aiur on 2012-02-19.
Copyright (c) 2012 4aiur. All rights reserved.
"""


import os
import codecs
import ConfigParser
import logging
import logging.config
import argparse
import traceback
import datetime
import smtplib
import re
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import Utils
import sys
reload(sys)
sys.setdefaultencoding("utf-8")
sys.path.append("libs")
abspath = os.path.dirname(__file__)
os.chdir(abspath)
sys.path.append(os.path.join(abspath, "libs"))

from xlrd import open_workbook


__version__ = "0.90"
logging.config.fileConfig(os.path.join("conf", "logging.cfg"))


class MainConfig(object):

    def __init__(self, config_file):
        self.config = ConfigParser.RawConfigParser()
        self.config.optionxform = str
        self.config.readfp(codecs.open(config_file, "r", "utf-8-sig"))
        return

    def get_configs(self):
        self.sheet_names = self.config.get("main", "sheet_names").split(",")
        column_names = self.config.get("main", "column_names").split(",")
        self.column_names = map(unicode, column_names)
        self.employee_name = unicode(self.config.get("main", "employee_name"))
        self.employee_mail = unicode(self.config.get("main", "employee_mail"))
        self.max_cells = self.config.getint("mail", "max_cells")
        self.subject = self.config.get("mail", "subject")
        self.from_addr = self.config.get("mail", "from_addr")
        self.smtp_server = self.config.get("mail", "smtp_server")
        self.starttls = self.config.getboolean("mail", "starttls")
        self.require_auth = self.config.getboolean("mail", "require_auth")
        self.smtp_username = self.config.get("mail", "smtp_username")
        self.smtp_password = self.config.get("mail", "smtp_password")
        return


def validate_email(email):
    if isinstance(email, unicode) and len(email) > 7:
        if re.match("^.+\\@(\\[?)[a-zA-Z0-9\\-\\.]+\\.([a-zA-Z]{2,3}|[0-9]{1,3})(\\]?)$", str(email)) != None:
            return True
    return False


def send_email(config, to_addr, message):
    msg = MIMEMultipart()
    msg["Subject"] = config.subject
    msg["From"] = config.from_addr
    msg["To"] = to_addr
    msg["Date"] = Utils.formatdate(localtime = 1)
    msg["Message-ID"] = Utils.make_msgid()
    body = MIMEText(message, "html", _charset="utf-8")
    msg.attach(body)
    smtp = smtplib.SMTP()
    #smtp.set_debuglevel(1)
    smtp.connect(config.smtp_server)
    ehlo_host = config.from_addr.split("@")[1]
    smtp.ehlo(ehlo_host)
    if config.starttls:
        try:
            smtp.starttls()
            smtp.ehlo()
        except:
            pass
    if config.require_auth:
        try:
            smtp.login(config.smtp_username, config.smtp_password)
        except:
            pass
    smtp.sendmail(msg["From"], msg["To"], msg.as_string())
    smtp.quit()
    return


class Employee(object):
    def __init__(self, name, email, data):
        self.name = name
        self.email = email
        self.data = data
        return


class Excel(object):
    def __init__(self, filename, config):
        self.logger = logging.getLogger(self.__class__.__name__)
        self.config = config
        self.key_column = set(map(unicode, [config.employee_name, config.employee_mail]))
        try:
            self.workbook = open_workbook(filename)
        except Exception, error:
            exstr = traceback.format_exc()
            self.logger.error(exstr)
        return

    def get_sheets(self):
        sheets = []
        for sheet in self.workbook.sheets():
            if sheet.name in self.config.sheet_names:
                self.logger.debug("add sheet name: %s" % (sheet.name))
                sheets.append(sheet)
        return sheets

    def get_column_names_line_number(self, sheet):
        for row_line_number in range(sheet.nrows):
            row_data = set()
            for row in sheet.row(row_line_number):
                cell = row.value
                if isinstance(cell, unicode):
                    row_data.add(cell)
            if self.key_column.issubset(row_data):
                self.logger.debug("column name define line is %d" % (row_line_number))
                return row_line_number
        self.logger.error("Can't find column name define line!")
        sys.exit()
        return sheet.nrows

    def get_columns(self, sheet):
        columns = []
        column_names_line_number = self.get_column_names_line_number(sheet)
        for index, column_name in enumerate(sheet.row(column_names_line_number)):
            if column_name.value in self.config.column_names:
                columns.append((index, column_name.value))
        column_values = ""
        for x, y in columns:
            column_values += "%s, " % (str(y))
        self.logger.debug("columns: %s" % (column_values))
        return columns

    def get_employees(self, sheet, columns):
        employees = []
        column_names_line_number = self.get_column_names_line_number(sheet)
        for row_number in range(column_names_line_number+1, sheet.nrows):
            row = sheet.row(row_number)
            employee = self.get_employee_data(columns, row)
            if employee:
                employees.append(employee)
        return employees

    def get_employee_data(self, columns, row):
        data = []
        is_valid = True
        for index, column_name in columns:
            cell = row[index].value
            if column_name == unicode(self.config.employee_name):
                if isinstance(cell, unicode):
                    name = cell
                else:
                    self.logger.error("illegal compellation: %s" % (cell))
                    is_valid = False
                data.append((column_name, cell))
            elif column_name == unicode(self.config.employee_mail):
                try:
                    cell = cell.strip()
                    if validate_email(cell):
                        mail = cell
                    else:
                        self.logger.error("illegal email address: %s" % (cell))
                        is_valid = False
                except:
                    is_valid = False
            else:
                data.append((column_name, cell))
        column_values = ""
        for x, y in data:
            column_values += "%s: %s, " % (str(x), str(y))
        self.logger.debug("columns: %s" % (column_values))
        if is_valid:
            employee = Employee(name, mail, data)
            return employee
        else:
            return

    def generate_message(self, employee):
        today = datetime.date.today()
        message = """
<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01//EN">
<html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <style type="text/css">
    table, th, td {
    border: 1px solid #D4E0EE;
    border-collapse: collapse;
    font-family: "Trebuchet MS", Arial, sans-serif;
    color: #555;
    }

    td, th { padding: 4px; }

    tbody tr { background: #FCFDFE; }

    tbody tr.odd { background: #F7F9FC; }
    </style>
    <title>payslip</title>
  </head>
<body>
<table>
  <tbody>"""
        count=0
        odd = []
        even = []
        column_name = []
        column_value = []
        for name, value in employee.data:
            count += 1
            column_name.append(name)
            column_value.append(value)
            if count % self.config.max_cells == 0:
                odd.append(column_name)
                even.append(column_value)
                column_name = []
                column_value = []
        if count % self.config.max_cells != 0 and count > self.config.max_cells:
            for x in range(count%self.config.max_cells, self.config.max_cells):
                column_name.append("&nbsp;")
                column_value.append("&nbsp;")
        odd.append(column_name)
        even.append(column_value)
        for index, column_name in enumerate(odd):
            message += "\n    <tr class=\"odd\">\n      "
            for value in column_name:
                message += "<td nowrap=\"nowrap\">%s</td>" % (value)
            message += "\n    </tr>"
            column_value = even[index]
            message += "\n    <tr>\n      "
            for value in column_value:
                message += "<td nowrap=\"nowrap\">%s</td>" % (value)
            message += "\n    </tr>"
        message += """
  </tbody>
</table>
<p>%s</p>
<p>Best Regards,</p>
</body>
</html>""" % (today)
        return message

    def send_receipt(self, columns, employees):
        for employee in employees:
            message = self.generate_message(employee)
            try:
                to_addr = employee.email
                self.logger.info("send email from_addr: %s, to_addr: %s" %
                                (self.config.from_addr, to_addr))
                send_email(self.config, to_addr, message)
            except Exception, error:
                self.logger.error("send paysip failure name: %s email: %s" % (employee.name, employee.email))
                exstr = traceback.format_exc()
                self.logger.error(exstr)
        return

    def run(self):
        for sheet in self.get_sheets():
            columns = self.get_columns(sheet)
            employees = self.get_employees(sheet, columns)
            self.send_receipt(columns, employees)
        return

def main():
    """Spawn Many Commands"""
    parser = argparse.ArgumentParser(description="send payslip email to employees")
    parser.add_argument("--version", action="version", version="%s" % (__version__))
    parser.add_argument("-r", metavar="read", dest="filename", default="example.xls",
                    help="Specify source payslip excel filename, default name is example.xls.")
    args = parser.parse_args()
    main_config = MainConfig(os.path.join("conf","main.cfg"))
    main_config.get_configs()
    excel = Excel(args.filename, main_config)
    excel.run()
    return


if __name__ == "__main__":
    sys.exit(main())

