Payslip
=======

Usage
-----

Configuration
-------------

Update main.cfg("utf-8" encode) in conf folder.

Change sheet_names, column_names, employee_name, employee_mail......

employee_name and employee_mail is a alias in your column, must be given.

Usage
-----

This script only support Excel version less than 2004, default is read "example.xls", you can input `send_payroll.py -r "your_excel_filename"` to run it.

    usage: send_payroll.py [-h] [--version] [-r read]

    send payroll email to employees

    optional arguments:
      -h, --help  show this help message and exit
      --version   show program's version number and exit
      -r read     Specify source payroll excel filename, default name is
                  example.xls.

Installation
------------

Prerequisites:

* Python >= 2.7
* excel file <= 2004
