# Payslip

## Usage

### Configuration
Modify conf directory's main.cfg("utf-8" encode).

Change sheet_names, column_names, employee_name, employee_mail......

employee_name and employee_mail is a alias in your column, must be give.

### Command arguments

This script only support Excel version less than 2004, default is read "example.xls", you can input send_payslip.py -r "your_excel_filename" to run it.

    usage: send_payslip.py [-h] [--version] [-r read]

    send payslip email to employees

    optional arguments:
      -h, --help  show this help message and exit
      --version   show program's version number and exit
      -r read     Specify source payslip excel filename, default name is
                  example.xls.

## Installation

Prerequisites:

* Python >= 2.7
* excel file <= 2004

