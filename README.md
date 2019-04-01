# CRDashboard

Automate the gathering of select metrics for University of Minnesota's Chrome River Rollout.

# Guide

In order to run the script, 5 excel spreadsheets need to be included in the same folder as the dashboard.py 

Those reports are generated from the Chrome River Analytics Tool and must be formatted in a specific way.


# Report Formatting

## Standard Report > Expense Analysis

* Change Dates to be 3/1/2018 - Today's Date
* Add the following  7 columns in order:
    * Expense Owner
    * Approval Status
    * Export Status
    * Export Date
    * Is Firm Paid
    * RRC Code
    * Affliation
* Run file as Excel
* Name of the file must be `expense_analysis.xlsx`


## Standard Report > Reference > Person Report

* Add the following column
    * RRC Code
* Run to Excel
* Name of the file must be `reference-person_report.xlsx`

## UM > My Content > Who Was Delegates Set Up
* Run Report, automatically outputs an Excel File
* Name of the file must be `Who_has_delegates_set_up.xlsx`

## Standard Report > Submitted Reports
* Change Dates to be 3/1/2018 - Today's Date
* Add the following 2 columns in order:
    * Expense Creator Name
    * RRC Code
* Name of the file must be `expense-submitted_reports.xlsx`

## UM > My Content > Copied from Beth: Approval Method
* Run report, automatically outputs an Excel file
* Name of the file must be `Beth_Approval_Method.xlsx`