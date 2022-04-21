# Case_Analysis_Bot

This is a simple program that pulls cases from Salesforce within a certain date range. Then it pulls info on the cases and writes that info to a spreadsheet for future review.

**Built With**

- [Simple-Salesforce](https://github.com/simple-salesforce/simple-salesforce) To access Salesforce
- [Gspread](https://github.com/burnash/gspread) To read and write data to Google Sheets
- [Python Decouple] (https://pypi.org/project/python-decouple/) For reading Variables from .env file

**Usage**

This script was made to help with case reviews. It has 2 built-in date ranges for the 2 kinds of common case reviews I do, Bi-Weekly Check-in and Historical Reviews. Also has an option for a custom time range in case I need to review cases in a different time range. Once the date range is set, an age is created to select all cases from that date to today. The cases are then found using a SOQL query. Info gathered from the cases is written to a spreadsheet using Gspread. The exact spreadsheet and the page on the spreadsheet are selected by the user at runtime. This does require setting up a Google Service Account and saving the credentials somewhere

**Roadmap**

- Save spreadsheet info between runs (Most likely using pickling)
- Allow users to select different fields to select info from
- Add a GUI (Using Tkinter)
- Add Salesforce login prompt
