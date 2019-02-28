# Microsoft Excel Reporting and Analysis Dashboard for Wild Apricot

This project helps Wild Apricot administrators with common reporting and data extraction needs. Excel is employed to call the Wild Apricot API to extract various parts of the database and make a copy of the data in a structured Excel worksheet. Data can be transformed and modified in a variety of ways.

The project is built in Visual Basic for Applications and requires access to the Microsoft XML library (MSXML).

# Usage

1. Open the Excel dashboard in Microsoft Excel for Windows.
2. [Enable editing and macros](https://support.office.com/en-us/article/enable-or-disable-macros-in-office-files-12b036fd-d140-4e74-b45e-16fed1a7e5c6)
3. Enter an valid Wild Apricot API key into the spreadsheet.
4. Use the filters and buttons to run reports.

A Wild Apricot API key can be created in the Settings/Security/Authorized Applications settings on free and licensed Wild Apricot instances. [Documentation on acquiring a key is available.](http://gethelp.wildapricot.com/en/articles/180)

*Note:* Some datasets can be large and may take a while to run. The application is limited to about 100-200 records per minute.

# Features

The dashboard provides several features:

1. The contact database can be download into the spreadsheet. A variety of filters can be applied before downloading the data. The resulting worksheet can be used to bulk edit data for import into Wild Apricot.
2. The [Membership Summary](https://newpathconsulting.wildapricot.org/Sys/InHelp#summary) report can be run and duplicated into an Excel worksheet. This data can be saved over time and used for trending analysis.
3. The Wild Apricot [Audit Log](http://gethelp.wildapricot.com/en/articles/68) can be exported over a custom period of time. This data can be used for auditing and financial reconciliation purposes. It can also be used to summarize various events such as event registrations and membership renewals. Payment Service approved/failure codes are also downloaded in a structured way, which can be used to reconcile with any payment service.

# System Requirements

The dashboard is only supported on Excel for Windows 2007 and later. It is not supported in Office 365.

Unfortuantely due to limitations for XML parsing support on Excel for Mac, the dashboard is not support on Apple Macintosh computers.

# References

## Roadmap

[Future wishlist items](https://slack-files.com/T02FHPK8E-F9A9H5GCX-c36579e6ca) are found in NewPath's Slack

## Improvements
* [VBA-Web: VBA library that connects Office to REST API services](http://vba-tools.github.io/VBA-Web/)
