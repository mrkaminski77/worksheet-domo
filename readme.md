# Scorecard Domo Uploader

This tool automates sending a range from a worksheet to a specified email address.

## Description

There are two component modules required to complete an upload.

### Config Module

``` Powershell
# path to target Workbook
$WorkbookPath = ''
# name of sheet to capture data from
$WorksheetName = ''

# letter references for the columns to capture
$ColumnStart = ''
$ColumnEnd = ''
# number references to the rows to capture
$RowStart = 
$RowEnd = 

#email config
$subjectLine = 'CSS Scorecard Data'
$to = '04b9e7f931384a14b73e7264724a16d4@serco-ap-au.mail.domo.com'
$cc = 'david.leyden@serco-ap.com.au'
$from = 'david.leyden@serco-ap.com.au'

# array of fieldnames
$fieldnames = @()
```

Note that the number of fieldnames in the array must match the number of columns.
The script will check for this and return an error if they do not match.

### Uploader Module

The uploader module expects the variables defined in the Config Module.

### Usage

Create a Powershell script that dot sources the config file followed by the upload module.
``` Powershell
. .\config.ps1
. .\uploader.ps1
```

In this way you can have a separate config file for each worksheet and one copy of the uploader component.

## Technical Notes

- Each row is turned in to hash table using the fieldnames provided.
- The hash table is converted to a PSObject that can be exported using Export-Csv cmdlet.
- Using the _System.Net.Mail.MailMessage_ class we construct the email and attach the csv file.
- The file is sent using _System.Net.Mail.SmtpClient_ with the default credentials.





