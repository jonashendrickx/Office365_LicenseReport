## Synopsis

This will generate an Excel report of all active licenses in your Office 365 and then mail it to the specified e-mail address.

## Motivation

This was initially made for a client who wanted a weekly report. To reduce time spent on generating a report. I created this Powershell script to automate the process. This is only meant for use with Project Server, but the filter can also be removed to include all Office 365 licenses.

## Installation

### Requirements
  * A [Microsoft Online Services Sign-In Assistant](https://www.microsoft.com/en-us/download/details.aspx?id=41950)
  * A [Windows Azure Active Directory Module for Windows Powershell](http://go.microsoft.com/fwlink/p/?linkid=236297)
  * Windows 8.1 or Windows Server 2012 R2 and older need to install A [Windows Management Framework 5.0](https://www.microsoft.com/en-us/download/details.aspx?id=50395)

### Settings

Here is a list of parameter the script takes.

| Parameter | Required | Default value | Example | Description |
|:---------:|:--------:|:-------------:|:-------:|:------------|
| -O365Username   | Yes | / | you@example.onmicrosoft.com | Office 365 username, requires authorization to read licenses |
| -O365Password | Yes | / | yourpassword | Office 365 password |
| -FilePrefix | No | prefix | prefix | Prefix for the file generated. |
| -FileOutDir | Yes | / | C:\path\to\file\out | Directory where the report file will be temporarily saved. |
| -MailSender | Yes | / | you@example.com | E-mail the report file will be sent with.
| -MailPassword | Yes | / | yourpassword | Password of the e-mail you'll be using to send the e-mail. |
| -MailSmtpServer | Yes | smtp.office365.com | smtp.office365.com | SMTP server |
| -MailSmtpPort | No | 587 | 587 | SMTP port |
| -MailRecipients | Yes | / | person1@example.com, person2@example.com | Recipients |
| -MailCc | No | / | someone@example.com | E-mails in CC |

**Example:**
Powershell.exe -Command "& C:\path\to\script\Send-O365PsLicenseReport.ps1 -O365Username 'you@example.onmicrosoft.com' -O365Password 'yourpassword' -FilePrefix 'prefix' -FileOutDir 'C:\path\to\file\out' -MailSender 'you@example.com' -MailPassword 'yourpassword' -MailRecipients 'person1@example.com', 'person2@example.com' -MailCc 'carboncopy@example.com'"

**Note:**
The debugging settings section in the script can still be edited if desired.

### Task Scheduler (optional)

Now let's schedule the script.

  1. Open the Task Scheduler in Windows or Windows Server.
  2. Right-click **Task Scheduler (local)** in the left pane.
  3. Select **Create Task**.
  4. Give it a meaningful name.
  5. Select a user, preferably one that has permissions to save the Excel file to the path you entered in the script.
  6. Tick the radio button **Run whether the user is logged in or not**.
  7. In the **Triggers** tab, enter the desired trigger.
  8. In the **Actions** tab, make sure **Start a program** is selected in the **Action** combobox.
  9. Click the **Browse** button and select the Powershell executable: **C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe**
  10. Then in the **Add arguments** text box enter: **-Command "& C:\path\to\script\Send-O365PsLicenseReport.ps1 -O365Username 'you@example.onmicrosoft.com' -O365Password 'yourpassword' -FilePrefix 'prefix' -FileOutDir 'C:\path\to\file\out' -MailSender 'you@example.com' -MailPassword 'yourpassword' -MailRecipients 'person1@example.com', 'person2@example.com' -MailCc 'carboncopy@example.com'"**

## Contributors

Jonas Hendrickx

## License

MIT License
