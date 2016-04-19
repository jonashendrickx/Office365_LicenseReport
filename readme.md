## Synopsis

This will generate an Excel report of all active licenses in your Office 365 and then mail it to the specified e-mail address.

## Motivation

This was initially made for a client who wanted a weekly report. To reduce time spent on generating a report. I created this Powershell script to automate the process. This is only meant for use with Project Server, but the filter can also be removed to include all Office 365 licenses.

## Installation

### Requirements
  * A [Microsoft Online Services Sign-In Assistant](https://www.microsoft.com/en-us/download/details.aspx?id=41950)
  * A [Windows Azure Active Directory Module for Windows Powershell](http://go.microsoft.com/fwlink/p/?linkid=236297)
  * Windows 8.1 or Windows Server 2012 R2 and older need to install A [Windows Management Framework 5.0](https://www.microsoft.com/en-us/download/details.aspx?id=50395)
  * Have Microsoft Office 2013 or later installed.

### Settings

In the settings section of the Powershell script, first we will define where the file will be temporarily saved.

  * $workingDirectory: This is the location of where the Excel file should be saved. Make sure you have read-write rights for this directory.
  * $reportFile: Path and template filename of how the Excel file will be saved.

In the settings section of the Powershell script, we will define the Office 365 credentials.

  * $office365Login: Your Office 365 login
  * $office365Password: Your Office 365 password

In the settings section of the Powershell script, we will define the mailing credentials and settings.

  * $mailSender: Your e-mail login (Preferably one with a password that does not expire if you are scheduling the script.)
  * $mailPassword: Your e-mail password
  * $mailSmtpServer
  * $mailSmtpPort
  * $mailRecipients: Recipients of the e-mail. Put each recipient between double quotes, and separate them with a comma.
  * $mailCc: Recipients of the e-mail in carbon copy. Put each recipient between double quotes, and separate them with a comma.

### Task Scheduler (optional)

Now let's schedule the script.

  1. Open the Task Scheduler in Windows or Windows Server.
  2. Right-click ‘Task Scheduler (local)’ in the left pane.
  3. Select ‘Create Task’
  4. Give it a meaningful name.
  5. Select a user, preferably one that has permissions to save the Excel file to the path you entered in the script.
  6. Tick the radio button ‘Run whether the user is logged in or not’.
  7. In the ‘Triggers’ tab, enter the desired trigger.
  8. In the ‘Actions’ tab, make sure ‘Start a program’ is selected in the ‘Action’ combobox.
  9. Click the ‘Browse’ button and select the Powershell executable: **C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe**
  10. Then in the ‘Add arguments’ text box enter: **-file "C:\path\to\your\script\myscript.ps1"**

Now the last missing bit is to create these missing directories, to prevent an error with Excel automation.

  * On a 32-bit and 64-bit operating system: **C:\Windows\System32\config\systemprofile\Desktop**
  * On a 64-bit operating system: **C:\Windows\SysWOW64\config\systemprofile\Desktop**

## Contributors

Jonas Hendrickx

## License

MIT License