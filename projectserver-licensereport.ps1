############
# SETTINGS #
############
$workingDirectory = 'C:\Office365_LicensingReport\'
$date = Get-Date
$reportFile = $workingDirectory + 'prefix-' + $date.Ticks + '.xlsx'

# Office 365 Credentials
$office365Login = 'username@domain.onmicrosoft.com'
$office365Password = 'yourpassword'

# E-mail settings
$mailSender = 'sender@yourdomain.onmicrosoft.com'
$mailPassword = 'yourpassword'
$mailSmtpServer = 'smtp.office365.com'
$mailSmtpPort = '587'
$mailRecipients = "recipient1@domain.com", "recipient2@domain.com"
$mailCc = "cc@domain.com"

###############################
# Do not touch the code below #
###############################

function Should-Skip($skuPartNumber)
{
    If ($license.SkuPartNumber -ne 'PROJECTESSENTIALS' -and $license.SkuPartNumber -ne 'PROJECTONLINE_PLAN_2')
    {
        return $true
    }
    else
    {
        return $false
    }
}

function Get-LicenseName($sku_part_number)
{
    switch ($sku_part_number)
    {
        'PROJECTESSENTIALS'
        {
            return 'Project Lite'
        }
        'PROJECTONLINE_PLAN_1'
        {
            return 'Project Online (Plan 1)'
        }
        'PROJECTONLINE_PLAN_2'
        {
            return 'Project Online (Plan 2)'
        }
    }
    return $sku_part_number
}

function Excel-AutoFitColumn($worksheet)
{
    $usedRange = $worksheet.UsedRange						
    $usedRange.EntireColumn.AutoFit() | Out-Null
}

#Cleanup
$removeFilesFilter = $workingDirectory + '*.xlsx'
Remove-Item $removeFilesFilter

# Connect to Microsoft Online
Import-Module MSOnline

$office365SecurePassword = ConvertTo-SecureString -String $office365Password -AsPlainText -Force
$office365Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $office365Login, $office365SecurePassword

try
{
    Connect-MsolService -Credential $office365Credentials -ErrorAction Stop
}
catch
{
    Write-Error $_.Exception.Message
    exit
}

# Get a list of all licences that exist within the tenant
$licensetype = Get-MsolAccountSku | Where {$_.ConsumedUnits -ge 0}

$excel = New-Object -ComObject excel.application
$excel.visible = $false

# Overview
$workbook = $excel.Workbooks.Add()
$currentWorksheet = $workbook.Worksheets.Item(1)
$currentWorksheet.Name = 'Overview'
$currentWorksheet.Cells.Item(1, 1) = 'License'
$currentWorksheet.Cells.Item(1, 1).Font.Bold = $true
$currentWorksheet.Cells.Item(1, 1).Interior.ColorIndex = 48
$currentWorksheet.Cells.Item(1, 2) = 'Active Units'
$currentWorksheet.Cells.Item(1, 2).Font.Bold = $true
$currentWorksheet.Cells.Item(1, 2).Interior.ColorIndex = 48
$currentWorksheet.Cells.Item(1, 3) = 'Warning Units'
$currentWorksheet.Cells.Item(1, 3).Font.Bold = $true
$currentWorksheet.Cells.Item(1, 3).Interior.ColorIndex = 48
$currentWorksheet.Cells.Item(1, 4) = 'Consumed Units'
$currentWorksheet.Cells.Item(1, 4).Font.Bold = $true
$currentWorksheet.Cells.Item(1, 4).Interior.ColorIndex = 48
$overviewRow = 1

$worksheetIndex = 1
ForEach ($license in $licensetype) 
{
    # Check if we want to include in report.
    If (Should-Skip($license.SkuPartNumber))
    {
        Continue
    }

    # Print statistics
    $overviewRow++
    $currentWorksheet = $workbook.Worksheets.Item(1)
    $currentWorksheet.Cells.Item($overviewRow,1) = Get-LicenseName($license.SkuPartNumber)
    $currentWorksheet.Cells.Item($overviewRow,2) = $license.ActiveUnits
    $currentWorksheet.Cells.Item($overviewRow,3) = $license.WarningUnits
    $currentWorksheet.Cells.Item($overviewRow,4) = $license.ConsumedUnits
    Excel-AutoFitColumn($currentWorksheet)

    # Create new sheet
    $Workbook.Worksheets.Add([System.Reflection.Missing]::Value,$Workbook.Worksheets.Item($Workbook.Worksheets.count))
    $worksheetIndex++
    $currentWorksheet = $workbook.Worksheets.Item($worksheetIndex)
    $currentWorksheet.Name = Get-LicenseName($license.SkuPartNumber)
    
    $currentWorksheet.Cells.Item(1, 1) = 'Login'
    $currentWorksheet.Cells.Item(1, 1).Font.Bold = $true
    $currentWorksheet.Cells.Item(1, 1).Interior.ColorIndex = 48
    $currentWorksheet.Cells.Item(1, 2) = 'Name'
    $currentWorksheet.Cells.Item(1, 2).Font.Bold = $true
    $currentWorksheet.Cells.Item(1, 2).Interior.ColorIndex = 48

    # Process all users
    $users = Get-MsolUser -all | where {$_.isLicensed -and $_.licenses[0].accountskuid.tostring() -eq $license.accountskuid}
    $row = 2
    ForEach ($user in $users)
    {
        $currentWorksheet.Cells.Item($row, 1) = $user.UserPrincipalName
        $currentWorksheet.Cells.Item($row, 2) = $user.DisplayName
        
        $row++
    }

    # Post-Formatting
    Excel-AutoFitColumn($currentWorksheet)
}

# Save
$workbook.SaveAs($reportFile)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$excel) | Out-Null

Start-Sleep -s 5

$mailSecurePassword = ConvertTo-SecureString -String $mailPassword -AsPlainText -Force
$mailCredentials = New-Object System.Management.Automation.PSCredential ($mailSender, $mailSecurePassword)

Send-MailMessage -To $mailRecipients -Cc $mailCc -Attachments $reportFile -SmtpServer $mailSmtpServer -Credential $mailCredentials -UseSsl "Project Server - Office365 License Report" -Port $mailSmtpPort -Body "<p>Dear client,</p><p>This is an automatically generated message.<br/>Check the attachments for your generated weekly report.</p><p>Kind regards,</p><p><strong>ServiceDesk ProjectIn</strong><br/>servicedesk@projectin.be</p>" -From $mailSender -BodyAsHtml