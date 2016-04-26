################################################
# DEBUGGING SETTINGS, VARIABLES CAN BE CHANGED #
################################################

param(
    [string] $O365Username = 'username@yourdomain.onmicrosoft.com',
    [string] $O365Password = 'password',
    [string] $FilePrefix = 'prefix',
    [string] $FileOutDir = 'C:\Path\To\File\Out',
    [string] $MailSender = 'me@example.com',
    [string] $MailPassword = 'password',
    [string] $MailSmtpServer = 'smtp.office365.com',
    [string] $MailSmtpPort='587',
    [string[]] $MailRecipients,
    [string] $MailCc)

###############################
# Do not touch the code below #
###############################

$WorkingDirectory = Split-Path $invocation.MyCommand.Path
$Date = Get-Date
$ReportFile = $FileOutDir + '\' + $FilePrefix + '-' + $Date.Ticks + '.html'

function Should-Skip($SkuPartNumber)
{
    If ($License.SkuPartNumber -ne 'PROJECTESSENTIALS' -and $License.SkuPartNumber -ne 'PROJECTONLINE_PLAN_2')
    {
        return $true
    }
    else
    {
        return $false
    }
}

function Get-LicenseName($SkuPartNumber)
{
    switch ($SkuPartNumber)
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
    return $SkuPartNumber
}

#Cleanup
# $removeFilesFilter = $WorkingDirectory + '*.xlsx'
# Remove-Item $removeFilesFilter

# Connect to Microsoft Online
Import-Module MSOnline
$O365SecurePassword = ConvertTo-SecureString -String $O365Password -AsPlainText -Force
$O365Credentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $O365Username, $O365SecurePassword
try
{
    Connect-MsolService -Credential $O365Credentials -ErrorAction Stop
}
catch
{
    Write-Error $_.Exception.Message
    exit
}

# Get a list of all licences that exist within the tenant
$LicenseType = Get-MsolAccountSku | Where {$_.ConsumedUnits -ge 0}

$Html = '<html><head><title>Office 365 License Report</title></head><body>'
$Html += '<style>'
$CssFile = $WorkingDirectory + '\style.css'
$Html += Get-Content $CssFile
$Html += '</style>'

$HtmlOverview = '<h1>Overview</h1>'
$HtmlOverview += '<table>'
$HtmlOverview += '<tr><th>License</th><th>Active Units</th><th>Warning Units</th><th>Consumed Units</th></tr>'

$HtmlLicenseDetails = New-Object string[] $LicenseType.Count
$Index = -1

ForEach ($License in $LicenseType) 
{
    # Check if we want to include in report.
    If (Should-Skip($License.SkuPartNumber))
    {
        Continue
    }
    $Index++

    # Print statistics
    $LicenseName = Get-LicenseName($License.SkuPartNumber)
    $HtmlOverview += '<tr><td>' + $LicenseName + '</td><td>' + $License.ActiveUnits + '</td><td>' + $License.WarningUnits + '</td><td>' + $License.ConsumedUnits + '</td></tr>'

    # Create new sheet
    $HtmlLicenseDetails[$Index] = '<h1>' + $LicenseName + '</h1>'
    $HtmlLicenseDetails[$Index] += '<table>'
    $HtmlLicenseDetails[$Index] += '<tr><th>Login</th><th>Name</th></tr>'

    # Process all users
    $Users = Get-MsolUser -all | where {$_.isLicensed -and $_.licenses[0].accountskuid.tostring() -eq $License.accountskuid}
    ForEach ($User in $Users)
    {
        $HtmlLicenseDetails[$Index] += '<tr><td>' + $User.UserPrincipalName + '</td><td>' + $User.DisplayName + '</td></tr>'
    }

    $HtmlLicenseDetails[$Index] += '</table>'
}

$HtmlOverview += '</table>'
$Html += $HtmlOverview
foreach ($HtmlLicenseDetail in $HtmlLicenseDetails)
{
    $Html += $HtmlLicenseDetail
}
$Html += '</body></html>'

$Html | Out-File $ReportFile

Start-Sleep -s 5

$MailSecurePassword = ConvertTo-SecureString -String $MailPassword -AsPlainText -Force
$MailCredentials = New-Object System.Management.Automation.PSCredential ($MailSender, $MailSecurePassword)

Send-MailMessage -To $MailRecipients -Cc $MailCc -Attachments $ReportFile -SmtpServer $MailSmtpServer -Credential $MailCredentials -UseSsl "Project Server - Office365 License Report" -Port $MailSmtpPort -Body "<p>Dear client,</p><p>This is an automatically generated message.<br/>Check the attachments for your generated weekly report.</p><p>Kind regards,</p><p><strong>ServiceDesk ProjectIn</strong><br/>servicedesk@projectin.be</p>" -From $MailSender -BodyAsHtml