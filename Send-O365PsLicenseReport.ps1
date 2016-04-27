################################################
# DEBUGGING SETTINGS, VARIABLES CAN BE CHANGED #
################################################

param(
    [string] $O365Username = 'username@yourdomain.onmicrosoft.com',
    [string] $O365Password = 'yourpassword',
    [string] $FilePrefix = 'prefix',
    [string] $RunDir = 'C:\Path\To\File\Out',
    [string] $MailSender = 'me@example.com',
    [string] $MailPassword = 'password',
    [string] $MailSmtpServer = 'smtp.office365.com',
    [string] $MailSmtpPort='587',
    [string[]] $MailRecipients,
    [string] $MailCc,
    [boolean] $All = $false)

###############################
# Do not touch the code below #
###############################

$WorkingDirectory = Split-Path $script:MyInvocation.MyCommand.Path
$Date = Get-Date
$ReportFile = $RunDir + '\' + $FilePrefix + '-' + $Date.Ticks + '.html'
$LicensesFile = $WorkingDirectory + '\licenses.xml'
$LicensesConfigFile = $RunDir + '\licenses_config.xml'
$CssFile = $WorkingDirectory + '\style.css'

function Should-Skip($SkuPartNumber)
{
    if ($All -eq $true)
    {
        return $false
    }

    [xml]$LicensesConfigContent = Get-Content $LicensesConfigFile
    foreach ($LicenseConfig in $LicensesConfigContent.Licenses.License)
    {
        if ($SkuPartNumber -eq $LicenseConfig.SkuPartNumber)
        {
            return $false
        }
    }
    return $true
}

function Get-LicenseName($SkuPartNumber)
{
    [xml]$LicensesFileContent = Get-Content $LicensesFile
    foreach ($Item in $LicensesFileContent.Licenses.License)
    {
        if ($SkuPartNumber -eq $Item.SkuPartNumber)
        {
            return $Item.Name
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

$Html = '<html><head><title>Office 365 License Report</title></head><body>'
$Html += '<style>'
$Html += Get-Content $CssFile
$Html += '</style>'

$HtmlOverview = '<h1>Overview</h1>'
$HtmlOverview += '<table>'
$HtmlOverview += '<tr><th>License</th><th>Active Units</th><th>Warning Units</th><th>Consumed Units</th></tr>'

# Get a list of all licences that exist within the tenant
$Licenses = Get-MsolAccountSku | Where {$_.ConsumedUnits -ge 0}
$HtmlLicenseDetails = New-Object string[] $Licenses.Count

$HtmlLicenseDetailsIndex = -1

ForEach ($License in $Licenses) 
{
    # Check if we want to include in report.
    If (Should-Skip($License.SkuPartNumber))
    {
        Continue
    }
    $HtmlLicenseDetailsIndex++

    # Print statistics
    $LicenseName = Get-LicenseName($License.SkuPartNumber)
    $HtmlOverview += '<tr><td>' + $LicenseName + '</td><td>' + $License.ActiveUnits + '</td><td>' + $License.WarningUnits + '</td><td>' + $License.ConsumedUnits + '</td></tr>'

    # Create new sheet
    $HtmlLicenseDetails[$HtmlLicenseDetailsIndex] = '<h1>' + $LicenseName + '</h1>'
    $HtmlLicenseDetails[$HtmlLicenseDetailsIndex] += '<table>'
    $HtmlLicenseDetails[$HtmlLicenseDetailsIndex] += '<tr><th>#</th><th>Login</th><th>Name</th></tr>'

    # Process all users
    $Users = Get-MsolUser -all | where {$_.isLicensed -and $_.licenses[0].accountskuid.tostring() -eq $License.accountskuid}

    $Index = 0
    ForEach ($User in $Users)
    {
        $Index++
        $HtmlLicenseDetails[$HtmlLicenseDetailsIndex] += '<tr><td>' + $Index + '</td><td>' + $User.UserPrincipalName + '</td><td>' + $User.DisplayName + '</td></tr>'
    }

    $HtmlLicenseDetails[$HtmlLicenseDetailsIndex] += '</table>'
}

$HtmlOverview += '</table>'
$Html += $HtmlOverview
foreach ($HtmlLicenseDetail in $HtmlLicenseDetails)
{
    $Html += $HtmlLicenseDetail
}
$Html += '</body></html>'

$Html | Out-File $ReportFile

# Required, otherwise file may not be saved.
Start-Sleep -s 5

$MailSecurePassword = ConvertTo-SecureString -String $MailPassword -AsPlainText -Force
$MailCredentials = New-Object System.Management.Automation.PSCredential ($MailSender, $MailSecurePassword)

Send-MailMessage -To $MailRecipients -Cc $MailCc -Attachments $ReportFile -SmtpServer $MailSmtpServer -Credential $MailCredentials -UseSsl "Project Server - Office365 License Report" -Port $MailSmtpPort -Body "<p>Dear client,</p><p>This is an automatically generated message.<br/>Check the attachments for your generated weekly report.</p><p>Kind regards,</p><p><strong>ServiceDesk ProjectIn</strong><br/>servicedesk@projectin.be</p>" -From $MailSender -BodyAsHtml