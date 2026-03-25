<#
.SYNOPSIS
    Sends a scheduled report via email with CSV attachment.

.DESCRIPTION
    Executes a stored procedure, calculates summary statistics,
    generates a styled HTML email with gradient summary cards and
    data tables, attaches a CSV of the full dataset, and sends
    via SMTP.

    Follows the standard reporting template pattern:
    - Stored procedure returns data only (no HTML, no email)
    - PowerShell handles all formatting, statistics, and delivery

.EXAMPLE
    .\Send-YourReport.ps1

.NOTES
    Author:  Your Name
    Version: 1.0
    Created: yyyy-MM-dd

    Requirements:
    - dbatools PowerShell module (auto-installed if missing)
    - Custom mail module (update Import-Module path below)
    - SQL credentials stored securely on disk
    - Stored procedure deployed: dbo.usp_YourReportName

    Scheduling:
    - Use Windows Task Scheduler or SQL Agent
    - Action: powershell.exe -ExecutionPolicy Bypass -File "C:\Scripts\Send-YourReport.ps1"
#>

# ============================================
# Module Imports
# ============================================

# Import your custom mail/utility module (update path)
# Import-Module -Name "C:\Scripts\Modules\yourmodule.psm1" -Verbose

# Check if dbatools module is installed, and install it if missing
if (!(Get-Module -ListAvailable -Name dbatools)) {
    Write-Host "dbatools module does not exist, installing module to scope: CurrentUser" -ForegroundColor Yellow
    Install-Module -Name dbatools -Scope CurrentUser -Force
    Write-Host "✓ dbatools module installed successfully" -ForegroundColor Green
}

# Import dbatools module
Import-Module -Name dbatools -Verbose

# ============================================
# Configuration
# ============================================

# -- SQL Connection --
$serverName = "YOUR-SQL-SERVER"
$sqlDatabase = "YourDatabase"
$sqlInstance = "YOUR-SQL-SERVER.domain.local"
$sqlUsername = "svc_YourServiceAccount"
$sqlPassword = Get-Content "C:\Scripts\Secure\YourPassword.txt" | ConvertTo-SecureString -Key (Get-Content "C:\Scripts\Secure\YourAesKey.txt")
$sqlCredential = New-Object System.Management.Automation.PSCredential ($sqlUsername, $sqlPassword)

# -- Email --
$recipientEmail = "recipient@yourdomain.com"
$senderEmail = "reporting@yourdomain.com"
$alertsSender = "alerts@yourdomain.com"
$alertsEmail = "alerts@yourdomain.com"

# -- Report --
$reportName = "Your Report Name"
$storedProc = "EXEC dbo.usp_YourReportName"
# If your proc takes parameters:
# $storedProc = "EXEC dbo.usp_YourReportName @Year = @Year, @Month = @Month"
# $sqlParams = @{ Year = 2026; Month = 3 }

# ============================================
# Report Date Info
# ============================================
$today = Get-Date
$reportDate = $today.ToString('dd/MM/yyyy')

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "$reportName" -ForegroundColor Cyan
Write-Host "Report Date: $reportDate" -ForegroundColor Cyan
Write-Host "Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# CSV File Path
$csvFileName = "YourReport_$(Get-Date -Format 'yyyy-MM-dd').csv"
$csvFilePath = "$env:TEMP\$csvFileName"

# ============================================
# Execute SQL Query
# ============================================
Write-Host "`nExecuting stored procedure..." -ForegroundColor Cyan

try {
    # Basic execution (no parameters)
    $sqlResults = Invoke-DbaQuery -SqlInstance $sqlInstance `
        -Database $sqlDatabase `
        -Query $storedProc `
        -SqlCredential $sqlCredential `
        -EnableException `
        -QueryTimeout 120

    # If your proc takes parameters, use this instead:
    # $sqlResults = Invoke-DbaQuery -SqlInstance $sqlInstance `
    #                                -Database $sqlDatabase `
    #                                -Query $storedProc `
    #                                -SqlParameters $sqlParams `
    #                                -SqlCredential $sqlCredential `
    #                                -EnableException `
    #                                -QueryTimeout 120

    Write-Host "✓ Retrieved $($sqlResults.Count) records" -ForegroundColor Green

}
catch {
    Write-Host "✗ SQL execution failed: $($_.Exception.Message)" -ForegroundColor Red

    $errorSubject = "$reportName Failed - SQL Error"
    $errorBody = @"
The $reportName failed during SQL execution.

Error: $($_.Exception.Message)
Time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
Server: $serverName
Database: $sqlDatabase

Please investigate and re-run the report.
"@

    # Send alert email (update to match your mail function)
    # Send-YourMailFunction -From $alertsSender -To $alertsEmail -Subject $errorSubject -Body $errorBody
    Send-MailMessage -From $alertsSender -To $alertsEmail -Subject $errorSubject -Body $errorBody -SmtpServer "your-smtp-server"
    exit 1
}

# ============================================
# Calculate Statistics
# ============================================
$totalRecords = $sqlResults.Count

# -- Customise these to match your data columns --
# Example: numeric aggregations
# $totalValue = [math]::Round(($sqlResults.YourNumericColumn | Measure-Object -Sum).Sum, 1)
# $avgValue   = if ($totalRecords -gt 0) { [math]::Round(($sqlResults.YourNumericColumn | Measure-Object -Average).Average, 1) } else { 0 }

# Example: group/category distribution
# $categoryGroups = $sqlResults | Group-Object -Property YourCategoryColumn | Sort-Object -Property Count -Descending

Write-Host "`nReport Statistics:" -ForegroundColor Cyan
Write-Host "  Total Records: $totalRecords" -ForegroundColor White
# Write-Host "  Total Value: $totalValue" -ForegroundColor White
# Write-Host "  Categories: $($categoryGroups.Count)" -ForegroundColor White

# ============================================
# Generate HTML Email
# ============================================
Write-Host "`nGenerating HTML email..." -ForegroundColor Cyan

# Customise card values to match your statistics
$card1Label = "TOTAL RECORDS"; $card1Value = $totalRecords
$card2Label = "METRIC 2"; $card2Value = "0"
$card3Label = "METRIC 3"; $card3Value = "0"
$card4Label = "METRIC 4"; $card4Value = "0"

$emailBody = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!--[if mso]>
    <style type="text/css">
        table { border-collapse: collapse; }
        .gradient-header { background-color: #2563eb !important; }
        .gradient-card-1 { background-color: #2563eb !important; }
        .gradient-card-2 { background-color: #10b981 !important; }
        .gradient-card-3 { background-color: #6366f1 !important; }
        .gradient-card-4 { background-color: #f59e0b !important; }
    </style>
    <![endif]-->
</head>
<body style="margin: 0; padding: 0; font-family: 'Segoe UI', Arial, sans-serif; background-color: #f5f5f5;">

    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: #f5f5f5; padding: 20px 0;">
        <tr>
            <td align="center">
                <table width="900" cellpadding="0" cellspacing="0" border="0" style="background-color: #ffffff; border-radius: 10px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">

                    <!-- Header with Gradient -->
                    <tr>
                        <td class="gradient-header" style="background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%); background-color: #2563eb; padding: 35px 30px; text-align: center; border-radius: 10px 10px 0 0;">
                            <h1 style="margin: 0; color: #ffffff; font-size: 32px; font-weight: 700; letter-spacing: -0.5px;">📊 $reportName</h1>
                            <p style="margin: 12px 0 0 0; color: #dbeafe; font-size: 18px; font-weight: 500;">Report Date: $reportDate</p>
                        </td>
                    </tr>

                    <!-- Summary Cards -->
                    <tr>
                        <td style="padding: 25px; background-color: #f8f9fa;">
                            <table width="100%" cellpadding="8" cellspacing="8" border="0">
                                <tr>
                                    <!-- Card 1 - Blue -->
                                    <td width="23%" align="center" class="gradient-card-1" style="background: linear-gradient(135deg, #2563eb 0%, #1e40af 100%); background-color: #2563eb; border-radius: 10px; padding: 22px 15px; box-shadow: 0 2px 4px rgba(37,99,235,0.3);">
                                        <p style="margin: 0 0 10px 0; color: #dbeafe; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px;">$card1Label</p>
                                        <p style="margin: 0; color: #ffffff; font-size: 36px; font-weight: 700; line-height: 1;">$card1Value</p>
                                    </td>

                                    <!-- Card 2 - Green -->
                                    <td width="23%" align="center" class="gradient-card-2" style="background: linear-gradient(135deg, #10b981 0%, #059669 100%); background-color: #10b981; border-radius: 10px; padding: 22px 15px; box-shadow: 0 2px 4px rgba(16,185,129,0.3);">
                                        <p style="margin: 0 0 10px 0; color: #d1fae5; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px;">$card2Label</p>
                                        <p style="margin: 0; color: #ffffff; font-size: 36px; font-weight: 700; line-height: 1;">$card2Value</p>
                                    </td>

                                    <!-- Card 3 - Indigo -->
                                    <td width="23%" align="center" class="gradient-card-3" style="background: linear-gradient(135deg, #6366f1 0%, #4f46e5 100%); background-color: #6366f1; border-radius: 10px; padding: 22px 15px; box-shadow: 0 2px 4px rgba(99,102,241,0.3);">
                                        <p style="margin: 0 0 10px 0; color: #e0e7ff; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px;">$card3Label</p>
                                        <p style="margin: 0; color: #ffffff; font-size: 36px; font-weight: 700; line-height: 1;">$card3Value</p>
                                    </td>

                                    <!-- Card 4 - Amber -->
                                    <td width="23%" align="center" class="gradient-card-4" style="background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); background-color: #f59e0b; border-radius: 10px; padding: 22px 15px; box-shadow: 0 2px 4px rgba(245,158,11,0.3);">
                                        <p style="margin: 0 0 10px 0; color: #fef3c7; font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px;">$card4Label</p>
                                        <p style="margin: 0; color: #ffffff; font-size: 36px; font-weight: 700; line-height: 1;">$card4Value</p>
                                    </td>
                                </tr>
                            </table>
                        </td>
                    </tr>

                    <!-- ================================================ -->
                    <!-- OPTIONAL: Summary/Distribution Table              -->
                    <!-- Uncomment and customise if you need a breakdown   -->
                    <!-- ================================================ -->
                    <!--
                    <tr>
                        <td style="padding: 25px 30px;">
                            <h2 style="margin: 0 0 18px 0; color: #1f2937; font-size: 20px; font-weight: 700;">📊 Distribution Breakdown</h2>

                            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse: collapse; background-color: #ffffff; border-radius: 6px; overflow: hidden; border: 1px solid #e5e7eb;">
                                <thead>
                                    <tr style="background-color: #f9fafb;">
                                        <th style="padding: 14px 12px; text-align: left; font-size: 12px; font-weight: 700; color: #6b7280; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #e5e7eb;">Category</th>
                                        <th style="padding: 14px 12px; text-align: center; font-size: 12px; font-weight: 700; color: #6b7280; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #e5e7eb;">Count</th>
                                        <th style="padding: 14px 12px; text-align: center; font-size: 12px; font-weight: 700; color: #6b7280; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #e5e7eb;">Percentage</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    DYNAMIC_DISTRIBUTION_ROWS_HERE
                                </tbody>
                            </table>
                        </td>
                    </tr>
                    -->

                    <!-- Detail Data Table -->
                    <tr>
                        <td style="padding: 0 30px 30px 30px;">
                            <h2 style="margin: 0 0 18px 0; color: #1f2937; font-size: 20px; font-weight: 700;">📋 Detail Data</h2>

                            <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse: collapse; background-color: #ffffff; border-radius: 6px; overflow: hidden; border: 1px solid #e5e7eb;">
                                <thead>
                                    <tr style="background-color: #f9fafb;">
                                        <!-- Customise column headers to match your stored proc output -->
                                        <th style="padding: 12px 8px; text-align: left; font-size: 11px; font-weight: 700; color: #6b7280; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #e5e7eb;">Column 1</th>
                                        <th style="padding: 12px 8px; text-align: left; font-size: 11px; font-weight: 700; color: #6b7280; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #e5e7eb;">Column 2</th>
                                        <th style="padding: 12px 8px; text-align: center; font-size: 11px; font-weight: 700; color: #6b7280; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #e5e7eb;">Column 3</th>
                                        <th style="padding: 12px 8px; text-align: right; font-size: 11px; font-weight: 700; color: #6b7280; text-transform: uppercase; letter-spacing: 0.5px; border-bottom: 2px solid #e5e7eb;">Column 4</th>
                                    </tr>
                                </thead>
                                <tbody>
"@

# ============================================
# Build Detail Table Rows
# ============================================
# Customise to match your stored proc output columns

foreach ($row in $sqlResults) {
    # Handle nulls/DBNull for each column as needed
    # $value = if ($row.YourColumn -is [DBNull] -or $null -eq $row.YourColumn) { "0.00" } else { $row.YourColumn.ToString('N2') }
    # $dateValue = if ([string]::IsNullOrWhiteSpace($row.YourDateColumn)) { "-" } else { $row.YourDateColumn }

    $emailBody += @"
                                    <tr style="border-bottom: 1px solid #f3f4f6;">
                                        <td style="padding: 10px 8px; font-size: 12px; color: #6b7280;">$($row.Column1)</td>
                                        <td style="padding: 10px 8px; font-size: 12px; font-weight: 600; color: #1f2937;">$($row.Column2)</td>
                                        <td style="padding: 10px 8px; text-align: center; font-size: 12px; color: #374151;">$($row.Column3)</td>
                                        <td style="padding: 10px 8px; text-align: right; font-size: 12px; font-weight: 700; color: #374151;">$($row.Column4)</td>
                                    </tr>
"@
}

# ============================================
# Optional: Build Distribution Table Rows
# ============================================
# Uncomment this block and the distribution table HTML above if needed
<#
$distributionRows = ""
foreach ($group in $categoryGroups) {
    $groupName  = if ([string]::IsNullOrWhiteSpace($group.Name)) { "Not Set" } else { $group.Name }
    $percentage = [math]::Round($group.Count / $totalRecords * 100, 1)

    $distributionRows += @"
                                    <tr style="border-bottom: 1px solid #f3f4f6;">
                                        <td style="padding: 14px 12px; font-size: 14px; color: #1f2937; font-weight: 600;">$groupName</td>
                                        <td style="padding: 14px 12px; text-align: center; font-size: 16px; font-weight: 700; color: #374151;">$($group.Count)</td>
                                        <td style="padding: 14px 12px; text-align: center; font-size: 14px; color: #6b7280; font-weight: 500;">$percentage%</td>
                                    </tr>
"@
}
# Then replace the placeholder in $emailBody:
# $emailBody = $emailBody -replace 'DYNAMIC_DISTRIBUTION_ROWS_HERE', $distributionRows
#>

$emailBody += @"
                                </tbody>
                            </table>
                        </td>
                    </tr>

                    <!-- Footer -->
                    <tr>
                        <td style="padding: 20px 30px; background-color: #f9fafb; border-top: 1px solid #e5e7eb; border-radius: 0 0 10px 10px;">
                            <p style="margin: 0 0 10px 0; color: #6b7280; font-size: 13px; line-height: 1.6;"><strong style="color: #374151;">📎 Full Dataset:</strong> Complete details attached as CSV file.</p>
                            <p style="margin: 0 0 10px 0; color: #6b7280; font-size: 13px; line-height: 1.6;"><strong style="color: #374151;">📊 Usage:</strong> Import CSV into Excel or Power BI for detailed analysis.</p>
                            <p style="margin: 0; color: #9ca3af; font-size: 12px;">This is an automated report. Generated $(Get-Date -Format 'dd/MM/yyyy HH:mm').</p>
                        </td>
                    </tr>

                </table>
            </td>
        </tr>
    </table>
</body>
</html>
"@

# ============================================
# Export CSV and Send Email
# ============================================
Write-Host "Exporting CSV to: $csvFilePath" -ForegroundColor Cyan
$sqlResults | Export-Csv -Path $csvFilePath -NoTypeInformation

$emailSubject = "$reportName ($reportDate)"
$Attachment = $csvFilePath

Write-Host "Sending email to: $recipientEmail" -ForegroundColor Cyan

# Update to match your mail function:
# Send-YourMailFunction -From $senderEmail -To $recipientEmail -Subject $emailSubject -Body $emailBody -Attachment $Attachment
Send-MailMessage -From $senderEmail -To $recipientEmail -Subject $emailSubject -Body $emailBody -BodyAsHtml -SmtpServer "your-smtp-server" -Attachments $Attachment

Write-Host "✓ Email sent successfully!" -ForegroundColor Green

# ============================================
# File Cleanup with Retry Logic
# ============================================
Write-Host "`nCleaning up CSV file..." -ForegroundColor Cyan

# If you have a Remove-FileWithRetry function in your module, use it:
$retryFunction = Get-Command Remove-FileWithRetry -ErrorAction SilentlyContinue

if ($retryFunction) {
    $success = Remove-FileWithRetry -FilePath $csvFilePath -MaxRetries 3 -RetryDelaySeconds 5

    if (-not $success) {
        $errorEmailSubject = "Failed to Delete Report CSV - $reportName - $reportDate"
        $errorEmailBody = "The script failed to delete the CSV report file at $csvFilePath on $serverName after 3 attempts. Please check the file and delete it manually."
        # Send-YourMailFunction -From $alertsSender -To $alertsEmail -Subject $errorEmailSubject -Body $errorEmailBody
        Send-MailMessage -From $alertsSender -To $alertsEmail -Subject $errorEmailSubject -Body $errorEmailBody -SmtpServer "your-smtp-server"
        Write-Warning "Failed to delete CSV file after 3 attempts. Alert email sent."
    }
    else {
        Write-Host "✓ CSV file cleaned up successfully" -ForegroundColor Green
    }
}
else {
    # Fallback: Simple removal with try-catch
    try {
        if (Test-Path $csvFilePath) {
            Remove-Item $csvFilePath -Force -ErrorAction Stop
            Write-Host "✓ CSV file cleaned up successfully" -ForegroundColor Green
        }
    }
    catch {
        $errorEmailSubject = "Failed to Delete Report CSV - $reportName - $reportDate"
        $errorEmailBody = "The script failed to delete the CSV report file at $csvFilePath on $serverName. Error: $($_.Exception.Message). Please check the file and delete it manually."
        # Send-YourMailFunction -From $alertsSender -To $alertsEmail -Subject $errorEmailSubject -Body $errorEmailBody
        Send-MailMessage -From $alertsSender -To $alertsEmail -Subject $errorEmailSubject -Body $errorEmailBody -SmtpServer "your-smtp-server"
        Write-Warning "Failed to delete CSV file. Alert email sent. Error: $($_.Exception.Message)"
    }
}

Write-Host "`n========================================" -ForegroundColor Green
Write-Host "✓ Report completed successfully!" -ForegroundColor Green
Write-Host "========================================`n" -ForegroundColor Green