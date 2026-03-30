<#
PowerShell Utility Module.

List of functions in this module:
- Send-SMTPMail
- Set-AlternatingRows
- Convert-SQLtoHTML
#>

function Send-SMTPMail {
  <#
    .SYNOPSIS
    Sends an email using configurable SMTP server settings.

    .DESCRIPTION
    The Send-SMTPMail function sends an email using the specified SMTP server. It allows you to specify the SMTP server, port, sender email address, recipient email address, subject, body, and an optional attachment.

    .PARAMETER SmtpServer
    The SMTP server to use for sending the email. The default value is "contoso-com.mail.protection.outlook.com".

    .PARAMETER SmtpPort
    The port number to use for the SMTP server. The default value is 25 and should not be changed.

    .PARAMETER SenderEmail
    The email address of the sender. The default value is "notifications@example.com" but can be changed.

    .PARAMETER RecipientEmail
    The email address of the recipient. This parameter is mandatory.

    .PARAMETER Subject
    The subject of the email. This parameter is mandatory.

    .PARAMETER Body
    The body of the email. This parameter is mandatory.

    .PARAMETER Attachment
    The path to supply optional file attachment(s). Can supply multiple attachments using comma separated paths.  eg "C:\Temp\upload.csv" or "C:\Temp\upload.csv,C:\Temp\upload2.csv"

    .EXAMPLE
    Send-SMTPMail -RecipientEmail "john@example.com" -Subject "Test Email" -Body "This is a test email"

    This example sends a test email to the recipient "john@example.com" with the subject "Test Email" and the body "This is a test email."
    #>

  param (
    [string]$SmtpServer = "contoso-com.mail.protection.outlook.com",
    [int]$SmtpPort = 25,
    [string]$SenderEmail = "notifications@example.com",

    [Parameter(Mandatory = $true)]
    [string]$RecipientEmail,

    [Parameter(Mandatory = $true)]
    [string]$Subject,

    [Parameter(Mandatory = $true)]
    [string]$Body,

    [Parameter(Mandatory = $false)]
    [string]$Attachment
  )

  # Create the mail message object
  $mailMessage = New-Object System.Net.Mail.MailMessage
  $mailMessage.From = $SenderEmail
  $mailMessage.To.Add($RecipientEmail)
  $mailMessage.Subject = $Subject
  $mailMessage.Body = $Body
  $mailMessage.IsBodyHtml = $true

  if ($Attachment) {
    # Create the attachment object(s) and add to the mail message
    $AttachmentPathArray = $Attachment -split ","
    try {
      foreach ($Attachment in $AttachmentPathArray) {
        $msgAttachment = New-Object System.Net.Mail.Attachment($Attachment)
        $mailMessage.Attachments.Add($msgAttachment)
      }
    }
    catch {
      Write-Host "Failed to attach file." -ForegroundColor Red
      write-host "$_" -ForegroundColor Yellow
    }
  }

  # Create the SMTP client object
  $smtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
  $smtpClient.Timeout = 5000
  $smtpClient.EnableSsl = $true

  try {
    # Send the email
    $smtpClient.Send($mailMessage)
    Write-Host "Email sent successfully." -ForegroundColor Green
  }
  catch {
    # Handle any errors that occur
    Write-Host "Email failed to send." -ForegroundColor Red
    Write-Host "$_" -ForegroundColor Yellow
    exit
  }
}

Function Set-AlternatingRows {
  <#
    .SYNOPSIS
      Simple function to alternate the row colors in an HTML table
    .DESCRIPTION
      This function accepts pipeline input from ConvertTo-HTML or any
      string with HTML in it.  It will then search for <tr> and replace 
      it with <tr class=(something)>.  With the combination of CSS it
      can set alternating colors on table rows.
      
      CSS requirements:
      .odd  { background-color:#ffffff; }
      .even { background-color:#dddddd; }
      
      Classnames can be anything and are configurable when executing the
      function.  Colors can, of course, be set to your preference.
      
      This function does not add CSS to your report, so you must provide
      the style sheet, typically part of the ConvertTo-HTML cmdlet using
      the -Head parameter.
    .PARAMETER Line
      String containing the HTML line, typically piped in through the
      pipeline.
    .PARAMETER CSSEvenClass
      Define which CSS class is your "even" row and color.
    .PARAMETER CSSOddClass
      Define which CSS class is your "odd" row and color.
    .EXAMPLE $Report | ConvertTo-HTML -Head $Header | Set-AlternateRows -CSSEvenClass even -CSSOddClass odd | Out-File HTMLReport.html
    
      $Header can be defined with a here-string as:
      $Header = @"
      <style>
      TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
      TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
      TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
      .odd  { background-color:#ffffff; }
      .even { background-color:#dddddd; }
      </style>
      "@
      
      This will produce a table with alternating white and grey rows.  Custom CSS
      is defined in the $Header string and included with the table thanks to the -Head
      parameter in ConvertTo-HTML.
    .NOTES
      Author:         Martin Pugh
      Twitter:        @thesurlyadm1n
      Spiceworks:     Martin9700
      Blog:           www.thesurlyadmin.com
      
      Changelog:
        1.1         Modified replace to include the <td> tag, as it was changing the class for the TH row as well.
        1.0         Initial function release
    .LINK
      http://community.spiceworks.com/scripts/show/1745-set-alternatingrows-function-modify-your-html-table-to-have-alternating-row-colors
      .LINK
          http://thesurlyadmin.com/2013/01/21/how-to-create-html-reports/
    #>

  [CmdletBinding()]
  Param(
    [Parameter(Mandatory, ValueFromPipeline)]
    [string]$Line,
       
    [Parameter(Mandatory)]
    [string]$CSSEvenClass,
       
    [Parameter(Mandatory)]
    [string]$CSSOddClass
  )
  Begin {
    $ClassName = $CSSEvenClass
  }
  Process {
    If ($Line.Contains("<tr><td>")) {
      $Line = $Line.Replace("<tr>", "<tr class=""$ClassName"">")
      If ($ClassName -eq $CSSEvenClass) {
        $ClassName = $CSSOddClass
      }
      Else {
        $ClassName = $CSSEvenClass
      }
    }
    Return $Line
  }
}

function Convert-SQLtoHTML {
  <#
    .SYNOPSIS
    Converts SQL query results (table, Sproc or Function) to HTML format.

    .DESCRIPTION
    The Convert-SQLtoHTML function connects to a SQL Server instance, executes the specified SQL query, and converts the results to HTML format. The HTML output includes a customizable report title and CSS styling.

    .PARAMETER sqlInstance
    The SQL Server instance to connect to. The default value is "sqlserver-01.corp.local".

    .PARAMETER showFooter
    Whether to include the date and time the report was generated. The default value is $true.

    .PARAMETER sqlQuery
    The SQL query to execute. This parameter is mandatory. Can be a Query or Stored Procedure.

    .PARAMETER sqlUsername
    The username to authenticate with the SQL Server. This parameter is mandatory.

    .PARAMETER sqlPassword
    The password to authenticate with the SQL Server. This parameter is mandatory and should be provided as a secure string. Plain text is not accepted.

    .PARAMETER reportTitle
    The title of the HTML report. This parameter is mandatory.

    .EXAMPLE
    Convert-SQLtoHTML -sqlQuery "SELECT * FROM People" -sqlUsername "admin" -sqlPassword $securePassword -reportTitle "People List Report"
    Connects to the default SQL Server instance, executes the specified SQL query, and generates an HTML report titled "People List Report".

    #>
  param (
    [string]$sqlInstance = "sqlserver-01.corp.local",
        
    [Boolean]$showFooter = $true,

    [Parameter(Mandatory = $true)]
    [string]$sqlQuery,

    [Parameter(Mandatory = $true)]
    [string]$sqlUsername,

    [Parameter(Mandatory = $true)]
    [securestring]$sqlPassword,

    [Parameter(Mandatory = $true)]
    [string]$reportTitle
  )

  # Check if the dbatools module is installed, and install it if missing
  if (!(Get-Module -ListAvailable -Name dbatools)) {
    Write-Host "Module does not exist, installing module to scope: CurrentUser"
    Install-Module -Name dbatools -Scope CurrentUser -Force
  }

  # User Variables
  $sqlCredential = New-Object PsCredential $sqlUsername, $sqlPassword

  # Connect to SQL Server and run the query
  try {
    Set-DbatoolsInsecureConnection -SessionOnly | Out-Null
    $sqlResults = Invoke-DbaQuery -SqlInstance $sqlInstance -SqlCredential $sqlCredential -Query $sqlQuery -EnableException
  }
  catch {
    Write-Host "Failed to execute the SQL query." -ForegroundColor Red
    Write-Host $_ -ForegroundColor Yellow
  }
    
  # Configure CSS for the HTML output
  $css = @"
    <head>
      <style>
        .even {
          background-color: white;
          color: black;
          font-weight: normal;
          font-family: Arial;
        }

        .odd {
          background-color: lightgray;
          color: black;
          font-weight: normal;
          font-family: Arial;
        }

        th {
          background-color: #3a464e;
          text-align: left;
          color: white;
          font-weight: bold;
          font-family: Arial;
          font-size: larger;
          padding-left: 5px;
          padding-right: 5px;
        }

        td {
          padding-left: 5px;
          padding-right: 5px;
        }

        table {
          width: fit-content;
          border: 3px solid;
          border-color: #39464e;
          border-radius: 10px;
          border-style: outset;
        }

        div {
          padding-bottom: 20px;
        }

        #header img {
          height: auto;
          width: 100px;
          margin-top: 10px;
          margin-left: 10px;
        }

        #header h1 {
          display: inline;
          margin-left: 50px;
        }

      </style>
    </head>
"@

  # Configure the HTML output title
  $title = @"
    <div id="header">
        <img src='https://your-domain.com/logo.png'>
        <h1>$reportTitle</h1>
    </div>
"@

  # Exclude properties that are not needed in the HTML output
  $excludeProperties = @("RowError", "RowState", "ItemArray", "HasErrors", "Table")
    
  # Combine the CSS and title to be added to the head of the HTML output
  $headerContent = $css + $title

  # Add the data generated footer (if required)
  if ($showFooter) {
    $dataGenerated = "<footer><p>Data Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p></footer>"
  }
  else {
    $dataGenerated = " "
  }

  # Convert the SQL results to HTML
  $outputHTML = $sqlResults | Select-Object * -ExcludeProperty $excludeProperties | ConvertTo-Html -Title $title -Head $headerContent -PostContent $dataGenerated | Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd

  return $outputHTML
}

Export-ModuleMember -Function Send-SMTPMail, Set-AlternatingRows, Convert-SQLtoHTML