#########################################################
# Please modify these variables according to your needs #
#########################################################
$thumbprint = <your certificate thumbprint>
$org = <your org>.onmicrosoft.com # this should be formatted like contosco.onmicrosoft.com
$appID = <your app id>
$SMTPServer = <your SMTP server>
$fromEmailError = <Error-Email-Admin@yourDomain.com> # make sure to include the <>
$toEmailError = yourEmail@yourDomain.com
$emailAdmin = <noreply@yourDomain.com> # make sure to include the <>
$workingDirectory = "C:\path\to\your\working\directory"
$HelpDesk = HelpDesk@yourDomain.com

$dateFile = $workingDirectory + "\date.txt"

function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$Message,

        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARN", "ERROR", "DEBUG")]
        [string]$Level = "INFO",

        [Parameter(Mandatory=$false)]
        [string]$LogFilePath = $workingDirectory + "\app.log"
    )

    Begin {
        # Check if log file directory exists, if not, create it
        $logFileDirectory = Split-Path -Path $LogFilePath -Parent
        if (-not (Test-Path -Path $logFileDirectory)) {
            New-Item -ItemType Directory -Path $logFileDirectory | Out-Null
        }
    }

    Process {
        # Format the log entry
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "$timestamp [$Level] $Message"

        # Write the log entry to the log file
        Add-Content -Path $LogFilePath -Value $logEntry
    }

    End {
        if ($Level -eq "ERROR") {
            Write-Host "An error has been logged to $LogFilePath" -ForegroundColor Red
        } else {
            Write-Host "Log entry added to $LogFilePath" -ForegroundColor Green
        }
    }
}

try {
    # Step 1: Connect to Exchange Online
    Connect-ExchangeOnline -CertificateThumbPrint $thumbprint -AppID $appID -Organization $org -ShowBanner:$false -ErrorAction Stop
    Write-Log -Message "Successfully Connected to Exchange Online"
} catch {
    Write-Log -Message "Failure in Connecting to Exchange Online: $_" -Level ERROR
    try {
        Send-MailMessage -From "Script Run Failure $fromEmailError" -To $toEmailError -Subject "Terminal Failure: Quarantined Email Digest" -Body $_ -SmtpServer $SMTPServer
        Write-Log -Message "Successfully sent ERROR email."
    } catch {
        Write-Log -Level ERROR -Message "Everything is broken. It's time to cry now. $_"
    }
    exit 1
}

# Get the current local time
$localTime = Get-Date

# Read the last run date from the file as local time
$lastRunContent = Get-Content -Path $dateFile -Raw

$lastRunLocal = [DateTime]::Parse($lastRunContent)

# Save the current local time for the next run
$localTime | Out-File -FilePath $dateFile
try {
    # Use the local times for querying
    $messages = Get-QuarantineMessage -StartReceivedDate $lastRunLocal.AddHours(-1) -EndReceivedDate $localTime -ErrorAction Stop
    Write-Log -Message "Successfully pulled recently quarantined messages."
} catch {
    Write-Log -Message "Error in pulling quarantined messages: $_" -Level ERROR
    try {
        Send-MailMessage -From "Script Run Failure $toEmailError" -To $toEmailError -Subject "Terminal Failure: Quarantined Email Digest" -Body $_ -SmtpServer $SMTPServer
        Write-Log -Message "Successfully sent ERROR email."
    } catch {
        Write-Log -Level ERROR -Message "Everything is broken. It's time to cry now. $_"
    }
    exit 1
}
# Check if there are any relevant quarantined messages
if ($messages.Count -eq 0) {
    Write-Log -Message "No quarantined messages found since the last run plus 1 hour."
    Start-Sleep 15
    Disconnect-ExchangeOnline -Confirm:$false
    return
}

# Step 3: Remove any messages that have already been processed
$logPath = $workingDirectory + "\messages.log"
$sentMessagesLog = Get-Content -Path $logPath
$newMessages = @()
foreach ($message in $messages) {
    if ($sentMessagesLog -NotContains $message.Identity) {
        $newMessages += $message
    }
}

# Check again if there are new quarantined messages
if ($newMessages.Count -eq 0) {
    Write-Log -Message "No new quarantined messages."
    Start-Sleep 15
    Disconnect-ExchangeOnline -Confirm:$false
    return
}

# Step 4: Process and format messages for each recipient
$uniqueRecipients = ($newMessages).RecipientAddress | Sort-Object | Get-Unique

try {
    foreach ($recipient in $uniqueRecipients) {
        $messageTable = @"
<table style="border-collapse: collapse; width: 100%;">
    <tr>
        <th style="border: 1px solid #dddddd; text-align: left; padding: 8px;">Received Time (MT)</th>
        <th style="border: 1px solid #dddddd; text-align: left; padding: 8px;">Sender Address</th>
        <th style="border: 1px solid #dddddd; text-align: left; padding: 8px;">Recipient Address</th>
        <th style="border: 1px solid #dddddd; text-align: left; padding: 8px;">Subject</th>
        <th style="border: 1px solid #dddddd; text-align: left; padding: 8px;">Quarantine Reason</th>
    </tr>
"@
        foreach ($message in $newMessages) {
            if ($message.RecipientAddress -eq $recipient) {
                $messageTable += "<tr><td style='border: 1px solid #dddddd; text-align: left; padding: 8px;'>$($message.ReceivedTime)</td><td style='border: 1px solid #dddddd; text-align: left; padding: 8px;'>$($message.SenderAddress)</td><td style='border: 1px solid #dddddd; text-align: left; padding: 8px;'>$($message.RecipientAddress)</td><td style='border: 1px solid #dddddd; text-align: left; padding: 8px;'>$($message.Subject)</td><td style='border: 1px solid #dddddd; text-align: left; padding: 8px;'>$($message.Type)</td></tr>"
            }
        }
        $messageTable += "</table>"

        # Enhancements for aesthetics
        $body = @"
<html>
<head>
<style>
body { font-family: Arial, sans-serif; line-height: 1.6; }
h2 { color: #333; }
p { margin: 16px 0; }
</style>
</head>
<body>
<h2>Quarantined Email Digest</h2>
<p>The following emails were quarantined in the last 24 hours:</p>
$messageTable
<p>To request a re-review or release of these messages please forward this email to $HelpDesk with a description of the business need for this email.</p>
<p>Thanks! -IT Security Team</p>
</body>
</html>
"@

        # This sends the emails to the users who have a quarantined message
        try {
            $recipientArray = $recipient -split ","
            Send-MailMessage -From "noreply $emailAdmin" -To $recipientArray -Subject "Quarantined Email Digest" -Body $body -BodyAsHtml -SmtpServer $SMTPServer
            Write-Log -Message "Successfully sent email to, $recipient"
        } catch {
            Write-Log -Level ERROR -Message "Error in sending email to ($recipient): $_"
        }
        

        ##these are for testing--simply dump into a log the emails that would've been sent
        #$output = $workingDirectory + "\$recipient" + ".html"
        #$body | Out-File $output
        #Write-Log -Message "Dropped HTML file to $output" -Level DEBUG

    }
    Write-Log -Message "Emails sending complete."
} catch {
    Write-Log -Message "Unknown error in sending email: $_" -Level ERROR
}

# Step 5: Add newly sent messages to messages.log
try {
    foreach ($message in $newMessages) {
        Add-Content -Path $logPath -Value $message.Identity
        $messageID = $message.Identity
        Write-Log -Level DEBUG -Message "Adding $messageID to $logPath"
    }
    Write-Log -Message "Message Identities added to $logPath"
} catch {
    Write-Log -Level ERROR -Message "Error in adding message identities to ($logPath): $_"
}

# Step 6: Removing old Message Identities

# Ppath to the temp file
$tempFilePath = $workingDirectory + "\temp.log"
New-Item -Path $tempFilePath -ItemType File -Force | Out-Null

# Read the file containing the Quarantined Message Identities
$quarantinedMessageIds = Get-Content $logPath

# Prepare an array to hold the valid lines
$validLines = @()

foreach ($id in $quarantinedMessageIds) {
    # Retrieve the quarantine message details
    try {
        $messageDetails = Get-QuarantineMessage -Identity $id
        if ($messageDetails) {
            # Calculate the days until expiration
            $daysUntilExpiration = ($messageDetails.Expires - $localTime).Days
            # If the message is more than 5 days away from expiring, add it to the valid lines
            if ($daysUntilExpiration -gt 5) {
                $validLines += $id
            } else {
                Write-Log -Message "Removing about to expire message id, $id from $logPath" -Level DEBUG
            }
        }
    } catch {
        Write-Log -Level ERROR -Message "Error thrown with identity, ($id): $_"
    }
}

# Write the valid lines to the output file
$validLines | Out-File $tempFilePath
Remove-Item -Path $logPath
Rename-Item -Path $tempFilePath -NewName $logPath

# Output completion message
Write-Log -Message "Process completed. Valid lines are saved to $logPath"


# Disconnect from Exchange Online session
Start-Sleep 30
try {
    Disconnect-ExchangeOnline -Confirm:$false
} catch {
    Write-Log -Level WARN -Message "Error in disconnecting Exchange Online: $_"
}

# step 7 compress old logs; delete the ones over a year old
# Define the log file path
$theLogFilePath = $workingDirectory + "\app.log"

$today = Get-Date -Format "yyyyMMdd"

# Temporary file paths
$tempFilePath = $workingDirectory + "\temp.log"
$finalFilePath = $workingDirectory + "\ArchivedLogs\app_log" + $today + ".zip"

# Get the total line count of the log file
$totalLines = (Get-Content $theLogFilePath).Count

# Calculate lines to skip (total lines - 1000)
$linesToSkip = $totalLines - 1000

$applog = $workingDirectory + "\app.log"

# Check if the file has more than 1000 lines
if ((Get-Item $applog).length -gt 100000000) {
    # Extract all but the last 1000 lines and save to a temporary file
    Get-Content $theLogFilePath | Select-Object -First $linesToSkip | Set-Content $tempFilePath

    # Compress the temporary file
    Compress-Archive -Path $tempFilePath -DestinationPath "$finalFilePath" -Force

    # Extract the last 1000 lines and overwrite the original log file
    Get-Content $theLogFilePath | Select-Object -Last 1000 | Set-Content $theLogFilePath

    Write-Log -Level INFO -Message "Log file size reduced; old logs compressed to $finalFilePath"
} else {
    Write-Log -Message "The log file has 1000 lines or fewer. No need to split and compress." -Level DEBUG
}

# Clean up the temporary file if it exists
if (Test-Path $tempFilePath) {
    Remove-Item $tempFilePath -Force
}

# Set the directory where your zip files are stored
$targetDirectory = $workingDirectory + "\ArchivedLogs"


# Calculate the date one year ago from today
$oneYearAgo = (Get-Date).AddYears(-1)

# Get a list of all zip files in the target directory
$zipFiles = Get-ChildItem -Path $targetDirectory -Filter "app_log*.zip"

foreach ($file in $zipFiles) {
    # Extract the date part of the file name (assuming format is "app_logyyyyMMdd.zip")
    $dateString = $file.BaseName -replace "app_log", ""

    # Parse the date string into a DateTime object
    try {
        $fileDate = [DateTime]::ParseExact($dateString, "yyyyMMdd", $null)

        # Check if the file date is older than one year ago
        if ($fileDate -lt $oneYearAgo) {
            # Delete the file
            Remove-Item $file.FullName -Force
            Write-Log -Message "Old Archival files found: Deleted file: $($file.FullName)" -Level INFO
        }
    }
    catch {
        Write-Log -Level ERROR -Message "Error in Archival Delete: Could not parse date for file: $($file.Name)"
    }
} 
