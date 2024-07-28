# Define the log file path with a .txt extension
$logFilePath = "C:\temp\MyLogfile.txt"

# Function for logging information
function Log-Command {
    param(
        [string]$Command,
        [string]$Output
    )

    $timeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logData = "[$timeStamp] Command: $Command`r`nOutput:`r`n$Output`r`n"

    try {
        # Append to the log file
        Add-Content -Path $logFilePath -Value $logData -Encoding utf8
    } catch {
        Write-Host "Error writing to log file: $_"
    }
}

# Logging information for script start
Log-Command -Command "Starting script" -Output "Script execution started"

# Init PowerShell GUI
Add-Type -AssemblyName System.Windows.Forms
$form = New-Object System.Windows.Forms.Form
$form.SuspendLayout()
$form.AutoScaleDimensions =  New-Object System.Drawing.SizeF(96, 96)
$form.AutoScaleMode  = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.Text = "Computer Information and BitLocker Status"
$form.ClientSize = New-Object System.Drawing.Size(900, 750)
$form.Font = New-Object System.Drawing.Font("Times New Roman",13)
# Get the computer's hostname
$hostname = $env:COMPUTERNAME

$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$transcriptPath = "C:\temp\adstatus.txt"

Start-Transcript -Path $transcriptPath

# Run the systeminfo | findstr /B "Domain" command and capture the output
$adStatusOutput = systeminfo | findstr /B "Domain"

# Check if the computer is bound to Active Directory
$adStatus = if ($adStatusOutput) {
    "Bound"
} else {
    "Not bound"
}

# Output AD status along with the command output to the transcript with timestamp
"[$timestamp] AD Status: $($adStatus)"
"[$timestamp] Command Output: $($adStatusOutput)"

Stop-Transcript


# Get the list of local administrators
$admins = Get-LocalGroupMember -Group "Administrators"
# Define the desired order of administrators
$desiredOrder = @(
    "CMP572JNW3\1899",
    "NAU\pg495",
    "NAU\CTS-Support",
    "NAU\DCIS Student UID Accounts",
    "NAU\domain admins",
    "NAU\ITS-IS-ARSA-PA",
    "NAU\ITS-IS-DCIS-PA",
    "NAU\its-is-dcis-students",
    "NAU\ITS-IS-Platform-PA",
    "NAU\ITSS-DS-PA",
    "NAU\SCCMadmin"
)
# Create a custom sorting order based on the desired order
$sortingOrder = @{}
for ($i = 0; $i -lt $desiredOrder.Length; $i++) {
    $sortingOrder[$desiredOrder[$i]] = $i
}
# Sort the administrators based on the custom sorting order
$admins = $admins | Sort-Object { $sortingOrder[$_.Name] }
# Create the formatted list
$adminList = $admins.Name -join "`r`n"

# Create a new instance of the Microsoft.Update.Session COM object
$updateSession = New-Object -ComObject Microsoft.Update.Session
# Set the ClientApplicationID property
$updateSession.ClientApplicationID = "MSDN Sample Script"
# Create an UpdateSearcher object
$updateSearcher = $updateSession.CreateUpdateSearcher()
# Search for all available software updates that are not installed and not hidden
$updates = $updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")

# Check if there are any Windows updates available
if ($updates.Updates.Count -gt 0) {
    # Display available updates
    $updateTitles = $updates.Updates | ForEach-Object { "$($_.Title) - $($_.HotFixID)" }
    $updatesAvailable = "$($updateTitles -join "`r`n")"

    # Check if any of the updates require a restart
    $updatesRequiringRestart = $updates.Updates | Where-Object { $_.RebootRequired -eq $true }

    if ($updatesRequiringRestart.Count -gt 0) {
        $windowsUpdateStatus = "Updates available, and the following updates require a restart"
        $windowsUpdateStatus += $updatesRequiringRestart | ForEach-Object { "$($_.Title) - $($_.HotFixID)" }
    } else {
        $windowsUpdateStatus = "UPDATES AVAILABLE, but none require a restart"
    }
} else {
    $windowsUpdateStatus = ""
    $updatesAvailable ="There are no updates available."    
}



# Get BitLocker status
$bitlockerStatus = (Get-BitlockerVolume -MountPoint "C:").ProtectionStatus
# Display BitLocker status
$bitlockerStatusText = if ($bitlockerStatus -like "Off") {
    "OFF" 
} else {
    "ON"
}

$updateLabel = New-Object System.Windows.Forms.Label
$updateLabel.Text = "Windows Update:"
$updateLabel.Location = New-Object System.Drawing.Point(0, 150)
$updateLabel.AutoSize = $true
$form.Controls.Add($updateLabel)
$updateText = New-Object System.Windows.Forms.TextBox
$updateText.AcceptsReturn = $true
$updateText.ReadOnly = $true
$updateText.Multiline = $true
$updateText.ScrollBars = "Vertical"
$updateText.Text = $updatesAvailable
$updateText.Width = 600
$updateText.Height = 80
$updateText.Location = New-Object System.Drawing.Point(200, 150)
$form.Controls.Add($updateText)

# Create a label for Updates Requiring Restart
$restartUpdatesLabel = New-Object System.Windows.Forms.Label
$restartUpdatesLabel.Text = "Updates:"
$restartUpdatesLabel.Location = New-Object System.Drawing.Point(0, 110)
$restartUpdatesLabel.AutoSize = $true
$form.Controls.Add($restartUpdatesLabel)

# Create a text box for Updates Requiring Restart
$restartUpdatesText = New-Object System.Windows.Forms.Label
$restartUpdatesText.Text = $windowsUpdateStatus
$restartUpdatesText.Width = 600
$restartUpdatesText.Height = 23
$restartUpdatesText.Location = New-Object System.Drawing.Point(200, 110)
$form.Controls.Add($restartUpdatesText)


# Check if the SCCM client is installed
$SCCMClientInstalled = Test-Path "C:\Windows\CCM"

# Add a label to display the SCCM client installation status
$sccmClientLabel = New-Object System.Windows.Forms.Label
$sccmClientLabel.Text = "SCCM Client :"
$sccmClientLabel.Location = New-Object System.Drawing.Point(0, 390)
$sccmClientLabel.AutoSize = $true
$form.Controls.Add($sccmClientLabel)

# Add a text box to display the SCCM client installation status
$sccmClientText = New-Object System.Windows.Forms.TextBox
$sccmClientText.AcceptsReturn = $false
$sccmClientText.ReadOnly = $true
$sccmClientText.Text = if ($SCCMClientInstalled) {"Installed"} else {"Not Installed"}
$sccmClientText.Width = 300
$sccmClientText.Location = New-Object System.Drawing.Point(200,380)
$form.Controls.Add($sccmClientText)

# Add a button to install Windows updates
$installButton = New-Object System.Windows.Forms.Button
$installButton.Text = "Install Windows Updates"
$installButton.Location = New-Object System.Drawing.Point(200, 540)  # Position the button below the update status
$installButton.Width = 200
$installButton.Height = 30
$installButton.Add_Click({
    # Search for and install Windows updates using Windows Update Agent
    $searcher = New-Object -ComObject Microsoft.Update.Searcher
    $session = New-Object -ComObject Microsoft.Update.Session
    $criteria = "IsInstalled=0 and Type='Software' and IsHidden=0"
    $searchResult = $searcher.Search($criteria)
    if ($searchResult.Updates.Count -gt 0) {
        $updateCollection = New-Object -ComObject Microsoft.Update.UpdateColl
        foreach ($update in $searchResult.Updates) {
            $updateCollection.Add($update)
        }
        $downloader = $session.CreateUpdateDownloader()
        $downloader.Updates = $updateCollection
        $downloader.Download()
        $installer = $session.CreateUpdateInstaller()
        $installer.Updates = $updateCollection
        $installationResult = $installer.Install()
        if ($installationResult.ResultCode -eq 2) {
            Write-Host "Updates were successfully installed."
            $updateText.Text = "All updates are installed."
        } else {
            Write-Host "Failed to install updates. Result Code: $($installationResult.ResultCode)"
        }
    } else {
        Write-Host "No updates available."
        $updateText.Text = "No updates available."
    }
})
$form.Controls.Add($installButton)
# Create a button to turn on BitLocker
$bitlockerButton = New-Object System.Windows.Forms.Button
$bitlockerButton.Text = "Turn On BitLocker"
$bitlockerButton.Location = New-Object System.Drawing.Point(600, 540)  # Position the button below the "Install Updates" button
$bitlockerButton.Width = 200
$bitlockerButton.Height = 30

# Create a function to display a message in a popup window
function Show-MessageBox($message) {
    [System.Windows.Forms.MessageBox]::Show($message, "Bitlocker status", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
}

$bitlockerButton.Add_Click({
    # Check the BitLocker status
    $bitlockerStatus = (Get-BitlockerVolume -MountPoint "C:").ProtectionStatus
    if ($bitlockerStatus -eq "Off") {
        # Turn on BitLocker
        Enable-BitLocker -MountPoint "C:" -RecoveryPasswordProtector -UsedSpaceOnly
        $bitlockerText.Text = "ON"
        Show-MessageBox "BitLocker is encrypting,please reboot your computer to finish encrypting your computer."
    } else {
        Show-MessageBox "BitLocker is already ON."
    }
})
$form.Controls.Add($bitlockerButton)

# Add a label to display the dell command update
$dellupdateLabel = New-Object System.Windows.Forms.Label
$dellupdateLabel.Text = "Dell command update :"
$dellupdateLabel.Location = New-Object System.Drawing.Point(0, 460)
$dellupdateLabel.AutoSize = $true
$form.Controls.Add($dellupdateLabel)

# Create a text box to display the script output
$outputTextBox = New-Object System.Windows.Forms.TextBox
$outputTextBox.Location = New-Object System.Drawing.Point(200, 420)
$outputTextBox.Size = New-Object System.Drawing.Size(600, 100)
$outputTextBox.Multiline = $true
$outputTextBox.ScrollBars = "Vertical"
$outputTextBox.ReadOnly = $true
$form.Controls.Add($outputTextBox)

# Define a function to execute your script
function ExecuteScript {
    # Define the path to dcu-cli.exe
    $dcuPath = "C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe"
    $scriptOutput = ""

    try {
        # Run dcu-cli.exe /scan and capture the output
        $dcuOutput = & $dcuPath /scan
        # Specify the update codes to search for
        $targetUpdateCodes = @("5WCHH", "6PJ7T", "MVJH7", "YYJM0")
        # Initialize a flag to check if any target update codes are found
        $targetCodesFound = $false

        # Iterate through each line and find the specified updates
        $updateLines = $dcuOutput -split "`n" | Where-Object { $_ -match "^($($targetUpdateCodes -join '|')):" }

        if ($updateLines.Count -eq 0) {
            $scriptOutput = "No updates available."
        } else {
            # Build the script output
            $scriptOutput = $updateLines -join "`r`n"
        }
    } catch {
        $scriptOutput = "Error running Dell Command Update: $_"
    }

    # Update the output text box with the script output
    $outputTextBox.Text = $scriptOutput
}

# Execute the script when the GUI is loaded
$null = $form.Add_Shown({ ExecuteScript })

# Function to run Dell Command Update
function RunDellCommandUpdate {
    $dcuPath = "C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe"
    $dcuOutput = & $dcuPath /applyUpdates
    $targetUpdateCodes = @("5WCHH", "6PJ7T", "MVJH7", "YYJM0")
        # Initialize a flag to check if any target update codes are found
        $targetCodesFound = $false

        # Iterate through each line and find the specified updates
        $updateLines = $dcuOutput -split "`n" | Where-Object { $_ -match "^($($targetUpdateCodes -join '|')):" }

        if ($updateLines.Count -eq 0) {
            $scriptOutput = "No updates available."
        } else {
            # Build the script output
            $scriptOutput = $updateLines -join "`r`n"
        }
     # Display a message about restarting and potential BitLocker status change
    $message = "Dell Command Update has completed. Please restart your computer to apply the updates. Note: BitLocker may be temporarily turned off during the update process."
    Show-MessageBox $message
    # Update the output text box with the DCU output
    $outputTextBox.Text = $scriptoutput
}
# Add a button to run Dell Command Update
$dellUpdateButton = New-Object System.Windows.Forms.Button
$dellUpdateButton.Text = "Run Dell Command Update"
$dellUpdateButton.Location = New-Object System.Drawing.Point(200, 600)  # Adjust the location as needed
$dellUpdateButton.Width = 250
$dellUpdateButton.Height = 30
$form.Controls.Add($dellUpdateButton)

# Add an event handler for the Dell Update button
$dellUpdateButton.Add_Click({
    RunDellCommandUpdate
})

# Create labels and text boxes for the GUI
$hostLabel = New-Object System.Windows.Forms.Label
$hostLabel.Text = "Hostname:"
$hostLabel.Location = New-Object System.Drawing.Point(0, 10)
$hostLabel.AutoSize = $true
$form.Controls.Add($hostLabel)

$hostText = New-Object System.Windows.Forms.TextBox
$hostText.AcceptsReturn = $false
$hostText.ReadOnly = $true
$hostText.Text = $hostname
$hostText.Width = 300
$hostText.Location = New-Object System.Drawing.Point(200, 10)
$form.Controls.Add($hostText)

$adLabel = New-Object System.Windows.Forms.Label
$adLabel.Text = "Active Directory Status:"
$adLabel.Location = New-Object System.Drawing.Point(0, 60)
$adLabel.AutoSize = $true
$form.Controls.Add($adLabel)

# Define the Active Directory status label with a specific location
$adStatusLabel = New-Object System.Windows.Forms.Label
$adStatusLabel.Text = $adStatus
$adStatusLabel.Location = New-Object System.Drawing.Point(200, 59)
$form.Controls.Add($adStatusLabel)

# Check the Active Directory status and set the text color
if ($adStatus -eq "Bound") {
    $adStatusLabel.ForeColor = [System.Drawing.Color]::Green
} else {
    $adStatusLabel.ForeColor = [System.Drawing.Color]::Red
}

$adminLabel = New-Object System.Windows.Forms.Label
$adminLabel.Text = "Local Administrators:"
$adminLabel.Location = New-Object System.Drawing.Point(0, 240)
$adminLabel.AutoSize = $true
$form.Controls.Add($adminLabel)

$adminText = New-Object System.Windows.Forms.TextBox
$adminText.AcceptsReturn = $true
$adminText.ReadOnly = $true
$adminText.Multiline = $true
$adminText.ScrollBars = "Vertical"
$adminText.Text = $adminList
$adminText.Width = 600
$adminText.Height = 90
$adminText.Location = New-Object System.Drawing.Point(200, 240)
$form.Controls.Add($adminText)

$bitlockerLabel = New-Object System.Windows.Forms.Label
$bitlockerLabel.Text = "BitLocker Status:"
$bitlockerLabel.Location = New-Object System.Drawing.Point(0, 340)
$bitlockerLabel.AutoSize = $true
$form.Controls.Add($bitlockerLabel)

$bitlockerText = New-Object System.Windows.Forms.Label
$bitlockerText.Location = New-Object System.Drawing.Point(200, 340)
$bitlockerText.AutoSize = $true
$bitlockerText.Text = $bitlockerStatusText

# Set the color based on the BitLocker status
if ($bitlockerStatus -like "Off") {
    $bitlockerText.ForeColor = [System.Drawing.Color]::Red
} else {
    $bitlockerText.ForeColor = [System.Drawing.Color]::Green
}
$form.Controls.Add($bitlockerText)
# Create a function to update all the components
function UpdateComponents {
    # Get the computer's hostname
    $hostname = $env:COMPUTERNAME
    $hostText.Text = $hostname
    
    # Run the systeminfo | findstr /B "Domain" command and capture the output
    $adStatusOutput = systeminfo | findstr /B "Domain"

    # Check if the computer is bound to Active Directory
    $adStatus = if ($adStatusOutput) {
    "Bound"
    } else {
    "Not bound"
    }
    $adStatusLabel.Text = $adStatus

    # Refresh the list of local administrators
    $admins = Get-LocalGroupMember -Group "Administrators"
    $admins = $admins | Sort-Object { $sortingOrder[$_.Name] }
    $adminText.Text = $admins.Name -join "`r`n"

    # Refresh Windows Update status
    $updateSearcher = $updateSession.CreateUpdateSearcher()
    $updates = $updateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
    $updatesAvailable = if ($updates.Updates.Count -gt 0) {
        $updateTitles = $updates.Updates | ForEach-Object { $_.Title }
        $updateTitles -join "`r`n"
    } else {
        "There are no updates available."
        $windowsUpdateStatus = ""
    }
    $updateText.Text = $updatesAvailable

    # Refresh BitLocker status
    $bitlockerStatus = (Get-BitlockerVolume -MountPoint "C:").ProtectionStatus
    $bitlockerText.Text = if ($bitlockerStatus -like "Off") {
        "OFF"
    } else {
        "ON"
    }
}   
# Create a button for refreshing components
$refreshButton = New-Object System.Windows.Forms.Button
$refreshButton.Text = "Refresh"
$refreshButton.Location = New-Object System.Drawing.Point(600, 600)  # Adjust button position
$refreshButton.Width = 100
$refreshButton.Height = 30
$refreshButton.Add_Click({
    # Call the function to update all components
    UpdateComponents
})
$form.Controls.Add($refreshButton)
# For example, after updating components or performing certain operations:
Log-Command -Command "Updated components" -Output "Components refreshed successfully"

# Finally, log the completion of the script
Log-Command -Command "Script completed" -Output "Script execution completed successfully"
$form.ResumeLayout()

$form.ShowDialog()