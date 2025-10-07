# Source helper functions
. .\HelperFunctions.ps1

# Initialize log file
$logFile = "RemoveUIExtensions.log"

Display-Message "Option 6 selected: Remove UI extensions"
Log-Message -Message "Option 6 selected: Remove UI extensions" -LogFile $logFile

# Prompt user for CSV file path
$csvFilePath = Get-FilePath -promptMessage "Please enter the full path to the CSV file:"

# Prompt user for the UI IDs to remove
$uiIdsToRemove = Read-Host "Please enter the UI IDs to remove (comma separated):"
$uiIdsToRemoveArray = $uiIdsToRemove -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }

if ($uiIdsToRemoveArray.Count -eq 0) {
    Display-Message "No UI extensions specified for removal."
    Log-Message -Message "No UI extensions specified for removal." -LogFile $logFile
    return
}

# Import CSV
try {
    $systems = Import-SystemCsv -Path $csvFilePath
    Display-Message "Found $($systems.Count) systems in the CSV file."
    Log-Message -Message "Found $($systems.Count) systems in the CSV file." -LogFile $logFile
} catch {
    $errorDetails = $_.Exception.Message
    Display-Message "Error importing CSV file. Response: $errorDetails"
    Log-Message -Message "Error importing CSV file. Response: $errorDetails" -LogFile $logFile
    return
}

# Confirm removal
if (-not (Get-Confirmation -promptMessage "Do you want to proceed with the removal?")) {
    Display-Message "Removal cancelled."
    Log-Message -Message "Removal cancelled." -LogFile $logFile
    return
}

# Perform the removal
$removalSummary = @()

foreach ($system in $systems) {
    $systemName = $system.'system name'
    $ipAddress = $system.'ip address'
    $username = $system.'username'
    $password = $system.'password'

    $systemSummary = [PSCustomObject]@{
        IPAddress          = $ipAddress
        SystemName         = $systemName
        TotalUIExtensions  = $uiIdsToRemoveArray.Count
        SuccessfulRemovals = 0
        FailedRemovals     = 0
    }

    foreach ($uiId in $uiIdsToRemoveArray) {
        $success = Remove-UiExtension -EndpointIp $ipAddress -Username $username -Password $password -UiId $uiId -LogFile $logFile -SystemName $systemName
        if ($success) {
            $systemSummary.SuccessfulRemovals++
        } else {
            $systemSummary.FailedRemovals++
        }
    }

    $removalSummary += $systemSummary
}

# Generate summary
$successfulSystems = $removalSummary | Where-Object { $_.SuccessfulRemovals -eq $_.TotalUIExtensions }
$partialSystems = $removalSummary | Where-Object { $_.SuccessfulRemovals -gt 0 -and $_.SuccessfulRemovals -lt $_.TotalUIExtensions }
$failedSystems = $removalSummary | Where-Object { $_.SuccessfulRemovals -eq 0 }

Display-Message "Removal completed."
Log-Message -Message "Removal completed." -LogFile $logFile

if ($successfulSystems.Count -gt 0) {
    $message = "Successfully removed all specified UI extensions from $($successfulSystems.Count) systems."
    Display-Message $message
    Log-Message -Message $message -LogFile $logFile
}

if ($partialSystems.Count -gt 0) {
    $message = "Partially removed some UI extensions from $($partialSystems.Count) systems."
    Display-Message $message
    Log-Message -Message $message -LogFile $logFile
}

if ($failedSystems.Count -gt 0) {
    $message = "$($failedSystems.Count) systems were unsuccessful for removing any UI extensions."
    Display-Message $message
    Log-Message -Message $message -LogFile $logFile
}
