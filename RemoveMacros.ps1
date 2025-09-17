# Source helper functions
. .\HelperFunctions.ps1

# Initialize log file
$logFile = "RemoveMacros.log"

Display-Message "Option 3 selected: Remove macros"
Log-Message -Message "Option 3 selected: Remove macros" -LogFile $logFile

# Prompt user for CSV file path
$csvFilePath = Get-FilePath -promptMessage "Please enter the full path to the CSV file:"

# Prompt user for the macros to remove
$macrosToRemove = Read-Host "Please enter the macros to remove (comma separated):"
$macrosToRemoveArray = $macrosToRemove -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }

if ($macrosToRemoveArray.Count -eq 0) {
    Display-Message "No macros specified for removal."
    Log-Message -Message "No macros specified for removal." -LogFile $logFile
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

    $macros = Get-SystemMacros -EndpointIp $ipAddress -Username $username -Password $password -LogFile $logFile -SystemName $systemName

    $systemSummary = [PSCustomObject]@{
        IPAddress         = $ipAddress
        SystemName        = $systemName
        TotalMacros       = $macrosToRemoveArray.Count
        SuccessfulRemovals = 0
        FailedRemovals     = 0
    }

    foreach ($macro in $macrosToRemoveArray) {
        if ($macros -contains $macro) {
            $success = Remove-SystemMacro -EndpointIp $ipAddress -Username $username -Password $password -MacroName $macro -LogFile $logFile -SystemName $systemName
            if ($success) {
                $systemSummary.SuccessfulRemovals++
            } else {
                $systemSummary.FailedRemovals++
            }
        } else {
            $message = "Macro $macro not found on ${ipAddress} ($systemName)."
            Display-Message $message
            Log-Message -Message $message -LogFile $logFile
            $systemSummary.FailedRemovals++
        }
    }

    $removalSummary += $systemSummary
}

# Generate summary
$successfulSystems = $removalSummary | Where-Object { $_.SuccessfulRemovals -eq $_.TotalMacros }
$partialSystems = $removalSummary | Where-Object { $_.SuccessfulRemovals -gt 0 -and $_.SuccessfulRemovals -lt $_.TotalMacros }
$failedSystems = $removalSummary | Where-Object { $_.SuccessfulRemovals -eq 0 }

Display-Message "Removal completed."
Log-Message -Message "Removal completed." -LogFile $logFile

if ($successfulSystems.Count -gt 0) {
    $message = "Successfully removed all specified macros from $($successfulSystems.Count) systems."
    Display-Message $message
    Log-Message -Message $message -LogFile $logFile
}

if ($partialSystems.Count -gt 0) {
    $message = "Partially removed some macros from $($partialSystems.Count) systems."
    Display-Message $message
    Log-Message -Message $message -LogFile $logFile
}

if ($failedSystems.Count -gt 0) {
    $message = "$($failedSystems.Count) systems were unsuccessful for removing any macros."
    Display-Message $message
    Log-Message -Message $message -LogFile $logFile
}
