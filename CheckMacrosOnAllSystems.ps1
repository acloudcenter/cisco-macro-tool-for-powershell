# Source helper functions
. .\HelperFunctions.ps1

# Initialize log file
$logFile = "CheckMacrosOnAllSystems.log"

Display-Message "Option 5 selected: Check macros on all systems"
Log-Message -Message "Option 5 selected: Check macros on all systems" -LogFile $logFile

# Prompt user for CSV file path
$csvFilePath = Get-FilePath -promptMessage "Please enter the full path to the CSV file:"

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

# Confirm check
if (-not (Get-Confirmation -promptMessage "Do you want to proceed with checking macros on all systems?")) {
    Display-Message "Check cancelled."
    Log-Message -Message "Check cancelled." -LogFile $logFile
    return
}

# Perform the check
$checkResults = @()

foreach ($system in $systems) {
    $systemName = $system.'system name'
    $ipAddress = $system.'ip address'
    $username = $system.'username'
    $password = $system.'password'

    $macros = Get-SystemMacros -EndpointIp $ipAddress -Username $username -Password $password -LogFile $logFile -SystemName $systemName

    if ($macros.Count -gt 0) {
        $message = "The following macros were found on ${ipAddress} ($systemName): $($macros -join ', ')"
    } else {
        $message = "No macros found on ${ipAddress} ($systemName)."
    }

    Display-Message $message
    Log-Message -Message $message -LogFile $logFile

    $checkResults += [PSCustomObject]@{
        IPAddress  = $ipAddress
        SystemName = $systemName
        Macros     = if ($macros.Count -gt 0) { $macros -join '; ' } else { '<None>' }
    }
}

Display-Message "Macro check completed."
Log-Message -Message "Macro check completed." -LogFile $logFile

if (Get-Confirmation -promptMessage "Would you like to export the results to a CSV file?") {
    $exportPath = Get-FilePath -promptMessage "Enter the full path for the export CSV file:"
    try {
        $checkResults | Export-Csv -Path $exportPath -NoTypeInformation
        $message = "Results exported to $exportPath"
        Display-Message $message
        Log-Message -Message $message -LogFile $logFile
    } catch {
        $errorDetails = $_.Exception.Message
        $message = "Failed to export results. Response: $errorDetails"
        Display-Message $message
        Log-Message -Message $message -LogFile $logFile
    }
}
