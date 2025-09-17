# Source helper functions
. .\HelperFunctions.ps1

# Initialize log file
$logFile = "CheckMacrosOnSystem.log"

Display-Message "Option 4 selected: Check macros on a specific system"
Log-Message -Message "Option 4 selected: Check macros on a specific system" -LogFile $logFile

# Prompt user for system details
$endpointIp = Read-Host "Please enter the IP address of the system"
$username = Read-Host "Please enter the username"
$password = Read-Host "Please enter the password" -AsSecureString
$passwordPlain = ConvertTo-PlainText -SecureString $password

$macros = Get-SystemMacros -EndpointIp $endpointIp -Username $username -Password $passwordPlain -LogFile $logFile -SystemName $endpointIp

if ($macros.Count -gt 0) {
    $message = "The following macros were found on ${endpointIp}: $($macros -join ', ')"
    Display-Message $message
    Log-Message -Message $message -LogFile $logFile
} else {
    $message = "No macros found on ${endpointIp}."
    Display-Message $message
    Log-Message -Message $message -LogFile $logFile
}

if (Get-Confirmation -promptMessage "Would you like to export the results to a CSV file?") {
    $exportPath = Get-FilePath -promptMessage "Enter the full path for the export CSV file:"
    try {
        $exportData = $macros | ForEach-Object { [PSCustomObject]@{ IPAddress = $endpointIp; Username = $username; Macro = $_ } }
        if ($exportData.Count -eq 0) {
            $exportData = [PSCustomObject]@{ IPAddress = $endpointIp; Username = $username; Macro = "<None>" }
        }
        $exportData | Export-Csv -Path $exportPath -NoTypeInformation
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
