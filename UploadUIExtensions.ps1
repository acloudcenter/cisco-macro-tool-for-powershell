# Source helper functions
. .\HelperFunctions.ps1

# Initialize log file
$logFile = "UploadUIExtensions.log"

Display-Message "Option 7 selected: Upload UI extensions"
Log-Message -Message "Option 7 selected: Upload UI extensions" -LogFile $logFile

# Prompt user for CSV file path
$csvFilePath = Get-FilePath -promptMessage "Please enter the full path to the CSV file:"

# Prompt user for the directory containing UI extension files
$panelDirectory = Get-FilePath -promptMessage "Please enter the full path to the directory containing the UI extension XML files:"

# List UI extension files
$panelFiles = Get-ChildItem -Path $panelDirectory -Filter *.xml -ErrorAction SilentlyContinue

if ($panelFiles.Count -eq 0) {
    Display-Message "No UI extension XML files were found in the specified directory."
    Log-Message -Message "No UI extension XML files were found in the directory $panelDirectory." -LogFile $logFile
    return
}

Display-Message "Found $($panelFiles.Count) UI extension files in the directory."
Log-Message -Message "Found $($panelFiles.Count) UI extension files in the directory." -LogFile $logFile

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

# Confirm upload
if (-not (Get-Confirmation -promptMessage "Do you want to proceed with uploading UI extensions?")) {
    Display-Message "Upload cancelled."
    Log-Message -Message "Upload cancelled." -LogFile $logFile
    return
}

# Perform the upload
$uploadSummary = @()

foreach ($system in $systems) {
    $systemName = $system.'system name'
    $ipAddress = $system.'ip address'
    $username = $system.'username'
    $password = $system.'password'

    $systemSummary = [PSCustomObject]@{
        IPAddress         = $ipAddress
        SystemName        = $systemName
        TotalExtensions   = $panelFiles.Count
        SuccessfulUploads = 0
        FailedUploads     = 0
    }

    foreach ($panelFile in $panelFiles) {
        try {
            $panelContent = Get-Content -Path $panelFile.FullName -Raw -ErrorAction Stop
            $panelContent = [System.Text.RegularExpressions.Regex]::Replace($panelContent, '^\s*<\?xml.*?\?>\s*', '', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        } catch {
            $errorDetails = $_.Exception.Message
            $message = "Error reading file $($panelFile.FullName). Response: $errorDetails"
            Display-Message $message
            Log-Message -Message $message -LogFile $logFile
            $systemSummary.FailedUploads++
            continue
        }

        $success = Save-UiExtension -EndpointIp $ipAddress -Username $username -Password $password -PanelContent $panelContent -LogFile $logFile -SystemName $systemName -SourceName $panelFile.Name

        if ($success) {
            $systemSummary.SuccessfulUploads++
        } else {
            $systemSummary.FailedUploads++
        }
    }

    $uploadSummary += $systemSummary
}

# Generate summary
$successfulSystems = $uploadSummary | Where-Object { $_.SuccessfulUploads -eq $_.TotalExtensions }
$partialSystems = $uploadSummary | Where-Object { $_.SuccessfulUploads -gt 0 -and $_.SuccessfulUploads -lt $_.TotalExtensions }
$failedSystems = $uploadSummary | Where-Object { $_.SuccessfulUploads -eq 0 }

Display-Message "UI extension upload completed."
Log-Message -Message "UI extension upload completed." -LogFile $logFile

if ($successfulSystems.Count -gt 0) {
    $message = "Successfully uploaded all UI extensions to $($successfulSystems.Count) systems."
    Display-Message $message
    Log-Message -Message $message -LogFile $logFile
}

if ($partialSystems.Count -gt 0) {
    $message = "Partially uploaded some UI extensions to $($partialSystems.Count) systems."
    Display-Message $message
    Log-Message -Message $message -LogFile $logFile
}

if ($failedSystems.Count -gt 0) {
    $message = "$($failedSystems.Count) systems were unsuccessful for all UI extensions."
    Display-Message $message
    Log-Message -Message $message -LogFile $logFile
}
