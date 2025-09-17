# Source helper functions
. .\HelperFunctions.ps1

# Initialize log file
$logFile = "UploadJsMacros.log"

Display-Message "Option 2 selected: Upload .js macros"
Log-Message -Message "Option 2 selected: Upload .js macros" -LogFile $logFile

# Prompt user for CSV file path
$csvFilePath = Get-FilePath -promptMessage "Please enter the full path to the CSV file:"

# Prompt user for the directory containing .js files
$jsDirectory = Get-FilePath -promptMessage "Please enter the full path to the directory containing the .js files:"

# List .js files
$jsFiles = Get-ChildItem -Path $jsDirectory -Filter *.js -ErrorAction SilentlyContinue

if ($jsFiles.Count -eq 0) {
    Display-Message "No .js files were found in the specified directory."
    Log-Message -Message "No .js files were found in the directory $jsDirectory." -LogFile $logFile
    return
}

Display-Message "Found $($jsFiles.Count) .js files in the directory."
Log-Message -Message "Found $($jsFiles.Count) .js files in the directory." -LogFile $logFile

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

# Ask to Enable/Disable EvaluateTranspiled option
$option = Read-Host "Do you want to enable transpile evaluation? (yes/no, default: no)"
$option = $option.ToLower()
if ([string]::IsNullOrWhiteSpace($option) -or $option -eq "no" -or $option -eq "n") {
    $evaluateTranspiled = "False"
} elseif ($option -eq "yes" -or $option -eq "y") {
    $evaluateTranspiled = "True"
} else {
    Display-Message "Invalid input. Defaulting to disabled."
    $evaluateTranspiled = "False"
}

# Confirm upload
if (-not (Get-Confirmation -promptMessage "Do you want to proceed with the upload?")) {
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

    Enable-MacrosMode -EndpointIp $ipAddress -Username $username -Password $password -LogFile $logFile -SystemName $systemName | Out-Null

    $transpileSuccess = Set-TranspileEvaluation -EndpointIp $ipAddress -Username $username -Password $password -LogFile $logFile -SystemName $systemName -EvaluateTranspiled $evaluateTranspiled
    if (-not $transpileSuccess) {
        Display-Message "Skipping uploads for $ipAddress ($systemName) due to configuration error"
        Log-Message -Message "Skipping uploads for $ipAddress ($systemName) due to configuration error" -LogFile $logFile
        continue
    }

    $systemSummary = [PSCustomObject]@{
        IPAddress         = $ipAddress
        SystemName        = $systemName
        TotalFiles        = $jsFiles.Count
        SuccessfulUploads = 0
        FailedUploads     = 0
    }

    foreach ($jsFile in $jsFiles) {
        try {
            $jsCode = Get-Content -Path $jsFile.FullName -Raw -ErrorAction Stop
        } catch {
            $errorDetails = $_.Exception.Message
            $message = "Error reading file $($jsFile.FullName). Response: $errorDetails"
            Display-Message $message
            Log-Message -Message $message -LogFile $logFile
            $systemSummary.FailedUploads++
            continue
        }

        $macroName = [System.IO.Path]::GetFileNameWithoutExtension($jsFile.Name)
        $saveSuccess = Save-SystemMacro -EndpointIp $ipAddress -Username $username -Password $password -MacroName $macroName -MacroBody $jsCode -LogFile $logFile -SystemName $systemName

        if ($saveSuccess) {
            $systemSummary.SuccessfulUploads++
            Enable-Macro -EndpointIp $ipAddress -Username $username -Password $password -MacroName $macroName -LogFile $logFile -SystemName $systemName | Out-Null
        } else {
            $systemSummary.FailedUploads++
        }
    }

    if ($systemSummary.SuccessfulUploads -gt 0) {
        Restart-MacroRuntime -EndpointIp $ipAddress -Username $username -Password $password -LogFile $logFile -SystemName $systemName | Out-Null
    }

    $uploadSummary += $systemSummary
}

# Generate summary
$successfulSystems = $uploadSummary | Where-Object { $_.SuccessfulUploads -eq $_.TotalFiles }
$partialSystems = $uploadSummary | Where-Object { $_.SuccessfulUploads -gt 0 -and $_.SuccessfulUploads -lt $_.TotalFiles }
$failedSystems = $uploadSummary | Where-Object { $_.SuccessfulUploads -eq 0 }

Display-Message "Upload completed."
Log-Message -Message "Upload completed." -LogFile $logFile

if ($successfulSystems.Count -gt 0) {
    $message = "Successfully uploaded all macros to $($successfulSystems.Count) systems."
    Display-Message $message
    Log-Message -Message $message -LogFile $logFile
}

if ($partialSystems.Count -gt 0) {
    $message = "Partially uploaded some macros to $($partialSystems.Count) systems."
    Display-Message $message
    Log-Message -Message $message -LogFile $logFile
}

if ($failedSystems.Count -gt 0) {
    $message = "$($failedSystems.Count) systems were unsuccessful for all macros."
    Display-Message $message
    Log-Message -Message $message -LogFile $logFile
}
