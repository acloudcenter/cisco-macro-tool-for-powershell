# Source helper functions
. .\HelperFunctions.ps1

# Function to log messages to a file
function Log-Message {
    param (
        [string]$message,
        [string]$logFile
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp - $message"
    Add-Content -Path $logFile -Value $logEntry
}

# Function to display messages in the console
function Display-Message {
    param (
        [string]$message
    )
    Write-Host $message
}

# Initialize log file
$logFile = "UploadJsMacros.log"

Display-Message "Option 2 selected: Upload .js macros"
Log-Message -message "Option 2 selected: Upload .js macros" -logFile $logFile

# Prompt user for CSV file path
$csvFilePath = Get-FilePath -promptMessage "Please enter the full path to the CSV file:"

# Prompt user for the directory containing .js files
$jsDirectory = Get-FilePath -promptMessage "Please enter the full path to the directory containing the .js files:"

# List .js files
$jsFiles = Get-ChildItem -Path $jsDirectory -Filter *.js
Display-Message "Found $($jsFiles.Count) .js files in the directory."
Log-Message -message "Found $($jsFiles.Count) .js files in the directory." -logFile $logFile

# Import CSV
$systems = Import-Csv -Path $csvFilePath
Display-Message "Found $($systems.Count) systems in the CSV file."
Log-Message -message "Found $($systems.Count) systems in the CSV file." -logFile $logFile

# Confirm upload
if (-not (Get-Confirmation -promptMessage "Do you want to proceed with the upload?")) {
    Display-Message "Upload cancelled."
    Log-Message -message "Upload cancelled." -logFile $logFile
    return
}

# Function to upload macros from a .js file
function Upload-MacroFromJs {
    param (
        [string]$endpointIp,
        [string]$username,
        [string]$password,
        [string]$jsFilePath,
        [string]$logFile
    )

    $message = "Attempting to upload macro from $jsFilePath to $endpointIp..."
    Display-Message $message
    Log-Message -message $message -logFile $logFile

    $macroName = [System.IO.Path]::GetFileNameWithoutExtension($jsFilePath)
    $jsCode = Get-Content -Path $jsFilePath -Raw

    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Content-Type", "application/xml")
    $headers.Add("Authorization", "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${username}:${password}")))

    $body = @"
<Command>
    <Macros>
        <Macro>
            <Save command="True">
                <name>$macroName</name>
                <body><![CDATA[$jsCode]]></body>
                <overWrite>True</overWrite>
                <Transpile>False</Transpile>
            </Save>
        </Macro>
    </Macros>
</Command>
"@

    try {
        $response = Invoke-RestMethod -Uri "https://$endpointIp/putxml" -Method 'POST' -Headers $headers -Body $body -SkipCertificateCheck -TimeoutSec 10
        $message = "Macro $macroName uploaded successfully to $endpointIp from $jsFilePath."
        Display-Message $message
        Log-Message -message $message -logFile $logFile
        return $true
    } catch {
        $errorDetails = $_.Exception.Message
        $message = "Error uploading macro $macroName to: $endpointIp from $jsFilePath. Response: $errorDetails"
        Display-Message $message
        Log-Message -message $message -logFile $logFile
        return $false
    }
}

# Function to restart macro runtime
function Restart-MacroRuntime {
    param (
        [string]$endpointIp,
        [string]$username,
        [string]$password,
        [string]$logFile
    )

    $message = "Attempting to restart macro runtime on $endpointIp..."
    Display-Message $message
    Log-Message -message $message -logFile $logFile

    $headers = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
    $headers.Add("Content-Type", "application/xml")
    $headers.Add("Authorization", "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${username}:${password}")))

    $body = @"
<Command>
    <Macros>
        <Runtime>
            <Restart command="True"/>
        </Runtime>
    </Macros>
</Command>
"@

    try {
        $response = Invoke-RestMethod -Uri "https://$endpointIp/putxml" -Method 'POST' -Headers $headers -Body $body -SkipCertificateCheck -TimeoutSec 10
        $message = "Macro runtime restarted successfully on $endpointIp."
        Display-Message $message
        Log-Message -message $message -logFile $logFile
    } catch {
        $errorDetails = $_.Exception.Message
        $message = "Error restarting macro runtime on $endpointIp. Response: $errorDetails"
        Display-Message $message
        Log-Message -message $message -logFile $logFile
    }
}

# Perform the upload
$uploadSummary = @()

foreach ($system in $systems) {
    $ipAddress = $system.'ip address'
    $username = $system.'username'
    $password = $system.'password'

    $systemSummary = [PSCustomObject]@{
        IPAddress = $ipAddress
        TotalFiles = $jsFiles.Count
        SuccessfulUploads = 0
        FailedUploads = 0
    }

    foreach ($jsFile in $jsFiles) {
        $success = Upload-MacroFromJs -endpointIp $ipAddress -username $username -password $password -jsFilePath $jsFile.FullName -logFile $logFile
        if ($success) {
            $systemSummary.SuccessfulUploads++
        } else {
            $systemSummary.FailedUploads++
        }
    }

    if ($systemSummary.SuccessfulUploads -gt 0) {
        Restart-MacroRuntime -endpointIp $ipAddress -username $username -password $password -logFile $logFile
    }

    $uploadSummary += $systemSummary
}

# Generate summary
$successfulSystems = $uploadSummary | Where-Object { $_.SuccessfulUploads -eq $_.TotalFiles }
$partialSystems = $uploadSummary | Where-Object { $_.SuccessfulUploads -gt 0 -and $_.SuccessfulUploads -lt $_.TotalFiles }
$failedSystems = $uploadSummary | Where-Object { $_.SuccessfulUploads -eq 0 }

Display-Message "Upload completed."
Log-Message -message "Upload completed." -logFile $logFile

if ($successfulSystems.Count -gt 0) {
    $message = "Successfully uploaded all macros to $($successfulSystems.Count) systems."
    Display-Message $message
    Log-Message -message $message -logFile $logFile
}

if ($partialSystems.Count -gt 0) {
    $message = "Partially uploaded some macros to $($partialSystems.Count) systems."
    Display-Message $message
    Log-Message -message $message -logFile $logFile
}

if ($failedSystems.Count -gt 0) {
    $message = "$($failedSystems.Count) systems were unsuccessful for all macros."
    Display-Message $message
    Log-Message -message $message -logFile $logFile
}
