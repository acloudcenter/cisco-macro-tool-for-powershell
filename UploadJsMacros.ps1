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
try {
    $systems = @(Import-Csv -Path $csvFilePath)
    if ($systems -eq $null -or $systems.Count -eq 0) {
        throw "CSV file is empty or improperly formatted."
    }
    Display-Message "Found $($systems.Count) systems in the CSV file."
    Log-Message -message "Found $($systems.Count) systems in the CSV file." -logFile $logFile
} catch {
    $errorDetails = $_.Exception.Message
    Display-Message "Error importing CSV file. Response: $errorDetails"
    Log-Message -message "Error importing CSV file. Response: $errorDetails" -logFile $logFile
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
        [string]$logFile,
        [string]$systemName
    )

    $macroName = [System.IO.Path]::GetFileNameWithoutExtension($jsFilePath)
    $jsCode = Get-Content -Path $jsFilePath -Raw

    $headers = @{
        "Content-Type"  = "application/xml"
        "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${username}:${password}"))
    }

    $encodedJsCode = [System.Security.SecurityElement]::Escape($jsCode)

    $bodyTemplate = @"
<Command>
    <Macros>
        <Macro>
            <Save>
                <Name>$macroName</Name>
                <body>$encodedJsCode</body>
                <overWrite>True</overWrite>
            </Save>
        </Macro>
    </Macros>
</Command>
"@

    $body = $bodyTemplate
    $bodyBytes = [Text.Encoding]::UTF8.GetBytes($body)

    $headers["Content-Length"] = $bodyBytes.Length

    $message = "Attempting to upload macro from $jsFilePath to $endpointIp ($systemName)..."
    Display-Message $message
    Log-Message -message $message -logFile $logFile

    try {
        $response = Invoke-RestMethod -Uri "https://$endpointIp/putxml" -Method 'POST' -Headers $headers -Body $bodyBytes -TimeoutSec 15 -SkipCertificateCheck
        $message = "Macro $macroName uploaded successfully to $endpointIp ($systemName) from $jsFilePath."
        Display-Message $message
        Log-Message -message $message -logFile $logFile

        # Enable the uploaded macro
        Enable-Macro -endpointIp $endpointIp -username $username -password $password -macroName $macroName -logFile $logFile -systemName $systemName

        return $true
    } catch {
        $errorDetails = $_.Exception.Message
        $message = "Error uploading macro $macroName to: $endpointIp ($systemName) from $jsFilePath. Response: $errorDetails"
        Display-Message $message
        Log-Message -message $message -logFile $logFile
        return $false
    }
}

# Function to enable a macro on a system
function Enable-Macro {
    param (
        [string]$endpointIp,
        [string]$username,
        [string]$password,
        [string]$macroName,
        [string]$logFile,
        [string]$systemName
    )

    $message = "Attempting to enable macro $macroName on $endpointIp ($systemName)..."
    Display-Message $message
    Log-Message -message $message -logFile $logFile

    $headers = @{
        "Content-Type"  = "application/xml"
        "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${username}:${password}"))
    }

    $body = @"
<Command>
    <Macros>
        <Macro>
            <Activate>
                <name>$macroName</name>
            </Activate>
        </Macro>
    </Macros>
</Command>
"@

    try {
        $response = Invoke-RestMethod -Uri "https://$endpointIp/putxml" -Method 'POST' -Headers $headers -Body $body -TimeoutSec 15 -SkipCertificateCheck
        $message = "Macro $macroName enabled successfully on $endpointIp ($systemName)."
        Display-Message $message
        Log-Message -message $message -logFile $logFile
    } catch {
        $errorDetails = $_.Exception.Message
        $message = "Error enabling macro $macroName on $endpointIp ($systemName). Response: $errorDetails"
        Display-Message $message
        Log-Message -message $message -logFile $logFile
    }
}

# Function to enable macros mode
function Enable-MacrosMode {
    param (
        [string]$endpointIp,
        [string]$username,
        [string]$password,
        [string]$logFile,
        [string]$systemName
    )

    $message = "Attempting to enable macros mode on $endpointIp ($systemName)..."
    Display-Message $message
    Log-Message -message $message -logFile $logFile

    $headers = @{
        "Content-Type"  = "application/xml"
        "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${username}:${password}"))
    }

    $body = "<Configuration><Macros><Mode>On</Mode></Macros></Configuration>"

    try {
        $response = Invoke-RestMethod -Uri "https://$endpointIp/putxml" -Method 'POST' -Headers $headers -Body $body -TimeoutSec 10 -SkipCertificateCheck
        $message = "Macros mode enabled on $endpointIp ($systemName)"
        Display-Message $message
        Log-Message -message $message -logFile $logFile
    } catch {
        $errorDetails = $_.Exception.Message
        $message = "Error enabling macros mode on $endpointIp ($systemName). Response: $errorDetails"
        Display-Message $message
        Log-Message -message $message -logFile $logFile
    }
}

# Function to restart macro runtime
function Restart-MacroRuntime {
    param (
        [string]$endpointIp,
        [string]$username,
        [string]$password,
        [string]$logFile,
        [string]$systemName
    )

    $message = "Attempting to restart macro runtime on $endpointIp ($systemName)..."
    Display-Message $message
    Log-Message -message $message -logFile $logFile

    $headers = @{
        "Content-Type"  = "application/xml"
        "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${username}:${password}"))
    }

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
        $response = Invoke-RestMethod -Uri "https://$endpointIp/putxml" -Method 'POST' -Headers $headers -Body $body -TimeoutSec 15 -SkipCertificateCheck
        $message = "Macro runtime restarted successfully on $endpointIp ($systemName)."
        Display-Message $message
        Log-Message -message $message -logFile $logFile
    } catch {
        $errorDetails = $_.Exception.Message
        $message = "Error restarting macro runtime on $endpointIp ($systemName). Response: $errorDetails"
        Display-Message $message
        Log-Message -message $message -logFile $logFile
    }
}

# Function to set transpile evaluation mode
function Set-TranspileEvaluation {
    param (
        [string]$endpointIp,
        [string]$username,
        [string]$password,
        [string]$logFile,
        [string]$systemName,
        [string]$evaluateTranspiled
    )

    $message = "Attempting to set transpile evaluation mode to $evaluateTranspiled on $endpointIp ($systemName)..."
    Display-Message $message
    Log-Message -message $message -logFile $logFile

    $headers = @{
        "Content-Type"  = "application/xml"
        "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${username}:${password}"))
    }

    $body = "<Configuration><Macros><EvaluateTranspiled>$evaluateTranspiled</EvaluateTranspiled></Macros></Configuration>"

    try {
        $response = Invoke-RestMethod -Uri "https://$endpointIp/putxml" -Method 'POST' -Headers $headers -Body $body -TimeoutSec 10 -SkipCertificateCheck
        $message = "Transpile evaluation mode set to $evaluateTranspiled on $endpointIp ($systemName)"
        Display-Message $message
        Log-Message -message $message -logFile $logFile
        return $true
    } catch {
        $errorDetails = $_.Exception.Message
        if ($errorDetails -match "Invalid value|Unknown command|Invalid path") {
            $message = "System $endpointIp ($systemName) does not support transpile evaluation settings. Continuing with default configuration."
            Display-Message $message
            Log-Message -message $message -logFile $logFile
            return $true
        } else {
            $message = "Error setting transpile evaluation mode on $endpointIp ($systemName). Response: $errorDetails"
            Display-Message $message
            Log-Message -message $message -logFile $logFile
            return $false
        }
    }
}

# Perform the upload
$uploadSummary = @()

foreach ($system in $systems) {
    $systemName = $system.'system name'
    $ipAddress = $system.'ip address'
    $username = $system.'username'
    $password = $system.'password'

    # Enable macros mode once per system before any operations
    Enable-MacrosMode -endpointIp $ipAddress -username $username -password $password -logFile $logFile -systemName $systemName

    # Set transpile evaluation mode
    $transpileSuccess = Set-TranspileEvaluation -endpointIp $ipAddress -username $username -password $password -logFile $logFile -systemName $systemName -evaluateTranspiled $evaluateTranspiled
    if (-not $transpileSuccess) {
        Display-Message "Skipping uploads for $ipAddress ($systemName) due to configuration error"
        continue
    }

    $systemSummary = [PSCustomObject]@{
        IPAddress = $ipAddress
        SystemName = $systemName
        TotalFiles = $jsFiles.Count
        SuccessfulUploads = 0
        FailedUploads = 0
    }

    foreach ($jsFile in $jsFiles) {
        # Upload macro
        $success = Upload-MacroFromJs -endpointIp $ipAddress -username $username -password $password -jsFilePath $jsFile.FullName -logFile $logFile -systemName $systemName
        if ($success) {
            $systemSummary.SuccessfulUploads++
        } else {
            $systemSummary.FailedUploads++
        }
    }

    if ($systemSummary.SuccessfulUploads -gt 0) {
        Restart-MacroRuntime -endpointIp $ipAddress -username $username -password $password -logFile $logFile -systemName $systemName
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
