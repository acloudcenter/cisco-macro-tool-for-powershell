# Source helper functions
. .\HelperFunctions.ps1

# Initialize log file
$logFile = "UploadPexipOTJMacros.log"

Display-Message "Option 1 selected: Upload Pexip OTJ macros"
Log-Message -Message "Option 1 selected: Upload Pexip OTJ macros" -LogFile $logFile

# Prompt user for CSV file path
$csvFilePath = Get-FilePath -promptMessage "Please enter the full path to the CSV file:"

# Prompt user for the directory containing ZIP files
$zipDirectory = Get-FilePath -promptMessage "Please enter the full path to the directory containing the Pexip OTJ ZIP files:"

# List ZIP files
$zipFiles = Get-ChildItem -Path $zipDirectory -Filter *-otj-macro-latest*.zip -ErrorAction SilentlyContinue

if ($zipFiles.Count -eq 0) {
    Display-Message "No Pexip OTJ ZIP files were found in the specified directory."
    Log-Message -Message "No Pexip OTJ ZIP files were found in the directory $zipDirectory." -LogFile $logFile
    return
}

# Import CSV
try {
    $systems = Import-SystemCsv -Path $csvFilePath
} catch {
    $errorDetails = $_.Exception.Message
    Display-Message "Error importing CSV file. Response: $errorDetails"
    Log-Message -Message "Error importing CSV file. Response: $errorDetails" -LogFile $logFile
    return
}

# Check for matching systems
$matchingSystems = @()
$nonMatchingSystems = 0

foreach ($system in $systems) {
    $systemName = $system.'system name'.Replace(" ", "")
    $matchingZipFiles = $zipFiles | Where-Object { $_.Name -like "*$($systemName.ToLower())*-otj-macro-latest*.zip" }
    if ($matchingZipFiles.Count -gt 0) {
        $matchingSystems += [PSCustomObject]@{
            SystemName = $system.'system name'
            IPAddress  = $system.'ip address'
            Username   = $system.'username'
            Password   = $system.'password'
            ZipFile    = $matchingZipFiles
        }
    } else {
        $nonMatchingSystems++
    }
}

Display-Message "Found $($matchingSystems.Count) systems with matching Pexip OTJ ZIP files."
Display-Message "Found $nonMatchingSystems systems with no matching Pexip OTJ ZIP files."
Log-Message -Message "Found $($matchingSystems.Count) systems with matching Pexip OTJ ZIP files." -LogFile $logFile
Log-Message -Message "Found $nonMatchingSystems systems with no matching Pexip OTJ ZIP files." -LogFile $logFile

# Confirm upload
if (-not (Get-Confirmation -promptMessage "Do you want to proceed with the upload?")) {
    Display-Message "Upload cancelled."
    Log-Message -Message "Upload cancelled." -LogFile $logFile
    return
}

# Function to upload macros from a ZIP file
function Invoke-PexipMacroUpload {
    param (
        [string]$EndpointIp,
        [string]$Username,
        [string]$Password,
        [string]$ZipFilePath,
        [string]$LogFile,
        [string]$SystemName
    )

    $message = "Attempting to upload macros from $ZipFilePath to $EndpointIp ($SystemName)..."
    Display-Message $message
    Log-Message -Message $message -LogFile $LogFile

    $tempDir = Join-Path -Path $env:TEMP -ChildPath ([System.IO.Path]::GetFileNameWithoutExtension([System.IO.Path]::GetRandomFileName()))
    New-Item -ItemType Directory -Path $tempDir | Out-Null

    try {
        Expand-Archive -Path $ZipFilePath -DestinationPath $tempDir -Force
        $jsFiles = Get-ChildItem -Path $tempDir -Filter *.js -ErrorAction Stop

        foreach ($jsFile in $jsFiles) {
            $macroName = [System.IO.Path]::GetFileNameWithoutExtension($jsFile.Name)
            $jsCode = Get-Content -Path $jsFile.FullName -Raw

            $saveSuccess = Save-SystemMacro -EndpointIp $EndpointIp -Username $Username -Password $Password -MacroName $macroName -MacroBody $jsCode -LogFile $LogFile -SystemName $SystemName -NameElement 'name' -IncludeTranspileElement -IncludeCommandAttribute

            if ($saveSuccess) {
                Enable-Macro -EndpointIp $EndpointIp -Username $Username -Password $Password -MacroName $macroName -LogFile $LogFile -SystemName $SystemName | Out-Null
            }
        }

    } catch {
        $errorDetails = $_.Exception.Message
        $message = "Error processing file $ZipFilePath. Response: $errorDetails"
        Display-Message $message
        Log-Message -Message $message -LogFile $LogFile
    } finally {
        Remove-Item -Recurse -Force -Path $tempDir
    }
}

# Perform the upload and restart macro runtimes
$uploadedSystems = 0

foreach ($system in $matchingSystems) {
    Enable-MacrosMode -EndpointIp $system.IPAddress -Username $system.Username -Password $system.Password -LogFile $logFile -SystemName $system.SystemName | Out-Null

    foreach ($zipFile in $system.ZipFile) {
        Invoke-PexipMacroUpload -EndpointIp $system.IPAddress -Username $system.Username -Password $system.Password -ZipFilePath $zipFile.FullName -LogFile $logFile -SystemName $system.SystemName
        $uploadedSystems++
    }

    Restart-MacroRuntime -EndpointIp $system.IPAddress -Username $system.Username -Password $system.Password -LogFile $logFile -SystemName $system.SystemName | Out-Null
}

$message = "Upload completed. Successfully processed $uploadedSystems ZIP packages."
Display-Message $message
Log-Message -Message $message -LogFile $logFile
