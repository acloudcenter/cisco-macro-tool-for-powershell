# Function to prompt user for file paths
function Get-FilePath {
    param (
        [string]$promptMessage
    )
    Write-Host $promptMessage
    return Read-Host
}

# Function to log messages to a file
function Log-Message {
    param (
        [string]$Message,
        [string]$LogFile
    )

    if (-not [string]::IsNullOrWhiteSpace($LogFile)) {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "$timestamp - $Message"
        Add-Content -Path $LogFile -Value $logEntry
    }
}

# Function to display messages in the console
function Display-Message {
    param (
        [string]$Message
    )
    Write-Host $Message
}

# Function to prompt user for confirmation
function Get-Confirmation {
    param (
        [string]$promptMessage
    )
    Write-Host "$promptMessage (y/n)"
    $response = Read-Host
    return $response -eq 'y'
}

# Function to import systems from CSV with validation
function Import-SystemCsv {
    param (
        [string]$Path
    )

    $systems = @(Import-Csv -Path $Path)
    if ($systems -eq $null -or $systems.Count -eq 0) {
        throw "CSV file is empty or improperly formatted."
    }

    return $systems
}

# Convert secure string to plain text
function ConvertTo-PlainText {
    param (
        [System.Security.SecureString]$SecureString
    )

    if ($null -eq $SecureString) {
        return [string]::Empty
    }

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
    }
    finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

# Build common headers for Cisco xAPI requests
function New-CiscoXapiHeader {
    param (
        [string]$Username,
        [string]$Password
    )

    return @{
        "Content-Type"  = "application/xml"
        "Authorization" = "Basic " + [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${Username}:${Password}"))
    }
}

# Wrapper for Invoke-RestMethod calls to Cisco endpoints
function Invoke-CiscoXapiRequest {
    param (
        [string]$EndpointIp,
        [string]$Username,
        [string]$Password,
        [object]$Body,
        [int]$TimeoutSec = 15
    )

    $headers = New-CiscoXapiHeader -Username $Username -Password $Password

    return Invoke-RestMethod -Uri "https://$EndpointIp/putxml" -Method 'POST' -Headers $headers -Body $Body -TimeoutSec $TimeoutSec -SkipCertificateCheck -ErrorAction Stop
}

# Retrieve macros available on a system
function Get-SystemMacros {
    param (
        [string]$EndpointIp,
        [string]$Username,
        [string]$Password,
        [string]$LogFile,
        [string]$SystemName
    )

    $target = if ($SystemName) { "$EndpointIp ($SystemName)" } else { $EndpointIp }
    $message = "Attempting to get macros from $target..."
    Display-Message $message
    Log-Message -Message $message -LogFile $LogFile

    $body = @"
<Command>
    <Macros>
        <Macro>
            <Get/>
        </Macro>
    </Macros>
</Command>
"@

    try {
        $response = Invoke-CiscoXapiRequest -EndpointIp $EndpointIp -Username $Username -Password $Password -Body $body -TimeoutSec 10
        $xmlResponse = [xml]$response
        $macros = $xmlResponse.Command.MacroGetResult.Macro | ForEach-Object { $_.Name }

        if ($null -eq $macros) {
            return @()
        }

        return @($macros)
    }
    catch {
        $errorDetails = $_.Exception.Message
        $message = "Error getting macros from $target. Response: $errorDetails"
        Display-Message $message
        Log-Message -Message $message -LogFile $LogFile
        return @()
    }
}

# Save macro content to a system
function Save-SystemMacro {
    param (
        [string]$EndpointIp,
        [string]$Username,
        [string]$Password,
        [string]$MacroName,
        [string]$MacroBody,
        [string]$LogFile,
        [string]$SystemName,
        [string]$NameElement = 'Name',
        [switch]$IncludeTranspileElement,
        [string]$TranspileValue = 'False',
        [switch]$IncludeCommandAttribute
    )

    $target = if ($SystemName) { "$EndpointIp ($SystemName)" } else { $EndpointIp }
    $message = "Attempting to upload macro $MacroName to $target..."
    Display-Message $message
    Log-Message -Message $message -LogFile $LogFile

    $encodedMacro = [System.Security.SecurityElement]::Escape($MacroBody)
    $commandAttribute = if ($IncludeCommandAttribute.IsPresent) { ' command="True"' } else { '' }
    $transpileLine = if ($IncludeTranspileElement.IsPresent) { "`n                <Transpile>$TranspileValue</Transpile>" } else { '' }

    $body = @"
<Command>
    <Macros>
        <Macro>
            <Save$commandAttribute>
                <$NameElement>$MacroName</$NameElement>
                <body>$encodedMacro</body>
                <overWrite>True</overWrite>$transpileLine
            </Save>
        </Macro>
    </Macros>
</Command>
"@

    try {
        Invoke-CiscoXapiRequest -EndpointIp $EndpointIp -Username $Username -Password $Password -Body $body
        $message = "Macro $MacroName uploaded successfully to $target."
        Display-Message $message
        Log-Message -Message $message -LogFile $LogFile
        return $true
    }
    catch {
        $errorDetails = $_.Exception.Message
        $message = "Error uploading macro $MacroName to $target. Response: $errorDetails"
        Display-Message $message
        Log-Message -Message $message -LogFile $LogFile
        return $false
    }
}

# Enable a macro on a system
function Enable-Macro {
    param (
        [string]$EndpointIp,
        [string]$Username,
        [string]$Password,
        [string]$MacroName,
        [string]$LogFile,
        [string]$SystemName
    )

    $target = if ($SystemName) { "$EndpointIp ($SystemName)" } else { $EndpointIp }
    $message = "Attempting to enable macro $MacroName on $target..."
    Display-Message $message
    Log-Message -Message $message -LogFile $LogFile

    $body = @"
<Command>
    <Macros>
        <Macro>
            <Activate>
                <name>$MacroName</name>
            </Activate>
        </Macro>
    </Macros>
</Command>
"@

    try {
        Invoke-CiscoXapiRequest -EndpointIp $EndpointIp -Username $Username -Password $Password -Body $body
        $message = "Macro $MacroName enabled successfully on $target."
        Display-Message $message
        Log-Message -Message $message -LogFile $LogFile
        return $true
    }
    catch {
        $errorDetails = $_.Exception.Message
        $message = "Error enabling macro $MacroName on $target. Response: $errorDetails"
        Display-Message $message
        Log-Message -Message $message -LogFile $LogFile
        return $false
    }
}

# Enable macros mode on a system
function Enable-MacrosMode {
    param (
        [string]$EndpointIp,
        [string]$Username,
        [string]$Password,
        [string]$LogFile,
        [string]$SystemName
    )

    $target = if ($SystemName) { "$EndpointIp ($SystemName)" } else { $EndpointIp }
    $message = "Attempting to enable macros mode on $target..."
    Display-Message $message
    Log-Message -Message $message -LogFile $LogFile

    $body = "<Configuration><Macros><Mode>On</Mode></Macros></Configuration>"

    try {
        Invoke-CiscoXapiRequest -EndpointIp $EndpointIp -Username $Username -Password $Password -Body $body -TimeoutSec 10
        $message = "Macros mode enabled on $target."
        Display-Message $message
        Log-Message -Message $message -LogFile $LogFile
        return $true
    }
    catch {
        $errorDetails = $_.Exception.Message
        $message = "Error enabling macros mode on $target. Response: $errorDetails"
        Display-Message $message
        Log-Message -Message $message -LogFile $LogFile
        return $false
    }
}

# Restart macro runtime on a system
function Restart-MacroRuntime {
    param (
        [string]$EndpointIp,
        [string]$Username,
        [string]$Password,
        [string]$LogFile,
        [string]$SystemName
    )

    $target = if ($SystemName) { "$EndpointIp ($SystemName)" } else { $EndpointIp }
    $message = "Attempting to restart macro runtime on $target..."
    Display-Message $message
    Log-Message -Message $message -LogFile $LogFile

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
        Invoke-CiscoXapiRequest -EndpointIp $EndpointIp -Username $Username -Password $Password -Body $body
        $message = "Macro runtime restarted successfully on $target."
        Display-Message $message
        Log-Message -Message $message -LogFile $LogFile
        return $true
    }
    catch {
        $errorDetails = $_.Exception.Message
        $message = "Error restarting macro runtime on $target. Response: $errorDetails"
        Display-Message $message
        Log-Message -Message $message -LogFile $LogFile
        return $false
    }
}

# Configure transpile evaluation mode
function Set-TranspileEvaluation {
    param (
        [string]$EndpointIp,
        [string]$Username,
        [string]$Password,
        [string]$LogFile,
        [string]$SystemName,
        [string]$EvaluateTranspiled
    )

    $target = if ($SystemName) { "$EndpointIp ($SystemName)" } else { $EndpointIp }
    $message = "Attempting to set transpile evaluation mode to $EvaluateTranspiled on $target..."
    Display-Message $message
    Log-Message -Message $message -LogFile $LogFile

    $body = "<Configuration><Macros><EvaluateTranspiled>$EvaluateTranspiled</EvaluateTranspiled></Macros></Configuration>"

    try {
        Invoke-CiscoXapiRequest -EndpointIp $EndpointIp -Username $Username -Password $Password -Body $body -TimeoutSec 10
        $message = "Transpile evaluation mode set to $EvaluateTranspiled on $target."
        Display-Message $message
        Log-Message -Message $message -LogFile $LogFile
        return $true
    }
    catch {
        $errorDetails = $_.Exception.Message
        if ($errorDetails -match "Invalid value|Unknown command|Invalid path") {
            $message = "System $target does not support transpile evaluation settings. Continuing with default configuration."
            Display-Message $message
            Log-Message -Message $message -LogFile $LogFile
            return $true
        }

        $message = "Error setting transpile evaluation mode on $target. Response: $errorDetails"
        Display-Message $message
        Log-Message -Message $message -LogFile $LogFile
        return $false
    }
}

# Remove a macro from a system
function Remove-SystemMacro {
    param (
        [string]$EndpointIp,
        [string]$Username,
        [string]$Password,
        [string]$MacroName,
        [string]$LogFile,
        [string]$SystemName
    )

    $target = if ($SystemName) { "$EndpointIp ($SystemName)" } else { $EndpointIp }
    $message = "Attempting to remove macro $MacroName from $target..."
    Display-Message $message
    Log-Message -Message $message -LogFile $LogFile

    $body = @"
<Command>
    <Macros>
        <Macro>
            <Remove>
                <Name>$MacroName</Name>
            </Remove>
        </Macro>
    </Macros>
</Command>
"@

    try {
        Invoke-CiscoXapiRequest -EndpointIp $EndpointIp -Username $Username -Password $Password -Body $body -TimeoutSec 10
        $message = "Macro $MacroName removed successfully from $target."
        Display-Message $message
        Log-Message -Message $message -LogFile $LogFile
        return $true
    }
    catch {
        $errorDetails = $_.Exception.Message
        $message = "Error removing macro $MacroName from $target. Response: $errorDetails"
        Display-Message $message
        Log-Message -Message $message -LogFile $LogFile
        return $false
    }
}

# Remove a UI extension from a system
function Remove-UiExtension {
    param (
        [string]$EndpointIp,
        [string]$Username,
        [string]$Password,
        [string]$UiId,
        [string]$LogFile,
        [string]$SystemName
    )

    $target = if ($SystemName) { "$EndpointIp ($SystemName)" } else { $EndpointIp }
    $message = "Attempting to remove UI extension $UiId from $target..."
    Display-Message $message
    Log-Message -Message $message -LogFile $LogFile

    $body = @"
<Command>
    <UserInterface>
        <Extensions>
            <Panel>
                <Remove>
                    <PanelId>$UiId</PanelId>
                </Remove>
            </Panel>
        </Extensions>
    </UserInterface>
</Command>
"@

    try {
        Invoke-CiscoXapiRequest -EndpointIp $EndpointIp -Username $Username -Password $Password -Body $body -TimeoutSec 10
        $message = "UI extension $UiId removed successfully from $target."
        Display-Message $message
        Log-Message -Message $message -LogFile $LogFile
        return $true
    }
    catch {
        $errorDetails = $_.Exception.Message
        $message = "Error removing UI extension $UiId from $target. Response: $errorDetails"
        Display-Message $message
        Log-Message -Message $message -LogFile $LogFile
        return $false
    }
}

# Upload a UI extension definition to a system
function Save-UiExtension {
    param (
        [string]$EndpointIp,
        [string]$Username,
        [string]$Password,
        [string]$PanelContent,
        [string]$LogFile,
        [string]$SystemName,
        [string]$SourceName
    )

    $target = if ($SystemName) { "$EndpointIp ($SystemName)" } else { $EndpointIp }
    $panelIdMatch = [regex]::Match($PanelContent, '<PanelId>([^<]+)</PanelId>')
    $panelId = if ($panelIdMatch.Success) { $panelIdMatch.Groups[1].Value } else { $SourceName }

    $message = "Attempting to upload UI extension $panelId from $SourceName to $target..."
    Display-Message $message
    Log-Message -Message $message -LogFile $LogFile

    $body = @"
<Command>
    <UserInterface>
        <Extensions>
            <Set>
$PanelContent
            </Set>
        </Extensions>
    </UserInterface>
</Command>
"@

    try {
        Invoke-CiscoXapiRequest -EndpointIp $EndpointIp -Username $Username -Password $Password -Body $body
        $message = "UI extension $panelId uploaded successfully to $target."
        Display-Message $message
        Log-Message -Message $message -LogFile $LogFile
        return $true
    }
    catch {
        $errorDetails = $_.Exception.Message
        $message = "Error uploading UI extension $panelId to $target. Response: $errorDetails"
        Display-Message $message
        Log-Message -Message $message -LogFile $LogFile
        return $false
    }
}
