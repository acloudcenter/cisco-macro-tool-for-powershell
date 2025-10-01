# Cisco Macro Tool For PowerShell

> **Note**: This is a community tool and is not officially supported by Pexip or Cisco.

> **Note**: Cisco is making changes to their Macro engine. This means that transpilation will be deprecated. Depending on your device's firmware version, you may or may not need to change the transpile settings when uploading macros. Check your device's firmware documentation for JavaScript runtime support.


A PowerShell-based utility for managing macros and UI extensions on Cisco video systems. This tool simplifies the process of deploying, managing, and removing macros across single or multiple Cisco endpoints.

## Features

### Primary Functions
- **Upload Pexip OTJ Macros**: Deploy Pexip One-Touch-Join macros from ZIP packages
- **Upload JS Macros**: Deploy custom JavaScript macros to systems
  - Optional transpile evaluation configuration
  - Automatic macro mode enabling
  - Supports both modern and legacy Cisco firmware
- **Remove Macros**: Bulk remove specific macros from systems
- **Remove UI Extensions**: Clean up UI extensions from systems

### Secondary Functions
- **System Checks**: 
  - Verify macro status on a single system
  - Audit macros across multiple systems
- **Logging**: Detailed operation logs for troubleshooting

### Recent Updates
- Added automatic macro mode enabling before uploads
- Added transpile evaluation configuration option
  - Users can choose to enable/disable transpile evaluation
  - Gracefully handles systems that don't support this feature
  - Default setting is disabled
- Improved error handling and logging for macro mode operations
- Optimized upload process by enabling macro mode once per system

## Prerequisites

### System Requirements
- PowerShell 7.5.0 or later (PowerShell 5.1 not supported)
- Windows operating system (MacOS not supported)
- Network access to Cisco endpoints (port 443)

### Network Requirements
- Direct network connectivity to Cisco systems
- Administrative access to target Cisco endpoints

### Files Needed
- CSV file with system details (template provided in `macro_template_pwsh.csv`)
  - Required columns: system name, ip address, username, password
  - One row per system to be managed
- Macro files (.js format or Pexip OTJ ZIP packages)

## Security Notes

The tool uses `-SkipCertificateCheck` when connecting to systems to handle self-signed certificates. If certificate validation is required:
1. Remove the `-SkipCertificateCheck` parameter from `Invoke-RestMethod` calls
2. Ensure proper certificates are installed on endpoints

## Installation

1. Clone the repository:
   ```powershell
   git clone https://github.com/Josh-E-S/Cisco-Macro-Tool-For-PowerShell.git
   ```
2. Navigate to the tool directory:
   ```powershell
   cd Cisco-Macro-Tool-For-PowerShell
   ```

## Usage

1. Prepare your CSV file using the provided template (`macro_template_pwsh.csv`)
2. Run the main script:
   ```powershell
   .\MainMenu.ps1
   ```
3. Choose from the available options:
   - 1: Upload Pexip OTJ macros
   - 2: Upload .js macros
   - 3: Remove macros
   - 4: Check macros on a specific system
   - 5: Check macros on all systems
   - 6: Remove UI extensions
   - 7: Exit

## Logging

- Each operation generates a detailed log file
- Logs are saved in the same directory as the scripts
- Log files follow the naming pattern: `Operation.log` (e.g., `UploadJsMacros.log`)

## Contributing

Contributions are welcome! Please feel free to submit pull requests.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
