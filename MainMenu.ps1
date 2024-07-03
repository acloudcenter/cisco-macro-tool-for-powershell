# Main script execution
while ($true) {
    Write-Host ""
    Write-Host "Choose an option:"
    Write-Host "1. Upload Pexip OTJ macros"
    Write-Host "2. Upload .js macros"
    Write-Host "3. Remove macros"
    Write-Host "4. Exit"
    $option = Read-Host "Enter your choice (1, 2, 3, or 4)"

    switch ($option) {
        "1" {
            . .\UploadPexipOTJMacros.ps1
        }
        "2" {
            . .\UploadJsMacros.ps1
        }
        "3" {
            . .\RemoveMacros.ps1
        }
        "4" {
            Write-Host "Exiting..."
            exit
        }
        default {
            Write-Host "Invalid option selected. Please choose a valid option."
        }
    }
}
