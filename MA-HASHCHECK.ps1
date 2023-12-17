<#
.SYNOPSIS
Mandiant Advantage Hash Checker Script

.DESCRIPTION
This script pulls down hashes from Mandiant Advantage based on Malware Family and Malware Family Associations.

.CREDIT
Script developed by Mario Simon
Date: 2023-12-17
Contact: https://www.linkedin.com/in/mario-r-simon/
#>

param (
    [int]$C,
    [string]$M,
    [switch]$h,
    [switch]$help
)

if ($C) {
    $choice = $C
}
function Show-Help {
    $helpText = @"
Mandiant Advantage Hash Checker

DESCRIPTION:
This script pulls down hashes from Mandiant Advantage based on Malware Family and Malware Family Associations.

USAGE:
.\MAVT-Single.ps1 [-C <choice>] [-M <malwareFamily>] [-h]

OPTIONS:
-C <choice>           Specify the script choice (1, 2, 3, 0).
-M <malwareFamily>    Specify the malware family name.
-h or -help           Display this help message.

EXAMPLES:
.\MAVT-Single.ps1 -C 1 -M MyMalwareFamily -V 1
.\MAVT-Single.ps1 -C 1 -M AnotherMalwareFamily
.\MAVT-Single.ps1 -C 0

HASH TYPES:
1. MD5
2. SHA256
3. SHA1
4. ALL

"@

    Write-Host $helpText
}

# Check if the script is launched with a help switch
if ($help.IsPresent -or $h.IsPresent) {
    Show-Help
    return
}


# Check if PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "[ ! ] This script requries PowerShell 7 or higher. Launching PowerShell 7..." -ForegroundColor Yellow
    Start-Process pwsh.exe -ArgumentList "-c ./MA-HASHCHECK.ps1" -Verb RunAs
    return
} 

# Declare a global variable
$excelFilePath = $null
$bearerToken = $null


# Initialize a variable to count the hash values
$hashCount = 0

# Define the path to the config file
$configFilePath = Join-Path -Path $PSScriptRoot -ChildPath "MAConfig.json"

Write-Host "Script Current Directory: $($PSScriptRoot)" -ForegroundColor Green
Write-Host "Config File Path: $($configFilePath)" -ForegroundColor Green



function RequestBearerToken {
    <#
    .SYNOPSIS
    Requests a bearer token from the Mandiant Advantages API.
    
    .DESCRIPTION
    Reads the client ID and client secret from the config file and sends a request to the Mandiant Advantages API to obtain a bearer token.
    
    .PARAMETER configFilePath
    The path to the config file.
    
    .OUTPUTS
    [string]
    Returns the bearer token.
    #>

    param (
        [Parameter(Mandatory = $true)]
        [string]$configFilePath
    )
    # Check if the config file exists
    if (Test-Path $configFilePath) { 
        # Read the config file
        $configFile = Get-Content $configFilePath | ConvertFrom-Json
        # Get the client ID and client secret from the config file
        $apiKey = $configFile.advantageApiKey
        $apiSecret = $configFile.advantageApiSecret
    }

    # Build the URL for the Mandiant Advantages API
    $tokenUrl = "https://api.intelligence.mandiant.com/token"
    $appName = "<ENTER APP NAME>"

    # Create the authorization header
    $base64Auth = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${apiKey}:${apiSecret}"))

    # Define the grant_type
    $grantType = "client_credentials"

    # Build the request body
    $body = @{
        "grant_type" = $grantType
    }

    # Create an empty array to store the key-value pairs
    $keyValuePairs = @()

    # Convert the body to a query string
    foreach ($key in $body.Keys) {
        $keyValuePairs += "$($key)=$($body[$key])"
    }

    # Join the key-value pairs with an ampersand
    $bodyQueryString = $keyValuePairs -join "&"

    # Create headers with the API key
    $headers = @{
        "Content-Type" = "application/x-www-form-urlencoded"
        "Accept" = "application/json"
        "X-App-Name" = "$appName"
        "Authorization" = "Basic $base64Auth"
    }

    # Make a request to the Mandiant Advantages API
    $response = Invoke-RestMethod -Uri $tokenUrl -Headers $headers -Body $bodyQueryString -Method Post

    # Extract the and print the access token
    $response.access_token
}

function ExportToExcel {
    <#
    .SYNOPSIS
    Exports hash data to an Excel file.
    
    .DESCRIPTION
    Opens an Excel file, updates the Y/N column, hash column, and malware family column with the provided data, and saves the changes.
    
    .PARAMETER excelFilePath
    The path to the Excel file.
    
    .PARAMETER malwareFamilyColumn
    The column letter for the malware family column.
    
    .PARAMETER ynColumn
    The column letter for the Y/N column.
    
    .PARAMETER hashColumn
    The column letter for the hash column.
    
    .PARAMETER hashes
    An array of hash values.
    
    .PARAMETER malwareFamily
    The name of the malware family.
    
    .EXAMPLE
    ExportToExcel -excelFilePath "C:\Data\Hashes.xlsx" -malwareFamilyColumn "A" -ynColumn "B" -hashColumn "C" -hashes @("hash1", "hash2") -malwareFamily "MalwareFamily"
    #>
    
    param(
        [string]$excelFilePath,
        [string]$malwareFamilyColumn,
        [string]$ynColumn,
        [string]$hashColumn,
        [array]$hashes,
        [string]$malwareFamily 
    )
    # Create a new Excel application object
    $excel = New-Object -ComObject Excel.Application
    $excel.ScreenUpdating = $false

    try {
        # Open the Excel Workbook
        $workbook = $excel.Workbooks.Open($excelFilePath)
        # Get the first worksheet
        $worksheet = $workbook.Worksheets.Item(1)
        # Convert the column letters to numbers
        $hashColumnIndex = [System.Convert]::ToUInt32([System.Convert]::ToChar($hashColumn)) - 64
        $malwareFamilyColumnIndex = [System.Convert]::ToUInt32([System.Convert]::ToChar($malwareFamilyColumn)) - 64
        $ynColumnIndex = [System.Convert]::ToUInt32([System.Convert]::ToChar($ynColumn)) - 64
        # Calculate the range for Y/N column
        $ynColumnRange = $worksheet.range($ynColumn + "2:" + $ynColumn + ($hashes.Count + 1))
        # Set "Y" for the entire Y/N range in one operation
        $ynColumnRange.Value2 = "Y"
        # Calculate the range for the hash column
        $hashRange = $worksheet.range($hashColumn + "2:" + $hashColumn + ($hashes.Count + 1))

        if ($hashes.Count -eq 0) { 
            # Set the hash values in the hash column
            $hashRange.Value2 = "N/A"
        }
        else {
            # Set the value for the entire hash range in one operation
            for ($i = 1; $i -le $hashes.Count; $i++) {
                $hashRange.Item($i).Value2 = $hashes[$i - 1]
            }
        }
        # Calculate the range for the malware family column
        $malwareFamilyRange = $worksheet.range($malwareFamilyColumn + "2:" + $malwareFamilyColumn + ($hashes.Count + 1))
        # Set the malware family name in the malware family column
        $malwareFamilyRange.Value2 = $malwareFamily.ToUpper()
    } catch {
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Red
    } finally {
        # Re-enable screen updating
        $excel.ScreenUpdating = $true
        # Save and close the excel workbook
        if ($null -ne $workbook) {
            $workbook.Save()
            $workbook.Close()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
        # Release the Excel application object
        if ($null -ne $excel) {
            $excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function OpenFileDialogBox {
    <#
    .SYNOPSIS
    Opens a file dialog to select an Excel file.
    
    .DESCRIPTION
    Prompts the user to select an Excel file using a file dialog box.
    
    .OUTPUTS
    [string]
    Returns the path of the selected Excel file.
    #>

    while (-not $excelFilePath -or -not (Test-Path $excelFilePath -PathType Leaf)) {
        # Add-Type -AssemblyName System.Windows.Forms to use the OpenFileDialog
        Add-Type -AssemblyName System.Windows.Forms

        # Prompt the user to select the Excel file
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
        $fileDialog.Title = "Select the Excel File"

        # Show the dialog box and check if the user clicked OK
        if ($fileDialog.ShowDialog() -eq "OK") {
            $excelFilePath = $fileDialog.FileName

            # Continue with the rest of the script using $excelFilePath
            Write-Host "`nSelected Excel File: $excelFilePath`n" -ForegroundColor Green
        } else {
            Write-Host "User cancelled the file selection." -ForegroundColor Red
            return  # Exit the function if the user cancels
        }

        # Remove any surrounding double quotes and trim spaces
        $excelFilePath = $excelFilePath.Trim('"')

        # Check if the file path exists
        if (-not $excelFilePath) {
            Write-Host "Please enter a valid file path." -ForegroundColor Red
        } elseif (-not (Test-Path $excelFilePath -PathType Leaf)) {
            Write-Host "File not found. Please make sure the file exists and try again." -ForegroundColor Red
        }
    }
    # Return the selected file path
    return $excelFilePath
}

function ReleaseExcelObject {
	<#
    .SYNOPSIS
    Releases Excel objects.
    
    .DESCRIPTION
    This function releases the Excel application, workbook, and worksheet objects to free up system resources.
    
    .PARAMETER excel
    The Excel application object.
    
    .PARAMETER workbook
    The Excel workbook object.
    
    .PARAMETER worksheet
    The Excel worksheet object.
    #>
    
    param (
        [object]$excel,
        [object]$workbook,
        [object]$worksheet
    )

    # Release the Excel workbook object
    if ($null -ne $workbook) {
        $workbook.Save()
        $workbook.Close()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    }

    # Release the Excel worksheet object
    if ($null -ne $worksheet) {
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    }

    # Release the Excel application object
    if ($null -ne $excel) {
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

function Get-ColumnLetters {
    <#
    .SYNOPSIS
    Prompts the user to enter a column letter.
    
    .DESCRIPTION
    This function prompts the user to enter a column letter and validates the input using a regular expression pattern.
    
    .PARAMETER prompt
    The prompt message to display to the user. The default prompt is "Enter a column letter".
    
    .PARAMETER default
    The default column letter to use if no input is provided.
    
    .OUTPUTS
    [string]
    The entered column letter.
    #>

    param (
        [string]$prompt,
        [string]$default = ''
    )

    do {
        $columnLetter = Read-Host $prompt
        $columnLetter = if ([string]::IsNullOrEmpty($columnLetter)) { $default } else { $columnLetter.ToUpper() }

        # Define the regex pattern for a single letter from a to z (case-insensitive)
        $regexPattern = "^[A-Za-z]$"

        # Check if the input matches the regex pattern
        if ($columnLetter -notmatch $regexPattern) {
            Write-Host "Error: Invalid input. Please enter a single letter from A to Z." -ForegroundColor Red
        } else {
            break  # Exit the loop if the input is valid
        }
    } while ($true)

    return $columnLetter
}

<#
.SYNOPSIS
This script provides a menu-driven interface to perform various operations related to malware analysis. It allows the user to run a script to perform a Mandiant Advantage check on Malware Families and Associated Malware Families.
 
.DESCRIPTION
The script presents a menu with different options and prompts the user to enter their choice. Based on the selected option, the script performs the corresponding operation. The script utilizes various functions to handle different tasks, such as requesting a bearer token, retrieving hash types, selecting an Excel file, and exporting data to Excel.
 
.NOTES
- This script requires the following functions to be defined: RequestBearerToken, OpenFileDialogBox, Get-ColumnLetters, ExportToExcel.
- The script assumes the existence of an API endpoint for the Mandiant Advantages API.
 
#>

# Main Menu Loop
do {
    Write-Host "`nMenu Options:`n------------"
    Write-Host "1. Run Script - Performs Mandiant Advantage Check - Step 1" -ForegroundColor Green
    Write-Host "0. Exit"

    if (-not $C) {
        $choice = Read-Host "Enter your choice`n"
    }

    switch ($choice) {
        1 { 
            # Request the bearer token and store it
            $bearerToken = RequestBearerToken -configFilePath $configFilePath

            # Prompt the user for the malware family
            Write-Host "The script takes time to execute. It will save the output to the selected file template"
            Write-Host "NOTE - No output after the script completes means no hashes were found. Check manually to verify`n" -ForegroundColor Red
            
            # If -M parameter is provided, set $malwareFamily
            if ($M) {
                $malwareFamily = $M
            } else {
                # If -M parameter is not provided, prompt the user for input
                $malwareFamily = Read-Host "Enter the malware family name"
            }
            #malwareFamily = Read-Host "Enter the malware family name"

            # Mapping of numeric values to hash types
            $hashTypeMappings = @{
                "1" = "md5"
                "2" = "sha256"
                "3" = "sha1"
                "4" = "all"
            }

            if ($V) {
                # If -HashType is provided, get the value from the command line arguments
                $selectedHashTypeKey = $V
            } else {
                # If -HashType is not provided, prompt the user
                $selectedHashTypeKey = Read-Host @"
            Enter the numeric value to retrieve hash type:
            (Default set to 'SHA256' hit Enter to proceed with the default)
            - Enter '1' for MD5 hashes
            - Enter '2' for SHA256 hashes
            - Enter '3' for SHA1 hashes
            - Enter '4' for ALL hash types
"@
            
                # Set the default value to '2' (SHA256)
                $defaultHashTypeKey = '2'
            
                # If the user didn't enter anything, use the default
                if (-not $selectedHashTypeKey) {
                    $selectedHashTypeKey = $defaultHashTypeKey
                }
            }

            # Get the selected hash type based on the user's input
            $selectedHashTypes = $hashTypeMappings[$selectedHashTypeKey]
   
            # Call OpenFileDialogBox function to select the Excel file
            $excelFilePath = OpenFileDialogBox
            
            # Prompt the user to enter the column letters for excel using Get-ColumnLetters function
            $malwareFamilyColumn = Get-ColumnLetters -prompt "Enter the column letter for the malware family name (e.g., A, default is A)" -default 'A'
            $ynColumn = Get-ColumnLetters -prompt "Enter the column letter for 'Available in Mandiant Advantage? (Y/N)' (e.g., B, default is B)" -default 'B'
            $hashColumn = Get-ColumnLetters -prompt "Enter the column letter for the hash column (e.g., E, default is E)" -default 'E'

            # Build the URL for the Mandiant Advantages API
            $apiEndpoint = "https://api.intelligence.mandiant.com/v4/malware/$malwareFamily/indicators?limit=1000"
            $appName = "<ENTER APP NAME HERE>"

            # Create headers with the bearer token
            $headers = @{
                "Authorization" = "Bearer $bearerToken"
                "Accept" = "application/json"
                "X-App-Name" = "$appName"
            }

            # Create an array to store the hash values
            $hashes = @()
            
            try {
                # Make a request to the Mandiant Advantages API
                $response = Invoke-RestMethod -Uri $apiEndpoint -Headers $headers -Method Get

                # Get the first indicator (loop through all indicators as needed)
                $indicators = $response.indicators
                foreach ($indicator in $indicators) {
                    #$id = $indicator.id # Not used in this script, uncomment for debugging
                    #$name = $indicator.name # Not used in this script, uncomment for debugging

                    # Access the associated_hashes field if it exists
                    if ($indicator.PSObject.Properties["associated_hashes"]) {
                        $associatedHashes = $indicator.associated_hashes
                        foreach ($hash in $associatedHashes) {
                            # Access Properties within associated_hashes
                            $hashValue = $hash.value

                            if ($selectedHashTypes -eq "all" -or 
                            ($hashValue.Length -eq 32 -and $selectedHashTypes -eq "md5") -or 
                            ($hashValue.Length -eq 64 -and $selectedHashTypes -eq "sha256") -or 
                            ($hashValue.Length -eq 40 -and $selectedHashTypes -eq "sha1")) {
                                # Append the hash value to the $hashes array
                                $hashes += $hashValue
                                # Increment the total hash count
                                $hashCount++
                            }
                        }
                    }
                }
            } catch {
                #  Print the status code 
                Write-Host "Error: $($_.Exception.Response.StatusCode.value__) The Malware Family in your search was found. Check to ensure you typed the name corectly, if you entered the name correctly then the Malware Family is not available" -ForegroundColor Red
                }
            # Display the total number of hashes found
            Write-Host "`nHashes pulled for malware family: $($malwareFamily.ToUpper())" -ForegroundColor Green
            Write-Host "`nTotal Number of Hashes: $hashCount" -ForegroundColor Green

            # Save all the hashes to a Excel passing $hashes array and $malwareFamily
            ExportToExcel -excelFilePath $excelFilePath -malwareFamilyColumn $malwareFamilyColumn -ynColumn $ynColumn -hashColumn $hashColumn -hashes $hashes -malwareFamily $malwareFamily

            # Make $C and $M null to avoid re-running the script with the same parameters
            $C = $null
            $M = $null
        }
        0 { Write-Host "Exiting the script."; break }
        default { Write-Host "Invalid choice. Please select a valid option" }
        }
} while ($choice -ne "0")