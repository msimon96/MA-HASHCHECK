# Mandiant Advantage Hash Checker Script

## Overview

This PowerShell script interacts with Mandiant Advantage API to retrieve hash values associated with specified Malware Families. The script allows you to perform various operations related to malware analysis and export the results to an Excel file.

## Prerequisites

- PowerShell 7 or higher [Link to Install PowerShell 7](https://learn.microsoft.com/en-us/shows/it-ops-talk/how-to-install-powershell-7)
- Mandiant Advantage API Key and Secret
- Excel application installed on your machine (required for exporting data to Excel)
- Excel Template

### Download Instructions
Download the script using Git:

```bash
git clone https://github.com/msimon96/MA-HASHCECK.git
```

## Configuration

1. Open `MAConfig.json` file and update the API key and secret:

    ```json
    {
        "advantageApiKey": "<ENTER YOUR API KEY HERE>",
        "advantageApiSecret": "<ENTER YOUR API SECRET HERE>"
    }
    ```

2. Update lines 116 and 468 in the script with your desired application name:

    ```powershell
    $appName = "<YOUR APP NAME>"
    ```

## Usage

### Command Line

Run the script from the command line with the following options:

    ```powershell
    ./MAVT-Single.ps1 [-C <choice>] [-M <malwareFamily>] [-h]
    ```

### Options:

    -C <choice>: Specify the script choice (1, 2, 3, 0).
    -M <malwareFamily>: Specify the malware family name.
    -h or -help: Display help message.

### Examples:

    ```powershell
    ./MAVT-Single.ps1 -C 1 -M MyMalwareFamily
    ./MAVT-Single.ps1 -C 0
    ```

## Notes

- The script uses PowerShell 7 or higher. If your version is lower, the script will attempt to launch PowerShell 7.