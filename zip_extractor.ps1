# Load required .NET assemblies for GUI and ZIP extraction
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO.Compression.FileSystem

# Define the path for the error log file, located in the same folder as this script
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$logFile = Join-Path $scriptDir 'zip_extractor_errors.log'

# Function to write error messages with timestamps to the log file
function Log-Error {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    # Append the timestamped error message to the log file with UTF8 encoding
    "$timestamp ERROR: $Message" | Out-File -FilePath $logFile -Append -Encoding UTF8
}

# Core extraction function that handles different archive formats
function Extract-File {
    param(
        [string]$ZipPath,       # Full path to the archive file
        [string]$DestPath       # Destination folder to extract contents into
    )

    # Check if the source archive file exists
    if (-not (Test-Path $ZipPath)) {
        $err = "Source file not found: $ZipPath"
        Log-Error $err
        Write-Host $err -ForegroundColor Red
        return $false
    }

    # Ensure the destination directory exists, create it if missing
    if (-not (Test-Path $DestPath)) {
        try {
            New-Item -ItemType Directory -Path $DestPath -Force | Out-Null
        } catch {
            $err = "Failed to create destination folder '$DestPath': $_"
            Log-Error $err
            Write-Host $err -ForegroundColor Red
            return $false
        }
    }

    try {
        # Determine the archive extension (e.g. .zip, .tar, .gz)
        $ext = [System.IO.Path]::GetExtension($ZipPath).ToLower()

        # Switch logic to handle extraction differently based on file extension
        switch ($ext) {
            ".zip" {
                # Use .NET ZipFile class to extract ZIP archives (overwrite enabled)
                [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $DestPath, $true)
            }
            ".tar" {
                # Use tar.exe (native in Windows 10+) to extract TAR archives
                tar -xf $ZipPath -C $DestPath
            }
            ".gz" | ".tgz" {
                # Extract gzipped tarballs (tar.gz, tgz)
                tar -xzf $ZipPath -C $DestPath
            }
            default {
                # For other archive formats, try PowerShell's Expand-Archive cmdlet
                Expand-Archive -LiteralPath $ZipPath -DestinationPath $DestPath -Force
            }
        }

        return $true
    }
    catch {
        # Log and display extraction errors
        $err = "Extraction error for '$ZipPath': $($_.Exception.Message)"
        Log-Error $err
        Write-Host $err -ForegroundColor Red
        return $false
    }
}

# Accept optional command-line parameters for terminal use
param(
    [string]$Source,          # Archive file path
    [string]$Destination      # Extraction folder path (optional)
)

# If the script is run with parameters, perform terminal extraction
if ($Source) {
    # Default destination is current directory + archive filename without extension
    if (-not $Destination) {
        $Destination = Join-Path (Get-Location) ([System.IO.Path]::GetFileNameWithoutExtension($Source))
    }
    $success = Extract-File -ZipPath $Source -DestPath $Destination
    if ($success) {
        Write-Host "Extraction completed to: $Destination" -ForegroundColor Green
    } else {
        Write-Host "Check error log at $logFile for details." -ForegroundColor Yellow
    }
    exit
}

# --- GUI Mode starts here ---

# Create the main form (window)
$form = New-Object System.Windows.Forms.Form
$form.Text = "Archive Extractor"
$form.Size = New-Object System.Drawing.Size(400, 200)
$form.StartPosition = "CenterScreen"

# Label to display selected archive file path
$lblFilePath = New-Object System.Windows.Forms.Label
$lblFilePath.Text = "No file selected"
$lblFilePath.AutoSize = $true
$lblFilePath.Location = New-Object System.Drawing.Point(140, 25)
$form.Controls.Add($lblFilePath)

# Label to display selected destination folder path
$lblFolderPath = New-Object System.Windows.Forms.Label
$lblFolderPath.Text = "No destination selected"
$lblFolderPath.AutoSize = $true
$lblFolderPath.Location = New-Object System.Drawing.Point(140, 65)
$form.Controls.Add($lblFolderPath)

# Button to open file picker dialog for archive selection
$btnSelectFile = New-Object System.Windows.Forms.Button
$btnSelectFile.Text = "Select Archive"
$btnSelectFile.Location = New-Object System.Drawing.Point(10, 20)
$btnSelectFile.Width = 120
$form.Controls.Add($btnSelectFile)

# Button to open folder picker dialog for destination selection
$btnSelectFolder = New-Object System.Windows.Forms.Button
$btnSelectFolder.Text = "Select Destination"
$btnSelectFolder.Location = New-Object System.Drawing.Point(10, 60)
$btnSelectFolder.Width = 120
$form.Controls.Add($btnSelectFolder)

# Button to start extraction process
$btnExtract = New-Object System.Windows.Forms.Button
$btnExtract.Text = "Extract"
$btnExtract.Location = New-Object System.Drawing.Point(10, 100)
$btnExtract.Width = 120
$form.Controls.Add($btnExtract)

# Open file dialog setup with archive file filters
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "Archives|*.zip;*.tar;*.tar.gz;*.tgz;*.tar.bz2|All Files|*.*"

# Folder browser dialog for destination selection
$folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog

# Event handler: When "Select Archive" clicked, open file picker dialog
$btnSelectFile.Add_Click({
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        # Update label with selected file path
        $lblFilePath.Text = $openFileDialog.FileName
    }
})

# Event handler: When "Select Destination" clicked, open folder picker dialog
$btnSelectFolder.Add_Click({
    if ($folderBrowserDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        # Update label with selected folder path
        $lblFolderPath.Text = $folderBrowserDialog.SelectedPath
    }
})

# Event handler: When "Extract" clicked, validate selections and extract
$btnExtract.Add_Click({
    # Validate if both file and folder are selected
    if (($lblFilePath.Text -eq "No file selected") -or ($lblFolderPath.Text -eq "No destination selected")) {
        # Show warning message box if missing selections
        [System.Windows.Forms.MessageBox]::Show(
            "Please select both a file and a destination folder.",
            "Warning",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Attempt extraction
    $success = Extract-File -ZipPath $lblFilePath.Text -DestPath $lblFolderPath.Text
    if ($success) {
        # Show success message box with destination path
        [System.Windows.Forms.MessageBox]::Show(
            "Extraction completed.`nPath: $($lblFolderPath.Text)",
            "Success",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    } else {
        # On failure, instruct user to check the log
        [System.Windows.Forms.MessageBox]::Show(
            "Extraction failed. Check the log file for details:`n$logFile",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
})

# Show the form and start the GUI event loop
[void]$form.ShowDialog()
