<#MIT License

Copyright (c) 2025 jojerd

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.#>

[CmdletBinding(DefaultParameterSetName = 'GUI')]
param(
    [Parameter(ParameterSetName = 'CommandLine', Mandatory)]
    [ValidateSet('Verify', 'VerifyDisabledDevices', 'DisableDevices', 'CleanDisabledDevices', 'CleanDevices')]
    [string]$Operation,
    [Parameter(ParameterSetName = 'CommandLine')]
    [int]$ThresholdDays = 90,
    [Parameter(ParameterSetName = 'CommandLine')]
    [string]$OutputPath,
    [parameter(ParameterSetName = 'CommandLine')]
    [switch]$Force,
    [Parameter(ParameterSetName = 'CommandLine')]
    [switch]$UseDeviceCode,
    [Parameter(ParameterSetName = 'CommandLine')]
    [string]$ClientId,
    [Parameter(ParameterSetName = 'CommandLine')]
    [string]$TenantId,
    [Parameter(ParameterSetName = 'CommandLine')]
    [string]$ClientSecret
) 
# Script configuration
$script:Config = @{
    DefaultThreshold   = $ThresholdDays
    ExportPath         = if ($OutputPath) { $OutputPath } else { $PSScriptRoot }
    DateFormat         = "{0:s}"
    ModuleRequirements = @{
        'ImportExcel'                                  = @()  # Empty array means import all commands
        'Microsoft.Graph.Authentication'               = @('Connect-MgGraph')
        'Microsoft.Graph.Identity.DirectoryManagement' = @(
            'Get-MgDevice',
            'Update-MgDevice',
            'Remove-MgDevice'
        )
    }
}
$transcriptPath = Join-Path ($(if ($OutputPath) { $OutputPath } else { $PSScriptRoot })) "DeviceCleanup_$(Get-Date -Format 'yyyyMMdd')_transcript.log"
Start-Transcript -Path $transcriptPath
Function Show-Header {
    Write-Host ""
    Write-Host "Entra AD Device Cleanup Script" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Version: 2.0" -ForegroundColor Green
    Write-Host ("Date: " + $(Get-Date).ToString("F", [System.Globalization.CultureInfo]::CurrentCulture)) -ForegroundColor Yellow
    Write-Host "Created by: Josh Jerdon" -ForegroundColor Green
    Write-Host "Email: jojerd@microsoft.com" -ForegroundColor Green
    Write-Host "GitHub: https://github.com/jojerd/EntraAdDeviceCleanup" -ForegroundColor Yellow
    Write-Host ""
}
Function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to console with color
    switch ($Level) {
        'Info' { Write-Host $logMessage -ForegroundColor Green }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Error' { Write-Host $logMessage -ForegroundColor Red }
    }
    
    # If running automated, write to log file
    if ($PSCmdlet.ParameterSetName -eq 'CommandLine') {
        $logFile = Join-Path $script:Config.ExportPath "DeviceCleanup_$(Get-Date -Format 'yyyyMMdd').log"
        $logMessage | Out-File -FilePath $logFile -Append
    }
}
<#
    ============================
    Entra Application Registration Instructions
    ============================
    1. Go to https://entra.microsoft.com/ (Azure Portal).
    2. Navigate to "Entra Active Directory" > "App registrations".
    3. Click "New registration".
    4. Name your app (e.g., "Entra AD Device Cleanup Script").
    5. Set "Supported account types" as needed.
    6. Redirect URI is NOT required for device code flow or service principal.
    7. Click "Register".
    8. After registration, copy the "Application (client) ID" and "Directory (tenant) ID".
    9. For service principal: Go to "Certificates & secrets" > "New client secret". Copy the value.
    10. Go to "API permissions" > "Add a permission" > "Microsoft Graph" > "Application permissions".
        - Add: Device.Read.All, Device.ReadWrite.All
    11. Click "Grant admin consent" for your organization.
#>

Function Connect-GraphWithDeviceCode {
    param(
        [Parameter(Mandatory)]
        [string]$ClientId,
        [Parameter(Mandatory)]
        [string]$TenantId
    )
    try {
        Write-Log "Connecting to Microsoft Graph using device code flow..." -Level Info
        Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Scopes "Device.Read.All", "Device.ReadWrite.All" -UseDeviceCode -ErrorAction Stop
        Write-Log "Connected to Microsoft Graph successfully (device code flow)." -Level Info
        return $true
    }
    catch {
        Write-Log "Failed to connect (device code flow): $($_.Exception.Message)" -Level Error
        return $false
    }
}
Function Connect-GraphWithServicePrincipal {
    param(
        [Parameter(Mandatory)]
        [string]$ClientId,
        [Parameter(Mandatory)]
        [string]$TenantId,
        [Parameter(Mandatory)]
        [string]$ClientSecret
    )
    try {
        Write-Log "Connecting to Microsoft Graph using service principal..." -Level Info
        Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -ClientSecret $ClientSecret -Scopes "Device.Read.All", "Device.ReadWrite.All" -ErrorAction Stop
        Write-Log "Connected to Microsoft Graph successfully (service principal)." -Level Info
        return $true
    }
    catch {
        Write-Log "Failed to connect (service principal): $($_.Exception.Message)" -Level Error
        return $false
    }
}
Function Initialize-RequiredModules {
    try {
        foreach ($module in $script:Config.ModuleRequirements.Keys) {
            if (-not (Get-Module -ListAvailable -Name $module)) {
                Write-Log "Installing $module..." -Level Warning
                Install-Module -Name $module -Force -Scope CurrentUser -AllowClobber -ErrorAction Stop
            }

            if ($script:Config.ModuleRequirements[$module].Count -gt 0) {
                Import-Module -Name $module -Function $script:Config.ModuleRequirements[$module] -ErrorAction Stop
            }
            else {
                Import-Module -Name $module -ErrorAction Stop
            }
            Write-Log "$module loaded successfully."
        }
        return $true
    }
    catch {
        Write-Log "Failed to initialize modules: $($_.Exception.Message)" -Level Error
        return $false
    }
}

Function Show-MenuForm {
    Add-Type -AssemblyName System.Windows.Forms
    $MenuForm = New-Object System.Windows.Forms.Form
    $MenuForm.Text = "Entra AD Device Cleanup"
    $MenuForm.Size = New-Object System.Drawing.Size(400, 320)
    $MenuForm.StartPosition = "CenterScreen"

    # Define available actions
    $actions = @(
        @{Label = "Verify (List stale devices)"; Switch = "Verify" }
        @{Label = "VerifyDisabledDevices (List stale disabled devices)"; Switch = "VerifyDisabledDevices" }
        @{Label = "DisableDevices (Disable stale devices)"; Switch = "DisableDevices" }
        @{Label = "CleanDisabledDevices (Remove stale disabled devices)"; Switch = "CleanDisabledDevices" }
        @{Label = "CleanDevices (Remove all stale devices)"; Switch = "CleanDevices" }
    )

    # Create radio buttons for actions
    $radioButtons = @()
    $y = 20
    foreach ($action in $actions) {
        $rb = New-Object System.Windows.Forms.RadioButton
        $rb.Text = $action.Label
        $rb.Tag = $action.Switch
        $rb.Location = New-Object System.Drawing.Point(20, $y)
        $rb.Size = New-Object System.Drawing.Size(350, 24)
        $MenuForm.Controls.Add($rb)
        $radioButtons += $rb
        $y += 30
    }
    $radioButtons[0].Checked = $true

    # Add threshold input
    $thresholdLabel = New-Object System.Windows.Forms.Label
    $thresholdLabel.Text = "Threshold Days (default $($script:Config.DefaultThreshold)):"
    $thresholdLabel.Location = New-Object System.Drawing.Point(20, ($y + 10))
    $thresholdLabel.Size = New-Object System.Drawing.Size(180, 20)
    $MenuForm.Controls.Add($thresholdLabel)

    $thresholdBox = New-Object System.Windows.Forms.TextBox
    $thresholdBox.Location = New-Object System.Drawing.Point(200, ($y + 8))
    $thresholdBox.Size = New-Object System.Drawing.Size(60, 20)
    $thresholdBox.Text = $script:Config.DefaultThreshold.ToString()
    $MenuForm.Controls.Add($thresholdBox)

    # Add buttons
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Location = New-Object System.Drawing.Point(220, ($y + 50))
    $okButton.Add_Click({
            # Validate threshold input
            if (-not [int]::TryParse($thresholdBox.Text, [ref]$null)) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Please enter a valid number for Threshold Days.",
                    "Input Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                return
            }

            # Store selected operation and threshold
            $selectedRadio = $radioButtons | Where-Object { $_.Checked }
            if ($selectedRadio) {
                $script:SelectedOperation = $selectedRadio.Tag
                $script:Config.DefaultThreshold = [int]$thresholdBox.Text
                $MenuForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $MenuForm.Close()
            }
        })
    $MenuForm.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancelButton.Location = New-Object System.Drawing.Point(100, ($y + 50))
    $cancelButton.Add_Click({ $MenuForm.Close() })
    $MenuForm.Controls.Add($cancelButton)

    # Set default button and accept/cancel keys
    $MenuForm.AcceptButton = $okButton
    $MenuForm.CancelButton = $cancelButton
    $MenuForm.KeyPreview = $true

    return $MenuForm.ShowDialog()
}

if ($PSCmdlet.ParameterSetName -eq 'CommandLine') {
    if ($SaveCredentials) {
        if (-not $CredentialPath) {
            $CredentialPath = Join-Path $script:Config.ExportPath "AADDeviceCleanup.cred"
        }
        if (-not (Save-Credentials -CredentialPath $CredentialPath)) {
            throw "Failed to save credentials"
        }
        exit 0
    }
    
    if ($CredentialPath -and -not $Credential) {
        $Credential = Get-SavedCredentials -CredentialPath $CredentialPath
        if (-not $Credential) {
            throw "Failed to load credentials from $CredentialPath"
        }
    }
}
Function Initialize-GraphConnection {
    try {
        $null = Invoke-WebRequest -Uri https://adminwebservice.microsoftonline.com/ProvisioningService.svc -ErrorAction Stop
        if ($UseDeviceCode) {
            if (-not $ClientId -or -not $TenantId) {
                throw "ClientId and TenantId are required for device code flow. See script comments for instructions."
            }
            return Connect-GraphWithDeviceCode -ClientId $ClientId -TenantId $TenantId
        }
        elseif ($ClientId -and $TenantId -and $ClientSecret) {
            return Connect-GraphWithServicePrincipal -ClientId $ClientId -TenantId $TenantId -ClientSecret $ClientSecret
        }
        else {
            # Default to interactive login if nothing else specified
            Connect-MgGraph -Scopes "Device.Read.All", "Device.ReadWrite.All" -ErrorAction Stop
            Write-Log "Connected to Microsoft Graph interactively."
            return $true
        }
    }
    catch {
        Write-Log "Failed to connect: $($_.Exception.Message)" -Level Error
        return $false
    }
}
Function Get-StaleDevices {
    param(
        [bool]$OnlyDisabled = $false
    )
    try {
        Write-Log "Retrieving devices from Entra AD..." -Level Info
        $devices = Get-MgDevice -All -ErrorAction Stop
        $lastLogon = (Get-Date).AddDays(-$script:Config.DefaultThreshold)
        
        Write-Log "Filtering devices older than $lastLogon..." -Level Info
        $filtered = $devices | Where-Object {
            $_.ApproximateLastSignInDateTime -lt $lastLogon -and
            ($OnlyDisabled -eq $false -or $_.AccountEnabled -eq $false)
        }
        Write-Log "Found $($filtered.Count) devices matching criteria." -Level Info
        return $filtered
    }
    catch {
        Write-Log "Failed to retrieve devices: $($_.Exception.Message)" -Level Error
        return $null
    }
}

Function Export-DeviceReport {
    param(
        [Parameter(Mandatory)]
        $Devices,
        [Parameter(Mandatory)]
        [string]$Operation
    )
    
    try {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $fileName = "${Operation}_${timestamp}.xlsx"
        $filePath = Join-Path $script:Config.ExportPath $fileName

        # If we're disabling devices, update their status in the report
        $reportData = if ($Operation -eq 'Disable') {
            $Devices | Select-Object @{
                Name = 'DisplayName'; Expression = { $_.DisplayName }
            },
            @{
                Name = 'Id'; Expression = { $_.Id }
            },
            @{
                Name = 'DeviceId'; Expression = { $_.DeviceId }
            },
            @{
                Name = 'OperatingSystem'; Expression = { $_.OperatingSystem }
            },
            @{
                Name = 'OperatingSystemVersion'; Expression = { $_.OperatingSystemVersion }
            },
            @{
                # Force AccountEnabled to false for disabled devices
                Name = 'AccountEnabled'; Expression = { $false }
            },
            @{
                Name = 'ApproximateLastSignInDateTime'; Expression = { $_.ApproximateLastSignInDateTime }
            },
            @{
                Name = 'TrustType'; Expression = { $_.TrustType }
            },
            @{
                Name = 'IsCompliant'; Expression = { $_.IsCompliant }
            },
            @{
                Name = 'IsManaged'; Expression = { $_.IsManaged }
            }
        }
        else {
            $Devices | Select-Object `
            DisplayName,
            Id,
            DeviceId,
            OperatingSystem,
            OperatingSystemVersion,
            AccountEnabled,
            ApproximateLastSignInDateTime,
            TrustType,
            IsCompliant,
            IsManaged
        }

        $reportData | Export-Excel -Path $filePath -WorksheetName "Devices" -AutoSize -TableName "Devices" -ErrorAction Stop
        Write-Log "Report exported to: $filePath" -Level Info
        return $fileName
    }
    catch {
        Write-Log "Failed to export report: $($_.Exception.Message)" -Level Error
        return $null
    }
}

Function Invoke-DeviceOperation {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('List', 'Disable', 'Remove')]
        [string]$Operation,
        
        [Parameter(Mandatory)]
        [bool]$OnlyDisabled
    )

    # Get devices based on criteria
    $devices = Get-StaleDevices -OnlyDisabled $OnlyDisabled
    if (-not $devices -or $devices.Count -eq 0) {
        Write-Log "No devices found matching the criteria." -Level Warning
        return
    }

    Write-Log "Processing $($devices.Count) devices..." -Level Info
    $successCount = 0
    $failCount = 0

    # Process devices based on operation
    if ($Operation -ne 'List') {
        foreach ($device in $devices) {
            try {
                switch ($Operation) {
                    'Disable' {
                        if ($device.AccountEnabled) {
                            Update-MgDevice -DeviceId $device.Id -BodyParameter @{accountEnabled = $false } -ErrorAction Stop
                            Write-Log "Disabled device: $($device.DisplayName)" -Level Info
                            $successCount++
                        }
                        else {
                            Write-Log "Device already disabled: $($device.DisplayName)" -Level Warning
                        }
                    }
                    'Remove' {
                        Remove-MgDevice -DeviceId $device.Id -ErrorAction Stop
                        Write-Log "Removed device: $($device.DisplayName)" -Level Info
                        $successCount++
                    }
                }
            }
            catch {
                Write-Log "Failed to $Operation device $($device.DisplayName): $($_.Exception.Message)" -Level Error
                $failCount++
            }
        }
    }
    else {
        $successCount = $devices.Count
    }

    # Export results
    $reportFile = Export-DeviceReport -Devices $devices -Operation $Operation
    
    # Display summary
    Write-Log "Operation Summary:" -Level Info
    Write-Log "Total devices processed: $($devices.Count)" -Level Info
    Write-Log "Successful operations: $successCount" -Level Info
    if ($failCount -gt 0) {
        Write-Log "Failed operations: $failCount" -Level Error
    }
    if ($reportFile) {
        Write-Log "Report file: $reportFile" -Level Info
    }
}

# Main script execution logic
try {
    Clear-Host
    Show-Header
    # Initialize required modules
    if (-not (Initialize-RequiredModules)) {
        throw "Failed to initialize required modules"
    }
    # Handle GUI/CommandLine modes
    if ($PSCmdlet.ParameterSetName -eq 'GUI') {
        Add-Type -AssemblyName System.Windows.Forms
        $result = Show-MenuForm
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
            Write-Log "Operation cancelled by user." -Level Warning
            exit
        }
        $Operation = $script:SelectedOperation
    }
    elseif (-not $Force) {
        # Confirm destructive operations in command-line mode
        if ($Operation -match 'Clean|Disable') {
            $confirm = Read-Host "Are you sure you want to $Operation devices? (Y/N)"
            if ($confirm -ne 'Y') {
                Write-Log "Operation cancelled by user." -Level Warning
                exit
            }
        }
    }

    # Initialize Graph connection
    if (-not (Initialize-GraphConnection)) {
        throw "Failed to connect to Graph API"
    }

    # Execute requested operation
    Write-Log "Executing operation: $Operation"
    switch ($Operation) {
        "Verify" { 
            Invoke-DeviceOperation -Operation "List" -OnlyDisabled $false 
        }
        "VerifyDisabledDevices" { 
            Invoke-DeviceOperation -Operation "List" -OnlyDisabled $true 
        }
        "DisableDevices" { 
            Invoke-DeviceOperation -Operation "Disable" -OnlyDisabled $false 
        }
        "CleanDisabledDevices" { 
            Invoke-DeviceOperation -Operation "Remove" -OnlyDisabled $true 
        }
        "CleanDevices" { 
            Invoke-DeviceOperation -Operation "Remove" -OnlyDisabled $false 
        }
        default {
            throw "Invalid operation selected: $Operation"
        }
    }
}
catch {
    Write-Log "Script execution failed: $($_.Exception.Message)" -Level Error
    exit 1
}
finally {
    Stop-Transcript
    Write-Log "Script execution completed." -Level Info
}