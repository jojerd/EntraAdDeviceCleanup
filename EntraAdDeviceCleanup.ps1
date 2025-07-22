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
SOFTWARE.

.SYNOPSIS
    Entra AD Device Cleanup Script

.DESCRIPTION
    This script provides functionality to manage stale devices in Entra Active Directory (Azure AD).


.PARAMETER
    ThresholdDays
    Specifies the period of the last login.
    Note: The default value is 90 days if this parameter is not configured.

.PARAMETER
    Verify
    Verifies the affected devices that will be deleted when running the PowerShell with 'CleanDevices' parameter.

.PARAMETER
    VerifyDisabledDevices
    Verifies disabled devices that will be deleted when running the PowerShell with 'CleanDisabledDevices' parameter.

.PARAMETER
    DisableDevices
    Disables the stale devices as per the configured threshold.

.PARAMETER
    CleanDisabledDevices
    Removes the stale disabled devices as per the configured threshold.

.PARAMETER
    CleanDevices
    Removed the stale devices as per the configured threshold.

.PARAMETER
    UseParallel
    Uses parallel processing for bulk operations. Automatically enabled for 25+ devices unless explicitly disabled.

.PARAMETER
    ThrottleLimit
    Number of concurrent operations when using parallel processing (1-50). Default is 10.

.Example Device Code Flow
    .\EntraAdDeviceCleanup.ps1 -Operation Verify -Threshold 90 -UseDeviceCode -ClientId "your-client-id" -TenantId "your-tenant-id"

    This example connects to Microsoft Graph using device code flow and verifies stale devices.

.Example Service Principal Flow
    .\EntraAdDeviceCleanup.ps1 -Operation DisableDevices -Threshold 90 -ClientId "your-client-id" -TenantId "your-tenant-id" -ClientSecret "your-client-secret"

    This example connects to Microsoft Graph using service principal and disables stale devices.

.Example Parallel Processing
    .\EntraAdDeviceCleanup.ps1 -Operation CleanDevices -Force -threshold 90 -UseParallel -ThrottleLimit 15 -ClientId "Your-Client-Id" -TenantId "Your-Tenant-Id" -ClientSecret "Your-Client-Secret"

    This example connects to Microsoft Graph using service principal and removes stale devices using parallel processing with 15 concurrent operations.

#>
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
    [string]$ClientSecret,
    [Parameter(ParameterSetName = 'CommandLine')]
    [switch]$UseParallel,
    [Parameter(ParameterSetName = 'CommandLine')]
    [int]$ThrottleLimit = 10
) 
# Script configuration
$script:Config = @{
    DefaultThreshold   = $ThresholdDays
    ExportPath         = if ($OutputPath) { $OutputPath } else { $PSScriptRoot }
    DateFormat         = "{0:s}"
    ParallelProcessing = @{
        DefaultThrottleLimit = 10
        MinDevicesForParallel = 25  # Only use parallel for 25+ devices
        ProgressUpdateInterval = 10  # Update progress every N completed operations
    }
    RetryConfiguration = @{
        MaxRetries = 3
        InitialDelaySeconds = 2
        BackoffMultiplier = 2  # Exponential backoff
        RateLimitRetryDelaySeconds = 60  # Wait time for 429 errors
    }
    Performance = @{
        BatchSize = 100  # Process devices in batches for very large datasets
        MaxConcurrentBatches = 5  # Max batches to process simultaneously
    }
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
    Write-Host "Version: 2.1 (Enterprise Edition)" -ForegroundColor Green
    Write-Host ("Date: " + $(Get-Date).ToString("F", [System.Globalization.CultureInfo]::CurrentCulture)) -ForegroundColor Yellow
    Write-Host "Created by: Josh Jerdon" -ForegroundColor Green
    Write-Host "Email: jojerd@microsoft.com" -ForegroundColor Green
    Write-Host "GitHub: https://github.com/jojerd/EntraAdDeviceCleanup" -ForegroundColor Yellow
    Write-Host ""
    
    # Environment information
    Write-Host "Environment Information:" -ForegroundColor Cyan
    Write-Host "PowerShell Version: $($PSVersionTable.PSVersion)" -ForegroundColor White
    Write-Host "OS: $((Get-CimInstance Win32_OperatingSystem).Caption) $((Get-CimInstance Win32_OperatingSystem).Version)" -ForegroundColor White
    if ($PSCmdlet.ParameterSetName -eq 'CommandLine' -and $UseParallel) {
        Write-Host "Parallel Processing: Enabled (Throttle: $ThrottleLimit)" -ForegroundColor Green
    }
    Write-Host ""
}
# Write Log Function. Writes to screen during interactive mode and to a log file during command line execution.
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

# Retry function with exponential backoff for Graph API operations
Function Invoke-GraphOperationWithRetry {
    param(
        [Parameter(Mandatory)]
        [scriptblock]$Operation,
        [string]$OperationName = "Graph Operation",
        [int]$MaxRetries = $script:Config.RetryConfiguration.MaxRetries,
        [int]$InitialDelay = $script:Config.RetryConfiguration.InitialDelaySeconds
    )
    
    $attempt = 1
    $delay = $InitialDelay
    
    while ($attempt -le ($MaxRetries + 1)) {
        try {
            return & $Operation
        }
        catch {
            $isRateLimit = $_.Exception.Message -match "429|Too Many Requests|throttle"
            $isTransient = $_.Exception.Message -match "timeout|temporary|service unavailable|502|503|504"
            
            if ($attempt -le $MaxRetries -and ($isRateLimit -or $isTransient)) {
                if ($isRateLimit) {
                    $waitTime = $script:Config.RetryConfiguration.RateLimitRetryDelaySeconds
                    Write-Log "Rate limit encountered for $OperationName. Waiting $waitTime seconds before retry $attempt/$MaxRetries..." -Level Warning
                    Start-Sleep -Seconds $waitTime
                } else {
                    Write-Log "Transient error for $OperationName. Retrying in $delay seconds (attempt $attempt/$MaxRetries)..." -Level Warning
                    Start-Sleep -Seconds $delay
                    $delay *= $script:Config.RetryConfiguration.BackoffMultiplier
                }
                $attempt++
            } else {
                # Re-throw the exception if max retries exceeded or non-retryable error
                throw
            }
        }
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
# Connect to Graph API using device code.
# Requires ClientId and TenantId.
# Note: This is useful for interactive scenarios where user consent is required.
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
# Connect to Graph API using service principal.
# Requires ClientId, TenantId, and ClientSecret.
# Note: This is useful for automated scripts where user interaction is not possible.
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
# Check if the required modules are installed and import them.
# If not installed, attempt to install them.
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
# Show a GUI form for selecting operations.
# This is used when the script is run in GUI mode.
Function Show-MenuForm {
    Add-Type -AssemblyName System.Windows.Forms
    $MenuForm = New-Object System.Windows.Forms.Form
    $MenuForm.Text = "Entra AD Device Cleanup"
    $MenuForm.Size = New-Object System.Drawing.Size(420, 400)
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

    # Add parallel processing checkbox
    $parallelCheckbox = New-Object System.Windows.Forms.CheckBox
    $parallelCheckbox.Text = "Use parallel processing (for 25+ devices)"
    $parallelCheckbox.Location = New-Object System.Drawing.Point(20, ($y + 40))
    $parallelCheckbox.Size = New-Object System.Drawing.Size(250, 20)
    $parallelCheckbox.Checked = $true
    $MenuForm.Controls.Add($parallelCheckbox)

    # Add throttle limit input
    $throttleLabel = New-Object System.Windows.Forms.Label
    $throttleLabel.Text = "Concurrent operations:"
    $throttleLabel.Location = New-Object System.Drawing.Point(280, ($y + 40))
    $throttleLabel.Size = New-Object System.Drawing.Size(120, 20)
    $MenuForm.Controls.Add($throttleLabel)

    $throttleBox = New-Object System.Windows.Forms.TextBox
    $throttleBox.Location = New-Object System.Drawing.Point(320, ($y + 60))
    $throttleBox.Size = New-Object System.Drawing.Size(40, 20)
    $throttleBox.Text = $script:Config.ParallelProcessing.DefaultThrottleLimit.ToString()
    $MenuForm.Controls.Add($throttleBox)

    # Add buttons
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Location = New-Object System.Drawing.Point(220, ($y + 100))
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

            # Validate throttle limit input
            if (-not [int]::TryParse($throttleBox.Text, [ref]$null) -or [int]$throttleBox.Text -lt 1 -or [int]$throttleBox.Text -gt 50) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Please enter a valid number for Concurrent Operations (1-50).",
                    "Input Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                )
                return
            }

            # Store selected operation and settings
            $selectedRadio = $radioButtons | Where-Object { $_.Checked }
            if ($selectedRadio) {
                $script:SelectedOperation = $selectedRadio.Tag
                $script:Config.DefaultThreshold = [int]$thresholdBox.Text
                $script:UseParallelFromGUI = $parallelCheckbox.Checked
                $script:ThrottleLimitFromGUI = [int]$throttleBox.Text
                $MenuForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $MenuForm.Close()
            }
        })
    $MenuForm.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancelButton.Location = New-Object System.Drawing.Point(100, ($y + 100))
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
# Initialize Graph connection
# This function will handle device code and service principal authentication as well as interactive login.
# It will also check if the required permissions are granted.
Function Initialize-GraphConnection {
    try {
        $null = Invoke-WebRequest -Uri https://adminwebservice.microsoftonline.com/ProvisioningService.svc -ErrorAction Stop
        if ($UseDeviceCode) {
            if (-not $ClientId -or -not $TenantId) {
                throw "ClientId and TenantId are required for device code flow. See script comments for instructions."
            }
            $connected = Connect-GraphWithDeviceCode -ClientId $ClientId -TenantId $TenantId
        }
        elseif ($ClientId -and $TenantId -and $ClientSecret) {
            $connected = Connect-GraphWithServicePrincipal -ClientId $ClientId -TenantId $TenantId -ClientSecret $ClientSecret
        }
        else {
            # Default to interactive login if nothing else specified
            Connect-MgGraph -Scopes "Device.Read.All", "Device.ReadWrite.All" -ErrorAction Stop
            Write-Log "Connected to Microsoft Graph interactively."
            $connected = $true
        }
        
        if ($connected) {
            # Validate connection and permissions
            return Test-GraphConnection
        }
        return $false
    }
    catch {
        Write-Log "Failed to connect: $($_.Exception.Message)" -Level Error
        return $false
    }
}

# Test Graph connection and validate required permissions
Function Test-GraphConnection {
    try {
        Write-Log "Validating Graph connection and permissions..." -Level Info
        
        # Test basic connectivity
        $context = Get-MgContext -ErrorAction Stop
        if (-not $context) {
            throw "No Graph context found"
        }
        
        Write-Log "Connected to tenant: $($context.TenantId)" -Level Info
        Write-Log "Using account: $($context.Account)" -Level Info
        
        # Test device read permissions
        try {
            $testDevice = Get-MgDevice -Top 1 -ErrorAction Stop
            Write-Log "Device read permissions validated." -Level Info
        }
        catch {
            throw "Missing Device.Read.All permission: $($_.Exception.Message)"
        }
        
        # Test device write permissions (non-destructive test)
        try {
            $testDeviceId = (Get-MgDevice -Top 1 -ErrorAction Stop)[0].Id
            if ($testDeviceId) {
                # This is a read operation but requires write scope
                Get-MgDevice -DeviceId $testDeviceId -ErrorAction Stop | Out-Null
                Write-Log "Device write permissions validated." -Level Info
            }
        }
        catch {
            Write-Log "Warning: Device write permissions may be limited: $($_.Exception.Message)" -Level Warning
        }
        
        return $true
    }
    catch {
        Write-Log "Graph connection validation failed: $($_.Exception.Message)" -Level Error
        return $false
    }
}
# Function to retrieve stale devices based on the configured threshold.
# It filters devices based on their last sign-in date and whether they are disabled.
Function Get-StaleDevices {
    param(
        [bool]$OnlyDisabled = $false
    )
    try {
        Write-Log "Retrieving devices from Entra AD..." -Level Info
        
        # Use retry logic for device retrieval
        $devices = Invoke-GraphOperationWithRetry -Operation {
            Get-MgDevice -All -ErrorAction Stop
        } -OperationName "Get-MgDevice"
        
        $lastLogon = (Get-Date).AddDays(-$script:Config.DefaultThreshold)
        
        Write-Log "Filtering devices older than $lastLogon..." -Level Info
        $filtered = $devices | Where-Object {
            # Handle null/empty last sign-in dates
            $lastSignIn = $_.ApproximateLastSignInDateTime
            if (-not $lastSignIn -or $lastSignIn -eq [DateTime]::MinValue) {
                # If no sign-in date, consider it stale
                $isStale = $true
            } else {
                $isStale = $lastSignIn -lt $lastLogon
            }
            
            $isStale -and ($OnlyDisabled -eq $false -or $_.AccountEnabled -eq $false)
        }
        
        Write-Log "Found $($filtered.Count) devices matching criteria." -Level Info
        
        # Log breakdown by OS for better visibility
        if ($filtered.Count -gt 0) {
            $osBreakdown = $filtered | Group-Object OperatingSystem | Sort-Object Count -Descending
            Write-Log "Device breakdown by OS:" -Level Info
            foreach ($os in $osBreakdown) {
                Write-Log "  $($os.Name): $($os.Count) devices" -Level Info
            }
        }
        
        return $filtered
    }
    catch {
        Write-Log "Failed to retrieve devices: $($_.Exception.Message)" -Level Error
        return $null
    }
}
# Function to export device report to Excel.
# It formats the device data and saves it to an Excel file.
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
# Function to perform device operations (list, disable, remove) with optional parallel processing.
# It retrieves stale devices and performs the specified operation on them.
Function Invoke-DeviceOperation {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('List', 'Disable', 'Remove')]
        [string]$Operation,
        
        [Parameter(Mandatory)]
        [bool]$OnlyDisabled,
        
        [switch]$UseParallel,
        [int]$ThrottleLimit = 0
    )

    # Get devices based on criteria
    $devices = Get-StaleDevices -OnlyDisabled $OnlyDisabled
    if (-not $devices -or $devices.Count -eq 0) {
        Write-Log "No devices found matching the criteria." -Level Warning
        return
    }

    # Determine if we should use parallel processing
    $shouldUseParallel = $UseParallel -and 
                        ($devices.Count -ge $script:Config.ParallelProcessing.MinDevicesForParallel) -and
                        ($Operation -ne 'List')

    if ($ThrottleLimit -le 0) {
        $ThrottleLimit = $script:Config.ParallelProcessing.DefaultThrottleLimit
    }

    Write-Log "Processing $($devices.Count) devices..." -Level Info
    
    # For very large datasets, use batch processing
    if ($devices.Count -gt $script:Config.Performance.BatchSize -and $Operation -ne 'List') {
        Write-Log "Large dataset detected. Using batch processing..." -Level Info
        Invoke-DeviceOperationBatched -Devices $devices -Operation $Operation -UseParallel $shouldUseParallel -ThrottleLimit $ThrottleLimit
    }
    elseif ($shouldUseParallel) {
        Write-Log "Using parallel processing with $ThrottleLimit concurrent operations." -Level Info
        Invoke-DeviceOperationParallel -Devices $devices -Operation $Operation -ThrottleLimit $ThrottleLimit
    }
    else {
        if ($UseParallel -and $devices.Count -lt $script:Config.ParallelProcessing.MinDevicesForParallel) {
            Write-Log "Device count below threshold for parallel processing. Using sequential processing." -Level Info
        }
        Invoke-DeviceOperationSequential -Devices $devices -Operation $Operation
    }
}

# Batch processing function for very large datasets
Function Invoke-DeviceOperationBatched {
    param(
        [Parameter(Mandatory)]
        $Devices,
        [Parameter(Mandatory)]
        [string]$Operation,
        [bool]$UseParallel,
        [int]$ThrottleLimit
    )
    
    $batchSize = $script:Config.Performance.BatchSize
    $totalDevices = $Devices.Count
    $totalBatches = [math]::Ceiling($totalDevices / $batchSize)
    $overallSuccessCount = 0
    $overallFailCount = 0
    
    Write-Log "Processing $totalDevices devices in $totalBatches batches of $batchSize..." -Level Info
    
    for ($batchIndex = 0; $batchIndex -lt $totalBatches; $batchIndex++) {
        $startIndex = $batchIndex * $batchSize
        $endIndex = [math]::Min($startIndex + $batchSize - 1, $totalDevices - 1)
        $currentBatch = $Devices[$startIndex..$endIndex]
        
        Write-Log "Processing batch $($batchIndex + 1) of $totalBatches ($($currentBatch.Count) devices)..." -Level Info
        
        # Process current batch
        $batchResults = if ($UseParallel) {
            Invoke-DeviceOperationParallel -Devices $currentBatch -Operation $Operation -ThrottleLimit $ThrottleLimit -ReturnResults
        } else {
            Invoke-DeviceOperationSequential -Devices $currentBatch -Operation $Operation -ReturnResults
        }
        
        if ($batchResults) {
            $overallSuccessCount += $batchResults.SuccessCount
            $overallFailCount += $batchResults.FailCount
        }
        
        # Brief pause between batches to be respectful to the API
        if ($batchIndex -lt ($totalBatches - 1)) {
            Start-Sleep -Seconds 2
        }
    }
    
    # Show overall summary
    Write-Log "Batch Processing Complete:" -Level Info
    Write-Log "Total devices processed: $totalDevices" -Level Info
    Write-Log "Overall successful operations: $overallSuccessCount" -Level Info
    Write-Log "Overall failed operations: $overallFailCount" -Level Info
    
    # Export results for all devices
    Show-OperationSummary -Devices $Devices -Operation $Operation -SuccessCount $overallSuccessCount -FailCount $overallFailCount
}

# Sequential processing function (original implementation)
Function Invoke-DeviceOperationSequential {
    param(
        [Parameter(Mandatory)]
        $Devices,
        [Parameter(Mandatory)]
        [string]$Operation,
        [switch]$ReturnResults
    )

    $successCount = 0
    $failCount = 0

    # Process devices based on operation
    if ($Operation -ne 'List') {
        foreach ($device in $Devices) {
            try {
                switch ($Operation) {
                    'Disable' {
                        if ($device.AccountEnabled) {
                            # Use retry logic for device operations
                            Invoke-GraphOperationWithRetry -Operation {
                                Update-MgDevice -DeviceId $device.Id -BodyParameter @{accountEnabled = $false } -ErrorAction Stop
                            } -OperationName "Disable device $($device.DisplayName)"
                            
                            Write-Log "Disabled device: $($device.DisplayName)" -Level Info
                            $successCount++
                        }
                        else {
                            Write-Log "Device already disabled: $($device.DisplayName)" -Level Warning
                            $successCount++  # Count as success since desired state is achieved
                        }
                    }
                    'Remove' {
                        # Use retry logic for device operations
                        Invoke-GraphOperationWithRetry -Operation {
                            Remove-MgDevice -DeviceId $device.Id -ErrorAction Stop
                        } -OperationName "Remove device $($device.DisplayName)"
                        
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
        $successCount = $Devices.Count
    }

    # Return results if requested (for batch processing), otherwise show summary
    if ($ReturnResults) {
        return @{
            SuccessCount = $successCount
            FailCount = $failCount
        }
    } else {
        Show-OperationSummary -Devices $Devices -Operation $Operation -SuccessCount $successCount -FailCount $failCount
    }
}

# Parallel processing function using RunspacePool (PowerShell 5.1 compatible)
Function Invoke-DeviceOperationParallel {
    param(
        [Parameter(Mandatory)]
        $Devices,
        [Parameter(Mandatory)]
        [string]$Operation,
        [Parameter(Mandatory)]
        [int]$ThrottleLimit,
        [switch]$ReturnResults
    )

    # Create runspace pool
    $initialSessionState = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
    
    # Import Microsoft Graph module into runspace pool
    try {
        $mgModulePath = (Get-Module Microsoft.Graph.Identity.DirectoryManagement -ListAvailable | Select-Object -First 1).ModuleBase
        if ($mgModulePath) {
            $initialSessionState.ImportPSModule(@('Microsoft.Graph.Identity.DirectoryManagement'))
        }
    }
    catch {
        Write-Log "Warning: Could not pre-import Microsoft Graph module into runspaces. Modules will be imported per runspace." -Level Warning
    }

    $runspacePool = [runspacefactory]::CreateRunspacePool(1, $ThrottleLimit, $initialSessionState, $Host)
    $runspacePool.Open()

    # Enhanced script block with retry logic
    $scriptBlock = {
        param($DeviceId, $DisplayName, $Operation, $AccountEnabled, $RetryConfig)
        
        $result = @{
            DeviceId = $DeviceId
            DisplayName = $DisplayName
            Success = $false
            ErrorMessage = $null
            AlreadyDisabled = $false
            Retries = 0
        }

        # Retry function within runspace
        function Invoke-OperationWithRetry {
            param($Operation, $MaxRetries = 3, $InitialDelay = 2)
            
            $attempt = 1
            $delay = $InitialDelay
            
            while ($attempt -le ($MaxRetries + 1)) {
                try {
                    return & $Operation
                }
                catch {
                    $isRateLimit = $_.Exception.Message -match "429|Too Many Requests|throttle"
                    $isTransient = $_.Exception.Message -match "timeout|temporary|service unavailable|502|503|504"
                    
                    if ($attempt -le $MaxRetries -and ($isRateLimit -or $isTransient)) {
                        if ($isRateLimit) {
                            Start-Sleep -Seconds 60  # Wait longer for rate limits
                        } else {
                            Start-Sleep -Seconds $delay
                            $delay *= 2  # Exponential backoff
                        }
                        $attempt++
                    } else {
                        throw
                    }
                }
            }
        }

        try {
            # Import required modules if not already loaded
            if (-not (Get-Command Update-MgDevice -ErrorAction SilentlyContinue)) {
                Import-Module Microsoft.Graph.Identity.DirectoryManagement -Force -ErrorAction Stop
            }

            switch ($Operation) {
                'Disable' {
                    if ($AccountEnabled) {
                        Invoke-OperationWithRetry -Operation {
                            Update-MgDevice -DeviceId $DeviceId -BodyParameter @{accountEnabled = $false } -ErrorAction Stop
                        } -MaxRetries $RetryConfig.MaxRetries -InitialDelay $RetryConfig.InitialDelaySeconds
                        $result.Success = $true
                    }
                    else {
                        $result.Success = $true
                        $result.AlreadyDisabled = $true
                    }
                }
                'Remove' {
                    Invoke-OperationWithRetry -Operation {
                        Remove-MgDevice -DeviceId $DeviceId -ErrorAction Stop
                    } -MaxRetries $RetryConfig.MaxRetries -InitialDelay $RetryConfig.InitialDelaySeconds
                    $result.Success = $true
                }
            }
        }
        catch {
            $result.ErrorMessage = $_.Exception.Message
        }
        
        return $result
    }

    # Create and start jobs
    $jobs = @()
    $jobIndex = 0
    
    foreach ($device in $Devices) {
        $powershell = [powershell]::Create()
        $powershell.RunspacePool = $runspacePool
        
        $null = $powershell.AddScript($scriptBlock)
        $null = $powershell.AddParameter("DeviceId", $device.Id)
        $null = $powershell.AddParameter("DisplayName", $device.DisplayName)
        $null = $powershell.AddParameter("Operation", $Operation)
        $null = $powershell.AddParameter("AccountEnabled", $device.AccountEnabled)
        $null = $powershell.AddParameter("RetryConfig", $script:Config.RetryConfiguration)
        
        $jobs += [PSCustomObject]@{
            Index = $jobIndex++
            PowerShell = $powershell
            AsyncResult = $powershell.BeginInvoke()
            Device = $device
            StartTime = Get-Date
        }
    }

    Write-Log "Started $($jobs.Count) concurrent operations. Waiting for completion..." -Level Info

    # Monitor and collect results
    $successCount = 0
    $failCount = 0
    $completed = 0
    $progressUpdateInterval = $script:Config.ParallelProcessing.ProgressUpdateInterval
    
    try {
        while ($jobs) {
            $completedJobs = $jobs | Where-Object { $_.AsyncResult.IsCompleted }
            
            foreach ($job in $completedJobs) {
                try {
                    $result = $job.PowerShell.EndInvoke($job.AsyncResult)
                    
                    if ($result -and $result.Count -gt 0) {
                        $jobResult = $result[0]
                        
                        if ($jobResult.Success) {
                            if (-not $jobResult.AlreadyDisabled) {
                                Write-Log "$Operation device: $($jobResult.DisplayName)" -Level Info
                            }
                            else {
                                Write-Log "Device already disabled: $($jobResult.DisplayName)" -Level Warning
                            }
                            $successCount++
                        }
                        else {
                            Write-Log "Failed to $Operation device $($jobResult.DisplayName): $($jobResult.ErrorMessage)" -Level Error
                            $failCount++
                        }
                    }
                    else {
                        Write-Log "Failed to $Operation device $($job.Device.DisplayName): No result returned" -Level Error
                        $failCount++
                    }
                }
                catch {
                    Write-Log "Failed to $Operation device $($job.Device.DisplayName): $($_.Exception.Message)" -Level Error
                    $failCount++
                }
                finally {
                    $job.PowerShell.Dispose()
                }
                
                $completed++
                
                # Show progress updates
                if ($completed % $progressUpdateInterval -eq 0 -or $completed -eq $Devices.Count) {
                    $percentComplete = [math]::Round(($completed / $Devices.Count) * 100, 1)
                    Write-Log "Progress: $completed of $($Devices.Count) devices processed ($percentComplete%)" -Level Info
                }
            }
            
            # Remove completed jobs from monitoring list
            $jobs = $jobs | Where-Object { -not $_.AsyncResult.IsCompleted }
            
            # Small delay to prevent excessive CPU usage while monitoring
            if ($jobs) {
                Start-Sleep -Milliseconds 100
            }
        }
    }
    finally {
        # Cleanup: Dispose any remaining PowerShell instances
        foreach ($job in $jobs) {
            try {
                $job.PowerShell.Dispose()
            }
            catch {
                # Ignore disposal errors
            }
        }
        
        # Close and dispose runspace pool
        try {
            $runspacePool.Close()
            $runspacePool.Dispose()
        }
        catch {
            Write-Log "Warning: Error disposing runspace pool: $($_.Exception.Message)" -Level Warning
        }
    }

    # Return results if requested (for batch processing), otherwise show summary
    if ($ReturnResults) {
        return @{
            SuccessCount = $successCount
            FailCount = $failCount
        }
    } else {
        Show-OperationSummary -Devices $Devices -Operation $Operation -SuccessCount $successCount -FailCount $failCount
    }
}

# Helper function to display operation summary and export results
Function Show-OperationSummary {
    param(
        [Parameter(Mandatory)]
        $Devices,
        [Parameter(Mandatory)]
        [string]$Operation,
        [Parameter(Mandatory)]
        [int]$SuccessCount,
        [Parameter(Mandatory)]
        [int]$FailCount
    )

    # Calculate execution time
    $executionTime = (Get-Date) - $script:OperationStartTime
    
    # Export results
    $reportFile = Export-DeviceReport -Devices $Devices -Operation $Operation
    
    # Display comprehensive summary
    Write-Log "========================================" -Level Info
    Write-Log "Operation Summary:" -Level Info
    Write-Log "========================================" -Level Info
    Write-Log "Operation: $Operation" -Level Info
    Write-Log "Total devices processed: $($Devices.Count)" -Level Info
    Write-Log "Successful operations: $SuccessCount" -Level Info
    if ($FailCount -gt 0) {
        Write-Log "Failed operations: $FailCount" -Level Error
        $successRate = [math]::Round(($SuccessCount / ($SuccessCount + $FailCount)) * 100, 1)
        Write-Log "Success rate: $successRate%" -Level Warning
    } else {
        Write-Log "Success rate: 100%" -Level Info
    }
    
    Write-Log "Execution time: $($executionTime.ToString('hh\:mm\:ss'))" -Level Info
    
    if ($Devices.Count -gt 0) {
        $avgTimePerDevice = $executionTime.TotalSeconds / $Devices.Count
        Write-Log "Average time per device: $([math]::Round($avgTimePerDevice, 2)) seconds" -Level Info
    }
    
    if ($reportFile) {
        Write-Log "Report file: $reportFile" -Level Info
    }
    Write-Log "========================================" -Level Info
}

# Main script execution logic
# This section handles both GUI and command-line modes.
# It initializes required modules, connects to Graph API, and executes the requested operation.   
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
        $UseParallel = $script:UseParallelFromGUI
        $ThrottleLimit = $script:ThrottleLimitFromGUI
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
    $script:OperationStartTime = Get-Date
    Write-Log "Executing operation: $Operation"
    switch ($Operation) {
        "Verify" { 
            Invoke-DeviceOperation -Operation "List" -OnlyDisabled $false -UseParallel:$UseParallel -ThrottleLimit $ThrottleLimit
        }
        "VerifyDisabledDevices" { 
            Invoke-DeviceOperation -Operation "List" -OnlyDisabled $true -UseParallel:$UseParallel -ThrottleLimit $ThrottleLimit
        }
        "DisableDevices" { 
            Invoke-DeviceOperation -Operation "Disable" -OnlyDisabled $false -UseParallel:$UseParallel -ThrottleLimit $ThrottleLimit
        }
        "CleanDisabledDevices" { 
            Invoke-DeviceOperation -Operation "Remove" -OnlyDisabled $true -UseParallel:$UseParallel -ThrottleLimit $ThrottleLimit
        }
        "CleanDevices" { 
            Invoke-DeviceOperation -Operation "Remove" -OnlyDisabled $false -UseParallel:$UseParallel -ThrottleLimit $ThrottleLimit
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