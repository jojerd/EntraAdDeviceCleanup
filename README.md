# Entra AD Device Cleanup
Script to list, disable and cleanup Entra AD stale devices.
Inspiration and original idea was developed by Mohammad Zmaili: https://github.com/mzmaili/AzureADDeviceCleanup

# Reasons for complete rewrite
1.) AzureADDeviceCleanup is no longer able to connect to Entra due to deprecation of MSOnline. 
https://techcommunity.microsoft.com/blog/microsoft-entra-blog/action-required-msonline-and-azuread-powershell-retirement---2025-info-and-resou/4364991

2.) There is still a need for customers to have access to a utility to be able to programmatically list, disable, and cleanup stale devices in their tenant.

3.) Added additional functions and features into this utility after re-writing it to support Microsoft Graph connectivity. (Device Code flow, and Service Principal sign-in support.)

4.) Support for multithreaded operations for large operations (1000+ devices).

As an example:

Sequential:
1000 devices ≈ 15-30 minutes

Parallel (10 threads): 
1000 devices ≈ 3-5 minutes

# Very Important Notes
This source code is freeware and is provided on an "as is" basis without warranties of any kind, whether express or implied, including without limitation warranties that the code is free of defect, fit for a particular purpose or non-infringing. The entire risk as to the quality and performance of the code is with the end user.

It is not advisable to immediately delete a device that appears to be stale because you can't undo a deletion in the case of false positives. As a best practice, disable a device for a grace period before deleting it. In your policy, define a timeframe to disable a device before deleting it.

When configured, BitLocker keys for Windows 10 or newer devices are stored on the device object in Microsoft Entra ID. If you delete a stale device, you also delete the BitLocker keys that are stored on the device. Confirm that your cleanup policy aligns with the actual lifecycle of your device before deleting a stale device.

For more information, please visit: https://docs.microsoft.com/en-us/azure/active-directory/devices/manage-stale-devices

# Script Execution Methods
Executing as a single use, you are presented with a pop-up of options to select as well as the default threshold of 90 days.
90 day threshold can be changed with a custom threshold to whatever works best for you when dealing with stale devices.

<img width="395" height="386" alt="image" src="https://github.com/user-attachments/assets/a112e787-5bc1-4368-a305-e0379816c655" />

# Multithreaded Operation 
Multithreaded operations can be set via GUI as well as Command line. The script will also automatically choose multithreaded (Parellel) operations if more than 25 devices are returned during initial lookup.


# Device Code Flow Execution

.\EntraAdDeviceCleanup.ps1 -Operation Verify -Threshold 90 -UseDeviceCode -ClientID "Your-Client-Id" -TenantId "You-Tenant-Id"

# Service Principal Execution
This is used for an automated scheduled task. When wanting to perform a clean you must use the -Force switch in the command line as I am asking for confirmation on removal during interactive execution and -force bypasses that.

.\EntraAdDeviceCleanup.ps1 -Operation Verify -Threshold 90 -ClientId "Your-Client-Id" -TenantId "Your-Tenant-Id" -ClientSecret "Your-Client-Secret"

To clean a device via service principal scheduled task

.\EntraAdDeviceCleanup.ps1 -Operation CleanDevices -Force -threshold 90 -ClientId "Your-Client-Id" -TenantId "Your-Tenant-Id" -ClientSecret "Your-Client-Secret"

Parallel Processing

.\EntraAdDeviceCleanup.ps1 -Operation CleanDevices -Force -threshold 90 -UseParallel -ThrottleLimit 15 -ClientId "Your-Client-Id" -TenantId "Your-Tenant-Id" -ClientSecret "Your-Client-Secret"

# Command Line

The script can be executed via command line for all options listed in the screen shot above. Just make sure to include -Operation with the desired switch (Verify, VerifyDisabledDevices, DisableDevices, CleanDisabledDevices, CleanDevices)


# Logging

The script has the same capability and requirement to use ImportExcel to output stale and disabled devices depending on switch that is used.
It also creates a transcript for troubleshooting purposes as well as actions taken via Write-Log function that writes to the screen during interactive or writes to a log file when executed via command line.
