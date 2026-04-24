[CmdletBinding()]
param (
    [switch]$TestMode
)

# Suppress the Spectre Console text encoding warning
$env:IgnoreSpectreEncoding = $true
<<<<<<< HEAD
=======

# ==========================================
# 0. PREREQUISITES & MODULES
# ==========================================
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.DeviceManagement", "PwshSpectreConsole")
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing missing module: $module..." -ForegroundColor Yellow
        Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
    }
    Import-Module $module -ErrorAction Stop
}

# Clear previous cached sessions
try { Disconnect-MgGraph -Confirm:$false -ErrorAction SilentlyContinue } catch { }

# Connect to Microsoft Graph using Device Code flow AND force it into memory (Process scope)
Connect-MgGraph -Scopes "DeviceManagementManagedDevices.ReadWrite.All", "User.Read.All", "Directory.Read.All" -NoWelcome -UseDeviceAuthentication -ContextScope Process
>>>>>>> d874abec0309bd3c130528a64bbaea4c6b3fe2e0

# ==========================================
# 0. PREREQUISITES & MODULES
# ==========================================
<<<<<<< HEAD
$requiredModules = @("Microsoft.Graph.Authentication", "Microsoft.Graph.DeviceManagement", "PwshSpectreConsole")
foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing missing module: $module..." -ForegroundColor Yellow
        
        # Temporarily trust PSGallery to prevent confirmation prompts from hanging the script
        $currentPolicy = (Get-PSRepository -Name 'PSGallery' -ErrorAction SilentlyContinue).InstallationPolicy
        Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted -ErrorAction SilentlyContinue
        
        Install-Module -Name $module -Force -AllowClobber -Scope CurrentUser
        
        # Restore previous policy
        if ($currentPolicy) { Set-PSRepository -Name 'PSGallery' -InstallationPolicy $currentPolicy -ErrorAction SilentlyContinue }
    }
    Import-Module $module -ErrorAction Stop
=======
if (-not $TestMode) {
    $allDevices = Invoke-SpectreCommandWithStatus -Title "Downloading Device List..." -Spinner "Dots2" -Color Cyan -ScriptBlock {
        $devices = Get-MgDeviceManagementManagedDevice -All -Property "id,deviceName,serialNumber,userPrincipalName,model,operatingSystem,osVersion,complianceState,lastSyncDateTime,managementAgent,totalStorageSpaceInBytes,freeStorageSpaceInBytes,azureADDeviceId"
        return $devices
    }
    
    if (-not $allDevices) {
        Write-Host "`n[ERROR] Failed to retrieve devices from Intune. The list is empty." -ForegroundColor Red
        Write-Host "Please verify your Graph permissions and ensure devices exist in the target tenant." -ForegroundColor Red
        exit
    }
} else {
    # SANITIZED FOR GITHUB: Replaced identifying initials with generic Contoso data
    $allDevices = @(
        [PSCustomObject]@{ Id = "11111111-2222-3333-4444-555555555555"; deviceName = "LAPTOP-CORP-01"; serialNumber = "PF3B99X"; userPrincipalName = "adelev@contoso.com"; model = "Surface Laptop 5"; operatingSystem = "Windows"; osVersion = "10.0.22631"; complianceState = "compliant"; lastSyncDateTime = (Get-Date).AddMinutes(-45); managementAgent = "intune"; totalStorageSpaceInBytes = 256GB; freeStorageSpaceInBytes = 45GB; azureADDeviceId = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" }
        [PSCustomObject]@{ Id = "66666666-7777-8888-9999-000000000000"; deviceName = "DESKTOP-DEV-02"; serialNumber = "MXL1234"; userPrincipalName = "meganb@contoso.com"; model = "OptiPlex 7090"; operatingSystem = "Windows"; osVersion = "10.0.19045"; complianceState = "noncompliant"; lastSyncDateTime = (Get-Date).AddDays(-3); managementAgent = "intune"; totalStorageSpaceInBytes = 512GB; freeStorageSpaceInBytes = 12GB; azureADDeviceId = "bbbbbbbb-cccc-dddd-eeee-ffffffffffff" }
    )
>>>>>>> d874abec0309bd3c130528a64bbaea4c6b3fe2e0
}
# Clear previous cached sessions
try { Disconnect-MgGraph -Confirm:$false -ErrorAction SilentlyContinue } catch { }

# Connect to Microsoft Graph using the interactive pop-up
Write-SpectreHost "Prompting for authentication..." -Color Yellow
Connect-MgGraph -Scopes "DeviceManagementManagedDevices.ReadWrite.All", "User.Read.All", "Directory.Read.All" -ContextScope Process

# Validate Connection so we don't proceed without auth
$graphContext = Get-MgContext
if ($graphContext) {
    Write-SpectreHost "Connected as: $($graphContext.Account)" -Color Green
    Start-Sleep -Seconds 2
} else {
    Write-SpectreHost "Failed to authenticate to Microsoft Graph. Exiting..." -Color Red
    exit
}


# ==========================================
# 1. HELPER FUNCTIONS (API Calls)
# ==========================================
function Get-HDDeviceApps {
    param([string]$DeviceId)
    try {
        $appUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$DeviceId/detectedApps?`$top=5"
        $appResponse = Invoke-MgGraphRequest -Method GET -Uri $appUri -ErrorAction Stop
        return $appResponse.value | Select-Object @{n='Name';e={$_.displayName}}, @{n='Ver';e={$_.version}}
    } catch { return $null }
}

function Get-HDEntraGroups {
    param([string]$AzureADDeviceId)
    if ([string]::IsNullOrWhiteSpace($AzureADDeviceId)) { return $null }
    try {
        $entraUri = "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$AzureADDeviceId'&`$select=id"
        $entraDevice = Invoke-MgGraphRequest -Method GET -Uri $entraUri -ErrorAction Stop
        
        if ($entraDevice.value) {
            $entraId = $entraDevice.value[0].id
            $groupsUri = "https://graph.microsoft.com/v1.0/devices/$entraId/memberOf?`$select=displayName"
            $groupsResponse = Invoke-MgGraphRequest -Method GET -Uri $groupsUri -ErrorAction Stop
            return $groupsResponse.value | Select-Object @{n='Group';e={$_.displayName}}
        }
    } catch { return $null }
}

function Get-HDDeviceConfig {
    param([string]$DeviceId)
    try {
        $configUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$DeviceId/deviceConfigurationStates?`$top=5"
        $configResponse = Invoke-MgGraphRequest -Method GET -Uri $configUri -ErrorAction Stop
        return $configResponse.value | Select-Object @{n='Policy';e={$_.displayName}}, @{n='State';e={$_.state}}
    } catch { return $null }
}


# ==========================================
# 2. MAIN APPLICATION LOOP
# ==========================================
while ($true) {
    Clear-Host
<<<<<<< HEAD
=======
    
>>>>>>> d874abec0309bd3c130528a64bbaea4c6b3fe2e0
    Write-SpectreFigletText -Text "Intune Help Desk" -Color Cyan
    "----------------------------------------------------------------" | Out-SpectreHost
    Write-Host ""
    
<<<<<<< HEAD
    # ------------------------------------------
    # SEARCH & SELECTION (Optimized for scale)
    # ------------------------------------------
    if ($TestMode) {
        $searchTerm = Read-SpectreText -Message "Enter search term (Test Mode: press ENTER for all, or 'exit')"
        if ($searchTerm -eq 'exit') { break }
        
        $devices = @(
            [PSCustomObject]@{ Id = "11111111-2222-3333-4444-555555555555"; deviceName = "LAPTOP-CORP-01"; serialNumber = "PF3B99X"; userPrincipalName = "adelev@contoso.com"; model = "Surface Laptop 5"; operatingSystem = "Windows"; osVersion = "10.0.22631"; complianceState = "compliant"; lastSyncDateTime = (Get-Date).AddMinutes(-45); managementAgent = "intune"; totalStorageSpaceInBytes = 256GB; freeStorageSpaceInBytes = 45GB; azureADDeviceId = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" }
            [PSCustomObject]@{ Id = "66666666-7777-8888-9999-000000000000"; deviceName = "DESKTOP-DEV-02"; serialNumber = "MXL1234"; userPrincipalName = "meganb@contoso.com"; model = "OptiPlex 7090"; operatingSystem = "Windows"; osVersion = "10.0.19045"; complianceState = "noncompliant"; lastSyncDateTime = (Get-Date).AddDays(-3); managementAgent = "intune"; totalStorageSpaceInBytes = 512GB; freeStorageSpaceInBytes = 12GB; azureADDeviceId = "bbbbbbbb-cccc-dddd-eeee-ffffffffffff" }
        )
    } else {
        $searchTerm = Read-SpectreText -Message "Enter a device name or serial number (or type 'exit')"
        if ($searchTerm -eq 'exit' -or [string]::IsNullOrWhiteSpace($searchTerm)) { break }

        $devices = Invoke-SpectreCommandWithStatus -Title "Searching Intune..." -Spinner "Dots2" -Color Cyan -ScriptBlock {
            Get-MgDeviceManagementManagedDevice -Filter "startswith(deviceName, '$searchTerm') or serialNumber eq '$searchTerm'" -Property "id,deviceName,serialNumber,userPrincipalName,model,operatingSystem,osVersion,complianceState,lastSyncDateTime,managementAgent,totalStorageSpaceInBytes,freeStorageSpaceInBytes,azureADDeviceId"
        }
    }

    if (-not $devices) {
        Write-SpectreHost "`nNo devices found matching '$searchTerm'." -Color Yellow
        Start-Sleep -Seconds 2
        continue
    }
=======
    $choices = $allDevices | ForEach-Object { "$($_.deviceName) | $($_.serialNumber) | $($_.Id)" }
    $selectionString = Read-SpectreSelection -Title "Search Device (ESC to exit)" -Choices $choices -EnableSearch -PageSize 15
    
    if (-not $selectionString) { break } 

    $selectedId = ($selectionString -split "\|")[-1].Trim()
    $targetDevice = $allDevices | Where-Object { $_.Id -eq $selectedId } | Select-Object -First 1

    Clear-Host
    "Diagnostics: [bold cyan]$($targetDevice.deviceName)[/]" | Out-SpectreHost
    "----------------------------------------------------------------" | Out-SpectreHost
    Write-Host ""

    # ------------------------------------------
    # ROW 1: IDENTITY & HEALTH PANELS
    # ------------------------------------------
    $idText = "Serial: [bold]$($targetDevice.serialNumber)[/]`nUser: [deepskyblue1]$($targetDevice.userPrincipalName)[/]`nModel: $($targetDevice.model)"
    $identityPanel = $idText | Format-SpectrePanel -Header "Identity"

    $statusColor = if($targetDevice.complianceState -eq "compliant") { "green" } else { "red" }
    $healthText = "Status: [$statusColor]$($targetDevice.complianceState)[/]`nSync: $($targetDevice.lastSyncDateTime)"
    $healthPanel = $healthText | Format-SpectrePanel -Header "Health"

    @($identityPanel, $healthPanel) | Format-SpectreColumns | Out-SpectreHost
    Write-Host ""

    # ------------------------------------------
    # ROW 2: STORAGE CHART
    # ------------------------------------------
    try {
        $total = [Math]::Round($targetDevice.totalStorageSpaceInBytes / 1GB, 0)
        $free = [Math]::Round($targetDevice.freeStorageSpaceInBytes / 1GB, 0)
        $used = $total - $free
        
        $chartItems = @(
            (New-SpectreChartItem -Label "Used ($used GB)" -Value $used -Color Red),
            (New-SpectreChartItem -Label "Free ($free GB)" -Value $free -Color Green)
        )
        $chartItems | Format-SpectreBreakdownChart | Format-SpectrePanel -Header "Storage" | Out-SpectreHost
    } catch { 
        "Storage metrics unavailable." | Format-SpectrePanel -Header "Storage" | Out-SpectreHost
    }
    Write-Host ""

    # ------------------------------------------
    # ROW 3: LIVE GRAPH DATA GATHERING
    # ------------------------------------------
    $appData = $null
    $groupData = $null
    $configData = $null

    if ($TestMode) {
        $appData = @([PSCustomObject]@{ Name = "Chrome"; Ver = "124" })
        $groupData = @([PSCustomObject]@{ Group = "SG-IT-Admins" })
        $configData = @([PSCustomObject]@{ Policy = "Win10_SecBaseline"; State = "compliant" })
    } else {
        # 1. Fetch Recent Detected Apps
        try {
            $appUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($targetDevice.Id)/detectedApps?`$top=5"
            $appResponse = Invoke-MgGraphRequest -Method GET -Uri $appUri -ErrorAction Stop
            $appData = $appResponse.value | Select-Object @{n='Name';e={$_.displayName}}, @{n='Ver';e={$_.version}}
        } catch { }

        # 2. Fetch Entra ID Groups
        try {
            if ($targetDevice.azureADDeviceId) {
                $entraUri = "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$($targetDevice.azureADDeviceId)'&`$select=id"
                $entraDevice = Invoke-MgGraphRequest -Method GET -Uri $entraUri -ErrorAction Stop
                
                if ($entraDevice.value) {
                    $entraId = $entraDevice.value[0].id
                    $groupsUri = "https://graph.microsoft.com/v1.0/devices/$entraId/memberOf?`$select=displayName"
                    $groupsResponse = Invoke-MgGraphRequest -Method GET -Uri $groupsUri -ErrorAction Stop
                    $groupData = $groupsResponse.value | Select-Object @{n='Group';e={$_.displayName}}
                }
            }
        } catch { }

        # 3. Fetch Device Configurations (Policies)
        try {
            $configUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($targetDevice.Id)/deviceConfigurationStates?`$top=5"
            $configResponse = Invoke-MgGraphRequest -Method GET -Uri $configUri -ErrorAction Stop
            $configData = $configResponse.value | Select-Object @{n='Policy';e={$_.displayName}}, @{n='State';e={$_.state}}
        } catch { }
    }

    # Format into Spectre Tables
    $appsTable = if ($appData) { $appData | Format-SpectreTable -Title "Detected Apps" } else { "No apps detected" | Format-SpectrePanel -Header "Detected Apps" }
    $groupsTable = if ($groupData) { $groupData | Format-SpectreTable -Title "Entra ID Groups" } else { "No groups found" | Format-SpectrePanel -Header "Entra ID Groups" }
    $configTable = if ($configData) { $configData | Format-SpectreTable -Title "Config Policies" } else { "No policies found" | Format-SpectrePanel -Header "Config Policies" }

    # Display in a unified side-by-side grid
    @($appsTable, $groupsTable, $configTable) | Format-SpectreColumns | Out-SpectreHost

    Write-Host ""
    "----------------------------------------------------------------" | Out-SpectreHost
    $portalUrl = "https://intune.microsoft.com/#view/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/~/overview/managedDeviceId/$($targetDevice.Id)"
    " > [blue][link=$portalUrl]Open in Intune Portal[/][/]" | Out-SpectreHost
    Write-Host ""
>>>>>>> d874abec0309bd3c130528a64bbaea4c6b3fe2e0

    # If multiple devices match, let user pick. If only one, select it automatically.
    if ($devices.Count -gt 1) {
        $choices = $devices | ForEach-Object { "$($_.deviceName) | $($_.serialNumber) | $($_.Id)" }
        $selectionString = Read-SpectreSelection -Title "Multiple matches found. Select a device:" -Choices $choices -EnableSearch -PageSize 15
        if (-not $selectionString) { continue } 
        $selectedId = ($selectionString -split "\|")[-1].Trim()
        $targetDevice = $devices | Where-Object { $_.Id -eq $selectedId } | Select-Object -First 1
    } else {
        $targetDevice = $devices[0]
    }

    # Inner loop to keep the user on the current device dashboard after taking actions
    $viewingDevice = $true
    while ($viewingDevice) {
        Clear-Host
        "Diagnostics: [bold cyan]$($targetDevice.deviceName)[/]" | Out-SpectreHost
        "----------------------------------------------------------------" | Out-SpectreHost
        Write-Host ""

        # ------------------------------------------
        # ROW 1: IDENTITY & HEALTH PANELS
        # ------------------------------------------
        $idText = "Serial: [bold]$($targetDevice.serialNumber)[/]`nUser: [deepskyblue1]$($targetDevice.userPrincipalName)[/]`nModel: $($targetDevice.model)"
        $identityPanel = $idText | Format-SpectrePanel -Header "Identity"

        $statusColor = if($targetDevice.complianceState -eq "compliant") { "green" } else { "red" }
        $healthText = "Status: [$statusColor]$($targetDevice.complianceState)[/]`nSync: $($targetDevice.lastSyncDateTime)"
        $healthPanel = $healthText | Format-SpectrePanel -Header "Health"

        @($identityPanel, $healthPanel) | Format-SpectreColumns | Out-SpectreHost
        Write-Host ""

        # ------------------------------------------
        # ROW 2: STORAGE CHART
        # ------------------------------------------
        try {
            $total = [Math]::Round($targetDevice.totalStorageSpaceInBytes / 1GB, 0)
            $free  = [Math]::Round($targetDevice.freeStorageSpaceInBytes / 1GB, 0)
            $used  = $total - $free
            
            $chartItems = @(
                (New-SpectreChartItem -Label "Used ($used GB)" -Value $used -Color Red),
                (New-SpectreChartItem -Label "Free ($free GB)" -Value $free -Color Green)
            )
            $chartItems | Format-SpectreBreakdownChart | Format-SpectrePanel -Header "Storage" | Out-SpectreHost
        } catch { 
            "Storage metrics unavailable." | Format-SpectrePanel -Header "Storage" | Out-SpectreHost
        }
        Write-Host ""

        # ------------------------------------------
        # ROW 3: LIVE GRAPH DATA GATHERING
        # ------------------------------------------
        if ($TestMode) {
            $appData    = @([PSCustomObject]@{ Name = "Chrome"; Ver = "124" })
            $groupData  = @([PSCustomObject]@{ Group = "SG-IT-Admins" })
            $configData = @([PSCustomObject]@{ Policy = "Win10_SecBaseline"; State = "compliant" })
        } else {
            $appData    = Get-HDDeviceApps -DeviceId $targetDevice.Id
            $groupData  = Get-HDEntraGroups -AzureADDeviceId $targetDevice.azureADDeviceId
            $configData = Get-HDDeviceConfig -DeviceId $targetDevice.Id
        }

        # Format into Spectre Tables
        $appsTable   = if ($appData) { $appData | Format-SpectreTable -Title "Detected Apps" } else { "No apps detected" | Format-SpectrePanel -Header "Detected Apps" }
        $groupsTable = if ($groupData) { $groupData | Format-SpectreTable -Title "Entra ID Groups" } else { "No groups found" | Format-SpectrePanel -Header "Entra ID Groups" }
        $configTable = if ($configData) { $configData | Format-SpectreTable -Title "Config Policies" } else { "No policies found" | Format-SpectrePanel -Header "Config Policies" }

        # Display in a unified side-by-side grid
        @($appsTable, $groupsTable, $configTable) | Format-SpectreColumns | Out-SpectreHost
        Write-Host "`n----------------------------------------------------------------"

        # Flush Buffer
        while ([console]::KeyAvailable) { $null = [console]::ReadKey($true) }
        Start-Sleep -Milliseconds 500

        # ------------------------------------------
        # ROW 4: ACTION MENU
        # ------------------------------------------
        $action = Read-SpectreSelection -Title "Device Actions" -Choices @("New Search", "Sync Device", "Reboot Device", "Open in Intune", "Exit")

        switch ($action) {
            "New Search" {
                $viewingDevice = $false # Break inner loop, returns to search
            }
            "Sync Device" {
                if (-not $TestMode) {
                    Invoke-SpectreCommandWithStatus -Title "Sending Sync Command..." -Spinner "Dots2" -ScriptBlock {
                        Sync-MgDeviceManagementManagedDevice -ManagedDeviceId $targetDevice.Id
                    }
                }
                Write-SpectreHost "Sync initiated successfully!" -Color Green
                Start-Sleep -Seconds 2
            }
            "Reboot Device" {
                $confirm = Read-SpectreSelection -Title "Are you sure you want to REBOOT $($targetDevice.deviceName)?" -Choices @("No", "Yes")
                if ($confirm -eq "Yes") {
                    if (-not $TestMode) {
                        Invoke-SpectreCommandWithStatus -Title "Sending Reboot Command..." -Spinner "Dots2" -ScriptBlock {
                            Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$($targetDevice.Id)/rebootNow"
                        }
                    }
                    Write-SpectreHost "Reboot command sent to device." -Color Green
                    Start-Sleep -Seconds 2
                }
            }
            "Open in Intune" {
                $portalUrl = "https://intune.microsoft.com/#view/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/~/overview/managedDeviceId/$($targetDevice.Id)"
                Start-Process $portalUrl
            }
            "Exit" {
                exit
            }
        }
    }
}

"Goodbye!" | Out-SpectreHost
