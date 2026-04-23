[CmdletBinding()]
param (
    [switch]$TestMode
)

Import-Module PwshSpectreConsole

# ==========================================
# 1. AUTHENTICATION & DATA GATHERING
# ==========================================
if (-not $TestMode) {
    $allDevices = Invoke-SpectreCommandWithStatus -Title "Connecting..." -Spinner "Dots2" -Color Cyan -ScriptBlock {
        Write-Host "Checking Authentication..."
        Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All", "User.Read.All" -NoWelcome
        Write-Host "Downloading Device List..."
        $devices = Get-MgDeviceManagementManagedDevice -All -Property "id,deviceName,serialNumber,userPrincipalName,model,operatingSystem,osVersion,complianceState,lastSyncDateTime,managementAgent,totalStorageSpaceInBytes,freeStorageSpaceInBytes"
        return $devices
    }
} else {
    $allDevices = @(
        [PSCustomObject]@{ Id = "111-222"; deviceName = "LAPTOP-JLR-01"; serialNumber = "PF3B99X"; userPrincipalName = "jrice@domain.com"; model = "Surface Laptop 5"; operatingSystem = "Windows"; osVersion = "10.0.22631"; complianceState = "compliant"; lastSyncDateTime = (Get-Date).AddMinutes(-45); managementAgent = "intune"; totalStorageSpaceInBytes = 256GB; freeStorageSpaceInBytes = 45GB }
        [PSCustomObject]@{ Id = "333-444"; deviceName = "DESKTOP-DEV-02"; serialNumber = "MXL1234"; userPrincipalName = "krice@domain.com"; model = "OptiPlex 7090"; operatingSystem = "Windows"; osVersion = "10.0.19045"; complianceState = "noncompliant"; lastSyncDateTime = (Get-Date).AddDays(-3); managementAgent = "intune"; totalStorageSpaceInBytes = 512GB; freeStorageSpaceInBytes = 12GB }
    )
}

# ==========================================
# 2. MAIN APPLICATION LOOP
# ==========================================
while ($true) {
    Clear-Host
    
    # Header - No Justification parameter to avoid InvalidOperation error
    Write-SpectreFigletText -Text "Intune" -Color Cyan
    "----------------------------------------------------------------" | Out-SpectreHost
    Write-Host ""
    
    $choices = $allDevices | ForEach-Object { "$($_.deviceName) | $($_.serialNumber)" }
    $selectionString = Read-SpectreSelection -Title "Search Device (ESC to exit)" -Choices $choices -EnableSearch
    
    if (-not $selectionString) { break } 

    $selectedName = ($selectionString -split "\|")[0].Trim()
    $targetDevice = $allDevices | Where-Object { $_.deviceName -eq $selectedName }

    Clear-Host
    "Diagnostics: [bold cyan]$($targetDevice.deviceName)[/]" | Out-SpectreHost
    "----------------------------------------------------------------" | Out-SpectreHost
    Write-Host ""

    # DASHBOARD PANELS
    $idText = "Serial: [bold]$($targetDevice.serialNumber)[/]`nUser: [deepskyblue1]$($targetDevice.userPrincipalName)[/]`nModel: $($targetDevice.model)"
    $identityPanel = $idText | Format-SpectrePanel -Header "Identity"

    $statusColor = if($targetDevice.complianceState -eq "compliant") { "green" } else { "red" }
    $healthText = "Status: [$statusColor]$($targetDevice.complianceState)[/]`nSync: $($targetDevice.lastSyncDateTime)"
    $healthPanel = $healthText | Format-SpectrePanel -Header "Health"

    @($identityPanel, $healthPanel) | Format-SpectreColumns | Out-SpectreHost
    Write-Host ""

    # STORAGE CHART
    try {
        $total = [Math]::Round($targetDevice.totalStorageSpaceInBytes / 1GB, 0)
        $free = [Math]::Round($targetDevice.freeStorageSpaceInBytes / 1GB, 0)
        $used = $total - $free
        
        $chartItems = @(
            (New-SpectreChartItem -Label "Used ($used GB)" -Value $used -Color Red),
            (New-SpectreChartItem -Label "Free ($free GB)" -Value $free -Color Green)
        )
        $chartItems | Format-SpectreBreakdownChart | Format-SpectrePanel -Header "Storage" | Out-SpectreHost
    } catch { }
    Write-Host ""

    # RECENT APPS
    if ($TestMode) {
        $appData = @([PSCustomObject]@{ Name = "Chrome"; Ver = "124"; Size = "300MB" })
    } else {
        $apps = Get-MgDeviceManagementManagedDeviceDetectedApp -ManagedDeviceId $targetDevice.Id -Top 3
        $appData = $apps | Select-Object @{n='Name';e={$_.DisplayName}}, @{n='Ver';e={$_.Version}}
    }
    $appData | Format-SpectreTable -Title "Recent Apps" | Out-SpectreHost

    Write-Host ""
    "----------------------------------------------------------------" | Out-SpectreHost
    $portalUrl = "https://intune.microsoft.com/#view/Microsoft_Intune_Devices/DeviceSettingsMenuBlade/~/overview/managedDeviceId/$($targetDevice.Id)"
    # Removed the brackets around the > to prevent the .ctor crash
    " > [blue][link=$portalUrl]Open in Intune Portal[/][/]" | Out-SpectreHost
    Write-Host ""

    # Flush Buffer
    while ([console]::KeyAvailable) { $null = [console]::ReadKey($true) }
    Start-Sleep -Seconds 1

    $action = Read-SpectreSelection -Title "Next Step" -Choices @("New Search", "Exit")
    if ($action -match "Exit") { break }
}

"Goodbye!" | Out-SpectreHost