# C:\Scripts\PowerPoller.ps1
# Automatically switches power plan and toggles dark/light mode
# Works on PowerShell 5.1 and 7+

# ==== CONFIGURATION ====
$BatteryPlanGuid = "b317b94d-7c0f-4afa-886a-dfa131234c05"   # Your battery plan GUID
$ACPlanGuid      = "6f740897-a4ae-4f69-9c92-a049307f5bbe"   # Your AC plan GUID
$UseBatterySaverToggle = $true
$PollIntervalSec = 15   # how often to check (seconds)

# ==== FUNCTIONS ====
function Get-OnAC {
    try {
        $battery = Get-WmiObject -Class Win32_Battery -ErrorAction Stop
        return ($battery.BatteryStatus -eq 2 -or $battery.BatteryStatus -eq 3)
    } catch {
        return $true
    }
}

function Set-DarkMode {
    Write-Output "$(Get-Date) - Enabling Dark Mode"
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name AppsUseLightTheme -Value 0
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name SystemUsesLightTheme -Value 0
    Stop-Process -Name explorer -Force
}

function Set-LightMode {
    Write-Output "$(Get-Date) - Enabling Light Mode"
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name AppsUseLightTheme -Value 1
    Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name SystemUsesLightTheme -Value 1
    Stop-Process -Name explorer -Force
}

function Apply-PowerProfile($onAC) {
    if ($onAC) {
        Write-Output "$(Get-Date) - On AC: switching to AC plan"
        powercfg /setactive $ACPlanGuid | Out-Null
        if ($UseBatterySaverToggle) {
            powercfg /setdcvalueindex SCHEME_CURRENT SUB_ENERGYSAVER ESBATTTHRESHOLD 0 | Out-Null
            powercfg /setactive SCHEME_CURRENT | Out-Null
        }
        Set-LightMode
    } else {
        Write-Output "$(Get-Date) - On Battery: switching to Battery plan"
        powercfg /setactive $BatteryPlanGuid | Out-Null
        if ($UseBatterySaverToggle) {
            powercfg /setdcvalueindex SCHEME_CURRENT SUB_ENERGYSAVER ESBATTTHRESHOLD 100 | Out-Null
            powercfg /setactive SCHEME_CURRENT | Out-Null
        }
        Set-DarkMode
    }
}

# ==== MAIN LOOP ====
Write-Output "PowerPoller started. Polling every $PollIntervalSec seconds..."
$prevStatus = $null

while ($true) {
    $onAC = Get-OnAC
    if ($prevStatus -ne $onAC) {
        Apply-PowerProfile $onAC
        $prevStatus = $onAC
    }
    Start-Sleep -Seconds $PollIntervalSec
}
