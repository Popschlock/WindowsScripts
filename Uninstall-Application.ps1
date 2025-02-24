# Generic Uninstaller for most programs
#   Provide any process names to force stop, and the name of the application(s) as shown in programs and features
#   Example included below would uninstall globalprotect. Works with EXE and MSI based uninstalls as long as they support running silently.

$ProcessNames = @("PanGPA", "PanGPS")
$ApplicationNames = @("GlobalProtect")

# Force-stop processes
foreach ($ProcessName in $ProcessNames) {
    Get-Process -Name $ProcessName -ErrorAction SilentlyContinue | Stop-Process -Force
}

# Paths to look for uninstall keys
$RegPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

# Look for uninstall entries and run uninstall command
foreach ($ApplicationName in $ApplicationNames) {
    foreach ($Path in $RegPaths) {
        $Result = Get-ItemProperty $Path -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -eq "$ApplicationName" }
        if ($null -ne $Result) {
            Write-Output "$ApplicationName found in: $Path"
            $UninstallString = $Result.UninstallString
            Write-Output "Original Uninstall String: $UninstallString"
            
            if ($UninstallString -match "msiexec") {
                if ($UninstallString -match "/I") {
                    $UninstallString = $UninstallString -replace "/I", "/X"
                }
                if ($UninstallString -notmatch "/qn") {
                    $UninstallString += " /qn"
                }
            }
            else {
                $UninstallString += " /quiet"
            }
            
            Write-Output "Executing: $UninstallString"
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c", "$UninstallString" -WindowStyle Hidden -Wait
        }
        else {
            Write-Output "$ApplicationName not found in: $Path"
        }
    }
}
