# Ensure the script is running as Administrator
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole] "Administrator"))
{
    Write-Warning "This script needs to be run as an Administrator. Exiting."
    exit
}

Write-Host "Starting Webex (Cisco) uninstall procedure..."

#############################
# 1. Stop Processes
#############################

Write-Host "Stopping any running processes with 'cisco' or 'webex' in the name..."
Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.Name -match '(?i)cisco|webex' } | ForEach-Object {
    try {
        Stop-Process -Id $_.Id -Force -ErrorAction Stop
        Write-Host "Stopped process: $($_.Name) (ID: $($_.Id))"
    }
    catch {
        Write-Warning "Failed to stop process: $($_.Name) (ID: $($_.Id))"
    }
}

#############################
# 2. Remove Folders from All User Profiles
#############################

Write-Host "Removing Webex/Cisco folders from user profiles..."

# Get all user profile directories (excluding common system profiles)
$UserProfileRoot = "C:\Users"
$UserFolders = Get-ChildItem -Path $UserProfileRoot -Directory | Where-Object {
    $_.Name -notmatch '^(Default|Public|All Users)$'
}

foreach ($user in $UserFolders) {
    # Ensure we work with a string for the profile path
    $userProfile = [string]$user.FullName

    # Compute each folder path separately
    $folder1 = Join-Path -Path $userProfile -ChildPath "AppData\Local\Programs\Cisco Spark"
    $folder2 = Join-Path -Path $userProfile -ChildPath "AppData\Local\CiscoSparkLauncher"
    $folder3 = Join-Path -Path $userProfile -ChildPath "AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Webex"
    
    # Combine the paths into an array
    $pathsToRemove = @($folder1, $folder2, $folder3)
    
    foreach ($path in $pathsToRemove) {
        if (Test-Path $path) {
            try {
                Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
                Write-Host "Removed folder: $path"
            }
            catch {
                Write-Warning "Failed to remove folder: $path"
            }
        }
        else {
            Write-Host "Folder not found (skipping): $path"
        }
    }
}

#############################
# 3. Remove CiscoWebex Registry Key for All Users
#############################

Write-Host "Removing CiscoWebex registry key from all user hives..."

# Iterate over each user registry hive in HKEY_USERS
Get-ChildItem "Registry::HKEY_USERS" | ForEach-Object {
    $userHive = $_.PSChildName
    # The registry key path to remove
    $regKeyPath = "Software\Microsoft\Windows\CurrentVersion\Uninstall\CiscoWebex"
    $fullRegPath = "Registry::HKEY_USERS\$userHive\$regKeyPath"
    
    if (Test-Path $fullRegPath) {
        try {
            Remove-Item -Path $fullRegPath -Recurse -Force -ErrorAction Stop
            Write-Host "Removed registry key: $fullRegPath"
        }
        catch {
            Write-Warning "Failed to remove registry key: $fullRegPath"
        }
    }
    else {
        Write-Host "Registry key not found in hive $userHive. Skipping."
    }
}

#############################
# 4. Remove Desktop Shortcuts for All Users
#############################

Write-Host "Removing Webex desktop shortcuts for all users..."

# Iterate over each user registry hive to determine the Desktop folder path
Get-ChildItem "Registry::HKEY_USERS" | ForEach-Object {
    $userHive = $_.PSChildName
    $shellFoldersRegPath = "Registry::HKEY_USERS\$userHive\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    
    if (Test-Path $shellFoldersRegPath) {
        try {
            $shellFolders = Get-ItemProperty -Path $shellFoldersRegPath
            $desktopPath = $shellFolders.Desktop
            if ($desktopPath -and (Test-Path $desktopPath)) {
                # Look for desktop shortcuts with "Webex" in the name (typically .lnk files)
                $shortcuts = Get-ChildItem -Path $desktopPath -Filter "*Webex*.lnk" -ErrorAction SilentlyContinue
                foreach ($shortcut in $shortcuts) {
                    try {
                        Remove-Item $shortcut.FullName -Force -ErrorAction Stop
                        Write-Host "Removed desktop shortcut: $($shortcut.FullName) from user hive $userHive"
                    }
                    catch {
                        Write-Warning "Failed to remove desktop shortcut: $($shortcut.FullName) from user hive $userHive"
                    }
                }
            }
            else {
                Write-Host "Desktop path not found or invalid in user hive $userHive"
            }
        }
        catch {
            Write-Warning "Error reading Shell Folders for user hive $userHive"
        }
    }
    else {
        Write-Host "Shell Folders registry path not found for user hive $userHive"
    }
}

Write-Host "Webex (Cisco) uninstall procedure completed."
