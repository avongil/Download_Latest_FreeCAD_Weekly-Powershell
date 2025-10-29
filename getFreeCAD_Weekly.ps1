# PowerShell script to download the latest FreeCAD weekly build for Windows x86_64 (.7z),
# verify its hash, extract it to a portable directory (handling archives with or without root folder),
# optionally enable true portable mode with a launcher, and update a shortcut.
# Switch for true portable mode (set to $true to store user data in $targetPath\.FreeCAD; $false to use system AppData)
$enablePortableMode = $false
# Variable for the portable directory (change this if needed)
$portableDir = "C:\Software-Portable"
# Downloads directory (using environment variable for reliability in admin mode)
$downloadsDir = "$env:USERPROFILE\Downloads"
# Kill any running FreeCAD processes
Get-Process -Name FreeCAD -ErrorAction SilentlyContinue | Stop-Process -Force
Write-Host "Terminated any running FreeCAD processes."
# Function to get the latest weekly release that has the Windows x86_64 .7z asset
function Get-LatestWeeklyReleaseWithAsset {
    $releasesUrl = "https://api.github.com/repos/FreeCAD/FreeCAD/releases"
    $releases = Invoke-RestMethod -Uri $releasesUrl -Method Get
   
    $weeklyReleases = $releases | Where-Object { $_.tag_name -like "weekly-*" } |
        Sort-Object -Property published_at -Descending
   
    if ($weeklyReleases.Count -eq 0) {
        Write-Error "No weekly releases found."
        return $null
    }
   
    foreach ($release in $weeklyReleases) {
        $assets = $release.assets
        $asset = $assets | Where-Object { $_.name -match "Windows-x86_64.*\.7z$" -and $_.name -notmatch "\.txt$" } | Select-Object -First 1
        if ($null -ne $asset) {
            return $release
        }
    }
   
    Write-Error "No weekly release with Windows x86_64 .7z asset found."
    return $null
}
# Get the latest release with the asset
$latestRelease = Get-LatestWeeklyReleaseWithAsset
if ($null -eq $latestRelease) {
    exit 1
}
$tag = $latestRelease.tag_name
$publishedAt = $latestRelease.published_at
Write-Host "Using weekly release: $tag ($publishedAt) (latest with available Windows .7z asset)"
# Get assets
$assets = $latestRelease.assets
# Find the Windows x86_64 .7z asset
$asset = $assets | Where-Object { $_.name -match "Windows-x86_64.*\.7z$" -and $_.name -notmatch "\.txt$" } | Select-Object -First 1
if ($null -eq $asset) {
    Write-Error "No Windows x86_64 .7z asset found for $tag"
    exit 1
}
# Find the corresponding hash asset (adjusted to match potential naming variations)
$hashAsset = $assets | Where-Object { $_.name -like "$($asset.name)*SHA256.txt" -or $_.name -like "$($asset.name -replace '_', '-').SHA256.txt" -or $_.name -like "$($asset.name -replace '-', '_').SHA256.txt" } | Select-Object -First 1
if ($null -eq $hashAsset) {
    Write-Warning "No SHA256 hash asset found for $($asset.name). Proceeding without hash verification."
}
$fileName = $asset.name
$hashFileName = if ($null -ne $hashAsset) { $hashAsset.name } else { $null }
$filePath = Join-Path $downloadsDir $fileName
$hashPath = if ($null -ne $hashFileName) { Join-Path $downloadsDir $hashFileName } else { $null }
$downloadNeeded = $true
if (Test-Path $filePath) {
    Write-Host "File $fileName already exists in Downloads."
    if ($null -ne $hashAsset -and (Test-Path $hashPath)) {
        # Verify hash
        $computedHash = (Get-FileHash $filePath -Algorithm SHA256).Hash.ToUpper()
        $hashContent = Get-Content $hashPath
        $expectedHash = ($hashContent -split '\s+')[0].ToUpper()
       
        if ($computedHash -eq $expectedHash) {
            Write-Host "Hash matches. Skipping download."
            $downloadNeeded = $false
        } else {
            Write-Host "Hash mismatch. Will redownload."
        }
    } else {
        Write-Host "Hash file missing or not available. Skipping download as per instructions."
        $downloadNeeded = $false
    }
}
if ($downloadNeeded) {
    Write-Host "Downloading $fileName from $($asset.browser_download_url)"
    Invoke-WebRequest -Uri $asset.browser_download_url -OutFile "$filePath"
    Write-Host "Download complete: $fileName"
   
    if ($null -ne $hashAsset) {
        Write-Host "Downloading $hashFileName from $($hashAsset.browser_download_url)"
        Invoke-WebRequest -Uri $hashAsset.browser_download_url -OutFile "$hashPath"
        Write-Host "Download complete: $hashFileName"
       
        # Verify hash after download
        $computedHash = (Get-FileHash $filePath -Algorithm SHA256).Hash.ToUpper()
        $hashContent = Get-Content $hashPath
        $expectedHash = ($hashContent -split '\s+')[0].ToUpper()
       
        if ($computedHash -ne $expectedHash) {
            Write-Error "Downloaded file hash mismatch. File may be corrupted."
            exit 1
        }
        Write-Host "Hash verification passed."
    }
}
# Locate 7-Zip executable
$sevenZipPath = $null
if (Test-Path "C:\Program Files\7-Zip\7z.exe") {
    $sevenZipPath = "C:\Program Files\7-Zip\7z.exe"
} elseif (Test-Path "C:\Program Files (x86)\7-Zip\7z.exe") {
    $sevenZipPath = "C:\Program Files (x86)\7-Zip\7z.exe"
} else {
    $sevenZipPath = (Get-Command "7z.exe" -ErrorAction SilentlyContinue).Source
}
if ($null -eq $sevenZipPath) {
    Write-Error "7-Zip not found. Please install 7-Zip or ensure it's in your PATH."
    exit 1
}
# Determine the target folder name (based on filename without .7z; optionally simplify to "FreeCAD_$tag" for a shorter name)
$extractedFolderName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
# Alternative shorter name: $extractedFolderName = "FreeCAD_$tag"
$targetPath = Join-Path $portableDir $extractedFolderName
# Remove existing target folder if it exists
if (Test-Path $targetPath) {
    Write-Host "Removing existing folder $targetPath"
    Remove-Item "$targetPath" -Recurse -Force
}
# Create the target folder (ensures it's empty)
New-Item -ItemType Directory -Path "$targetPath" -Force | Out-Null
# Create a temporary extraction folder
$tempPath = Join-Path $downloadsDir "FreeCAD_extract_temp"
if (Test-Path $tempPath) {
    Remove-Item "$tempPath" -Recurse -Force
}
New-Item -ItemType Directory -Path "$tempPath" -Force | Out-Null
# Extract the .7z file to the temp folder
Write-Host "Extracting $fileName to $tempPath"
& $sevenZipPath x "$filePath" -o"$tempPath" -y
# Check the contents of the temp folder to handle root folder if present
$items = Get-ChildItem -Path $tempPath
if ($items.Count -eq 1 -and $items[0].PSIsContainer) {
    # Archive has a root folder; strip it by moving its contents
    Write-Host "Detected root folder in archive. Stripping it."
    $rootFolderPath = $items[0].FullName
    Move-Item -Path (Join-Path $rootFolderPath "*") -Destination $targetPath -Force
} else {
    # No root folder; move all contents directly
    Write-Host "No root folder detected. Moving contents directly."
    Move-Item -Path (Join-Path $tempPath "*") -Destination $targetPath -Force
}
# Clean up the temp folder
Remove-Item "$tempPath" -Recurse -Force
# Verify FreeCAD.exe exists
$exePath = Join-Path $targetPath "bin\FreeCAD.exe"
if (-not (Test-Path $exePath)) {
    Write-Error "FreeCAD.exe not found in $targetPath\bin. Structure may have changed."
    exit 1
}
# Optionally create launch.bat for portable mode
$launcherPath = $null
if ($enablePortableMode) {
    $launcherPath = Join-Path $targetPath "launch.bat"
    $launcherContent = @"
@echo off
set FREECAD_USER_HOME=%~dp0.FreeCAD
"%~dp0bin\FreeCAD.exe" %*
"@
    Set-Content -Path "$launcherPath" -Value $launcherContent
}
# Update the shortcut
$shortcutPath = Join-Path $portableDir "freecad.exe.lnk"
if (Test-Path $shortcutPath) {
    Write-Host "Deleting existing shortcut $shortcutPath"
    Remove-Item "$shortcutPath" -Force
}
$wsh = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut($shortcutPath)
$shortcut.TargetPath = if ($enablePortableMode) { $launcherPath } else { $exePath }
$shortcut.WorkingDirectory = $targetPath
$shortcut.Description = "FreeCAD Weekly Build" + $(if ($enablePortableMode) { " (Portable)" } else { "" })
$shortcut.IconLocation = $exePath
$shortcut.Save()
Write-Host "Shortcut created at $shortcutPath"
# Delete older FreeCAD folders if installation was successful
$olderFolders = Get-ChildItem -Path $portableDir -Directory | Where-Object { $_.Name -like "FreeCAD*" -and $_.FullName -ne $targetPath }
foreach ($folder in $olderFolders) {
    Write-Host "Deleting older folder: $($folder.FullName)"
    try {
        Remove-Item $folder.FullName -Recurse -Force -ErrorAction Stop
    } catch {
        Write-Warning "Could not delete older folder $($folder.FullName): $_"
        Write-Host "Please ensure no processes (like FreeCAD) are using files in this folder and try deleting manually or run the script again after closing them."
    }
}
Write-Host "Update complete. Launch FreeCAD from the shortcut." + $(if ($enablePortableMode) { " User data will be in $targetPath\.FreeCAD." } else { " User data will be in system AppData." })
Write-Host "Opening the portable directory for you to pin the shortcut to the taskbar."
explorer $portableDir