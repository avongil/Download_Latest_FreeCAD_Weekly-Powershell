# PowerShell script to download the latest FreeCAD weekly build for Windows x86_64 (.7z),
# verify its hash, extract it, move the extracted folder to a portable directory, and update a shortcut.

# Variable for the portable directory (change this if needed)
$portableDir = "C:\Software-Portable"

# Downloads directory (using environment variable for reliability in admin mode)
$downloadsDir = "$env:USERPROFILE\Downloads"

# Function to get the latest weekly release
function Get-LatestWeeklyRelease {
    $releasesUrl = "https://api.github.com/repos/FreeCAD/FreeCAD/releases"
    $releases = Invoke-RestMethod -Uri $releasesUrl -Method Get
    
    $weeklyReleases = $releases | Where-Object { $_.tag_name -like "weekly-*" } | 
        Sort-Object -Property published_at -Descending
    
    if ($weeklyReleases.Count -eq 0) {
        Write-Error "No weekly releases found."
        return $null
    }
    
    return $weeklyReleases[0]
}

# Get the latest release
$latestRelease = Get-LatestWeeklyRelease
if ($null -eq $latestRelease) {
    exit 1
}

$tag = $latestRelease.tag_name
$publishedAt = $latestRelease.published_at
Write-Host "Latest weekly release: $tag ($publishedAt)"

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
    Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $filePath
    Write-Host "Download complete: $fileName"
    
    if ($null -ne $hashAsset) {
        Write-Host "Downloading $hashFileName from $($hashAsset.browser_download_url)"
        Invoke-WebRequest -Uri $hashAsset.browser_download_url -OutFile $hashPath
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

# Extract the .7z file to Downloads directory
Write-Host "Extracting $fileName to $downloadsDir"
& $sevenZipPath x $filePath -o$downloadsDir -y

# Determine the extracted folder name (assuming it's the file name without .7z)
$extractedFolderName = [System.IO.Path]::GetFileNameWithoutExtension($fileName)
$extractedPath = Join-Path $downloadsDir $extractedFolderName

if (-not (Test-Path $extractedPath -PathType Container)) {
    Write-Error "Extracted folder $extractedFolderName not found. Extraction may have failed or structure changed."
    exit 1
}

# Target path in portable directory
$targetPath = Join-Path $portableDir $extractedFolderName

# Remove existing folder if it exists
if (Test-Path $targetPath) {
    Write-Host "Removing existing folder $targetPath"
    Remove-Item $targetPath -Recurse -Force
}

# Move the extracted folder
Write-Host "Moving $extractedFolderName to $portableDir"
Move-Item -Path $extractedPath -Destination $portableDir

# Verify FreeCAD.exe exists
$exePath = Join-Path $targetPath "bin\FreeCAD.exe"
if (-not (Test-Path $exePath)) {
    Write-Error "FreeCAD.exe not found in $targetPath\bin. Structure may have changed."
    exit 1
}

# Update the shortcut
$shortcutPath = Join-Path $portableDir "freecad.exe.lnk"
if (Test-Path $shortcutPath) {
    Write-Host "Deleting existing shortcut $shortcutPath"
    Remove-Item $shortcutPath -Force
}

$wsh = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $exePath
$shortcut.WorkingDirectory = $targetPath
$shortcut.Description = "FreeCAD Weekly Build"
$shortcut.IconLocation = $exePath
$shortcut.Save()

Write-Host "Shortcut created at $shortcutPath"

Write-Host "Update complete. You can launch FreeCAD from the shortcut."