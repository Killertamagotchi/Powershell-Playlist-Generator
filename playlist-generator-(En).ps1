[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# ADJUST THE DRIVE LETTER OF YOUR MICROSD CARD HERE:
$sdCard = "E:\"

# Ensure the path ends with a backslash
if (-not $sdCard.EndsWith("\")) { $sdCard += "\" }

# Change location to the SD card root
Set-Location $sdCard

# Load Windows Shell COM object to read metadata
$shell = New-Object -ComObject Shell.Application

# Function to replace forbidden Windows characters with allowed, similar-looking Unicode characters
function Get-SafeFileName ($name) {
    if (-not $name) { return "Unknown_Album" }
    
    # Safe character replacements using Hex codes (immune to encoding issues)
    $name = $name -replace ':', [char]0xA789  # Replaces colon with U+A789 (Modifier Letter Colon)
    $name = $name -replace '\?', [char]0xFF1F # Replaces question mark with U+FF1F (Fullwidth Question Mark)
    $name = $name -replace '\*', [char]0x2217 # Replaces asterisk with U+2217 (Asterisk Operator)
    $name = $name -replace '/', [char]0x2215  # Replaces forward slash with U+2215 (Division Slash)
    $name = $name -replace '\\', [char]0x2216 # Replaces backslash with U+2216 (Set Minus)
    $name = $name -replace '"', [char]0x2033  # Replaces quote with U+2033 (Double Prime)
    $name = $name -replace '<', [char]0x2039  # Replaces less-than with U+2039 (Single Left Angle Quote)
    $name = $name -replace '>', [char]0x203A  # Replaces greater-than with U+203A (Single Right Angle Quote)
    $name = $name -replace '\|', [char]0x2502 # Replaces pipe with U+2502 (Box Drawings Light Vertical)

    return $name.Trim()
}

# Function to read the Album tag from the first audio file in a folder
function Get-AlbumTagFromFolder ($folderPath) {
    $firstTrack = Get-ChildItem -Path $folderPath -File | Where-Object { $_.Extension -match "\.(mp3|flac|wav|m4a)$" } | Select-Object -First 1
    if (-not $firstTrack) { return $null }

    $shellFolder = $shell.NameSpace($folderPath)
    $shellFile = $shellFolder.ParseName($firstTrack.Name)
    
    # Attribute ID 14 corresponds to "Album" in Windows
    $albumTag = $shellFolder.GetDetailsOf($shellFile, 14)
    
    if ([string]::IsNullOrWhiteSpace($albumTag)) {
        return $null
    }
    return $albumTag.Trim()
}

# Function to locate the actual parent album directory (bypassing CD 1, CD 2, etc.)
function Get-ActualAlbumFolder ($folder) {
    if ($folder.Name -match "^CD\s*\d+$|^Disc\s*\d+$|^FLAC$|^MP3$" -and $folder.Parent) {
        return Get-ActualAlbumFolder $folder.Parent
    }
    return $folder
}

Write-Host "Scanning SD card for new albums... Please wait..." -ForegroundColor Yellow

# Fetch all directories containing audio files
$allFolders = Get-ChildItem -Path $sdCard -Recurse -Directory
$albumPaths = @{}

foreach ($folder in $allFolders) {
    $hasMusic = Get-ChildItem -Path $folder.FullName -File | Where-Object { $_.Extension -match "\.(mp3|flac|wav|m4a)$" }
    if ($hasMusic) {
        $actualAlbumFolder = Get-ActualAlbumFolder $folder
        
        # Pre-check using the folder name to speed up the process
        $possibleSafeFolderName = Get-SafeFileName $actualAlbumFolder.Name
        if (Test-Path (Join-Path -Path $sdCard -ChildPath "$possibleSafeFolderName.m3u")) {
            continue 
        }

        # Attempt to read the Album tag
        $albumName = Get-AlbumTagFromFolder $folder.FullName
        if (-not $albumName) {
            $albumName = Get-AlbumTagFromFolder $actualAlbumFolder.FullName
        }
        if (-not $albumName) {
            $albumName = $actualAlbumFolder.Name
        }

        # Skip if a playlist with the cleaned tag name already exists
        $safePlaylistName = Get-SafeFileName $albumName
        if (Test-Path (Join-Path -Path $sdCard -ChildPath "$safePlaylistName.m3u")) {
            continue
        }

        $albumPaths[$actualAlbumFolder.FullName] = $albumName
    }
}

# Generate playlists in the ROOT directory (only for newly discovered albums)
$createdCount = 0
foreach ($albumPath in $albumPaths.Keys) {
    $albumName = $albumPaths[$albumPath]
    $safePlaylistName = Get-SafeFileName $albumName
    $playlistPath = Join-Path -Path $sdCard -ChildPath "$safePlaylistName.m3u"
    
    # Collect all audio tracks inside the album directory (including CD 1, CD 2)
    $tracks = Get-ChildItem -Path $albumPath -Recurse -File | Where-Object { 
        $_.Extension -match "\.(mp3|flac|wav|m4a)$" 
    } | Sort-Object FullName

    if ($tracks.Count -gt 0) {
        Write-Host "NEWLY DISCOVERED: Creating playlist for $safePlaylistName" -ForegroundColor Green
        $createdCount++
        
        # StreamWriter using UTF-8 without BOM (essential for Linux-based audio players)
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $writer = New-Object System.IO.StreamWriter($playlistPath, $false, $utf8NoBom)
        
        $writer.WriteLine("#EXTM3U")
        
        foreach ($track in $tracks) {
            # Strip the drive letter prefix to create a relative path
            $relativePath = $track.FullName.Substring($sdCard.Length)
            $writer.WriteLine($relativePath)
        }
        
        $writer.Close()
    }
}

# Clean up the COM object
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null

if ($createdCount -gt 0) {
    Write-Host "Done! Successfully created $createdCount new playlists with safe Unicode characters." -ForegroundColor Cyan
} else {
    Write-Host "Done! No new albums found. All playlists are already up to date." -ForegroundColor Yellow
}
