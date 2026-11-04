[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
# ADJUST THE DRIVE LETTER OF YOUR MICROSD CARD HERE:
$sdCard = "" # Beispiel $sdCard = "E:/"

# Stellt sicher, dass der Pfad mit einem Backslash endet
if (-not $sdCard.EndsWith("\")) { $sdCard += "\" }

# Wechselt auf die SD-Karte
Set-Location $sdCard

# Windows Shell-Objekt laden, um Metadaten auszulesen
$shell = New-Object -ComObject Shell.Application

# Funktion, um verbotene Windows-Zeichen durch erlaubte, ähnlich aussehende Unicode-Zeichen zu ersetzen
function Get-SafeFileName ($name) {
    if (-not $name) { return "Unbekanntes_Album" }
    
    # Ersetzungen über Hex-Codes (Sicher gegen falsche Datei-Codierung)
    $name = $name -replace ':', [char]0xA789  # Ersetzt Doppelpunkt durch Modifier Letter Colon
    $name = $name -replace '\?', [char]0xFF1F # Ersetzt Fragezeichen durch Fullwidth Question Mark
    $name = $name -replace '\*', [char]0x2217 # Ersetzt Stern durch Asterisk Operator
    $name = $name -replace '/', [char]0x2215  # Ersetzt Schrägstrich durch Division Slash
    $name = $name -replace '\\', [char]0x2216 # Ersetzt Backslash durch Set Minus
    $name = $name -replace '"', [char]0x2033  # Ersetzt Anführungszeichen durch Double Prime
    $name = $name -replace '<', [char]0x2039  # Ersetzt Kleiner-Als durch Single Left Angle Quote
    $name = $name -replace '>', [char]0x203A  # Ersetzt Größer-Als durch Single Right Angle Quote
    $name = $name -replace '\|', [char]0x2502 # Ersetzt Pipe durch Box Drawings Light Vertical

    return $name.Trim()
}

# Funktion, um das Album-Tag aus der ersten Datei eines Ordners zu lesen
function Get-AlbumTagFromFolder ($folderPath) {
    $firstTrack = Get-ChildItem -Path $folderPath -File | Where-Object { $_.Extension -match "\.(mp3|flac|wav|m4a)$" } | Select-Object -First 1
    if (-not $firstTrack) { return $null }

    $shellFolder = $shell.NameSpace($folderPath)
    $shellFile = $shellFolder.ParseName($firstTrack.Name)
    
    # Attribut-ID 14 ist das "Album" in Windows
    $albumTag = $shellFolder.GetDetailsOf($shellFile, 14)
    
    if ([string]::IsNullOrWhiteSpace($albumTag)) {
        return $null
    }
    return $albumTag.Trim()
}

# Funktion, um den übergeordneten Hauptordner (das eigentliche Album) zu finden
function Get-ActualAlbumFolder ($folder) {
    if ($folder.Name -match "^CD\s*\d+$|^Disc\s*\d+$|^FLAC$|^MP3$" -and $folder.Parent) {
        return Get-ActualAlbumFolder $folder.Parent
    }
    return $folder
}

Write-Host "Scanne SD-Karte nach neuen Alben... Bitte warten..." -ForegroundColor Yellow

# Holt alle Ordner, die Musikdateien enthalten
$allFolders = Get-ChildItem -Path $sdCard -Recurse -Directory
$albumPaths = @{}

foreach ($folder in $allFolders) {
    $hasMusic = Get-ChildItem -Path $folder.FullName -File | Where-Object { $_.Extension -match "\.(mp3|flac|wav|m4a)$" }
    if ($hasMusic) {
        $actualAlbumFolder = Get-ActualAlbumFolder $folder
        
        # Vorab-Prüfung mit dem Ordnernamen, um den Prozess zu beschleunigen
        $possibleSafeFolderName = Get-SafeFileName $actualAlbumFolder.Name
        if (Test-Path (Join-Path -Path $sdCard -ChildPath "$possibleSafeFolderName.m3u")) {
            continue 
        }

        # Versuche das Album-Tag zu lesen
        $albumName = Get-AlbumTagFromFolder $folder.FullName
        if (-not $albumName) {
            $albumName = Get-AlbumTagFromFolder $actualAlbumFolder.FullName
        }
        if (-not $albumName) {
            $albumName = $actualAlbumFolder.Name
        }

        # Falls die Playlist mit dem bereinigten Tag-Namen bereits existiert, überspringen
        $safePlaylistName = Get-SafeFileName $albumName
        if (Test-Path (Join-Path -Path $sdCard -ChildPath "$safePlaylistName.m3u")) {
            continue
        }

        $albumPaths[$actualAlbumFolder.FullName] = $albumName
    }
}

# Erstellt die Playlists im ROOT (nur für die neu gefundenen Alben)
$createdCount = 0
foreach ($albumPath in $albumPaths.Keys) {
    $albumName = $albumPaths[$albumPath]
    $safePlaylistName = Get-SafeFileName $albumName
    $playlistPath = Join-Path -Path $sdCard -ChildPath "$safePlaylistName.m3u"
    
    # Alle Tracks im Album-Ordner sammeln
    $tracks = Get-ChildItem -Path $albumPath -Recurse -File | Where-Object { 
        $_.Extension -match "\.(mp3|flac|wav|m4a)$" 
    } | Sort-Object FullName

    if ($tracks.Count -gt 0) {
        Write-Host "NEU ENTDECKT: Erstelle Playlist fuer $safePlaylistName" -ForegroundColor Green
        $createdCount++
        
        # UTF-8 OHNE BOM StreamWriter (Wichtig für die Sonderzeichen!)
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $writer = New-Object System.IO.StreamWriter($playlistPath, $false, $utf8NoBom)
        
        $writer.WriteLine("#EXTM3U")
        
        foreach ($track in $tracks) {
            $relativePath = $track.FullName.Substring($sdCard.Length)
            $writer.WriteLine($relativePath)
        }
        
        $writer.Close()
    }
}

# Aufräumen des COM-Objekts
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null

if ($createdCount -gt 0) {
    Write-Host "Fertig! Es wurden $createdCount neue Playlists mit schoenen Sonderzeichen erstellt." -ForegroundColor Cyan
} else {
    Write-Host "Fertig! Es wurden keine neuen Alben gefunden. Alle Playlists sind bereits aktuell." -ForegroundColor Yellow
}
