<#
.SYNOPSIS
    KSIS Gymnastics Results Parser (PowerShell Port)
.DESCRIPTION
    Scrapes gymnastics results, applies corrections, and exports to CSV.
    Supports dynamic columns, date range search, and data aggregation.
#>

param(
    [string]$PropId,
    [switch]$List,
    [switch]$DebugMode
)

# Enable ANSI escape sequences for color in standard console
if ($Host.UI.SupportsVirtualTerminal) { $Host.UI.RawUI.WindowTitle = "KSIS Parser" }

# --- Global Config ---
$Global:Debug = $DebugMode
$Global:ClubCorrections = @{}
$Global:AthleteCorrections = @{}
$Global:NameCache = @{}
# Session variable to persist cookies
$Global:KsisSession = New-Object Microsoft.PowerShell.Commands.WebRequestSession

# Drop these columns from output
$Global:DropCols = @('E', 'Bonus', 'Comp', 'ND', 'D', 'SV')

# Map HTML headers to CSV headers
# Note: PowerShell hash tables are case-insensitive, so 'Born' handles 'born' automatically
$Global:RenameMap = @{ 'Total' = 'Score'; 'SV' = 'Score'; 'Born' = 'YOB' }

# --- Helper Functions ---

function Write-Color {
    param([string]$Text, [ConsoleColor]$Color = [ConsoleColor]::White, [switch]$NoNewLine)
    if ($NoNewLine) { Write-Host $Text -ForegroundColor $Color -NoNewline }
    else { Write-Host $Text -ForegroundColor $Color }
}

function Write-DebugLog {
    param([string]$Message)
    if ($Global:Debug) { Write-Color "[DEBUG] $Message" -Color Magenta }
}

function Get-CleanText {
    param([string]$Html)
    if ([string]::IsNullOrWhiteSpace($Html)) { return "" }
    # Remove tags and decode HTML entities
    $text = $Html -replace '<[^>]+>', ' ' -replace '\s+', ' '
    return [System.Net.WebUtility]::HtmlDecode($text).Trim()
}

function Request-Url {
    param([string]$Url)
    
    # FORCE TLS 1.2 (Fixes many timeout/connection drop issues on older Windows)
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $maxRetries = 5
    $retryCount = 0
    $success = $false
    
    while (-not $success -and $retryCount -lt $maxRetries) {
        try {
            # Exponential Backoff: Wait 2s, 5s, 9s, 14s...
            $sleepTime = 2 + ($retryCount * 3) 
            if ($retryCount -gt 0) { 
                Write-DebugLog "Waiting ${sleepTime}s before retry..." 
                Start-Sleep -Seconds $sleepTime
            } else {
                Start-Sleep -Milliseconds 500 # Initial politeness delay
            }
            
            Write-DebugLog "Fetching: $Url (Attempt $($retryCount + 1))"
            
            $headers = @{
                "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
                "Accept" = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
                "Accept-Language" = "en-US,en;q=0.5"
                "Upgrade-Insecure-Requests" = "1"
                "Cache-Control" = "max-age=0"
            }

            # Use -WebSession to persist cookies across requests
            $response = Invoke-WebRequest -Uri $Url `
                                          -WebSession $Global:KsisSession `
                                          -Headers $headers `
                                          -UseBasicParsing `
                                          -Method Get `
                                          -TimeoutSec 120 `
                                          -ErrorAction Stop
            
            return $response.Content
        }
        catch {
            # Handle specific HTTP error codes if available
            if ($_.Exception.Response) {
                $errorCode = $_.Exception.Response.StatusCode.value__
                Write-DebugLog "Error fetching URL: $_ (Status: $errorCode)"
                
                # Retry on Gateway/Timeout errors (502, 503, 504, 522)
                if ($errorCode -in 502, 503, 504, 522) {
                    Write-Color "  ! Server busy/timeout ($errorCode). Retrying..." -Color Yellow
                    $retryCount++
                } else {
                    Write-Color "x Fatal Error fetching URL: $($_.Exception.Message)" -Color Red
                    return $null
                }
            } else {
                # Handle generic timeouts or connection drops
                Write-Color "x Connection Error: $($_.Exception.Message)" -Color Red
                
                if ($_.Exception.Message -match "Timeout" -or $_.Exception.Message -match "forcibly closed") {
                     Write-Color "  ! Retrying connection..." -Color Yellow
                     $retryCount++
                } else {
                     return $null
                }
            }
        }
    }
    
    Write-Color "x Failed to fetch URL after $maxRetries attempts." -Color Red
    return $null
}

# --- Correction Logic ---

function Load-Corrections {
    Write-DebugLog "Loading corrections..."
    
    # Load Club Corrections
    if (Test-Path "Club Name Corrections.csv") {
        $csv = Import-Csv "Club Name Corrections.csv" -Header "Original","Corrected"
        foreach ($row in $csv) {
            if ($row.Original -ne "Original") { 
                $Global:ClubCorrections[$row.Original.Trim()] = $row.Corrected.Trim()
            }
        }
        if ($Global:Debug) { Write-DebugLog "Loaded $($Global:ClubCorrections.Count) club corrections." }
    }

    # Load Athlete Corrections
    if (Test-Path "Athlete Name Corrections.csv") {
        $csv = Import-Csv "Athlete Name Corrections.csv" -Header "Original","Corrected"
        foreach ($row in $csv) {
            if ($row.Original -ne "Original") {
                $Global:AthleteCorrections[$row.Original.Trim()] = $row.Corrected.Trim()
            }
        }
        if ($Global:Debug) { Write-DebugLog "Loaded $($Global:AthleteCorrections.Count) athlete corrections." }
    }
}

function Save-AthleteCorrection {
    param([string]$Original, [string]$Corrected)
    
    $file = "Athlete Name Corrections.csv"
    $obj = [PSCustomObject]@{ Original = $Original; Corrected = $Corrected }
    
    # Update Memory
    $Global:AthleteCorrections[$Original] = $Corrected
    
    # Append to File
    try {
        $obj | Export-Csv -Path $file -Append -NoTypeInformation -Force
        Write-Color "  ✓ Saved correction to '$file'" -Color Green
    }
    catch {
        Write-Color "  ! Could not save correction to file (might be open)." -Color Red
    }
}

function Standardize-Club {
    param([string]$ClubName)
    if ([string]::IsNullOrWhiteSpace($ClubName)) { return "" }
    
    $clean = $ClubName.Trim()
    
    # 1. Check Dictionary
    if ($Global:ClubCorrections.ContainsKey($clean)) {
        $clean = $Global:ClubCorrections[$clean]
    }
    
    # 2. Suffix Regex Clean (Inc., ON, etc)
    $clean = $clean -replace '\s+(?:Inc\.?\s+ON|ON\s+Inc\.?|Inc\.?|ON)$', ''
    return $clean.Trim()
}

function Reorder-Name {
    param([string]$Name)
    $Name = $Name -replace '\s+', ' '
    $Name = $Name.Trim()

    # Check caches
    if ($Global:AthleteCorrections.ContainsKey($Name)) { return $Global:AthleteCorrections[$Name] }
    if ($Global:NameCache.ContainsKey($Name)) { return $Global:NameCache[$Name] }

    $parts = $Name.Split(' ')
    $result = $Name

    if ($parts.Count -eq 2) {
        $result = "$($parts[1]) $($parts[0])"
    }
    elseif ($parts.Count -gt 2) {
        Write-Host "`n----------------------------------------" -ForegroundColor Yellow
        Write-Host "Multiple-word name detected: " -NoNewline; Write-Host $Name -ForegroundColor White
        Write-Host "This name is in 'Last First' format. Where does the LAST name end?" -ForegroundColor Cyan
        
        for ($i = 1; $i -lt $parts.Count; $i++) {
            $last = ($parts[0..($i-1)] -join ' ')
            $first = ($parts[$i..($parts.Count-1)] -join ' ')
            Write-Host "$i. " -NoNewline -ForegroundColor Cyan
            Write-Host "Last: $last, First: $first -> " -NoNewline
            Write-Host "$first $last" -ForegroundColor Green
        }

        while ($true) {
            $choice = Read-Host -Prompt "Enter choice"
            if ($choice -match '^\d+$' -and [int]$choice -gt 0 -and [int]$choice -lt $parts.Count) {
                $idx = [int]$choice
                $last = ($parts[0..($idx-1)] -join ' ')
                $first = ($parts[$idx..($parts.Count-1)] -join ' ')
                $result = "$first $last"
                
                Save-AthleteCorrection -Original $Name -Corrected $result
                break
            }
            Write-Color "Invalid choice." -Color Red
        }
        Write-Host "----------------------------------------`n" -ForegroundColor Yellow
    }

    $Global:NameCache[$Name] = $result
    return $result
}

# --- Parsing Logic ---

function Parse-RowData {
    param($RowHtml, $Headers)
    
    # Extract Cells (Regex to find <td> or <th> content)
    $cells = [regex]::Matches($RowHtml, '<td.*?>(.*?)</td>', 'IgnoreCase') | ForEach-Object { $_.Groups[1].Value }
    
    if ($cells.Count -lt 4) { return $null }
    
    # Clean cell text
    $cellValues = $cells | ForEach-Object { Get-CleanText $_ }
    
    $rowDict = [Ordered]@{}
    
    # 1. Identify Name/Club Column
    $athIdx = -1
    for ($i=0; $i -lt $Headers.Count; $i++) {
        if ($Headers[$i] -match 'Name|Gymnast') { $athIdx = $i; break }
    }
    if ($athIdx -eq -1 -and $cells.Count -gt 2) { $athIdx = 2 } # Fallback

    # 2. Identify YOB Column
    $yobIdx = -1
    for ($i=0; $i -lt $Headers.Count; $i++) {
        if ($Headers[$i] -match 'Born|YOB') { $yobIdx = $i; break }
    }
    # Fallback YOB check (is it a 4 digit year?)
    if ($yobIdx -eq -1 -and $cells.Count -gt 3) {
        if ($cellValues[3] -match '^\d{4}$') { $yobIdx = 3 }
    }

    # Process Athlete Column (Split by <br>)
    if ($athIdx -ne -1 -and $athIdx -lt $cells.Count) {
        $rawHtml = $cells[$athIdx]
        $parts = $rawHtml -split '<br\s*/?>'
        
        $rawName = Get-CleanText $parts[0]
        $rawClub = if ($parts.Count -gt 1) { Get-CleanText $parts[1] } else { "" }
        
        $rowDict['Name'] = Reorder-Name $rawName
        $rowDict['Club'] = Standardize-Club $rawClub
    } else {
        $rowDict['Name'] = "Unknown"
        $rowDict['Club'] = "Unknown"
    }

    # Process YOB
    if ($yobIdx -ne -1 -and $yobIdx -lt $cells.Count) {
        $rowDict['YOB'] = $cellValues[$yobIdx]
    }

    # Process Other Columns
    for ($i=0; $i -lt $Headers.Count; $i++) {
        if ($i -eq $athIdx -or $i -eq $yobIdx) { continue }
        if ($i -ge $cellValues.Count) { break }

        $h = $Headers[$i].Trim()
        
        # Rename logic
        if ($Global:RenameMap.ContainsKey($h)) { $h = $Global:RenameMap[$h] }
        
        # Skip if empty or extracted elsewhere
        if ([string]::IsNullOrWhiteSpace($h)) { $h = "Col_$i" }
        
        $rowDict[$h] = $cellValues[$i]
    }

    return $rowDict
}

function Fetch-CompetitionData {
    param([string]$PropId)
    
    Load-Corrections
    $url = "https://ksis.eu/resultx.php?id_prop=$PropId"
    Write-Color "`nFetching competition data (prop_id: $PropId)..." -Color Cyan
    
    $html = Request-Url $url
    if (-not $html) { return $null, @(), @() }

    # Extract Comp Name (<h3>)
    $compName = "Unknown"
    if ($html -match '<h3>(.*?)</h3>') { 
        $compName = Get-CleanText $matches[1] 
        $compName = $compName -replace '[\\/*?:"<>|]', '-'
    }

    # Extract Date (<h4>)
    $compDate = (Get-Date).ToString("yyyy-MM-dd")
    if ($html -match '<h4>(.*?)</h4>') {
        $rawDate = Get-CleanText $matches[1]
        if ($rawDate -match '(\d{1,2})\.(\d{1,2})\.(\d{4})') {
            $compDate = "$($matches[3])-$($matches[2].PadLeft(2,'0'))-$($matches[1].PadLeft(2,'0'))"
        }
    }

    Write-Color "Parsing sessions for ""$compName""..." -Color Cyan

    # Extract Sessions
    # IMPROVED: Fallback logic if <select> not found by regex
    $sessions = @()
    if ($html -match '(?s)<select[^>]*id="id_sut"[^>]*>(.*?)</select>') {
        # Standard case: Found the specific select box
        $selectContent = $matches[1]
        $sessions = [regex]::Matches($selectContent, '<option value="([^"]+)">([^<]+)</option>')
    } else {
        # Fallback: Scan ENTIRE html for any option value that looks like an ID
        Write-DebugLog "Strict select tag not found. Using fallback scan."
        $sessions = [regex]::Matches($html, '<option value="(\d+)">([^<]+)</option>')
    }

    if ($sessions.Count -eq 0) {
        Write-Color "x No sessions found." -Color Red
        
        # Dump HTML for debugging
        if ($Global:Debug) {
            $dumpFile = "debug_page_dump.html"
            $html | Out-File $dumpFile
            Write-Color "  HTML dump saved to $dumpFile" -Color Magenta
        }
        return $compName, @(), @()
    }

    $allRows = @()
    $allHeaders = [System.Collections.Generic.HashSet[string]]::new()
    $allHeaders.Add("Competition") > $null
    $allHeaders.Add("Session") > $null
    $allHeaders.Add("Name") > $null
    $allHeaders.Add("Club") > $null
    $allHeaders.Add("Score") > $null
    $allHeaders.Add("Date") > $null
    
    foreach ($match in $sessions) {
        $val = $match.Groups[1].Value
        $sName = ($match.Groups[2].Value).Trim()
        
        if ($val -eq "0" -or $val -eq "") { continue }

        $resUrl = "https://ksis.eu/load_result_total_ksismg_art.php?lang=en&id_prop=$PropId&id_sut=$val&rn=null&mn=null&state=-1&age_group=&award=-1&nacinie=undefined"
        $resHtml = Request-Url $resUrl
        if (-not $resHtml) { continue }

        # Extract Table
        if ($resHtml -match '(?s)<table[^>]*id="myTablePrihlasky"[^>]*>(.*?)</table>') {
            $tableContent = $matches[1]
            
            # Extract Rows
            $rows = [regex]::Matches($tableContent, '(?s)<tr.*?>(.*?)</tr>')
            if ($rows.Count -eq 0) { continue }

            # Determine Headers (Check <thead> first, else first row)
            $headerRowMatch = $null
            if ($tableContent -match '(?s)<thead.*?>(.*?)</thead>') {
                # Get last TR in thead
                $theadRows = [regex]::Matches($matches[1], '(?s)<tr.*?>(.*?)</tr>')
                if ($theadRows.Count -gt 0) { $headerRowMatch = $theadRows[$theadRows.Count-1] }
            }
            
            # If no thead headers, use first row
            $dataStartIndex = 0
            if (-not $headerRowMatch) {
                $headerRowMatch = $rows[0]
                $dataStartIndex = 1 # Skip first row in data loop
            }

            # Parse Headers
            $rawHeaders = [regex]::Matches($headerRowMatch.Groups[1].Value, '<(?:th|td).*?>(.*?)</(?:th|td)>') | ForEach-Object { Get-CleanText $_.Groups[1].Value }
            
            # Add to Master Header List
            foreach ($h in $rawHeaders) {
                if ($Global:RenameMap.ContainsKey($h)) { $h = $Global:RenameMap[$h] }
                if (-not [string]::IsNullOrWhiteSpace($h) -and $h -notin $Global:DropCols) {
                    $allHeaders.Add($h) > $null
                }
            }

            # Parse Data Rows
            $count = 0
            for ($i = $dataStartIndex; $i -lt $rows.Count; $i++) {
                $rowData = Parse-RowData -RowHtml $rows[$i].Groups[1].Value -Headers $rawHeaders
                if ($rowData) {
                    $rowData['Competition'] = $compName
                    $rowData['Session'] = $sName
                    $rowData['Date'] = $compDate
                    
                    # Convert to PSCustomObject for export
                    $obj = [PSCustomObject]$rowData
                    $allRows += $obj
                    $count++
                }
            }

            if ($count -gt 0) {
                Write-Color "  $([char]0x2713) $sName`: $count athletes" -Color Green
            } else {
                Write-Color "  ! $sName`: No athletes found" -Color Yellow
            }
        }
    }

    return $compName, $allRows, $allHeaders
}

# --- Menu Functions ---

function Get-CompetitionsFromMenu {
    $url = "https://ksis.eu/menu.php?akcia=S&oblast=ARTW&country=CAN"
    Write-Color "`nFetching competition list..." -Color Cyan
    $html = Request-Url $url
    if (-not $html) { return @() }

    # Find all links with id_prop
    # This regex looks for the row <tr> containing the link to capture the date from the first column
    $rowMatches = [regex]::Matches($html, '(?s)<tr.*?>(.*?)</tr>')
    
    $comps = @()
    foreach ($row in $rowMatches) {
        $rowContent = $row.Groups[1].Value
        
        # Check if row has prop link
        if ($rowContent -match 'id_prop=(\d+)') {
            $propId = $matches[1]
            
            # Extract Name from Link
            $name = "Unknown"
            if ($rowContent -match '<a[^>]*>(.*?)</a>') {
                $name = Get-CleanText $matches[1]
            }

            # Extract Date (Look for DD.MM.YYYY in the whole row)
            $dateStr = $null
            if ($rowContent -match '(\d{2}\.\d{2}\.\d{4})') {
                $dParts = $matches[1].Split('.')
                $dateStr = "$($dParts[2])-$($dParts[1])-$($dParts[0])" # YYYY-MM-DD
            }

            $isLive = ($rowContent -match 'badge.*?>\s*LIVE')

            # Add to list if unique
            $exists = $comps | Where-Object { $_.id -eq $propId }
            if (-not $exists) {
                $comps += [PSCustomObject]@{ id = $propId; name = $name; date = $dateStr; isLive = $isLive }
            }
        }
    }
    return $comps
}

function Export-CsvData {
    param($Rows, $Headers, $Filename)
    
    if ($Rows.Count -eq 0) {
        Write-Color "No rows to write." -Color Yellow
        return
    }

    # Sort Headers: Priority first, then alphabetical
    $priority = @('Competition', 'Session', 'Name', 'YOB', 'Club', 'Score', 'Date')
    $otherHeaders = $Headers | Where-Object { $_ -notin $priority -and $_ -notin $Global:DropCols } | Sort-Object
    $finalHeaders = $priority + $otherHeaders

    # Select properties in order
    $exportData = $Rows | Select-Object $finalHeaders

    try {
        $exportData | Export-Csv -Path $Filename -NoTypeInformation -Encoding UTF8
        Write-Color "`n  $([char]0x2713) Successfully created $Filename with $($Rows.Count) records." -Color Green
    }
    catch {
        Write-Color "ERROR: Could not write to file. Is it open in Excel?" -Color Red
    }
}

function Search-ByDate {
    Write-Color "`n--- Date Range Search ---" -Color Cyan
    $startStr = Read-Host "Enter Start Date (YYYY-MM-DD)"
    $endStr   = Read-Host "Enter End Date   (YYYY-MM-DD)"

    try {
        $startDate = [DateTime]::ParseExact($startStr, "yyyy-MM-dd", $null)
        $endDate   = [DateTime]::ParseExact($endStr, "yyyy-MM-dd", $null)
    }
    catch {
        Write-Color "Invalid format. Use YYYY-MM-DD." -Color Red
        return
    }

    $comps = Get-CompetitionsFromMenu
    $matches = @()

    foreach ($c in $comps) {
        if ($c.date) {
            try {
                $cDate = [DateTime]::ParseExact($c.date, "yyyy-MM-dd", $null)
                if ($cDate -ge $startDate -and $cDate -le $endDate) {
                    $matches += $c
                }
            } catch {}
        }
    }

    if ($matches.Count -eq 0) {
        Write-Color "No competitions found in range." -Color Yellow
        return
    }

    Write-Color "`nFound $($matches.Count) competitions:" -Color Green
    foreach ($m in $matches) {
        Write-Host " - $($m.date): $($m.name) (ID: $($m.id))"
    }

    $conf = Read-Host "`nExport all $($matches.Count) competitions? (y/n)"
    if ($conf -eq 'y') {
        $allRows = @()
        $allHeaders = [System.Collections.Generic.HashSet[string]]::new()
        
        foreach ($m in $matches) {
            $res = Fetch-CompetitionData -PropId $m.id
            if ($res[1]) {
                $allRows += $res[1]
                foreach ($h in $res[2]) { $allHeaders.Add($h) > $null }
            }
        }
        
        $fname = "Aggregated_Results_${startStr}_to_${endStr}.csv"
        Export-CsvData -Rows $allRows -Headers $allHeaders -Filename $fname
    }
}

function Show-Menu {
    while ($true) {
        Write-Color "`n=======================================" -Color Cyan
        Write-Color "     KSIS Competition Results Tool     " -Color Cyan
        Write-Color "=======================================" -Color Cyan
        Write-Color "1. List all competitions" -Color Cyan
        Write-Color "2. List live competitions only" -Color Cyan
        Write-Color "3. Search competitions by keyword" -Color Cyan
        Write-Color "4. Export results by prop_id (Single or List)" -Color Cyan
        Write-Color "5. Search & Export by Date Range" -Color Cyan
        Write-Color "6. Exit" -Color Cyan

        $choice = Read-Host "`nEnter your choice (1-6)"
        
        switch ($choice) {
            '1' { $c = Get-CompetitionsFromMenu; Show-Comps $c }
            '2' { $c = Get-CompetitionsFromMenu | Where-Object { $_.isLive }; Show-Comps $c }
            '3' { 
                $k = Read-Host "Enter keyword"
                $c = Get-CompetitionsFromMenu | Where-Object { $_.name -match $k }
                Show-Comps $c 
            }
            '4' {
                $inputIds = Read-Host "Enter prop_id (comma separated for multiple)"
                if (-not [string]::IsNullOrWhiteSpace($inputIds)) {
                    # FIX: Force array type to avoid single string char iteration
                    $ids = @($inputIds -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -match '^\d+$' })
                    
                    if ($ids.Count -eq 1) {
                        $res = Fetch-CompetitionData -PropId $ids[0]
                        $ts = (Get-Date).ToString("yyyyMMddHHmm")
                        $fname = "$($res[0])-$ts.csv"
                        Export-CsvData $res[1] $res[2] $fname
                    } elseif ($ids.Count -gt 1) {
                        $allRows = @()
                        $allHeaders = [System.Collections.Generic.HashSet[string]]::new()
                        foreach ($id in $ids) {
                            $res = Fetch-CompetitionData -PropId $id
                            $allRows += $res[1]
                            foreach ($h in $res[2]) { $allHeaders.Add($h) > $null }
                        }
                        Export-CsvData $allRows $allHeaders "Aggregated_Manual_List.csv"
                    }
                }
            }
            '5' { Search-ByDate }
            '6' { Write-Color "Goodbye!" -Color Cyan; return }
            Default { Write-Color "Invalid Choice" -Color Red }
        }
    }
}

function Show-Comps {
    param($Comps)
    if ($Comps.Count -eq 0) { Write-Color "No competitions found." -Color Yellow; return }
    
    Write-Color "`nAvailable Competitions:" -Color White -NoNewLine; Write-Host ""
    Write-Host ("{0,-8} {1,-12} {2}" -f "ID","Date","Competition Name") -ForegroundColor Cyan
    Write-Host ("-" * 80)
    
    foreach ($c in $Comps) {
        $live = if ($c.isLive) { " [LIVE]" } else { "" }
        $color = if ($c.isLive) { "Red" } else { "Green" }
        # Only color the ID green/red, rest standard
        Write-Host ("{0,-8}" -f $c.id) -ForegroundColor Green -NoNewline
        Write-Host (" {0,-12} {1}" -f ($c.date ?? "Unknown"), $c.name) -NoNewline
        if ($c.isLive) { Write-Host $live -ForegroundColor Red } else { Write-Host "" }
    }
    Write-Host "`nTotal: $($Comps.Count) competitions" -ForegroundColor White
}

# --- Main Entry Point ---

if ($PropId) {
    # Direct Export Mode
    $res = Fetch-CompetitionData -PropId $PropId
    $ts = (Get-Date).ToString("yyyyMMddHHmm")
    $fname = "$($res[0])-$ts.csv"
    Export-CsvData $res[1] $res[2] $fname
}
elseif ($List) {
    $c = Get-CompetitionsFromMenu
    Show-Comps $c
}
else {
    Show-Menu
}
