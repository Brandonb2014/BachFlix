<#
.SYNOPSIS
  One-time move of open requests from the old "Movie/TV Show Requests" Google Sheet into Seerr.

.DESCRIPTION
  Two phases, so nothing is created until you've checked the matches:

  1. Prepare: reads the sheet CSV, keeps the open rows, looks each title up with Seerr's
     own search, guesses which Seerr user made it, and writes a review CSV.
       .\Import-SheetRequests.ps1 -Prepare -SheetCsv .\requests.csv

  2. Fix the review CSV in Excel/Sheets: correct TmdbId/MediaType where the match is wrong,
     fill in SeerrUserId (see seerr-users.csv), set Include to N for rows to skip.

  3. Import: dry-run first (default), then -Apply to create the requests.
       .\Import-SheetRequests.ps1 -Import
       .\Import-SheetRequests.ps1 -Import -Apply

  Requests created with the admin API key are auto-approved, so they show up for the
  requester as "Approved" (on your list) rather than "Pending".

  The API key is in Seerr under Settings > General. Pass -ApiKey or set $env:SEERR_API_KEY.
#>
[CmdletBinding(DefaultParameterSetName = 'Prepare')]
param(
    [Parameter(ParameterSetName = 'Prepare', Mandatory = $true)][switch]$Prepare,
    [Parameter(ParameterSetName = 'Prepare', Mandatory = $true)][string]$SheetCsv,
    # Sheet statuses treated as closed and skipped. Everything else (including blank) is open.
    [Parameter(ParameterSetName = 'Prepare')][string[]]$ClosedStatuses = @('Done', 'Not Available'),

    [Parameter(ParameterSetName = 'Import', Mandatory = $true)][switch]$Import,
    [Parameter(ParameterSetName = 'Import')][switch]$Apply,

    [string]$ReviewCsv = (Join-Path $PSScriptRoot 'import-review.csv'),
    [string]$SeerrUrl = $(if ($env:SEERR_URL) { $env:SEERR_URL } else { 'http://localhost:5055' }),
    [string]$ApiKey = $env:SEERR_API_KEY
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Invoke-Seerr {
    param([string]$Method = 'GET', [string]$Path, $Body)
    $uri = $SeerrUrl.TrimEnd('/') + '/api/v1' + $Path
    $req = @{ Method = $Method; Uri = $uri; Headers = @{ 'X-Api-Key' = $ApiKey }; UseBasicParsing = $true }
    if ($null -ne $Body) {
        $req.ContentType = 'application/json'
        $req.Body = [Text.Encoding]::UTF8.GetBytes(($Body | ConvertTo-Json -Depth 5 -Compress))
    }
    $resp = Invoke-WebRequest @req
    # Decode as UTF-8 explicitly; PS 5.1 otherwise mangles accented titles.
    $text = [Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    return $text | ConvertFrom-Json
}

function Get-ErrorText($err) {
    try {
        $stream = $err.Exception.Response.GetResponseStream()
        $reader = New-Object IO.StreamReader($stream)
        $body = $reader.ReadToEnd()
        if ($body) { return "$($err.Exception.Message) $body" }
    } catch { }
    return $err.Exception.Message
}

# Turns free-text like "Ted Lasso Season 3 (edited if possible, but no rush)" into
# a search query, a year, a movie/tv hint and requested seasons.
function ConvertTo-SearchHint([string]$raw) {
    $t = $raw -replace [char]0x2019, "'" -replace [char]0x2018, "'"
    $t = $t -replace '\s+', ' '
    $hint = [ordered]@{ Query = ''; Year = $null; Type = $null; Seasons = 'all' }

    if ($t -match '\b(19[2-9]\d|20[0-4]\d)\b') { $hint.Year = [int]$Matches[1] }

    if ($t -match '(?i)\b(seasons?|series|episodes?|tv show|show)\b' -and $t -notmatch '(?i)\btv movie\b') { $hint.Type = 'tv' }
    if ($t -match '(?i)\b(movie|film)\b') { $hint.Type = 'movie' }

    if ($t -match '(?i)\bseasons?\s+(\d{1,2})\s*(?:-|to|through|thru)\s*(\d{1,2})\b') {
        $hint.Seasons = @([int]$Matches[1]..[int]$Matches[2])
    } elseif ($t -match '(?i)\bseason\s+(\d{1,2})\b') {
        $hint.Seasons = @([int]$Matches[1])
    }

    $q = $t -replace [char]0x201C, '"' -replace [char]0x201D, '"'
    # 'Edit the movie "about time" when you have time' -> search just the quoted title.
    if ($q -match '"([^"]{2,})"' -and $q -notmatch '\([^)]*"[^"]*"[^)]*\)') { $q = $Matches[1] }
    $q = $q -replace '\([^)]*\)', ' '                                  # (anything in parens)
    $q = $q -replace '(?i)\b(netflix|amazon|prime video|prime|disney\+?|hulu|apple tv\+?|paramount\+?|peacock|max|hbo|angel studios|sub ?titles?)\b', ' '
    $q = $q -replace '(?i)\bseasons?\b.*$', ' '                        # "season 3 ..." to end
    $q = $q -replace '(?i)\b(with|w/|edited|no rush|if possible|please|thank you|thanks|released in|released|tv movie|disney channel movie|movie|film|the series|tv show|series|captions|subtitles)\b.*$', ' '
    $q = $q -replace '\s+-\s+.*$', ' '                                 # " - angel studios"
    $q = $q -replace '(?i)\btv\b', ' '                                 # "Wings tv 1990"
    $q = $q -replace '\b(19[2-9]\d|20[0-4]\d)\b', ' '                   # years
    $q = $q -replace '[^\p{L}\p{N}'':&!?.\s]', ' '
    $q = ($q -replace '\s+', ' ').Trim(' ', '.', ',', ':', '-')
    if (-not $q) { $q = $t.Trim() }
    $hint.Query = $q
    return [pscustomobject]$hint
}

function Get-ResultYear($r) {
    $d = if ($r.mediaType -eq 'tv') { $r.firstAirDate } else { $r.releaseDate }
    if ($d -and $d.Length -ge 4) { return [int]$d.Substring(0, 4) }
    return $null
}

function Select-BestMatch($results, $hint) {
    $media = @($results | Where-Object { $_.mediaType -eq 'movie' -or $_.mediaType -eq 'tv' })
    if ($media.Count -eq 0) { return $null }
    $scored = for ($i = 0; $i -lt $media.Count; $i++) {
        $r = $media[$i]
        $score = 100 - ($i * 5)                       # Seerr/TMDB relevance order
        $name = if ($r.mediaType -eq 'tv') { $r.name } else { $r.title }
        if ($name -and $name.ToLowerInvariant() -eq $hint.Query.ToLowerInvariant()) { $score += 30 }
        if ($hint.Type -and $r.mediaType -eq $hint.Type) { $score += 25 }
        $y = Get-ResultYear $r
        if ($hint.Year -and $y) {
            if ($y -eq $hint.Year) { $score += 40 } elseif ([math]::Abs($y - $hint.Year) -le 1) { $score += 20 } else { $score -= 15 }
        }
        [pscustomobject]@{ Score = $score; Result = $r }
    }
    return @($scored | Sort-Object Score -Descending)
}

function Format-Result($r) {
    $name = if ($r.mediaType -eq 'tv') { $r.name } else { $r.title }
    $y = Get-ResultYear $r
    return "{0}:{1} {2}{3}" -f $r.mediaType, $r.id, $name, $(if ($y) { " ($y)" } else { '' })
}

function Normalize-Name([string]$s) { return ($s.ToLowerInvariant() -replace '[^a-z ]', ' ' -replace '\s+', ' ').Trim() }

function Find-SeerrUser([string]$requester, $users) {
    $req = Normalize-Name ($requester -replace '(?i)\bfor\b.*$', '' -replace '-.*$', '')
    if (-not $req) { return $null }
    $first = ($req -split ' ')[0]
    $hits = @($users | Where-Object {
        $names = @($_.displayName, $_.plexUsername, $_.username, ($_.email -replace '@.*$', '')) | Where-Object { $_ } | ForEach-Object { Normalize-Name $_ }
        ($names -contains $req) -or ($names | Where-Object { $_ -eq $first -or $_.StartsWith("$first ") })
    })
    if ($hits.Count -eq 1) { return $hits[0] }
    return $null
}

if (-not $ApiKey) { throw 'No Seerr API key. Pass -ApiKey or set $env:SEERR_API_KEY (Seerr > Settings > General).' }

if ($Prepare) {
    $rows = Import-Csv -LiteralPath $SheetCsv -Encoding UTF8
    $usersResp = Invoke-Seerr -Path '/user?take=1000'
    $users = @($usersResp.results)
    $usersCsv = Join-Path (Split-Path $ReviewCsv) 'seerr-users.csv'
    $users | Select-Object id, displayName, plexUsername, email | Export-Csv -LiteralPath $usersCsv -NoTypeInformation -Encoding UTF8
    Write-Host "Seerr has $($users.Count) users (written to $usersCsv)."

    $out = New-Object System.Collections.Generic.List[object]
    $sheetRow = 1
    foreach ($row in $rows) {
        $sheetRow++
        $status = ($row.Status + '').Trim()
        if ($ClosedStatuses -contains $status) { continue }
        if (-not ($row.Title + '').Trim()) { continue }

        $hint = ConvertTo-SearchHint $row.Title
        $ranked = @()
        try {
            $search = Invoke-Seerr -Path ('/search?page=1&query=' + [Uri]::EscapeDataString($hint.Query))
            $ranked = @(Select-BestMatch $search.results $hint)
        } catch {
            Write-Warning "Search failed for '$($hint.Query)': $(Get-ErrorText $_)"
        }
        $best = if ($ranked.Count) { $ranked[0].Result } else { $null }
        $user = Find-SeerrUser $row.Name $users

        $seasonsText = if ($hint.Seasons -is [array]) { ($hint.Seasons -join ',') } else { 'all' }
        $out.Add([pscustomobject][ordered]@{
            Include      = if ($best -and $user) { 'Y' } else { '?' }
            SheetRow     = $sheetRow
            Timestamp    = $row.Timestamp
            Requester    = $row.Name
            SheetTitle   = $row.Title
            SheetStatus  = $status
            SheetNotes   = $row.Notes
            SearchedFor  = $hint.Query
            MediaType    = if ($best) { $best.mediaType } else { '' }
            TmdbId       = if ($best) { $best.id } else { '' }
            MatchedTitle = if ($best) { Format-Result $best } else { '(no match)' }
            Seasons      = if ($best -and $best.mediaType -eq 'tv') { $seasonsText } else { '' }
            SeerrUserId  = if ($user) { $user.id } else { '' }
            SeerrUser    = if ($user) { $user.displayName } else { '' }
            Alternatives = (@($ranked | Select-Object -Skip 1 -First 4 | ForEach-Object { Format-Result $_.Result }) -join ' | ')
        })
        Write-Host ("{0,-4} {1,-45} -> {2}" -f $sheetRow, ($row.Title.Substring(0, [Math]::Min(45, $row.Title.Length))), $(if ($best) { Format-Result $best } else { '(no match)' }))
    }
    $out | Export-Csv -LiteralPath $ReviewCsv -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "Wrote $($out.Count) open requests to $ReviewCsv"
    Write-Host "Review it: fix wrong matches (MediaType/TmdbId/Seasons), fill SeerrUserId, set Include=Y or N. Then run -Import."
    return
}

if ($Import) {
    $rows = @(Import-Csv -LiteralPath $ReviewCsv -Encoding UTF8)
    $todo = @($rows | Where-Object { $_.Include -eq 'Y' })
    $notReady = @($todo | Where-Object { -not $_.TmdbId -or -not $_.SeerrUserId -or ($_.MediaType -ne 'movie' -and $_.MediaType -ne 'tv') })
    if ($notReady.Count) {
        $notReady | ForEach-Object { Write-Warning "Row $($_.SheetRow) '$($_.SheetTitle)' is Include=Y but missing MediaType, TmdbId or SeerrUserId." }
        throw 'Fix the rows above (or set Include=N) and run again.'
    }
    $skipped = @($rows | Where-Object { $_.Include -ne 'Y' }).Count
    Write-Host "$($todo.Count) requests to create, $skipped rows skipped. Mode: $(if ($Apply) { 'APPLY' } else { 'dry run (add -Apply to create)' })"
    Write-Host ""

    $log = New-Object System.Collections.Generic.List[object]
    $verifiedUserId = $false
    foreach ($r in $todo) {
        $body = [ordered]@{ mediaType = $r.MediaType; mediaId = [int]$r.TmdbId; userId = [int]$r.SeerrUserId }
        if ($r.MediaType -eq 'tv') {
            $s = ($r.Seasons + '').Trim()
            if (-not $s -or $s -eq 'all') { $body.seasons = 'all' } else { $body.seasons = @($s -split '[,\s]+' | Where-Object { $_ } | ForEach-Object { [int]$_ }) }
        }
        $label = "row $($r.SheetRow): $($r.MatchedTitle) for $($r.SeerrUser) [user $($r.SeerrUserId)]"
        if (-not $Apply) {
            Write-Host "WOULD CREATE $label"
            continue
        }
        try {
            $created = Invoke-Seerr -Method POST -Path '/request' -Body $body
            $by = $created.requestedBy.id
            if ([int]$by -ne [int]$r.SeerrUserId) {
                # The API docs don't spell out whether an API key may file on behalf of
                # another user. If Seerr ignored userId, stop before filing everything as admin.
                Write-Warning "Request $($created.id) was filed under user $by, not $($r.SeerrUserId). Stopping. Reassign or delete it in Seerr."
                $log.Add([pscustomobject]@{ SheetRow = $r.SheetRow; Title = $r.MatchedTitle; Result = "WRONG USER ($by)"; RequestId = $created.id })
                break
            }
            if (-not $verifiedUserId) { Write-Host '(Confirmed: Seerr files requests under the chosen user.)'; $verifiedUserId = $true }
            Write-Host "CREATED  $label -> request $($created.id), status $($created.status)"
            $log.Add([pscustomobject]@{ SheetRow = $r.SheetRow; Title = $r.MatchedTitle; Result = 'created'; RequestId = $created.id })
        } catch {
            $msg = Get-ErrorText $_
            Write-Warning "FAILED   $label : $msg"
            $log.Add([pscustomobject]@{ SheetRow = $r.SheetRow; Title = $r.MatchedTitle; Result = "failed: $msg"; RequestId = '' })
        }
    }
    if ($Apply) {
        $logPath = Join-Path (Split-Path $ReviewCsv) ("import-log-{0:yyyyMMdd-HHmmss}.csv" -f (Get-Date))
        $log | Export-Csv -LiteralPath $logPath -NoTypeInformation -Encoding UTF8
        Write-Host ""
        Write-Host "Log written to $logPath"
    }
}
