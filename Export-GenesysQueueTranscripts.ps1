<#
.SYNOPSIS
    Exports all Genesys Cloud transcripts for a queue (last calendar month) to CSV.

.DESCRIPTION
    Region   : Australia (ap-southeast-2)  ->  login.mypurecloud.com.au / api.mypurecloud.com.au
    Auth     : OAuth2 Client Credentials grant (Client ID + Secret embedded below)
    Period   : Previous calendar month, local time, converted to UTC for the API
    Platform : Windows PowerShell 5.1 running in Constrained Language Mode (CLM)

    CLM rules honoured by this script:
      - No .NET static method calls  ([Convert]::, [Text.Encoding]::, [uri]::, [Guid]:: ...)
      - No ::new() constructors
      - No [pscustomobject] casts      (New-Object PSObject -Property + Select-Object for column order)
      - Base64 for the Basic auth header is implemented in pure PowerShell with -shl/-shr/-band/-bor
      - No manual URL encoding needed  (IDs are GUIDs, filtering is done server-side via JSON body)
      - HTTP status read via [int]$_.Exception.Response.StatusCode with string-match fallback

    API flow:
      1. POST /oauth/token                                             -> access_token
      2. POST /api/v2/analytics/conversations/details/query            -> conversations for the queue
         (interval is split into 7-day windows; 100 rows per page; all pages fetched)
      3. GET  /api/v2/speechandtextanalytics/conversations/{convId}/communications/{commId}/transcripturl
         -> pre-signed URL to the transcript JSON
      4. GET  <pre-signed url>                                         -> transcript phrases
      5. Export-Csv (phrase-level file + per-conversation summary file)

    OAuth client (role) permissions required:
      - analytics:conversationDetail:view
      - speechAndTextAnalytics:data:view
      - recording:recording:view   (needed by the transcript URL endpoint)

.NOTES
    Edit the CONFIG block below (ClientId, ClientSecret, QueueId, OutputFolder) then run:
        powershell.exe -ExecutionPolicy Bypass -File .\Export-GenesysQueueTranscripts.ps1
#>

# =====================================================================================
#  CONFIG  -  edit these values
# =====================================================================================
$ClientId     = 'PASTE-YOUR-CLIENT-ID-HERE'
$ClientSecret = 'PASTE-YOUR-CLIENT-SECRET-HERE'
$QueueId      = 'PASTE-YOUR-QUEUE-ID-HERE'          # GUID of the queue
$OutputFolder = 'C:\Temp\GenesysTranscripts'        # created if missing

# Australia (Sydney) region endpoints
$LoginBase = 'https://login.mypurecloud.com.au'
$ApiBase   = 'https://api.mypurecloud.com.au'

# Media types that can have a transcript
$TranscriptMediaTypes = @('voice', 'chat', 'message', 'email')

# Polite delay between transcript downloads (seconds). Raise if you hit many 429s.
$ThrottleSeconds = 0.2

# =====================================================================================
#  HELPERS
# =====================================================================================

function ConvertTo-Base64Clm {
    <#
        Pure-PowerShell Base64 encoder (CLM-safe).
        Client IDs / secrets are ASCII, so [int][char] == byte value.
    #>
    param([string]$Text)

    $alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/'

    $bytes = @()
    foreach ($c in $Text.ToCharArray()) {
        $bytes += [int][char]$c
    }

    $out = ''
    $i   = 0
    $len = $bytes.Count

    while ($i -lt $len) {
        $b0 = $bytes[$i]
        $b1 = 0
        $b2 = 0
        $has1 = $false
        $has2 = $false

        if (($i + 1) -lt $len) { $b1 = $bytes[$i + 1]; $has1 = $true }
        if (($i + 2) -lt $len) { $b2 = $bytes[$i + 2]; $has2 = $true }

        $triple = ($b0 -shl 16) -bor ($b1 -shl 8) -bor $b2

        $out += $alphabet.Substring((($triple -shr 18) -band 63), 1)
        $out += $alphabet.Substring((($triple -shr 12) -band 63), 1)

        if ($has1) { $out += $alphabet.Substring((($triple -shr 6) -band 63), 1) } else { $out += '=' }
        if ($has2) { $out += $alphabet.Substring(($triple -band 63), 1) }          else { $out += '=' }

        $i += 3
    }

    return $out
}

function Get-HttpStatusFromError {
    <#
        Returns the HTTP status code (int) from an Invoke-RestMethod error record,
        or 0 when it cannot be determined. CLM-safe (no deep .NET chains).
    #>
    param($ErrorRecord)

    $status = 0
    try {
        $status = [int]$ErrorRecord.Exception.Response.StatusCode
    } catch {
        $status = 0
    }

    if ($status -eq 0) {
        $msg = [string]$ErrorRecord.Exception.Message
        if     ($msg -like '*(401)*' -or $msg -like '*Unauthorized*')      { $status = 401 }
        elseif ($msg -like '*(403)*' -or $msg -like '*Forbidden*')         { $status = 403 }
        elseif ($msg -like '*(404)*' -or $msg -like '*Not Found*')         { $status = 404 }
        elseif ($msg -like '*(429)*' -or $msg -like '*Too Many Requests*') { $status = 429 }
        elseif ($msg -like '*(400)*' -or $msg -like '*Bad Request*')       { $status = 400 }
        elseif ($msg -like '*(5??)*')                                      { $status = 500 }
    }

    return $status
}

function Get-GenesysToken {
    <#
        OAuth2 client-credentials grant. Returns the access token string or throws.
    #>
    param([string]$Id, [string]$Secret)

    $basic = ConvertTo-Base64Clm -Text ($Id + ':' + $Secret)

    $headers = @{
        'Authorization' = 'Basic ' + $basic
        'Content-Type'  = 'application/x-www-form-urlencoded'
    }

    # Invoke-RestMethod form-encodes a hashtable body for us - no manual URL encoding needed.
    $body = @{ 'grant_type' = 'client_credentials' }

    $resp = $null
    try {
        $resp = Invoke-RestMethod -Method Post -Uri ($LoginBase + '/oauth/token') -Headers $headers -Body $body -ErrorAction Stop
    } catch {
        $st = Get-HttpStatusFromError -ErrorRecord $_
        throw ('Token request failed (HTTP ' + $st + '): ' + $_.Exception.Message)
    }

    if ($null -eq $resp -or $null -eq $resp.access_token -or [string]$resp.access_token -eq '') {
        throw 'Token request returned no access_token. Check Client ID / Secret and region.'
    }

    Write-Host ('Token acquired (expires in ' + [string]$resp.expires_in + ' s).')
    return [string]$resp.access_token
}

function Invoke-GenesysApi {
    <#
        Wrapper around Invoke-RestMethod with:
          - Bearer auth
          - 429 handling (honours Retry-After when readable, else 60 s)
          - 5xx retry with back-off
          - returns $null on 404 (caller decides what that means)
        $script:AccessToken must be set.
    #>
    param(
        [string]$Method,
        [string]$Uri,
        $Body = $null,
        [int]$MaxAttempts = 5
    )

    $attempt = 0
    while ($true) {
        $attempt++

        $headers = @{
            'Authorization' = 'Bearer ' + $script:AccessToken
            'Content-Type'  = 'application/json'
        }

        try {
            if ($null -ne $Body) {
                $json = $Body | ConvertTo-Json -Depth 20 -Compress
                return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -Body $json -ErrorAction Stop
            } else {
                return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -ErrorAction Stop
            }
        } catch {
            $st = Get-HttpStatusFromError -ErrorRecord $_

            if ($st -eq 404) {
                return $null
            }

            if ($st -eq 429) {
                $wait = 60
                try {
                    $ra = $_.Exception.Response.Headers['Retry-After']
                    if ($null -ne $ra -and [string]$ra -ne '') { $wait = [int]$ra + 1 }
                } catch { $wait = 60 }
                Write-Warning ('429 rate limited - sleeping ' + $wait + ' s (' + $Uri + ')')
                Start-Sleep -Seconds $wait
                if ($attempt -lt $MaxAttempts) { continue }
            }

            if ($st -eq 401 -and $attempt -lt $MaxAttempts) {
                Write-Warning '401 received - refreshing token and retrying.'
                $script:AccessToken = Get-GenesysToken -Id $ClientId -Secret $ClientSecret
                continue
            }

            if ($st -ge 500 -and $attempt -lt $MaxAttempts) {
                $wait = 5 * $attempt
                Write-Warning ('HTTP ' + $st + ' - retrying in ' + $wait + ' s (' + $Uri + ')')
                Start-Sleep -Seconds $wait
                continue
            }

            throw ('API call failed (HTTP ' + $st + ') ' + $Method + ' ' + $Uri + ' : ' + $_.Exception.Message)
        }
    }
}

function Format-GenesysUtc {
    param([datetime]$Value)
    # ISO-8601 UTC string expected by the analytics interval, e.g. 2026-09-01T00:00:00.000Z
    return $Value.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ss.fff') + 'Z'
}

# =====================================================================================
#  MAIN
# =====================================================================================

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory | Out-Null
}

# ---- Date range: previous calendar month (local), converted to UTC ------------------
$firstOfThisMonth = Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0 -Millisecond 0
$monthStart       = $firstOfThisMonth.AddMonths(-1)
$monthEnd         = $firstOfThisMonth
$monthLabel       = $monthStart.ToString('yyyy-MM')

Write-Host ('Period (local): ' + $monthStart.ToString('yyyy-MM-dd') + ' to ' + $monthEnd.ToString('yyyy-MM-dd') + ' (exclusive)')
Write-Host ('Period (UTC)  : ' + (Format-GenesysUtc -Value $monthStart) + ' / ' + (Format-GenesysUtc -Value $monthEnd))

# ---- Auth ---------------------------------------------------------------------------
$script:AccessToken = Get-GenesysToken -Id $ClientId -Secret $ClientSecret

# ---- Step 1: find all conversations that touched the queue --------------------------
$conversations = @()
$windowStart   = $monthStart

while ($windowStart -lt $monthEnd) {

    $windowEnd = $windowStart.AddDays(7)
    if ($windowEnd -gt $monthEnd) { $windowEnd = $monthEnd }

    $interval = (Format-GenesysUtc -Value $windowStart) + '/' + (Format-GenesysUtc -Value $windowEnd)
    Write-Host ('Querying conversations for interval ' + $interval)

    $pageNumber = 1
    while ($true) {

        $query = @{
            'interval'       = $interval
            'order'          = 'asc'
            'orderBy'        = 'conversationStart'
            'paging'         = @{ 'pageSize' = 100; 'pageNumber' = $pageNumber }
            'segmentFilters' = @(
                @{
                    'type'       = 'or'
                    'predicates' = @(
                        @{ 'type' = 'dimension'; 'dimension' = 'queueId'; 'operator' = 'matches'; 'value' = $QueueId }
                    )
                }
            )
        }

        $result = Invoke-GenesysApi -Method 'Post' -Uri ($ApiBase + '/api/v2/analytics/conversations/details/query') -Body $query

        if ($null -eq $result) {
            throw 'Analytics query returned 404 - check the API base URL / region.'
        }

        $page = @()
        if ($null -ne $result.conversations) { $page = @($result.conversations) }

        Write-Host ('  page ' + $pageNumber + ': ' + $page.Count + ' conversation(s)')

        foreach ($c in $page) { $conversations += $c }

        if ($page.Count -lt 100) { break }
        $pageNumber++
    }

    $windowStart = $windowEnd
}

# De-duplicate (a conversation spanning a window boundary can appear twice)
$seenConv = @{}
$uniqueConversations = @()
foreach ($c in $conversations) {
    $cid = [string]$c.conversationId
    if (-not $seenConv.ContainsKey($cid)) {
        $seenConv[$cid] = $true
        $uniqueConversations += $c
    }
}

Write-Host ('Total unique conversations for queue ' + $QueueId + ': ' + $uniqueConversations.Count)

if ($uniqueConversations.Count -eq 0) {
    Write-Warning 'No conversations found for this queue in the period. Nothing to export.'
    return
}

# ---- Step 2: download transcripts ---------------------------------------------------
$phraseRows  = @()
$summaryRows = @()
$seenTranscript = @{}
$convIndex = 0

foreach ($conv in $uniqueConversations) {

    $convIndex++
    $conversationId = [string]$conv.conversationId
    $convStart      = [string]$conv.conversationStart
    $convEnd        = [string]$conv.conversationEnd
    $direction      = [string]$conv.originatingDirection

    Write-Host ('[' + $convIndex + '/' + $uniqueConversations.Count + '] ' + $conversationId)

    # Collect candidate communication (session) IDs.
    # Transcripts hang off the customer/external communication; agent sessions are a fallback.
    $primarySessions  = @()
    $fallbackSessions = @()

    foreach ($p in @($conv.participants)) {
        $purpose = [string]$p.purpose
        foreach ($s in @($p.sessions)) {
            $mt = [string]$s.mediaType
            if ($TranscriptMediaTypes -notcontains $mt) { continue }
            if ($null -eq $s.sessionId -or [string]$s.sessionId -eq '') { continue }

            $entry = New-Object PSObject -Property @{
                'SessionId' = [string]$s.sessionId
                'MediaType' = $mt
                'Purpose'   = $purpose
            }

            if ($purpose -eq 'customer' -or $purpose -eq 'external') {
                $primarySessions += $entry
            } elseif ($purpose -eq 'agent' -or $purpose -eq 'user') {
                $fallbackSessions += $entry
            }
        }
    }

    $foundAny = $false

    foreach ($sessionList in @($primarySessions, $fallbackSessions)) {

        if ($foundAny) { break }

        foreach ($sess in @($sessionList)) {

            $commId = $sess.SessionId
            $urlApi = $ApiBase + '/api/v2/speechandtextanalytics/conversations/' + $conversationId + '/communications/' + $commId + '/transcripturl'

            $urlResp = $null
            try {
                $urlResp = Invoke-GenesysApi -Method 'Get' -Uri $urlApi
            } catch {
                Write-Warning ('  transcripturl failed for communication ' + $commId + ': ' + $_.Exception.Message)
                continue
            }

            if ($null -eq $urlResp -or $null -eq $urlResp.url -or [string]$urlResp.url -eq '') {
                continue   # 404 = no transcript for this communication
            }

            # Pre-signed URL: no Authorization header must be sent.
            $transcriptDoc = $null
            try {
                $transcriptDoc = Invoke-RestMethod -Method Get -Uri ([string]$urlResp.url) -ErrorAction Stop
            } catch {
                Write-Warning ('  transcript download failed for communication ' + $commId + ': ' + $_.Exception.Message)
                continue
            }

            if ($null -eq $transcriptDoc -or $null -eq $transcriptDoc.transcripts) {
                continue
            }

            foreach ($t in @($transcriptDoc.transcripts)) {

                $transcriptId = [string]$t.transcriptId
                if ($transcriptId -ne '' -and $seenTranscript.ContainsKey($transcriptId)) { continue }
                if ($transcriptId -ne '') { $seenTranscript[$transcriptId] = $true }

                $phrases   = @($t.phrases)
                $fullLines = @()
                $idx       = 0

                foreach ($ph in $phrases) {
                    $idx++
                    $speaker = [string]$ph.participantPurpose
                    $text    = [string]$ph.text
                    $durMs   = ''
                    if ($null -ne $ph.duration -and $null -ne $ph.duration.milliseconds) { $durMs = [string]$ph.duration.milliseconds }

                    $phraseRows += New-Object PSObject -Property @{
                        'ConversationId'    = $conversationId
                        'ConversationStart' = $convStart
                        'ConversationEnd'   = $convEnd
                        'Direction'         = $direction
                        'QueueId'           = $QueueId
                        'MediaType'         = [string]$transcriptDoc.mediaType
                        'CommunicationId'   = $commId
                        'TranscriptId'      = $transcriptId
                        'PhraseIndex'       = $idx
                        'Speaker'           = $speaker
                        'StartTimeMs'       = [string]$ph.startTimeMs
                        'DurationMs'        = $durMs
                        'Confidence'        = [string]$ph.confidence
                        'Text'              = $text
                    }

                    $fullLines += ($speaker + ': ' + $text)
                }

                $summaryRows += New-Object PSObject -Property @{
                    'ConversationId'    = $conversationId
                    'ConversationStart' = $convStart
                    'ConversationEnd'   = $convEnd
                    'Direction'         = $direction
                    'QueueId'           = $QueueId
                    'MediaType'         = [string]$transcriptDoc.mediaType
                    'CommunicationId'   = $commId
                    'TranscriptId'      = $transcriptId
                    'Status'            = 'OK'
                    'PhraseCount'       = $phrases.Count
                    'FullTranscript'    = ($fullLines -join "`r`n")
                }

                $foundAny = $true
            }

            Start-Sleep -Milliseconds ([int]($ThrottleSeconds * 1000))
        }
    }

    if (-not $foundAny) {
        $summaryRows += New-Object PSObject -Property @{
            'ConversationId'    = $conversationId
            'ConversationStart' = $convStart
            'ConversationEnd'   = $convEnd
            'Direction'         = $direction
            'QueueId'           = $QueueId
            'MediaType'         = ''
            'CommunicationId'   = ''
            'TranscriptId'      = ''
            'Status'            = 'NO_TRANSCRIPT'
            'PhraseCount'       = 0
            'FullTranscript'    = ''
        }
    }
}

# ---- Step 3: export -----------------------------------------------------------------
$phraseFile  = Join-Path -Path $OutputFolder -ChildPath ('Transcripts_Phrases_' + $monthLabel + '_' + $QueueId + '.csv')
$summaryFile = Join-Path -Path $OutputFolder -ChildPath ('Transcripts_Summary_' + $monthLabel + '_' + $QueueId + '.csv')

# -Property hashtables do not preserve key order, so fix the column order explicitly.
$phraseRows |
    Select-Object ConversationId, ConversationStart, ConversationEnd, Direction, QueueId, MediaType,
                  CommunicationId, TranscriptId, PhraseIndex, Speaker, StartTimeMs, DurationMs, Confidence, Text |
    Export-Csv -Path $phraseFile -NoTypeInformation -Encoding UTF8

$summaryRows |
    Select-Object ConversationId, ConversationStart, ConversationEnd, Direction, QueueId, MediaType,
                  CommunicationId, TranscriptId, Status, PhraseCount, FullTranscript |
    Export-Csv -Path $summaryFile -NoTypeInformation -Encoding UTF8

$okCount = @($summaryRows | Where-Object { $_.Status -eq 'OK' }).Count
$noCount = @($summaryRows | Where-Object { $_.Status -eq 'NO_TRANSCRIPT' }).Count

Write-Host ''
Write-Host ('Done.')
Write-Host ('  Conversations queried : ' + $uniqueConversations.Count)
Write-Host ('  Transcripts exported  : ' + $okCount)
Write-Host ('  No transcript found   : ' + $noCount)
Write-Host ('  Phrase rows           : ' + $phraseRows.Count)
Write-Host ('  Phrase-level CSV      : ' + $phraseFile)
Write-Host ('  Summary CSV           : ' + $summaryFile)
