# GetZoomLinks.ps1
# Exports Zoom meeting registrants (with unique join links) to CSV.
# Compatible with Windows PowerShell 5.1 and PowerShell 7+.
# Downloaded from https://github.com/sstoilovABLE/zoom-registrant-export


# --- SETUP ---
$scriptDir   = Split-Path -Parent $MyInvocation.MyCommand.Path
$secretsFile = Join-Path $scriptDir "zoom-secrets.json"



# --- FUNCTIONS ---



function Test-SecretsFileExists {
    param([string]$FilePath)
    return Test-Path -Path $FilePath
}



function Test-CredentialsObject {
    param($Credentials)

    if ($null -eq $Credentials) { return $false }

    $missing = @()
    if (-not $Credentials.AccountId)    { $missing += "account_id" }
    if (-not $Credentials.ClientId)     { $missing += "client_id" }
    if (-not $Credentials.ClientSecret) { $missing += "client_secret" }

    if ($missing.Count -gt 0) {
        Write-Host ""
        Write-Host "zoom-secrets.json is incomplete. Missing field(s): $($missing -join ', ')" -ForegroundColor Yellow
        return $false
    }

    return $true
}



function Load-Credentials {
    param([string]$FilePath)

    try {
        $raw = Get-Content -Path $FilePath -Raw -ErrorAction Stop

        if ([string]::IsNullOrWhiteSpace($raw)) {
            Write-Host ""
            Write-Host "zoom-secrets.json is empty." -ForegroundColor Yellow
            return $null
        }

        $secrets = $raw | ConvertFrom-Json -ErrorAction Stop

        return @{
            AccountId    = $secrets.account_id
            ClientId     = $secrets.client_id
            ClientSecret = $secrets.client_secret
        }

    } catch [System.ArgumentException] {
        Write-Host ""
        Write-Host "zoom-secrets.json contains invalid JSON and could not be parsed." -ForegroundColor Yellow
        return $null

    } catch {
        Write-Host ""
        Write-Host "Could not read zoom-secrets.json: $($_.Exception.Message)" -ForegroundColor Yellow
        return $null
    }
}



function Save-Credentials {
    param(
        [string]$FilePath,
        [string]$AccountId,
        [string]$ClientId,
        [string]$ClientSecret
    )

    $secretsObject = @{
        account_id    = $AccountId
        client_id     = $ClientId
        client_secret = $ClientSecret
    }

    $secretsObject | ConvertTo-Json | Set-Content -Path $FilePath -Force
    Write-Host "Credentials saved to $(Split-Path -Leaf $FilePath)" -ForegroundColor Green
}



function Prompt-ForCredentials {
    Write-Host ""
    Write-Host "=== Zoom Credentials Setup ===" -ForegroundColor Cyan
    Write-Host "Please enter your Zoom API credentials:"
    Write-Host ""

    $accountId    = Read-Host "Account ID"
    $clientId     = Read-Host "Client ID"
    $clientSecret = Read-Host "Client Secret" -AsSecureString

    # Compatible SecureString to plain text for both PS 5.1 and PS 7+
    $bstr              = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($clientSecret)
    $clientSecretPlain = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)

    return @{
        AccountId    = $accountId
        ClientId     = $clientId
        ClientSecret = $clientSecretPlain
    }
}



function Prompt-ForMeetingId {
    $meetingId = Read-Host "Enter Meeting ID"
    return $meetingId
}



function Update-GitIgnore {
    param([string]$ScriptDir)

    $gitIgnoreFile   = Join-Path $ScriptDir ".gitignore"
    $secretsFileName = "zoom-secrets.json"
    $csvPattern      = "*.csv"

    if (Test-Path -Path $gitIgnoreFile) {
        $gitIgnoreContent = Get-Content -Path $gitIgnoreFile -Raw
        if ($gitIgnoreContent -notmatch [regex]::Escape($secretsFileName)) {
            Add-Content -Path $gitIgnoreFile -Value "$secretsFileName"
            Write-Host "Added $secretsFileName to .gitignore" -ForegroundColor Green
        }
        if ($gitIgnoreContent -notmatch [regex]::Escape($csvPattern)) {
            Add-Content -Path $gitIgnoreFile -Value "$csvPattern"
            Write-Host "Added $csvPattern to .gitignore" -ForegroundColor Green
        }
    } else {
        $response = Read-Host "Create .gitignore and add zoom-secrets.json? (Y/n)"
        if ($response -ne 'n') {
            Set-Content -Path $gitIgnoreFile -Value $secretsFileName
            Add-Content -Path $gitIgnoreFile -Value $csvPattern
            Write-Host "Created .gitignore with $secretsFileName and $csvPattern" -ForegroundColor Green
        }
    }
}



function Reset-Credentials {
    param([string]$FilePath)

    $credentials = Prompt-ForCredentials
    Save-Credentials `
        -FilePath     $FilePath `
        -AccountId    $credentials.AccountId `
        -ClientId     $credentials.ClientId `
        -ClientSecret $credentials.ClientSecret

    return $credentials
}



function Get-ZoomAccessToken {
    param(
        [string]$AccountId,
        [string]$ClientId,
        [string]$ClientSecret,
        [string]$SecretsFile
    )

    $encodedCreds = [Convert]::ToBase64String(
        [Text.Encoding]::ASCII.GetBytes($ClientId + ":" + $ClientSecret)
    )

    $tokenUri = "https://zoom.us/oauth/token?grant_type=account_credentials" + "&account_id=" + $AccountId

    try {
        $response = Invoke-RestMethod `
            -Uri         $tokenUri `
            -Method      POST `
            -Headers     @{ Authorization = "Basic $encodedCreds" } `
            -ContentType "application/x-www-form-urlencoded"

        return $response.access_token

    } catch {
        $statusCode = $_.Exception.Response.StatusCode.value__

        Write-Host ""
        Write-Host "Authentication failed (HTTP $statusCode)." -ForegroundColor Red

        if ($statusCode -eq 400 -or $statusCode -eq 401) {
            Write-Host "One or more credentials in zoom-secrets.json are incorrect." -ForegroundColor Yellow

            $retry = Read-Host "Would you like to re-enter your credentials and try again? (Y/n)"
            if ($retry -ne 'n') {
                $newCreds = Reset-Credentials -FilePath $SecretsFile
                return Get-ZoomAccessToken `
                    -AccountId    $newCreds.AccountId `
                    -ClientId     $newCreds.ClientId `
                    -ClientSecret $newCreds.ClientSecret `
                    -SecretsFile  $SecretsFile
            }
        } else {
            Write-Host "Unexpected error: $($_.Exception.Message)" -ForegroundColor Red
        }

        return $null
    }
}



function Get-AllRegistrants {
    param(
        [string]$MeetingId,
        [string]$AccessToken
    )

    $allRegistrants = @()
    $nextPageToken  = ""
    $headers        = @{ Authorization = "Bearer $AccessToken" }

    try {
        do {
            $uri = "https://api.zoom.us/v2/meetings/" + $MeetingId + "/registrants?page_size=300"
            if ($nextPageToken) {
                $uri = $uri + "&next_page_token=" + $nextPageToken
            }

            $response      = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers
            $allRegistrants += $response.registrants
            $nextPageToken  = $response.next_page_token

        } while ($nextPageToken)

    } catch {
        $statusCode = $_.Exception.Response.StatusCode.value__

        # Try to extract Zoom's error body for a more specific message
        $errorBody = $null
        try {
            $stream    = $_.Exception.Response.GetResponseStream()
            $reader    = New-Object System.IO.StreamReader($stream)
            $errorBody = $reader.ReadToEnd() | ConvertFrom-Json
        } catch {}

        Write-Host ""
        Write-Host "Failed to fetch registrants (HTTP $statusCode)." -ForegroundColor Red

        if ($errorBody -and $errorBody.code -eq 124) {
            Write-Host "Invalid access token. Please re-run the script." -ForegroundColor Yellow
        } elseif ($statusCode -eq 404) {
            Write-Host "Meeting ID '$MeetingId' was not found. Please check the ID and try again." -ForegroundColor Yellow
        } elseif ($statusCode -eq 400) {
            Write-Host "Bad request. The Meeting ID may be invalid or the meeting may not have registration enabled." -ForegroundColor Yellow
        } elseif ($errorBody -and $errorBody.message) {
            Write-Host "Zoom error: $($errorBody.message)" -ForegroundColor Yellow
        } else {
            Write-Host "Unexpected error: $($_.Exception.Message)" -ForegroundColor Red
        }

        return $null
    }

    return $allRegistrants
}



# --- MAIN ---



# Load credentials if secrets file exists, validate, and prompt to fix if needed
if (Test-SecretsFileExists -FilePath $secretsFile) {
    $credentials = Load-Credentials -FilePath $secretsFile

    if (-not (Test-CredentialsObject -Credentials $credentials)) {
        Write-Host "zoom-secrets.json is invalid or incomplete." -ForegroundColor Yellow
        $recreate = Read-Host "Would you like to re-enter your credentials? (Y/n)"
        if ($recreate -eq 'n') {
            Write-Host "Exiting." -ForegroundColor Red
            exit 1
        }
        $credentials = Reset-Credentials -FilePath $secretsFile
    } else {
        Write-Host "Using saved credentials for account $($credentials.AccountId)" -ForegroundColor Green
    }

} else {
    $credentials = Prompt-ForCredentials
    Save-Credentials `
        -FilePath     $secretsFile `
        -AccountId    $credentials.AccountId `
        -ClientId     $credentials.ClientId `
        -ClientSecret $credentials.ClientSecret
    Update-GitIgnore -ScriptDir $scriptDir
}


# Prompt for Meeting ID and set output filename
$MEETING_ID  = Prompt-ForMeetingId
$timestamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$OUTPUT_FILE = Join-Path $scriptDir ($timestamp + "_registrants.csv")


# Get access token
Write-Host "Authenticating..." -ForegroundColor Cyan
$accessToken = Get-ZoomAccessToken `
    -AccountId    $credentials.AccountId `
    -ClientId     $credentials.ClientId `
    -ClientSecret $credentials.ClientSecret `
    -SecretsFile  $secretsFile

if (-not $accessToken) {
    Write-Host "Exiting -- no valid access token obtained." -ForegroundColor Red
    exit 1
}


# Fetch all registrants
Write-Host "Fetching registrants..." -ForegroundColor Cyan
$allRegistrants = Get-AllRegistrants -MeetingId $MEETING_ID -AccessToken $accessToken

if ($null -eq $allRegistrants) {
    Write-Host "Exiting -- no registrants retrieved." -ForegroundColor Red
    exit 1
}

if ($allRegistrants.Count -eq 0) {
    Write-Host "No registrants found for Meeting ID $MEETING_ID." -ForegroundColor Yellow
    exit 0
}


# Export to CSV
$allRegistrants |
    Select-Object email, first_name, last_name, join_url |
    Export-Csv -Path $OUTPUT_FILE -NoTypeInformation -Encoding UTF8

Write-Host "Done! $($allRegistrants.Count) registrants saved to $OUTPUT_FILE" -ForegroundColor Green