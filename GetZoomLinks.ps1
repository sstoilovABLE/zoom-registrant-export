# zoom_export.ps1
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


function Load-Credentials {
    param([string]$FilePath)

    if (Test-Path -Path $FilePath) {
        $secrets = Get-Content -Path $FilePath -Raw | ConvertFrom-Json
        return @{
            AccountId    = $secrets.account_id
            ClientId     = $secrets.client_id
            ClientSecret = $secrets.client_secret
        }
    }
    return $null
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
    Write-Host "=== First Run Setup ===" -ForegroundColor Cyan
    Write-Host "Please enter your Zoom credentials:"
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
            $gitIgnoreContent = @"
$secretsFileName
$csvPattern
"@
            Set-Content -Path $gitIgnoreFile -Value $gitIgnoreContent
            Write-Host "Created .gitignore with $secretsFileName and $csvPattern" -ForegroundColor Green
        }
    }
}


function Get-ZoomAccessToken {
    param(
        [string]$AccountId,
        [string]$ClientId,
        [string]$ClientSecret
    )

    $encodedCreds = [Convert]::ToBase64String(
        [Text.Encoding]::ASCII.GetBytes($ClientId + ":" + $ClientSecret)
    )

    $tokenUri = "https://zoom.us/oauth/token?grant_type=account_credentials" + "&account_id=" + $AccountId

    $response = Invoke-RestMethod `
        -Uri         $tokenUri `
        -Method      POST `
        -Headers     @{ Authorization = "Basic $encodedCreds" } `
        -ContentType "application/x-www-form-urlencoded"

    return $response.access_token
}


function Get-AllRegistrants {
    param(
        [string]$MeetingId,
        [string]$AccessToken
    )

    $allRegistrants = @()
    $nextPageToken  = ""
    $headers        = @{ Authorization = "Bearer $AccessToken" }

    do {
        $uri = "https://api.zoom.us/v2/meetings/" + $MeetingId + "/registrants?page_size=300"
        if ($nextPageToken) {
            $uri = $uri + "&next_page_token=" + $nextPageToken
        }

        $response      = Invoke-RestMethod -Uri $uri -Method GET -Headers $headers
        $allRegistrants += $response.registrants
        $nextPageToken  = $response.next_page_token

    } while ($nextPageToken)

    return $allRegistrants
}


# --- MAIN ---


# Load or prompt for credentials
if (Test-SecretsFileExists -FilePath $secretsFile) {
    $credentials = Load-Credentials -FilePath $secretsFile
    Write-Host "Using saved credentials for account $($credentials.AccountId)" -ForegroundColor Green
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
    -ClientSecret $credentials.ClientSecret

# Fetch all registrants
Write-Host "Fetching registrants..." -ForegroundColor Cyan
$allRegistrants = Get-AllRegistrants -MeetingId $MEETING_ID -AccessToken $accessToken

# Export to CSV
$allRegistrants |
    Select-Object email, first_name, last_name, join_url |
    Export-Csv -Path $OUTPUT_FILE -NoTypeInformation -Encoding UTF8

Write-Host "Done! $($allRegistrants.Count) registrants saved to $OUTPUT_FILE" -ForegroundColor Green