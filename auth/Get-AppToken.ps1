# =============================================================================
# Get-AppToken.ps1
# -----------------------------------------------------------------------------
# Acquires an application (client_credentials) access token for Microsoft Graph
# and copies it to the clipboard so it can be pasted into the SPFx demo
# launcher's "Manual bearer token" field.
#
# Usage:
#   1. Copy ./secrets.ps1.example to ./secrets.ps1 and fill in the secret.
#   2. From the repo root run:  ./auth/Get-AppToken.ps1
# =============================================================================

[CmdletBinding()]
param(
    [string] $SecretsFile = (Join-Path $PSScriptRoot 'secrets.ps1')
)

if (-not (Test-Path $SecretsFile)) {
    Write-Error "Secrets file not found: $SecretsFile`nCopy secrets.ps1.example to secrets.ps1 and fill in the values."
    exit 1
}

. $SecretsFile

if ([string]::IsNullOrWhiteSpace($ClientSecret) -or $ClientSecret -like '*paste-client-secret-here*') {
    Write-Error "ClientSecret is not set in $SecretsFile."
    exit 1
}

$body = @{
    client_id     = $ClientId
    scope         = 'https://graph.microsoft.com/.default'
    client_secret = $ClientSecret
    grant_type    = 'client_credentials'
}

$response = Invoke-RestMethod `
    -Method Post `
    -Uri  "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
    -Body $body

$response.access_token | Set-Clipboard

$expiresInMin = [math]::Round($response.expires_in / 60, 1)
Write-Host "Access token copied to clipboard (expires in ~$expiresInMin min)." -ForegroundColor Green
