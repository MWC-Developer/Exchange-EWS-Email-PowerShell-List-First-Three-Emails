# Ews-Powerhell-Appflow-ListFirstThreeEmails.ps1

<#
.SYNOPSIS
App-only OAuth (client credentials) + EWS Managed API: List the firsdt 3 emails in the Inbox.

.NOTES
- App-only requires the Entra app permission: full_access_as_app. [1](https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth)
- When using impersonation, you must set X-AnchorMailbox to the impersonated mailbox SMTP. [1](https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth)
- Download the EWS Managed API from here: https://www.microsoft.com/en-us/download/details.aspx?id=42022 or build from source it and put it into a folder.
- You may need to run this one time from PowerShell: Install-Module MSAL.PS -Scope CurrentUser
#>

$EwsDllPath = "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"
$EwsDllPath = "C:\Users\danba\source\repos\EwsEditor\EWSEditor\bin\Debug\Microsoft.Exchange.WebServices.dll"

$TenantId     = "dd55b8f6-4d6e-xxxx-8982-xxxxxxxxxxx"    # TODO: Set to the Tenant ID
$ClientId     = "729800b6-xxxx-441b-xxxx-28ace89daxxx"   # TODO: Set to the client ID
$ClientSecret = "f1x8Q~~NkVJYqlKXR................."     # TODO: Set to the client secret

$EwsUrl = "https://outlook.office365.com/EWS/Exchange.asmx"
 
$MailboxSmtp = "myuser@contoso.com"                          # TODO: Set to the smtp address of the mailbox to access.


# -------------------------
# Load EWS Managed API
# -------------------------
if (-not (Test-Path $EwsDllPath)) {
    throw "EWS DLL not found at: $EwsDllPath"
}
Add-Type -Path $EwsDllPath

# -------------------------
# Acquire app-only token (client credentials) using MSAL.PS + client secret
# For app-only EWS, scope is https://outlook.office365.com/.default [1](https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth)
# -------------------------
if (-not (Get-Module -ListAvailable -Name MSAL.PS)) {
    throw "MSAL.PS module not found. Install it: Install-Module MSAL.PS -Scope CurrentUser"
}

$secretSecure = ConvertTo-SecureString $ClientSecret -AsPlainText -Force

$token = Get-MsalToken `
    -TenantId $TenantId `
    -ClientId $ClientId `
    -ClientSecret $secretSecure `
    -Scopes "https://outlook.office365.com/.default"   # app-only EWS scope [1](https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth)

$accessToken = $token.AccessToken
if (-not $accessToken) { throw "Failed to obtain access token." }

# -------------------------
# Configure EWS service with OAuth token
# Set OAuthCredentials + impersonation for app-only [1](https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth)
# -------------------------
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService `
    ([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013_SP1)

$service.Url = [Uri]$EwsUrl
$service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.OAuthCredentials($accessToken)  # [1](https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth)

# App-only requires impersonation of the mailbox you want to access [1](https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth)
$service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId(
    [Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SmtpAddress,
    $MailboxSmtp
)

# When using impersonation, include X-AnchorMailbox set to that mailbox SMTP [1](https://learn.microsoft.com/en-us/exchange/client-developer/exchange-web-services/how-to-authenticate-an-ews-application-by-using-oauth)
if (-not $service.HttpHeaders) {
    $service.HttpHeaders = New-Object 'System.Collections.Generic.Dictionary[string,string]'
}
$service.HttpHeaders["X-AnchorMailbox"] = $MailboxSmtp

# -------------------------
# Bind Inbox and fetch first 3 items (Id + Subject)
# -------------------------
$inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind(
    $service,
    [Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox
)

$view = New-Object Microsoft.Exchange.WebServices.Data.ItemView(3)
$view.PropertySet = New-Object Microsoft.Exchange.WebServices.Data.PropertySet(
    [Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly,
    [Microsoft.Exchange.WebServices.Data.ItemSchema]::Subject,
    [Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived
)

# Explicit sort: newest first
$view.OrderBy.Add([Microsoft.Exchange.WebServices.Data.ItemSchema]::DateTimeReceived,
                  [Microsoft.Exchange.WebServices.Data.SortDirection]::Descending)

$results = $service.FindItems($inbox.Id, $view)

$results.Items | ForEach-Object {
    [pscustomobject]@{
        Id      = $_.Id.UniqueId
        Subject = $_.Subject
    }
} | Format-Table -AutoSize

 
