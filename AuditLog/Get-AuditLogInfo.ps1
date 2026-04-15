param(


    [datetime] $StartUtc = (Get-Date).ToUniversalTime().AddDays(-1),
    [datetime] $EndUtc = (Get-Date).ToUniversalTime(),

    [string] $OutCsv = ".\sp_list_audit.csv"
)

$configPath = ".\config.json"
$config = Get-Content -Raw -Path $configPath | ConvertFrom-Json
$TenantId = $config.TenantId
$ClientId = $config.AppId

$clientSecret = $config.ClientSecret
$SiteUrl = "https://caje77sharepoint.sharepoint.com/sites/M365Updates"
$ListServerRelUrl = "/sites/M365Updates/Lists/AzureAASiteRequest"

$contentType = "Audit.SharePoint"

# 1) Token for manage.office.com
$tokenResp = Invoke-RestMethod -Method POST `
    -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" `
    -ContentType "application/x-www-form-urlencoded" `
    -Body @{
    client_id     = $ClientId
    client_secret = $ClientSecret
    grant_type    = "client_credentials"
    scope         = "https://manage.office.com/.default"
}

$headers = @{ Authorization = "Bearer $($tokenResp.access_token)" }

# 2) Ensure subscription is started (safe to call repeatedly)
Invoke-RestMethod -Method POST `
    -Uri "https://manage.office.com/api/v1.0/$TenantId/activity/feed/subscriptions/start?contentType=$contentType" `
    -Headers $headers | Out-Null

# 3) List available content blobs for the time range
$st = $StartUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")
$et = $EndUtc.ToString("yyyy-MM-ddTHH:mm:ssZ")

$content = Invoke-RestMethod -Method GET `
    -Uri "https://manage.office.com/api/v1.0/$TenantId/activity/feed/subscriptions/content?contentType=$contentType&startTime=$st&endTime=$et" `
    -Headers $headers

# 4) Download blobs -> events, then filter to your list
$events = New-Object System.Collections.Generic.List[object]

foreach ($c in $content) {
    $blobEvents = Invoke-RestMethod -Method GET -Uri $c.contentUri -Headers $headers
    foreach ($e in $blobEvents) { $events.Add($e) }
}

# Common SharePoint “list item” operations are represented in Operation.
# (Exact operation names vary; see Purview “Audit log activities” for SharePoint operation names.) :contentReference[oaicite:1]{index=1}
$listItemOps = @(
    "ListItemCreated",
    "ListItemUpdated",
    "ListItemDeleted",
    "ListItemViewed"
)

# Filter logic:
# - Match site (SiteUrl when present)
# - Match your list path via ObjectId (often a URL) or other URL-like fields
$filtered = $events | Where-Object {
    ($_.SiteUrl -and ($_.SiteUrl.TrimEnd('/') -eq $SiteUrl.TrimEnd('/'))) -or
    ($_.ObjectId -and ($_.ObjectId -like "*$($SiteUrl.TrimEnd('/'))*"))
} | Where-Object {
    ($_.Operation -in $listItemOps) -or
    ($_.Operation -like "*ListItem*")
} | Where-Object {
    ($_.ObjectId -and ($_.ObjectId -like "*$ListServerRelUrl*")) -or
    ($_.SourceRelativeUrl -and ($_.SourceRelativeUrl -like "*$ListServerRelUrl*")) -or
    ($_.ItemUrl -and ($_.ItemUrl -like "*$ListServerRelUrl*"))
}

# Export a useful subset + keep raw JSON if you need it
$filtered |
Select-Object CreationTime, UserId, Operation, SiteUrl, ObjectId, ClientIP, UserAgent,
@{n = "RawEvent"; e = { $_ | ConvertTo-Json -Depth 30 -Compress } } |
Export-Csv -NoTypeInformation -Encoding UTF8 -Path $OutCsv


$filtered= $events | Where-Object { $_.Operation -like "*ListItem*" -and $_.ListUrl -like "*$ListServerRelUrl*" } 
#| Export-Csv -NoTypeInformation -Encoding UTF8 -Path ".\sp_list_audit_listitem.csv"

"Downloaded events: $($events.Count); Matched list events: $($filtered.Count); Exported: $OutCsv"
