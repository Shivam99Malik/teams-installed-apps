$tenantId = "a" # check these ids from admin azure center
 
$clientId = "b"
 
$clientSecret = "c"

$graphApiBaseUrl = "https://graph.microsoft.com/v1.0"
 
 $tokenUrl = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
# === Function to Get Auth Token ===
function Get-AccessToken {
    $token = Invoke-RestMethod -Method Post -Uri $tokenUrl -ContentType "application/x-www-form-urlencoded" -Body @{
        client_id     = $clientId
        scope         = "https://graph.microsoft.com/.default"
        client_secret = $clientSecret
        grant_type    = "client_credentials"
    }
    return $token.access_token
}
 
# === Global token + headers ===
$global:accessToken = Get-AccessToken
if (-not $global:accessToken) {
    Write-Host "⚠️ Error: Failed to obtain access token."
    return
}
$global:headers = @{ Authorization = "Bearer $global:accessToken" }
 
# === Function to make Graph requests and auto-refresh on 401 ===
function Invoke-GraphRequest {
    param (
        [string]$method = "GET",
        [string]$url
    )
    if (-not $url) {
        Write-Host "⚠️ Error: URL is null or empty."
        return $null
    }
    try {
        return Invoke-RestMethod -Method $method -Uri $url -Headers $global:headers
    } catch {
        if ($_.Exception.Response.StatusCode.Value__ -eq 401) {
            Write-Host "🔐 Token expired, reauthenticating..."
            $global:accessToken = Get-AccessToken
            if (-not $global:accessToken) {
                Write-Host "⚠️ Error: Failed to obtain access token."
                return $null
            }
            $global:headers = @{ Authorization = "Bearer $global:accessToken" }
            # Retry the request once
            return Invoke-RestMethod -Method $method -Uri $url -Headers $global:headers
        } else {
            throw $_
        }
    }
}
 
# === Function to make Graph requests with pagination ===
function Invoke-GraphRequestWithPagination {
    param (
        [string]$url
    )
    if (-not $url) {
        Write-Host "⚠️ Error: URL is null or empty."
        return $null
    }
    $results = @()
    do {
        $response = Invoke-GraphRequest -Method Get -Url $url
        if ($response -ne $null) {
            $results += $response.value # Append current page data
            $url = $response.'@odata.nextLink' # Get next page link
        } else {
            Write-Host "⚠️ Error occurred while fetching data. Exiting pagination."
            break
        }
    } while ($url) # Continue if there is a next page
    return $results
}
 
# === Fetch users from Azure AD ===
$users = @()
$url = "$graphApiBaseUrl/users"
if (-not $url) {
    Write-Host "⚠️ Error: URL is null or empty."
    return
}
 
do {
    $response = Invoke-GraphRequest -Method Get -Url $url
    if ($response -ne $null) {
        $users += $response.value
        $url = $response.'@odata.nextLink'
    } else {
        Write-Host "⚠️ Error occurred while fetching users. Exiting pagination."
        break
    }
} while ($url)
 
# === Split users into batches of 8,000 ===
$batchSize = 8000
$userBatches = @()
for ($i = 0; $i -lt $users.Count; $i += $batchSize) {
    $userBatches += ,($users[$i..([math]::Min($i + $batchSize - 1, $users.Count - 1))])
}
 
# === Process each batch separately ===
foreach ($batch in $userBatches) {
    $appUserMap = @{}
    foreach ($user in $batch) {
        $email = $user.mail
        $userId = $user.id
        if (![string]::IsNullOrEmpty($email)) {
            try {
                # Fetch installed apps for the user with pagination
                $url = "$graphApiBaseUrl/users/$userId/teamwork/installedApps?expand=teamsApp"
                if (-not $url) {
                    Write-Host "⚠️ Error: URL is null or empty."
                    continue
                }
                $installedApps = Invoke-GraphRequestWithPagination -url $url
                foreach ($app in $installedApps) {
                    $appId = $app.teamsApp.id
                    $appName = $app.teamsApp.displayName
                    if (![string]::IsNullOrEmpty($appId)) {
                        if ($appUserMap.ContainsKey($appId)) {
                            $appUserMap[$appId].Users += $email
                        } else {
                            $appUserMap[$appId] = @{
                                AppName = $appName
                                Users = @($email)
                            }
                        }
                    }
                }
            } catch {
                Write-Host "⚠️ Skipping $email due to error: $_"
            }
        } else {
            Write-Host "Skipping user with ID $userId because email is null or empty."
        }
    }
 
    # === Format for CSV output ===
    $results = @()
    foreach ($appId in $appUserMap.Keys) {
        $appInfo = $appUserMap[$appId]
        $userList = $appInfo.Users -join ", "
        # Split user list if it exceeds cell limit (32,767 characters)
        $maxCellLength = 32767
        $splitUserList = @()
        while ($userList.Length -gt $maxCellLength) {
            $splitUserList += $userList.Substring(0, $maxCellLength)
            $userList = $userList.Substring($maxCellLength)
        }
        $splitUserList += $userList
 
        foreach ($split in $splitUserList) {
            $results += [PSCustomObject]@{
                TeamsAppId = $appId
                AppName    = $appInfo.AppName
                Users      = $split
                UserIds    = ($batch | Where-Object { $_.mail -in $split }).id -join ", "
            }
        }
    }
 
    # === Export Results to CSV ===
    $maxFileSizeMB = 98
    $fileIndex = 1
    $currentFileSizeMB = 0
    $currentResults = @()
 
    foreach ($result in $results) {
        $currentResults += $result
        $currentFileSizeMB = ($currentResults | Export-Csv -Path "temp.csv" -NoTypeInformation | Measure-Object -Property Length -Sum).Sum / 1MB
        if ($currentFileSizeMB -ge $maxFileSizeMB) {
            $currentResults | Export-Csv -Path "C:\Users\RWA\OneDrive - Contoso\Documents\TeamsAPP\FFHOPEresult_batch_$($userBatches.IndexOf($batch))_$fileIndex.csv" -NoTypeInformation
            $currentResults = @()
            $fileIndex++
        }
    }
 
    # Export remaining results if any
    if ($currentResults.Count -gt 0) {
        $currentResults | Export-Csv -Path "C:\Users\RWA\OneDrive - Contoso\Documents\TeamsAPP\FFHOPEresult_batch_$($userBatches.IndexOf($batch))_$fileIndex.csv" -NoTypeInformation
    }
}