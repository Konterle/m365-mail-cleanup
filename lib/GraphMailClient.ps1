#Requires -Version 7.0

Set-StrictMode -Version Latest

$script:GraphMailClientConnected = $false
$script:GraphMailClientMaxRetries = 5
$script:GraphMailClientPageSize = 50
$script:GraphMailClientConcurrency = 4

function Get-GraphResponseProperty {
    [CmdletBinding()]
    param(
        [Parameter()]
        [object]$Response,

        [Parameter(Mandatory)]
        [string]$Name
    )

    if ($null -eq $Response) {
        return $null
    }

    if ($Response -is [System.Collections.IDictionary]) {
        if ($Response.Contains($Name)) {
            return $Response[$Name]
        }
        return $null
    }

    $property = $Response.PSObject.Properties[$Name]
    if ($property) {
        return $property.Value
    }

    return $null
}

function Get-GraphErrorStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object]$ErrorRecord
    )

    $result = [PSCustomObject]@{
        StatusCode = 0
        RetryAfter = 0
    }

    try {
        $response = $ErrorRecord.Exception.Response
        if ($response) {
            $result.StatusCode = [int]$response.StatusCode

            $retryAfterHeader = $null
            try {
                $retryAfterHeader = $response.Headers['Retry-After']
            }
            catch {
                $retryAfterHeader = $null
            }

            if ($retryAfterHeader) {
                $parsed = 0
                if ([int]::TryParse([string]$retryAfterHeader, [ref]$parsed)) {
                    $result.RetryAfter = $parsed
                }
            }
        }
    }
    catch {
        $result.StatusCode = 0
    }

    return $result
}

function Connect-GraphMailClient {
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$ConfigPath,

        [Parameter()]
        [string]$TenantId,

        [Parameter()]
        [string]$ClientId,

        [Parameter()]
        [string]$CertificateThumbprint,

        [Parameter()]
        [switch]$UseDelegatedAuth
    )

    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw 'Microsoft.Graph.Authentication module is not installed. Run: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop | Out-Null

    $existingContext = Get-MgContext -ErrorAction SilentlyContinue
    if ($existingContext) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }

    if ($UseDelegatedAuth) {
        Connect-MgGraph -Scopes 'Mail.ReadWrite' -NoWelcome -ErrorAction Stop | Out-Null
        Write-Verbose 'Connected to Microsoft Graph using delegated authentication.'
    }
    elseif ($ConfigPath) {
        if (-not (Test-Path -LiteralPath $ConfigPath)) {
            throw "Config file not found: $ConfigPath"
        }

        $config = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
        $TenantId = if ($TenantId) { $TenantId } else { $config.TenantId }
        $ClientId = if ($ClientId) { $ClientId } else { $config.ClientId }
        $CertificateThumbprint = if ($CertificateThumbprint) { $CertificateThumbprint } else { $config.CertificateThumbprint }

        if (-not $TenantId -or -not $ClientId -or -not $CertificateThumbprint) {
            throw 'App-only auth requires TenantId, ClientId, and CertificateThumbprint (via parameters or config file).'
        }

        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop | Out-Null
        Write-Verbose 'Connected to Microsoft Graph using app-only certificate authentication.'
    }
    elseif ($TenantId -and $ClientId -and $CertificateThumbprint) {
        Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop | Out-Null
        Write-Verbose 'Connected to Microsoft Graph using app-only certificate authentication.'
    }
    else {
        Connect-MgGraph -Scopes 'Mail.ReadWrite' -NoWelcome -ErrorAction Stop | Out-Null
        Write-Verbose 'Connected to Microsoft Graph using delegated authentication (default).'
    }

    $script:GraphMailClientConnected = $true
}

function Test-GraphMailClientConnection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter()]
        [string]$FolderPath = 'syncissues'
    )

    if (-not $script:GraphMailClientConnected) {
        throw 'Not connected to Microsoft Graph. Call Connect-GraphMailClient first.'
    }

    $encodedUpn = [uri]::EscapeDataString($UserPrincipalName)
    $uri = "https://graph.microsoft.com/v1.0/users/${encodedUpn}/mailFolders/${FolderPath}?`$select=id,displayName"
    Invoke-GraphMailRequest -Method GET -Uri $uri | Out-Null
    return $true
}

function Get-GraphMailFolderInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [string]$FolderPath
    )

    $encodedUpn = [uri]::EscapeDataString($UserPrincipalName)
    $uri = "https://graph.microsoft.com/v1.0/users/${encodedUpn}/mailFolders/${FolderPath}?`$select=id,displayName,parentFolderId,totalItemCount"

    $folder = Invoke-GraphMailRequest -Method GET -Uri $uri
    return [PSCustomObject]@{
        Id             = $folder.id
        DisplayName    = $folder.displayName
        WellKnownName  = $FolderPath
        FolderPath     = $FolderPath
        ParentFolderId = $folder.parentFolderId
        TotalItemCount = $folder.totalItemCount
    }
}

function Get-GraphEstimatedMinutes {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [int]$MessageCount,

        [Parameter()]
        [int]$Concurrency = $script:GraphMailClientConcurrency
    )

    if ($MessageCount -le 0) {
        return 0
    }

    return [Math]::Max(1, [Math]::Ceiling(($MessageCount / $Concurrency) * 0.5 / 60))
}

function Invoke-GraphMailRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('GET', 'DELETE', 'POST')]
        [string]$Method,

        [Parameter(Mandatory)]
        [string]$Uri,

        [Parameter()]
        [hashtable]$AdditionalHeaders,

        [Parameter()]
        [int]$MaxRetries = $script:GraphMailClientMaxRetries
    )

    if (-not $script:GraphMailClientConnected) {
        throw 'Not connected to Microsoft Graph. Call Connect-GraphMailClient first.'
    }

    $attempt = 0
    while ($true) {
        try {
            $params = @{
                Method      = $Method
                Uri         = $Uri
                OutputType  = 'PSObject'
                ErrorAction = 'Stop'
            }

            if ($AdditionalHeaders) {
                $params['Headers'] = $AdditionalHeaders
            }

            return Invoke-MgGraphRequest @params
        }
        catch {
            $attempt++
            $errorStatus = Get-GraphErrorStatus -ErrorRecord $_
            $isThrottle = ($errorStatus.StatusCode -eq 429) -or ($_.Exception.Message -match '429|Too Many Requests|throttl')
            $retryAfterSeconds = $errorStatus.RetryAfter

            if ($isThrottle -and $attempt -le $MaxRetries) {
                if ($retryAfterSeconds -le 0) {
                    $retryAfterSeconds = [Math]::Min(60, [Math]::Pow(2, $attempt))
                }

                Write-Verbose "Throttled (attempt $attempt/$MaxRetries). Waiting $retryAfterSeconds seconds..."
                Start-Sleep -Seconds $retryAfterSeconds
                continue
            }

            throw
        }
    }
}

function Get-GraphMailFolderMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [string]$FolderPath,

        [Parameter()]
        [string[]]$Select = @('id', 'subject', 'receivedDateTime')
    )

    $encodedUpn = [uri]::EscapeDataString($UserPrincipalName)
    $selectQuery = ($Select -join ',')
    $uri = "https://graph.microsoft.com/v1.0/users/${encodedUpn}/mailFolders/${FolderPath}/messages?`$top=$($script:GraphMailClientPageSize)&`$select=$selectQuery"

    $allMessages = [System.Collections.Generic.List[object]]::new()

    do {
        $response = Invoke-GraphMailRequest -Method GET -Uri $uri
        $items = Get-GraphResponseProperty -Response $response -Name 'value'

        foreach ($message in @($items)) {
            $allMessages.Add($message)
        }

        $uri = Get-GraphResponseProperty -Response $response -Name '@odata.nextLink'
    } while ($uri)

    return $allMessages.ToArray()
}

function Get-GraphMailFolderMessageCount {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [string]$FolderPath
    )

    $encodedUpn = [uri]::EscapeDataString($UserPrincipalName)
    $uri = "https://graph.microsoft.com/v1.0/users/${encodedUpn}/mailFolders/${FolderPath}/messages?`$count=true&`$top=1&`$select=id"
    $headers = @{ ConsistencyLevel = 'eventual' }

    try {
        $response = Invoke-GraphMailRequest -Method GET -Uri $uri -AdditionalHeaders $headers
        $count = Get-GraphResponseProperty -Response $response -Name '@odata.count'
        if ($null -ne $count) {
            return [int]$count
        }
    }
    catch {
        Write-Verbose "Count query failed for folder '$FolderPath'. Falling back to pagination count. $($_.Exception.Message)"
    }

    $messages = Get-GraphMailFolderMessages -UserPrincipalName $UserPrincipalName -FolderPath $FolderPath -Select @('id')
    return @($messages).Count
}

function Remove-GraphMailMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string[]]$MessageIds,

        [Parameter()]
        [switch]$HardDelete,

        [Parameter()]
        [int]$ThrottleLimit = $script:GraphMailClientConcurrency,

        [Parameter()]
        [int]$BatchSize = 100,

        [Parameter()]
        [int]$ProgressInterval = 50,

        [Parameter()]
        [scriptblock]$OnProgress
    )

    $encodedUpn = [uri]::EscapeDataString($UserPrincipalName)
    $allMessageIds = @($MessageIds)
    $totalCount = $allMessageIds.Count
    $totalDeleted = 0
    $totalFailed = 0
    $processed = 0
    $lastReported = 0

    for ($offset = 0; $offset -lt $totalCount; $offset += $BatchSize) {
        $endIndex = [Math]::Min($offset + $BatchSize - 1, $totalCount - 1)
        $batchIds = $allMessageIds[$offset..$endIndex]

        $batchResults = $batchIds | ForEach-Object -Parallel {
        $messageId = $_
        $encodedUser = $using:encodedUpn
        $hardDelete = [bool]$using:HardDelete

        function Invoke-LocalGraphRequest {
            param(
                [string]$Method,
                [string]$Uri
            )

            $attempt = 0
            while ($true) {
                try {
                    return Invoke-MgGraphRequest -Method $Method -Uri $Uri -OutputType PSObject -ErrorAction Stop
                }
                catch {
                    $attempt++
                    $retryAfterSeconds = 0
                    $isThrottle = $false

                    try {
                        $resp = $_.Exception.Response
                        if ($resp -and [int]$resp.StatusCode -eq 429) {
                            $isThrottle = $true
                            $retryAfterHeader = $resp.Headers['Retry-After']
                            if ($retryAfterHeader) {
                                [void][int]::TryParse([string]$retryAfterHeader, [ref]$retryAfterSeconds)
                            }
                        }
                    }
                    catch {
                        $isThrottle = $false
                    }

                    if ($_.Exception.Message -match '429|Too Many Requests|throttl') {
                        $isThrottle = $true
                    }

                    if ($isThrottle -and $attempt -le 5) {
                        if ($retryAfterSeconds -le 0) {
                            $retryAfterSeconds = [Math]::Min(60, [Math]::Pow(2, $attempt))
                        }

                        Start-Sleep -Seconds $retryAfterSeconds
                        continue
                    }

                    throw
                }
            }
        }

        $uri = if ($hardDelete) {
            "https://graph.microsoft.com/v1.0/users/$encodedUser/messages/$messageId/permanentDelete"
        }
        else {
            "https://graph.microsoft.com/v1.0/users/$encodedUser/messages/$messageId"
        }

        try {
            if ($hardDelete) {
                Invoke-LocalGraphRequest -Method POST -Uri $uri | Out-Null
            }
            else {
                Invoke-LocalGraphRequest -Method DELETE -Uri $uri | Out-Null
            }

            [PSCustomObject]@{
                MessageId = $messageId
                Success   = $true
                Error     = $null
            }
        }
        catch {
            [PSCustomObject]@{
                MessageId = $messageId
                Success   = $false
                Error     = $_.Exception.Message
            }
        }
    } -ThrottleLimit $ThrottleLimit

        $batchDeleted = @($batchResults | Where-Object { $_.Success }).Count
        $batchFailed = @($batchResults | Where-Object { -not $_.Success }).Count
        $totalDeleted += $batchDeleted
        $totalFailed += $batchFailed
        $processed += $batchIds.Count

        if ($OnProgress) {
            while ($lastReported + $ProgressInterval -le $processed) {
                $lastReported += $ProgressInterval
                & $OnProgress $lastReported $totalCount $totalDeleted $totalFailed
            }

            if ($processed -eq $totalCount -and $lastReported -lt $totalCount) {
                & $OnProgress $totalCount $totalCount $totalDeleted $totalFailed
            }
        }
    }

    return [PSCustomObject]@{
        Deleted = $totalDeleted
        Failed  = $totalFailed
        Total   = $totalCount
    }
}
