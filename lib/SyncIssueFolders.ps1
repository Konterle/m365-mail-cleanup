#Requires -Version 7.0

Set-StrictMode -Version Latest

$script:SyncIssueWellKnownFolders = @(
    'syncissues',
    'conflicts',
    'localfailures',
    'serverfailures'
)

function Get-SyncIssueFolders {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName
    )

    $encodedUpn = [uri]::EscapeDataString($UserPrincipalName)
    $discoveredFolders = [System.Collections.Generic.List[object]]::new()
    $seenFolderIds = [System.Collections.Generic.HashSet[string]]::new()

    foreach ($wellKnownName in $script:SyncIssueWellKnownFolders) {
        $folder = Get-SyncIssueFolderInfo -UserPrincipalName $UserPrincipalName -FolderPath $wellKnownName
        if ($folder) {
            Add-SyncIssueFolderIfNew -Folders $discoveredFolders -SeenIds $seenFolderIds -Folder $folder
        }
    }

    try {
        $uri = "https://graph.microsoft.com/v1.0/users/${encodedUpn}/mailFolders/syncissues/childFolders?`$select=id,displayName,parentFolderId"
        $nextLink = $uri

        do {
            $response = Invoke-GraphMailRequest -Method GET -Uri $nextLink
            $childItems = Get-GraphResponseProperty -Response $response -Name 'value'

            foreach ($childFolder in @($childItems)) {
                $folderInfo = [PSCustomObject]@{
                    Id             = $childFolder.id
                    DisplayName    = $childFolder.displayName
                    WellKnownName  = $null
                    FolderPath     = $childFolder.id
                    IsWellKnown    = $false
                    ParentFolderId = $childFolder.parentFolderId
                }

                Add-SyncIssueFolderIfNew -Folders $discoveredFolders -SeenIds $seenFolderIds -Folder $folderInfo
            }

            $nextLink = Get-GraphResponseProperty -Response $response -Name '@odata.nextLink'
        } while ($nextLink)
    }
    catch {
        Write-Warning "Could not enumerate child folders under syncissues: $($_.Exception.Message)"
    }

    return $discoveredFolders.ToArray()
}

function Get-SyncIssueFolderInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [string]$FolderPath
    )

    $encodedUpn = [uri]::EscapeDataString($UserPrincipalName)
    $uri = "https://graph.microsoft.com/v1.0/users/${encodedUpn}/mailFolders/${FolderPath}?`$select=id,displayName,parentFolderId"

    try {
        $folder = Invoke-GraphMailRequest -Method GET -Uri $uri
        return [PSCustomObject]@{
            Id              = $folder.id
            DisplayName     = $folder.displayName
            WellKnownName   = $FolderPath
            FolderPath      = $FolderPath
            IsWellKnown     = $true
            ParentFolderId  = $folder.parentFolderId
        }
    }
    catch {
        Write-Verbose "Folder '$FolderPath' not found or not accessible: $($_.Exception.Message)"
        return $null
    }
}

function Add-SyncIssueFolderIfNew {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[object]]$Folders,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$SeenIds,

        [Parameter(Mandatory)]
        [object]$Folder
    )

    if (-not $Folder.Id) {
        return
    }

    if ($SeenIds.Add([string]$Folder.Id)) {
        $Folders.Add($Folder)
    }
}

function Get-SyncIssueFolderSummary {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Folders
    )

    $summaries = foreach ($folder in $Folders) {
        $messageCount = Get-GraphMailFolderMessageCount -UserPrincipalName $UserPrincipalName -FolderPath $folder.FolderPath
        $estimatedMinutes = if ($messageCount -gt 0) {
            [Math]::Ceiling(($messageCount / $script:GraphMailClientConcurrency) * 0.5 / 60)
        }
        else {
            0
        }

        [PSCustomObject]@{
            DisplayName       = $folder.DisplayName
            WellKnownName     = $folder.WellKnownName
            FolderPath        = $folder.FolderPath
            MessageCount      = $messageCount
            EstimatedMinutes  = [Math]::Max(1, $estimatedMinutes)
        }
    }

    return @($summaries)
}

function Clear-SyncIssueFolderMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [object]$Folder,

        [Parameter()]
        [switch]$HardDelete,

        [Parameter()]
        [switch]$WhatIf,

        [Parameter()]
        [int]$ProgressInterval = 50,

        [Parameter()]
        [scriptblock]$OnProgress,

        [Parameter()]
        [scriptblock]$WriteLog
    )

    $messages = Get-GraphMailFolderMessages -UserPrincipalName $UserPrincipalName -FolderPath $folder.FolderPath
    $messageIds = @($messages | ForEach-Object { $_.id } | Where-Object { $_ })

    if ($messageIds.Count -eq 0) {
        if ($WriteLog) {
            & $WriteLog 'INFO' "Folder '$($folder.DisplayName)' is already empty."
        }

        return [PSCustomObject]@{
            FolderDisplayName = $folder.DisplayName
            Deleted           = 0
            Failed            = 0
            Total             = 0
        }
    }

    if ($WhatIf) {
        if ($WriteLog) {
            & $WriteLog 'INFO' "WhatIf: Would delete $($messageIds.Count) message(s) from '$($folder.DisplayName)'."
        }

        return [PSCustomObject]@{
            FolderDisplayName = $folder.DisplayName
            Deleted           = 0
            Failed            = 0
            Total             = $messageIds.Count
            Skipped           = $true
        }
    }

    if ($WriteLog) {
        $deleteMode = if ($HardDelete) { 'permanent' } else { 'soft (deleted items)' }
        & $WriteLog 'INFO' "Deleting $($messageIds.Count) message(s) from '$($folder.DisplayName)' using $deleteMode delete."
    }

    $result = Remove-GraphMailMessages `
        -UserPrincipalName $UserPrincipalName `
        -MessageIds $messageIds `
        -HardDelete:$HardDelete `
        -ProgressInterval $ProgressInterval `
        -OnProgress $OnProgress

    return [PSCustomObject]@{
        FolderDisplayName = $folder.DisplayName
        Deleted           = $result.Deleted
        Failed            = $result.Failed
        Total             = $result.Total
    }
}
