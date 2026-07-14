#Requires -Version 7.0

Set-StrictMode -Version Latest

function Get-MailFolderDryRunReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [string]$FolderPath,

        [Parameter()]
        [int]$PreviewCount = 25
    )

    $folder = Get-GraphMailFolderInfo -UserPrincipalName $UserPrincipalName -FolderPath $FolderPath
    $messageCount = Get-GraphMailFolderMessageCount -UserPrincipalName $UserPrincipalName -FolderPath $FolderPath
    $estimatedMinutes = Get-GraphEstimatedMinutes -MessageCount $messageCount

    $previewMessages = @()
    if ($messageCount -gt 0) {
        $messages = Get-GraphMailFolderMessages `
            -UserPrincipalName $UserPrincipalName `
            -FolderPath $FolderPath `
            -Select @('id', 'subject', 'receivedDateTime', 'from')

        $previewMessages = @($messages | Select-Object -First $PreviewCount | ForEach-Object {
            $fromAddress = $null
            if ($_.from -and $_.from.emailAddress) {
                $fromAddress = $_.from.emailAddress.address
            }

            [PSCustomObject]@{
                ReceivedDateTime = $_.receivedDateTime
                Subject          = if ($_.subject) { $_.subject } else { '(no subject)' }
                From             = if ($fromAddress) { $fromAddress } else { '(unknown)' }
            }
        })
    }

    return [PSCustomObject]@{
        Folder           = $folder
        MessageCount     = $messageCount
        EstimatedMinutes = $estimatedMinutes
        PreviewMessages  = $previewMessages
        RemainingCount   = [Math]::Max(0, $messageCount - $previewMessages.Count)
    }
}

function Clear-MailFolderMessages {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$UserPrincipalName,

        [Parameter(Mandatory)]
        [string]$FolderPath,

        [Parameter(Mandatory)]
        [switch]$HardDelete,

        [Parameter()]
        [int]$BatchSize = 100,

        [Parameter()]
        [int]$ProgressInterval = 50,

        [Parameter()]
        [scriptblock]$OnProgress,

        [Parameter()]
        [scriptblock]$WriteLog
    )

    $folder = Get-GraphMailFolderInfo -UserPrincipalName $UserPrincipalName -FolderPath $FolderPath
    $messages = Get-GraphMailFolderMessages -UserPrincipalName $UserPrincipalName -FolderPath $FolderPath -Select @('id')
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

    if ($WriteLog) {
        $deleteMode = if ($HardDelete) { 'permanent' } else { 'soft (deleted items)' }
        & $WriteLog 'INFO' "Deleting $($messageIds.Count) message(s) from '$($folder.DisplayName)' using $deleteMode delete."
    }

    $result = Remove-GraphMailMessages `
        -UserPrincipalName $UserPrincipalName `
        -MessageIds $messageIds `
        -HardDelete:$HardDelete `
        -BatchSize $BatchSize `
        -ProgressInterval $ProgressInterval `
        -OnProgress $OnProgress

    return [PSCustomObject]@{
        FolderDisplayName = $folder.DisplayName
        Deleted           = $result.Deleted
        Failed            = $result.Failed
        Total             = $result.Total
    }
}
