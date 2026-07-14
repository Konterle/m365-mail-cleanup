#Requires -Version 7.0

<#
.SYNOPSIS
    Deletes messages from Sync Issues folders in a user's mailbox via Microsoft Graph.

.DESCRIPTION
    Enumerates the syncissues area (syncissues, conflicts, localfailures, serverfailures,
    and additional child folders) and deletes all contained messages with throttling-safe
    parallel processing (max 4 concurrent requests per mailbox).

.PARAMETER UserPrincipalName
    Target mailbox UPN, for example max.mustermann@contoso.de.

.PARAMETER DryRun
    Only counts messages and shows an estimate. No deletions are performed.

.PARAMETER HardDelete
    Permanently deletes messages instead of moving them to Deleted Items.

.PARAMETER ConfigPath
    Optional path to config.json for app-only certificate authentication.

.PARAMETER UseDelegatedAuth
    Force delegated interactive authentication with Mail.ReadWrite scope.

.PARAMETER LogPath
    Optional path for a structured log file.

.EXAMPLE
    .\Remove-SyncIssuesMail.ps1 -UserPrincipalName max.mustermann@contoso.de -DryRun

.EXAMPLE
    .\Remove-SyncIssuesMail.ps1 -UserPrincipalName max.mustermann@contoso.de -Confirm

.EXAMPLE
    .\Remove-SyncIssuesMail.ps1 -UserPrincipalName max.mustermann@contoso.de -HardDelete -Confirm
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory)]
    [ValidateNotNullOrEmpty()]
    [string]$UserPrincipalName,

    [Parameter()]
    [switch]$DryRun,

    [Parameter()]
    [switch]$HardDelete,

    [Parameter()]
    [string]$ConfigPath,

    [Parameter()]
    [string]$TenantId,

    [Parameter()]
    [string]$ClientId,

    [Parameter()]
    [string]$CertificateThumbprint,

    [Parameter()]
    [switch]$UseDelegatedAuth,

    [Parameter()]
    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptRoot = $PSScriptRoot
. (Join-Path $scriptRoot 'lib\GraphMailClient.ps1')
. (Join-Path $scriptRoot 'lib\SyncIssueFolders.ps1')

function Write-CleanupLog {
    param(
        [Parameter(Mandatory)]
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level,

        [Parameter(Mandatory)]
        [string]$Message
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        'INFO' { Write-Host $line }
        'WARN' { Write-Warning $Message }
        'ERROR' { Write-Host $line -ForegroundColor Red }
    }

    if ($LogPath) {
        Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    }
}

try {
    if ($LogPath) {
        $logDirectory = Split-Path -Parent $LogPath
        if ($logDirectory -and -not (Test-Path -LiteralPath $logDirectory)) {
            New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
        }
    }

    Write-CleanupLog -Level INFO -Message "Starting Sync Issues cleanup for '$UserPrincipalName'."

    $connectParams = @{
        UseDelegatedAuth = [switch]$UseDelegatedAuth
    }

    if ($ConfigPath) { $connectParams['ConfigPath'] = $ConfigPath }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    if ($ClientId) { $connectParams['ClientId'] = $ClientId }
    if ($CertificateThumbprint) { $connectParams['CertificateThumbprint'] = $CertificateThumbprint }

    Connect-GraphMailClient @connectParams
    Test-GraphMailClientConnection -UserPrincipalName $UserPrincipalName | Out-Null
    Write-CleanupLog -Level INFO -Message 'Microsoft Graph connection verified.'

    $folders = @(Get-SyncIssueFolders -UserPrincipalName $UserPrincipalName)
    if ($folders.Count -eq 0) {
        throw 'No Sync Issues folders were found for the specified user.'
    }

    Write-CleanupLog -Level INFO -Message "Discovered $($folders.Count) folder(s) in the Sync Issues area."

    $summary = @(Get-SyncIssueFolderSummary -UserPrincipalName $UserPrincipalName -Folders $folders)
    $totalMessages = ($summary | Measure-Object -Property MessageCount -Sum).Sum
    $estimatedMinutes = if ($totalMessages -gt 0) {
        [Math]::Max(1, [Math]::Ceiling(($totalMessages / 4) * 0.5 / 60))
    }
    else {
        0
    }

    $summary | Format-Table DisplayName, WellKnownName, MessageCount, EstimatedMinutes -AutoSize | Out-String | ForEach-Object {
        Write-CleanupLog -Level INFO -Message $_.TrimEnd()
    }

    Write-CleanupLog -Level INFO -Message "Total messages: $totalMessages. Estimated duration: ~$estimatedMinutes minute(s)."

    if ($DryRun) {
        Write-CleanupLog -Level INFO -Message 'DryRun complete. No messages were deleted.'
        return
    }

    if ($totalMessages -eq 0) {
        Write-CleanupLog -Level INFO -Message 'No messages to delete. Exiting.'
        return
    }

    $deleteModeLabel = if ($HardDelete) { 'permanent delete' } else { 'soft delete (Deleted Items)' }
    $confirmationMessage = "Delete $totalMessages message(s) from Sync Issues folders for '$UserPrincipalName' using ${deleteModeLabel}?"

    if ($PSCmdlet.ShouldProcess($UserPrincipalName, $confirmationMessage)) {
        $folderResults = [System.Collections.Generic.List[object]]::new()
        $processedMessages = 0

        foreach ($folder in $folders) {
            $folderSummary = @($summary | Where-Object { $_.FolderPath -eq $folder.FolderPath })
            if ($folderSummary.Count -eq 0 -or $folderSummary[0].MessageCount -eq 0) {
                continue
            }

            $progressBlock = {
                param($Processed, $Total, $Deleted, $Failed)
                Write-CleanupLog -Level INFO -Message "Progress '$($folder.DisplayName)': $Processed / $Total ($Deleted deleted, $Failed failed)."
            }

            $result = Clear-SyncIssueFolderMessages `
                -UserPrincipalName $UserPrincipalName `
                -Folder $folder `
                -HardDelete:$HardDelete `
                -ProgressInterval 50 `
                -OnProgress $progressBlock `
                -WriteLog ${function:Write-CleanupLog}

            $folderResults.Add($result)
            $processedMessages += $result.Total

            Write-CleanupLog -Level INFO -Message (
                "Folder '$($result.FolderDisplayName)': deleted $($result.Deleted)/$($result.Total), failed $($result.Failed)."
            )
        }

        $deletedTotal = ($folderResults | Measure-Object -Property Deleted -Sum).Sum
        $failedTotal = ($folderResults | Measure-Object -Property Failed -Sum).Sum

        Write-CleanupLog -Level INFO -Message "Cleanup finished. Deleted: $deletedTotal, Failed: $failedTotal, Processed: $processedMessages."

        if ($failedTotal -gt 0) {
            Write-CleanupLog -Level WARN -Message 'Some messages could not be deleted. Re-run the script or inspect Graph permissions/throttling.'
        }
    }
    else {
        Write-CleanupLog -Level INFO -Message 'Operation cancelled by user.'
    }
}
catch {
    Write-CleanupLog -Level ERROR -Message $_.Exception.Message
    throw
}
finally {
    if (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue) {
        if (Get-MgContext -ErrorAction SilentlyContinue) {
            Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
        }
    }
}
