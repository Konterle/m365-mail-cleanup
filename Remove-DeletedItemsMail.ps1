#Requires -Version 7.0

<#
.SYNOPSIS
    Permanently empties the Deleted Items folder (Gelöschte Elemente) via Microsoft Graph.

.DESCRIPTION
    Lists and counts messages in the deleteditems folder with -DryRun.
    Permanently deletes all messages with -HardDelete using throttling-safe batch processing.

    Note: A normal DELETE on items already in Deleted Items only moves them to
    recoverableitemsdeletions. This script requires -HardDelete for actual cleanup.

.PARAMETER UserPrincipalName
    Target mailbox UPN.

.PARAMETER DryRun
    Count and preview messages. No deletions are performed.

.PARAMETER HardDelete
    Required for deletion. Permanently removes messages from the mailbox.

.PARAMETER PreviewCount
    Number of messages to preview in DryRun output. Default: 25.

.PARAMETER LogPath
    Optional path for a structured log file.

.EXAMPLE
    .\Remove-DeletedItemsMail.ps1 -UserPrincipalName max.mustermann@contoso.de -DryRun

.EXAMPLE
    .\Remove-DeletedItemsMail.ps1 -UserPrincipalName max.mustermann@contoso.de -HardDelete -Confirm
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
    [ValidateRange(1, 500)]
    [int]$PreviewCount = 25,

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

$script:DeletedItemsFolderPath = 'deleteditems'
$scriptRoot = $PSScriptRoot

. (Join-Path $scriptRoot 'lib\GraphMailClient.ps1')
. (Join-Path $scriptRoot 'lib\MailFolderCleanup.ps1')

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

    if (-not $DryRun -and -not $HardDelete) {
        throw 'Permanent deletion requires -HardDelete. Use -DryRun to preview without deleting.'
    }

    Write-CleanupLog -Level INFO -Message "Starting Deleted Items cleanup for '$UserPrincipalName'."

    $connectParams = @{
        UseDelegatedAuth = [switch]$UseDelegatedAuth
    }

    if ($ConfigPath) { $connectParams['ConfigPath'] = $ConfigPath }
    if ($TenantId) { $connectParams['TenantId'] = $TenantId }
    if ($ClientId) { $connectParams['ClientId'] = $ClientId }
    if ($CertificateThumbprint) { $connectParams['CertificateThumbprint'] = $CertificateThumbprint }

    Connect-GraphMailClient @connectParams
    Test-GraphMailClientConnection -UserPrincipalName $UserPrincipalName -FolderPath $script:DeletedItemsFolderPath | Out-Null
    Write-CleanupLog -Level INFO -Message 'Microsoft Graph connection verified.'

    Write-CleanupLog -Level INFO -Message 'Analyzing Deleted Items folder (deleteditems)...'
    $report = Get-MailFolderDryRunReport `
        -UserPrincipalName $UserPrincipalName `
        -FolderPath $script:DeletedItemsFolderPath `
        -PreviewCount $PreviewCount

    [PSCustomObject]@{
        DisplayName      = $report.Folder.DisplayName
        WellKnownName    = $script:DeletedItemsFolderPath
        MessageCount     = $report.MessageCount
        EstimatedMinutes = $report.EstimatedMinutes
    } | Format-Table -AutoSize | Out-String | ForEach-Object {
        Write-CleanupLog -Level INFO -Message $_.TrimEnd()
    }

    Write-CleanupLog -Level INFO -Message (
        "Total messages in Deleted Items: $($report.MessageCount). Estimated duration: ~$($report.EstimatedMinutes) minute(s)."
    )

    if ($report.MessageCount -gt 0 -and $report.PreviewMessages.Count -gt 0) {
        Write-CleanupLog -Level INFO -Message "Preview (first $($report.PreviewMessages.Count) message(s)):"
        $report.PreviewMessages | Format-Table ReceivedDateTime, From, Subject -AutoSize | Out-String | ForEach-Object {
            Write-CleanupLog -Level INFO -Message $_.TrimEnd()
        }

        if ($report.RemainingCount -gt 0) {
            Write-CleanupLog -Level INFO -Message "... and $($report.RemainingCount) more message(s) not shown."
        }
    }

    if ($DryRun) {
        Write-CleanupLog -Level INFO -Message 'DryRun complete. No messages were deleted.'
        return
    }

    if ($report.MessageCount -eq 0) {
        Write-CleanupLog -Level INFO -Message 'Deleted Items folder is already empty. Exiting.'
        return
    }

    $confirmationMessage = "Permanently delete $($report.MessageCount) message(s) from Deleted Items for '$UserPrincipalName'?"

    if ($PSCmdlet.ShouldProcess($UserPrincipalName, $confirmationMessage)) {
        $progressBlock = {
            param($Processed, $Total, $Deleted, $Failed)
            Write-CleanupLog -Level INFO -Message "Progress: $Processed / $Total processed ($Deleted deleted, $Failed failed)."
        }

        $result = Clear-MailFolderMessages `
            -UserPrincipalName $UserPrincipalName `
            -FolderPath $script:DeletedItemsFolderPath `
            -HardDelete `
            -ProgressInterval 50 `
            -OnProgress $progressBlock `
            -WriteLog ${function:Write-CleanupLog}

        Write-CleanupLog -Level INFO -Message (
            "Folder '$($result.FolderDisplayName)': deleted $($result.Deleted)/$($result.Total), failed $($result.Failed)."
        )

        if ($result.Failed -gt 0) {
            Write-CleanupLog -Level WARN -Message 'Some messages could not be deleted. Re-run with -DryRun or inspect Graph permissions/throttling.'
        }
        else {
            Write-CleanupLog -Level INFO -Message 'Deleted Items folder permanently emptied.'
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
