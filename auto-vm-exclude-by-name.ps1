<#
.SYNOPSIS
  Maintain Veeam VBR Global VM Exclusions for Hyper-V VMs based on a name pattern.
  - Include VMs where Name length > 20 AND Name contains "-control-plan".
  - Add new matches to Global VM Exclusions.
  - Remove previously excluded VMs if they no longer exist in inventory.
  - Send an SMTP notification if any changes were made.

.DESCRIPTION
  Resilient to unavailable Hyper-V sources: works even if some registered hosts/clusters/SCVMM are down.
  - Probes each source and skips unreachable ones (fast TCP probe).
  - Wraps per-server Find-VBRHvEntity calls in try/catch to avoid global failure.
  - Skips removal actions if *no* inventory source was successfully queried (prevents false removals).

.REQUIREMENTS
  - Run on the Veeam Backup & Replication Server (Windows PowerShell 5.1).
  - Veeam PowerShell module available (VBR 12.3+).
  - An SMTP relay (optional credentials) for notifications.

.NOTES
  Tested with Veeam Backup & Replication 12.3 (PowerShell on server).
#>

#region ----- User configuration -----

# Pattern & note
$NameContains        = '-control-plan'         # for the Azure-Local control-plane VMs - replace as needed
$MinNameLength       = 21                      # "longer than 20" -> >=21 - for the control-plane VMs - change as needed
$ExclusionNotePrefix = 'Auto-excluded by script'

# SMTP settings (edit to your environment)
$EnableEmail         = $true
$SmtpServer          = 'mail.mydomain.com'
$SmtpPort            = 25
$SmtpUseSsl          = $false
$MailFrom            = 'veeam@mydomain.com'
$MailTo              = 'admins@mydomain.com'
$MailSubject         = '[VBR] Global VM Exclusions updated'
# Optional: credentials for authenticated SMTP
$SmtpCredential      = $null   # e.g. (Get-Credential) or leave $null for anonymous

# Optional: basic logging
$LogFolder           = 'C:\Logs\VBR-GlobalExclusions' # change as needed
$EnableTranscript    = $true

#endregion

#region ----- Bootstrap -----

# Ensure log folder exists
if ($EnableTranscript) {
    try { New-Item -ItemType Directory -Path $LogFolder -ErrorAction SilentlyContinue | Out-Null } catch {}
    $logFile = Join-Path $LogFolder ("GlobalExclusions_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date))
    Start-Transcript -Path $logFile -ErrorAction SilentlyContinue | Out-Null
}

try {
    # Import Veeam PowerShell module (preferred in recent VBR versions)
    Import-Module Veeam.Backup.PowerShell -ErrorAction Stop
} catch {
    # Fallback to legacy snap-in if needed
    try { Add-PSSnapin VeeamPSSnapIn -ErrorAction Stop } catch {
        throw "Cannot load Veeam PowerShell (module or snap-in). Aborting. Details: $($_.Exception.Message)"
    }
}

# If executed on the VBR server, a Connect-VBRServer without parameters connects locally.
try {
    if (-not (Get-VBRServerSession -ErrorAction SilentlyContinue)) {
        Connect-VBRServer | Out-Null
    }
} catch {
    throw "Failed to connect to local VBR server. Details: $($_.Exception.Message)"
}

#endregion

#region ----- Collect Hyper-V inventory (resilient) -----

# Fast TCP probe without relying on Test-NetConnection (fine-grained timeout)
function Test-TcpPort {
    param(
        [Parameter(Mandatory)][string]$ComputerName,
        [int]$Port = 135,            # RPC is a good indicator for Hyper-V/WMI reachability
        [int]$TimeoutMs = 1500
    )
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $iar = $client.BeginConnect($ComputerName, $Port, $null, $null)
        $ok  = $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        if (-not $ok) {
            $client.Close()
            return $false
        }
        $client.EndConnect($iar) | Out-Null
        $client.Close()
        return $true
    } catch {
        return $false
    }
}

# Get all Hyper-V sources known to VBR: standalone hosts, clusters, (optional) SCVMM.
$hvServers = @()
$hvServers += Get-VBRServer -Type HvServer -ErrorAction SilentlyContinue
$hvServers += Get-VBRServer -Type HvCluster -ErrorAction SilentlyContinue
# OPTIONAL: include SCVMM as an inventory source (uncomment next line if desired)
# $hvServers += Get-VBRServer -Type Scvmm   -ErrorAction SilentlyContinue

# Deduplicate by Id and remove nulls
$hvServers = $hvServers | Where-Object { $_ } | Sort-Object Id -Unique

if (-not $hvServers) {
    Write-Host "No Hyper-V servers/clusters/SCVMM connected to VBR. Nothing to do."
    # Cleanup and exit gracefully
    if ($EnableTranscript) { try { Stop-Transcript | Out-Null } catch {} }
    return
}

# Optional knobs
$SkipUnreachableServers = $true
$ConnectivityPort       = 135      # RPC
$ConnectivityTimeoutMs  = 1500

# Collect inventory per server with isolation — use a plain PS array to avoid type cast issues
$hvEntities            = @()
$inventorySourcesOk    = 0
$skippedServers        = @()

foreach ($srv in $hvServers) {
    # Resolve a useful display name and probe target
    $serverName = $null
    if ($srv.PSObject.Properties['DnsName']) { $serverName = $srv.DnsName }
    if ([string]::IsNullOrWhiteSpace($serverName)) { $serverName = $srv.Name }

    $srvType = $srv.Type
    if (-not $srvType) { $srvType = $srv.GetType().Name }

    $shouldSkip = $false
    $skipReason = $null

    if ($SkipUnreachableServers) {
        $reachable = $false
        if ($serverName) {
            $reachable = Test-TcpPort -ComputerName $serverName -Port $ConnectivityPort -TimeoutMs $ConnectivityTimeoutMs
        }
        if (-not $reachable) {
            $shouldSkip = $true
            $skipReason = "No TCP $ConnectivityPort within ${ConnectivityTimeoutMs}ms"
        }
    }

    if ($shouldSkip) {
        $skippedServers += [PSCustomObject]@{ Server=$serverName; Type=$srvType; Reason=$skipReason }
        Write-Warning "Skipping $srvType '$serverName' ($skipReason)"
        continue
    }

    try {
        # Query this server only; do not let one failure break the whole discovery
        $e = Find-VBRHvEntity -Server $srv -HostsAndVMs -ErrorAction Stop
        if ($e) {
            $hvEntities += $e
            $inventorySourcesOk++
        }
    } catch {
        $skippedServers += [PSCustomObject]@{ Server=$serverName; Type=$srvType; Reason=$_.Exception.Message }
        Write-Warning "Failed to query $srvType '$serverName': $($_.Exception.Message)"
        continue
    }
}

# Keep VMs only (Find-VBRHvEntity returns mixed types)
$allHvVms = $hvEntities | Where-Object { $_.GetType().Name -eq 'CHvVmItem' }

if (-not $allHvVms) {
    Write-Warning "No Hyper-V VMs discovered. Queried sources OK: $inventorySourcesOk / $($hvServers.Count)."
}

# Keep a lookup set of existing VM names for later existence checks (case-insensitive)
$existingVmNames = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$allHvVms | ForEach-Object { $null = $existingVmNames.Add($_.Name) }

#endregion

#region ----- Select candidates by naming rule -----

$candidateVms = $allHvVms | Where-Object {
    ($_.Name.Length -ge $MinNameLength) -and ($_.Name -like "*$NameContains*")
}

# Deduplicate by Name to avoid duplicate Add-VBRVMExclusion attempts (e.g., host + cluster visibility)
$candidateVms = $candidateVms | Group-Object Name | ForEach-Object { $_.Group | Select-Object -First 1 }

#endregion

#region ----- Compute changes vs current Global Exclusions -----

# Get current Global VM Exclusions (all platforms), focus on Hyper-V entries we manage here.
$currentExclusions  = Get-VBRVMExclusion
$currentHvExcl      = $currentExclusions | Where-Object { $_.Platform -eq 'HyperV' }

# Set of already excluded VM names (Hyper-V only).
$alreadyExcludedNames = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$currentHvExcl | ForEach-Object { $null = $alreadyExcludedNames.Add($_.Name) }

# New additions: candidate VMs not yet in the exclusions by name.
$toAdd = @($candidateVms | Where-Object { -not $alreadyExcludedNames.Contains($_.Name) })

# Removals: exclusions we previously managed (pattern & length) but whose VM name no longer exists in inventory.
# Only compute removals if at least one inventory source responded OK, to avoid false removals.
if ($inventorySourcesOk -gt 0) {
    $toRemove = @($currentHvExcl | Where-Object {
        ($_.Name.Length -ge $MinNameLength) -and
        ($_.Name -like "*$NameContains*")   -and
        (-not $existingVmNames.Contains($_.Name))
    })
} else {
    Write-Warning "No reachable Hyper-V inventory source; skipping removal actions to avoid false positives."
    $toRemove = @()
}

#endregion

#region ----- Apply changes -----

$changeSummary = New-Object 'System.Collections.Generic.List[string]'
$addedNames    = @()
$removedNames  = @()

if ($toAdd.Count -gt 0) {
    $note = "{0} on {1:yyyy-MM-dd HH:mm} (pattern: '{2}', minLen: {3})" -f $ExclusionNotePrefix, (Get-Date), $NameContains, $MinNameLength
    try {
        Add-VBRVMExclusion -Entity $toAdd -Note $note | Out-Null
        $addedNames = $toAdd | Select-Object -ExpandProperty Name | Sort-Object
        $changeSummary.Add(("Added to Global Exclusions: {0}" -f ($addedNames -join ', '))) | Out-Null
    } catch {
        Write-Warning "Failed to add some VM(s) to Global Exclusions: $($_.Exception.Message)"
    }
}

if ($toRemove.Count -gt 0) {
    try {
        # NOTE: Remove-VBRVMExclusion does not support -Confirm; call directly
        Remove-VBRVMExclusion -Exclusion $toRemove
        $removedNames = $toRemove | Select-Object -ExpandProperty Name | Sort-Object
        $changeSummary.Add(("Removed from Global Exclusions (no longer exist): {0}" -f ($removedNames -join ', '))) | Out-Null
    } catch {
        Write-Warning "Failed to remove some Exclusion(s): $($_.Exception.Message)"
    }
}

#endregion

#region ----- Email notification on change -----

if ($EnableEmail -and ($addedNames.Count -gt 0 -or $removedNames.Count -gt 0)) {

    # Build a clear, sectioned email body
    $body = New-Object System.Text.StringBuilder
    [void]$body.AppendLine("Veeam VBR Global VM Exclusions were updated on $(hostname) at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss').")
    [void]$body.AppendLine()
    [void]$body.AppendLine("Rule: Name length >= $MinNameLength and name contains '$NameContains'")
    [void]$body.AppendLine("Managed platform: Hyper-V")
    [void]$body.AppendLine()

    [void]$body.AppendLine("=== Added to Global Exclusions ===")
    if ($addedNames.Count -gt 0) {
        $addedNames | ForEach-Object { [void]$body.AppendLine(" - $_") }
    } else {
        [void]$body.AppendLine(" - None")
    }
    [void]$body.AppendLine()

    [void]$body.AppendLine("=== Removed from Global Exclusions ===")
    if ($removedNames.Count -gt 0) {
        $removedNames | ForEach-Object { [void]$body.AppendLine(" - $_") }
    } else {
        [void]$body.AppendLine(" - None")
    }
    [void]$body.AppendLine()

    # Optional operational visibility: which sources were skipped/unreachable
    [void]$body.AppendLine("=== Skipped/Unreachable Hyper-V Sources ===")
    if ($skippedServers.Count -gt 0) {
        $skippedServers | ForEach-Object {
            [void]$body.AppendLine((" - {0} ({1}): {2}" -f $_.Server, $_.Type, $_.Reason))
        }
    } else {
        [void]$body.AppendLine(" - None")
    }
    [void]$body.AppendLine()

    [void]$body.AppendLine("This message was generated by an automated PowerShell task.")

    try {
        $mailParams = @{
            To         = $MailTo
            From       = $MailFrom
            Subject    = $MailSubject
            Body       = $body.ToString()
            SmtpServer = $SmtpServer
            Port       = $SmtpPort
            UseSsl     = $SmtpUseSsl
        }
        if ($SmtpCredential) { $mailParams.Credential = $SmtpCredential }
        Send-MailMessage @mailParams
    } catch {
        Write-Warning "Failed to send notification email: $($_.Exception.Message)"
    }
} else {
    Write-Host "No changes required." -ForegroundColor Green
}

#endregion

#region ----- Cleanup -----
if ($EnableTranscript) {
    try { Stop-Transcript | Out-Null } catch {}
}
#endregion