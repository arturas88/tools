<#
.SYNOPSIS
    Exchange Online Mailbox Manager - Cleanup and Reporting Tool
    
.DESCRIPTION
    Comprehensive Exchange Online mailbox management solution:
    - Comprehensive mailbox reporting (sizes, quotas, activity, folders)
    - Detailed folder statistics and analysis with CSV export
    - Graph API deletion (for regular folders) with safety confirmations
    - Compliance Search deletion (for Deleted Items and permanent deletion)
    - Retention policy management with safety confirmations
    - Quota issue handling
    - Multiple authentication methods
    - Built-in safeguards requiring explicit confirmation before deletion
    
    PERFORMANCE: Uses optimized Graph API batch operations (20 deletions per HTTP request)
    for up to 10-20x faster deletion compared to single-item deletion.
    
    Based on extensive testing and real-world scenarios.
    
.PARAMETER MailboxAddress
    Target mailbox email address (required)
    
.PARAMETER AdminUpn
    Admin UPN for Compliance Search operations (required for -UseComplianceSearch)
    
.PARAMETER FolderPath
    Specific folder to clean (leave empty for all folders)
    
.PARAMETER OlderThanDays
    Delete emails older than X days (default: 365)
    
.PARAMETER OlderThanDate
    Delete emails older than specific date (alternative to OlderThanDays)
    
.PARAMETER DateRangeStart
    Start date for date range deletion (e.g., "2023-01-01")
    Must be used together with -DateRangeEnd
    Range cannot exceed 365 days
    
.PARAMETER DateRangeEnd
    End date for date range deletion (e.g., "2023-12-31")
    Must be used together with -DateRangeStart
    Range cannot exceed 365 days
    
.PARAMETER ConfirmDelete
    Mandatory safety switch to enable actual deletion
    Without this switch, script will only run in analysis/dry-run mode
    This prevents accidental deletions
    
.PARAMETER UseComplianceSearch
    Use Microsoft Purview Compliance Search instead of Graph API
    Required for: Deleted Items, Recoverable Items, or when quota issues exist
    
.PARAMETER RemoveRetentionPolicy
    Automatically remove retention policies that block deletion
    
.PARAMETER CheckOnly
    Only check mailbox status, don't delete anything
    
.PARAMETER DryRun
    Show what would be deleted without actually deleting
    
.PARAMETER TenantId
    Azure AD Tenant ID (required for Graph API mode)
    
.PARAMETER ClientId
    Application Client ID (required for Graph API mode)
    
.PARAMETER ClientSecret
    Application Client Secret (required for Graph API mode)
    
.PARAMETER BatchSize
    Controls fetch batch size (default: 100)
    Actual fetch size is BatchSize * 4 (capped at 200) for optimal performance
    Messages are deleted in Graph API batches of 20 per HTTP request
    Example: -BatchSize 100 will fetch up to 200 messages and delete them in batches of 20
    
.PARAMETER SkipStatistics
    Skip statistics gathering for faster execution
    
.PARAMETER ForceWait
    Wait for Compliance Search to complete before exiting (can take hours)
    
.PARAMETER GenerateReport
    Generate comprehensive mailbox statistics report for all mailboxes or specific mailbox
    
.PARAMETER DetailedFolderReport
    Generate detailed folder statistics for a specific mailbox (requires -MailboxAddress)
    
.PARAMETER ReportAllMailboxes
    When used with -GenerateReport, reports on all mailboxes instead of just one
    
.EXAMPLE
    # Check mailbox status and identify issues
    .\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" -CheckOnly -AdminUpn "admin@domain.com"
    
.EXAMPLE
    # Generate comprehensive report for all mailboxes (exports CSV sorted by size DESC)
    .\ExchangeMailboxManager.ps1 -GenerateReport -ReportAllMailboxes -AdminUpn "admin@domain.com"
    
.EXAMPLE
    # Generate detailed folder report for specific mailbox (exports CSV sorted by size DESC)
    .\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" -DetailedFolderReport -AdminUpn "admin@domain.com"
    
.EXAMPLE
    # Generate both comprehensive and detailed reports for one mailbox
    .\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" -GenerateReport -DetailedFolderReport -AdminUpn "admin@domain.com"
    
.EXAMPLE
    # Clean regular folders using Graph API (will prompt for confirmation: type 'YES')
    .\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" -FolderPath "Sent Items" -OlderThanDays 365 -TenantId "xxx" -ClientId "yyy" -ClientSecret "zzz"
    
.EXAMPLE
    # Clean Deleted Items using Compliance Search (will prompt: type 'REMOVE' then 'DELETE')
    .\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" -FolderPath "Deleted Items" -UseComplianceSearch -AdminUpn "admin@domain.com" -RemoveRetentionPolicy
    
.EXAMPLE
    # Nuclear option: Remove everything using Compliance Search
    .\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" -UseComplianceSearch -AdminUpn "admin@domain.com" -RemoveRetentionPolicy -OlderThanDays 0
    
.EXAMPLE
    # Delete emails in a specific date range (e.g., all emails from 2023)
    .\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" -DateRangeStart "2023-01-01" -DateRangeEnd "2023-12-31" -ConfirmDelete -TenantId "xxx" -ClientId "yyy" -ClientSecret "zzz"
    
.EXAMPLE
    # Delete old emails before a specific date with mandatory confirmation
    .\ExchangeMailboxManager.ps1 -MailboxAddress "user@domain.com" -OlderThanDate "2024-01-01" -ConfirmDelete -UseComplianceSearch -AdminUpn "admin@domain.com"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$MailboxAddress,
    
    [Parameter(Mandatory=$false)]
    [string]$AdminUpn,
    
    [Parameter(Mandatory=$false)]
    [string]$FolderPath = "",
    
    [Parameter(Mandatory=$false)]
    [int]$OlderThanDays = 365,
    
    [Parameter(Mandatory=$false)]
    [datetime]$OlderThanDate,
    
    [Parameter(Mandatory=$false)]
    [datetime]$DateRangeStart,
    
    [Parameter(Mandatory=$false)]
    [datetime]$DateRangeEnd,
    
    [Parameter(Mandatory=$false)]
    [switch]$ConfirmDelete,
    
    [Parameter(Mandatory=$false)]
    [switch]$UseComplianceSearch,
    
    [Parameter(Mandatory=$false)]
    [switch]$RemoveRetentionPolicy,
    
    [Parameter(Mandatory=$false)]
    [switch]$CheckOnly,
    
    [Parameter(Mandatory=$false)]
    [switch]$DryRun,
    
    [Parameter(Mandatory=$false)]
    [string]$TenantId,
    
    [Parameter(Mandatory=$false)]
    [string]$ClientId,
    
    [Parameter(Mandatory=$false)]
    [string]$ClientSecret,
    
    [Parameter(Mandatory=$false)]
    [int]$BatchSize = 100,
    
    [Parameter(Mandatory=$false)]
    [switch]$SkipStatistics,
    
    [Parameter(Mandatory=$false)]
    [switch]$ForceWait,
    
    [Parameter(Mandatory=$false)]
    [switch]$GenerateReport,
    
    [Parameter(Mandatory=$false)]
    [switch]$DetailedFolderReport,
    
    [Parameter(Mandatory=$false)]
    [switch]$ReportAllMailboxes
)

#region Configuration
$ErrorActionPreference = "Stop"
$LogFile = "ExchangeMailboxManager_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$script:TotalProcessed = 0
$script:AccessToken = $null
$script:ScriptStartTime = Get-Date

# Force period as decimal separator (instead of comma in some locales)
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::InvariantCulture
[System.Threading.Thread]::CurrentThread.CurrentUICulture = [System.Globalization.CultureInfo]::InvariantCulture
#endregion

#region Logging Functions
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO","SUCCESS","WARNING","ERROR")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    $color = switch ($Level) {
        "SUCCESS" { "Green" }
        "WARNING" { "Yellow" }
        "ERROR" { "Red" }
        default { "White" }
    }
    
    Write-Host $logMessage -ForegroundColor $color
    Add-Content -Path $LogFile -Value $logMessage
}

function Write-Banner {
    param([string]$Title)
    Write-Log "========================================" "INFO"
    Write-Log $Title "INFO"
    Write-Log "========================================" "INFO"
}

function Write-LogAndHost {
    <#
    .SYNOPSIS
        Writes a message to both console and log file without timestamp formatting
    .DESCRIPTION
        Use this for formatting-specific output (tables, prompts, etc.) that needs to appear
        in the log file exactly as shown to the user
    #>
    param(
        [string]$Message,
        [string]$Color = "White"
    )
    
    Write-Host $Message -ForegroundColor $Color
    Add-Content -Path $LogFile -Value $Message
}

function Read-HostWithLogging {
    <#
    .SYNOPSIS
        Reads user input and logs both the prompt and response
    .DESCRIPTION
        Security-enhanced Read-Host that logs all interactive prompts and user responses
    #>
    param(
        [string]$Prompt
    )
    
    # Log the prompt
    Write-Log "USER PROMPT: $Prompt" "INFO"
    
    # Get user input
    $response = Read-Host $Prompt
    
    # Log the response (sanitized for safety)
    if ($response -match '^(YES|DELETE|REMOVE|NO|CANCEL)$') {
        Write-Log "USER RESPONSE: $response" "WARNING"
    } else {
        Write-Log "USER RESPONSE: [user provided response]" "INFO"
    }
    
    return $response
}

function ConvertTo-Bytes {
    <#
    .SYNOPSIS
        Safely converts ByteQuantifiedSize to bytes
    .DESCRIPTION
        Handles both objects with methods and deserialized objects without methods
    #>
    param($Size)
    
    if (-not $Size) { return 0 }
    
    try {
        # Check if Size has a Value property (common with Exchange objects)
        $sizeObject = $Size
        if ($Size.PSObject.Properties['Value']) {
            $sizeObject = $Size.Value
        }
        
        # Try calling ToBytes() method first (works for non-deserialized objects)
        if ($sizeObject.PSObject.Methods['ToBytes']) {
            return $sizeObject.ToBytes()
        }
        
        # For deserialized objects, parse from string
        $sizeString = $sizeObject.ToString()
        
        # Pattern 1: "45.23 GB (48,620,789,123 bytes)" - extract bytes from parentheses
        if ($sizeString -match '\(([0-9,]+)\s*bytes?\)') {
            return [double]($matches[1] -replace ',','')
        }
        
        # Pattern 2: "48620789123 bytes" - direct bytes value
        if ($sizeString -match '^([0-9,]+)\s*bytes?\s*$') {
            return [double]($matches[1] -replace ',','')
        }
        
        # Pattern 3: "45.23 GB" - convert from GB/MB/KB
        if ($sizeString -match '([0-9.,]+)\s*(KB|MB|GB|TB)') {
            $value = [double]($matches[1] -replace ',','')
            $unit = $matches[2]
            switch ($unit) {
                "KB" { return $value * 1KB }
                "MB" { return $value * 1MB }
                "GB" { return $value * 1GB }
                "TB" { return $value * 1TB }
                default { return $value }
            }
        }
        
        # Pattern 4: Just a number (assume bytes)
        if ($sizeString -match '^([0-9,]+)$') {
            return [double]($sizeString -replace ',','')
        }
    }
    catch {
        # Log parsing failure for debugging
        Write-Verbose "ConvertTo-Bytes: Failed to parse size: $($_.Exception.Message)"
    }
    
    return 0
}
#endregion

#region Mailbox Diagnostics
function Get-MailboxDiagnostics {
    param([string]$MailboxAddress)
    
    Write-Banner "Mailbox Diagnostics"
    
    try {
        # Ensure Exchange Online is connected - check for active session
        $exoConnected = $false
        try {
            $null = Get-OrganizationConfig -ErrorAction Stop
            $exoConnected = $true
        }
        catch {
            $exoConnected = $false
        }
        
        if (-not $exoConnected) {
            Write-Log "Connecting to Exchange Online..." "INFO"
            Connect-ExchangeOnline -UserPrincipalName $AdminUpn -ShowBanner:$false
        }
        else {
            Write-Log "Using existing Exchange Online connection" "INFO"
        }
        
        # Get mailbox properties
        $mailbox = Get-Mailbox $MailboxAddress
        $stats = Get-MailboxStatistics $MailboxAddress
        
        Write-Log "`nMailbox Information:" "INFO"
        Write-Log "  Display Name: $($mailbox.DisplayName)" "INFO"
        Write-Log "  Total Items: $($stats.ItemCount)" "INFO"
        Write-Log "  Total Size: $($stats.TotalItemSize)" "INFO"
        Write-Log "  Deleted Items: $($stats.DeletedItemCount)" "INFO"
        
        Write-Log "`nRetention & Hold Settings:" "INFO"
        Write-Log "  Retention Policy: $($mailbox.RetentionPolicy)" "INFO"
        Write-Log "  Litigation Hold: $($mailbox.LitigationHoldEnabled)" "INFO"
        Write-Log "  In-Place Holds: $($mailbox.InPlaceHolds -join ', ')" "INFO"
        Write-Log "  Retain Deleted Items: $($mailbox.RetainDeletedItemsFor)" "INFO"
        Write-Log "  Single Item Recovery: $($mailbox.SingleItemRecoveryEnabled)" "INFO"
        Write-Log "  Delay Hold: $($mailbox.DelayHoldApplied)" "INFO"
        
        # Get folder statistics
        Write-Log "`nTop Folders by Item Count:" "INFO"
        $folders = Get-MailboxFolderStatistics $MailboxAddress | 
            Where-Object {$_.ItemsInFolder -gt 0} | 
            Sort-Object ItemsInFolder -Descending | 
            Select-Object -First 10
        
        foreach ($folder in $folders) {
            Write-Log "  $($folder.Name): $($folder.ItemsInFolder) items ($($folder.FolderSize))" "INFO"
        }
        
        # Check for issues
        Write-Log "`nPotential Issues:" "WARNING"
        
        if ($mailbox.RetentionPolicy) {
            Write-Log "  ⚠️  Retention policy active: $($mailbox.RetentionPolicy)" "WARNING"
            Write-Log "     Recommendation: Use -RemoveRetentionPolicy flag" "WARNING"
        }
        
        if ($mailbox.LitigationHoldEnabled) {
            Write-Log "  ⚠️  Litigation hold enabled - deletion may be blocked" "WARNING"
        }
        
        if ($mailbox.DelayHoldApplied) {
            Write-Log "  ⚠️  Delay hold active - items won't delete for 30 days" "WARNING"
        }
        
        $deletedItemsFolder = $folders | Where-Object {$_.Name -like "*Deleted*"}
        if ($deletedItemsFolder -and $deletedItemsFolder.ItemsInFolder -gt 10000) {
            Write-Log "  ⚠️  Large Deleted Items folder: $($deletedItemsFolder.ItemsInFolder) items" "WARNING"
            Write-Log "     Recommendation: Use -UseComplianceSearch for Deleted Items" "WARNING"
        }
        
        $recoverableFolder = Get-MailboxFolderStatistics $MailboxAddress -FolderScope RecoverableItems | 
            Where-Object {$_.ItemsInFolder -gt 0}
        
        if ($recoverableFolder) {
            $totalRecoverable = ($recoverableFolder | Measure-Object -Property ItemsInFolder -Sum).Sum
            Write-Log "  ⚠️  Recoverable Items: $totalRecoverable items" "WARNING"
            Write-Log "     Recommendation: Use -UseComplianceSearch to clean" "WARNING"
        }
        
        # Check if mailbox size exceeds 30 GB and suggest cleanup commands
        $sizeBytes = ConvertTo-Bytes -Size $stats.TotalItemSize
        $sizeGB = [math]::Round($sizeBytes / 1GB, 2)
        
        Write-Log "`nMailbox Size Analysis:" "INFO"
        Write-Log "  Total Size: $($stats.TotalItemSize)" "INFO"
        Write-Log "  Parsed Size: $sizeGB GB ($sizeBytes bytes)" "INFO"
        
        if ($sizeGB -gt 30) {
            Write-Log "`n⚠️  MAILBOX SIZE ALERT: $sizeGB GB (exceeds 30 GB threshold)" "WARNING"
            Write-Log "`nSuggested Cleanup Commands:" "INFO"
            Write-Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" "INFO"
            
            Write-Log "`n1. Delete emails older than a specific date (e.g., before 2024-01-01):" "INFO"
            Write-LogAndHost "   .\ExchangeMailboxManager.ps1 -MailboxAddress '$MailboxAddress' ``" -Color Cyan
            Write-LogAndHost "     -OlderThanDate '2024-01-01' -ConfirmDelete ``" -Color Cyan
            Write-LogAndHost "     -UseComplianceSearch -AdminUpn '$AdminUpn'" -Color Cyan
            
            Write-Log "`n2. Delete emails from a specific year (e.g., all of 2023):" "INFO"
            Write-LogAndHost "   .\ExchangeMailboxManager.ps1 -MailboxAddress '$MailboxAddress' ``" -Color Cyan
            Write-LogAndHost "     -DateRangeStart '2023-01-01' -DateRangeEnd '2023-12-31' -ConfirmDelete ``" -Color Cyan
            Write-LogAndHost "     -UseComplianceSearch -AdminUpn '$AdminUpn'" -Color Cyan
            
            Write-Log "`n3. Delete emails from a specific 6-month period:" "INFO"
            Write-LogAndHost "   .\ExchangeMailboxManager.ps1 -MailboxAddress '$MailboxAddress' ``" -Color Cyan
            Write-LogAndHost "     -DateRangeStart '2023-01-01' -DateRangeEnd '2023-06-30' -ConfirmDelete ``" -Color Cyan
            Write-LogAndHost "     -UseComplianceSearch -AdminUpn '$AdminUpn'" -Color Cyan
            
            Write-Log "`n4. Delete old emails (older than 2 years) using Graph API:" "INFO"
            Write-LogAndHost "   .\ExchangeMailboxManager.ps1 -MailboxAddress '$MailboxAddress' ``" -Color Cyan
            Write-LogAndHost "     -OlderThanDays 730 -ConfirmDelete ``" -Color Cyan
            Write-LogAndHost "     -TenantId 'YOUR_TENANT_ID' -ClientId 'YOUR_CLIENT_ID' -ClientSecret 'YOUR_SECRET'" -Color Cyan
            
            Write-Log "`n5. Clean specific folder (e.g., Sent Items older than 1 year):" "INFO"
            Write-LogAndHost "   .\ExchangeMailboxManager.ps1 -MailboxAddress '$MailboxAddress' ``" -Color Cyan
            Write-LogAndHost "     -FolderPath 'Sent Items' -OlderThanDays 365 -ConfirmDelete ``" -Color Cyan
            Write-LogAndHost "     -TenantId 'YOUR_TENANT_ID' -ClientId 'YOUR_CLIENT_ID' -ClientSecret 'YOUR_SECRET'" -Color Cyan
            
            Write-Log "`n6. First run as DRY RUN (recommended - see what would be deleted):" "INFO"
            Write-LogAndHost "   .\ExchangeMailboxManager.ps1 -MailboxAddress '$MailboxAddress' ``" -Color Cyan
            Write-LogAndHost "     -DateRangeStart '2023-01-01' -DateRangeEnd '2023-12-31' -DryRun ``" -Color Cyan
            Write-LogAndHost "     -UseComplianceSearch -AdminUpn '$AdminUpn'" -Color Cyan
            
            Write-Log "`nIMPORTANT NOTES:" "WARNING"
            Write-Log "  • -ConfirmDelete switch is MANDATORY for actual deletion" "WARNING"
            Write-Log "  • Date ranges cannot exceed 365 days (1 year)" "WARNING"
            Write-Log "  • Use -DryRun first to see what would be deleted" "WARNING"
            Write-Log "  • Use -RemoveRetentionPolicy if retention policies block deletion" "WARNING"
            Write-Log "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━" "INFO"
        }
        
        return $true
    }
    catch {
        Write-Log "Error getting diagnostics: $($_.Exception.Message)" "ERROR"
        return $false
    }
}
#endregion

#region Graph API Functions
function Connect-GraphAPI {
    param($TenantId, $ClientId, $ClientSecret)
    
    Write-Log "Authenticating with Microsoft Graph..." "INFO"
    
    try {
        # Get access token
        $body = @{
            client_id     = $ClientId
            client_secret = $ClientSecret
            scope         = "https://graph.microsoft.com/.default"
            grant_type    = "client_credentials"
        }
        
        $tokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
        $response = Invoke-RestMethod -Uri $tokenUrl -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"
        
        $script:AccessToken = $response.access_token
        
        # Load required modules
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Mail)) {
            Write-Log "Installing Microsoft.Graph module..." "WARNING"
            Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
        }
        
        Import-Module Microsoft.Graph.Authentication -ErrorAction SilentlyContinue
        Import-Module Microsoft.Graph.Mail -ErrorAction SilentlyContinue
        Import-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue
        
        # Connect using token
        $secureToken = ConvertTo-SecureString $script:AccessToken -AsPlainText -Force
        Connect-MgGraph -AccessToken $secureToken -NoWelcome
        
        Write-Log "Graph API authentication successful" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Graph API authentication failed: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Get-FolderByPath {
    param($UserId, $Path)
    
    try {
        # Well-known folder IDs
        $wellKnownFolders = @{
            "Deleted Items" = "deleteditems"
            "Inbox" = "inbox"
            "Sent Items" = "sentitems"
            "Drafts" = "drafts"
        }
        
        if ($wellKnownFolders.ContainsKey($Path)) {
            $folderId = $wellKnownFolders[$Path]
            $folder = Get-MgUserMailFolder -UserId $UserId -MailFolderId $folderId -ErrorAction Stop
            return $folder
        }
        
        # Search by name
        $allFolders = Get-MgUserMailFolder -UserId $UserId -All
        $folder = $allFolders | Where-Object { $_.DisplayName -eq $Path } | Select-Object -First 1
        
        return $folder
    }
    catch {
        Write-Log "Error finding folder '$Path': $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function Remove-EmailsGraphAPI {
    param($UserId, $Folder, $CutoffDate, $DryRunMode, $BatchSize, $RangeStart, $RangeEnd)
    
    Write-Log "Processing folder: $($Folder.DisplayName)" "INFO"
    
    try {
        # Build filter based on date parameters
        if ($RangeStart -and $RangeEnd) {
            # Date range mode
            $startDateString = $RangeStart.ToString("yyyy-MM-ddTHH:mm:ssZ")
            $endDateString = $RangeEnd.ToString("yyyy-MM-ddTHH:mm:ssZ")
            $filter = "receivedDateTime ge $startDateString and receivedDateTime le $endDateString"
            Write-Log "  Date range: $($RangeStart.ToString('yyyy-MM-dd')) to $($RangeEnd.ToString('yyyy-MM-dd'))" "INFO"
        }
        else {
            # Older than mode (default)
            $cutoffDateString = $CutoffDate.ToString("yyyy-MM-ddTHH:mm:ssZ")
            $filter = "receivedDateTime lt $cutoffDateString"
        }
        
        # OPTIMIZED: Get count efficiently using $count parameter (no data transfer)
        try {
            $countUrl = "https://graph.microsoft.com/v1.0/users/$UserId/mailFolders/$($Folder.Id)/messages/`$count?`$filter=$([System.Web.HttpUtility]::UrlEncode($filter))"
            $messageCount = [int](Invoke-MgGraphRequest -Method GET -Uri $countUrl -ErrorAction Stop)
        }
        catch {
            # Fallback: Use minimal property fetch if $count fails
            try {
                $messages = Get-MgUserMailFolderMessage -UserId $UserId -MailFolderId $Folder.Id -Filter $filter -Top 999 -Property "id" -ErrorAction Stop
                $messageCount = ($messages | Measure-Object).Count
            }
            catch {
                Write-Log "  Error getting message count: $($_.Exception.Message)" "ERROR"
                return 0
            }
        }
        
        if ($messageCount -eq 0) {
            Write-Log "  No messages found matching criteria" "INFO"
            return 0
        }
        
        Write-Log "  Found $messageCount messages to process" "INFO"
        
        if ($DryRunMode) {
            Write-Log "  [DRY RUN] Would delete $messageCount messages" "WARNING"
            return $messageCount
        }
        
        # SAFETY CONFIRMATION - Added safeguard
        if ($messageCount -gt 0) {
            Write-Log "`n⚠️  WARNING: DELETION OPERATION ⚠️" "WARNING"
            Write-LogAndHost "`n⚠️  WARNING: DELETION OPERATION ⚠️" -Color Yellow
            Write-LogAndHost "You are about to delete $messageCount messages from '$($Folder.DisplayName)'" -Color Yellow
            Write-LogAndHost "Mailbox: $UserId" -Color Yellow
            if ($RangeStart -and $RangeEnd) {
                Write-LogAndHost "Date range: $($RangeStart.ToString('yyyy-MM-dd')) to $($RangeEnd.ToString('yyyy-MM-dd'))" -Color Yellow
            } else {
                Write-LogAndHost "Older than: $($CutoffDate.ToString('yyyy-MM-dd'))" -Color Yellow
            }
            
            $confirm = Read-HostWithLogging "`nType 'YES' to confirm deletion (or anything else to skip this folder)"
            
            if ($confirm -ne "YES") {
                Write-Log "  Operation cancelled by user - skipping folder" "WARNING"
                return 0
            }
            Write-Log "  User confirmed deletion - proceeding" "WARNING"
        }
        
        # OPTIMIZED: Process in batches with Graph API batch requests (up to 20 deletions per HTTP call)
        $deletedCount = 0
        $failedCount = 0
        $batchNum = 0
        $maxGraphBatchSize = 20  # Graph API batch request limit
        
        while ($deletedCount + $failedCount -lt $messageCount) {
            $batchNum++
            
            try {
                # OPTIMIZED: Fetch only IDs with increased batch size
                $fetchSize = [Math]::Min($BatchSize * 4, 200)  # Increase fetch size (4x)
                $messages = Get-MgUserMailFolderMessage -UserId $UserId -MailFolderId $Folder.Id -Filter $filter -Top $fetchSize -Property "id" -ErrorAction Stop
                
                if ($messages.Count -eq 0) {
                    break
                }
                
                # OPTIMIZED: Process deletions in Graph API batches (20 at a time)
                $messageIds = @($messages | Select-Object -ExpandProperty Id)
                
                for ($i = 0; $i -lt $messageIds.Count; $i += $maxGraphBatchSize) {
                    $endIndex = [Math]::Min($i + $maxGraphBatchSize - 1, $messageIds.Count - 1)
                    $batchIds = $messageIds[$i..$endIndex]
                    
                    # Build Graph API batch request
                    $batchRequests = @()
                    for ($j = 0; $j -lt $batchIds.Count; $j++) {
                        $batchRequests += @{
                            id     = "$j"
                            method = "DELETE"
                            url    = "/users/$UserId/messages/$($batchIds[$j])"
                        }
                    }
                    
                    $batchBody = @{
                        requests = $batchRequests
                    } | ConvertTo-Json -Depth 10
                    
                    # Execute batch deletion with retry logic
                    $retryCount = 0
                    $maxRetries = 3
                    $batchSuccess = $false
                    
                    while (-not $batchSuccess -and $retryCount -lt $maxRetries) {
                        try {
                            $batchResponse = Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/`$batch" -Body $batchBody -ContentType "application/json" -ErrorAction Stop
                            
                            # Count successes and failures
                            foreach ($response in $batchResponse.responses) {
                                if ($response.status -ge 200 -and $response.status -lt 300) {
                                    $deletedCount++
                                } else {
                                    $failedCount++
                                    if ($response.status -eq 429 -or $response.status -eq 503) {
                                        Write-Log "    Warning: Throttling detected in batch response (status: $($response.status))" "WARNING"
                                    }
                                }
                            }
                            $batchSuccess = $true
                        }
                        catch {
                            if ($_.Exception.Message -like "*429*" -or $_.Exception.Message -like "*throttl*") {
                                $retryCount++
                                $waitTime = [Math]::Pow(2, $retryCount) * 2
                                Write-Log "    Throttled - waiting $waitTime seconds (retry $retryCount/$maxRetries)..." "WARNING"
                                Start-Sleep -Seconds $waitTime
                            }
                            elseif ($_.Exception.Message -like "*403*" -or $_.Exception.Message -like "*quota*") {
                                Write-Log "    Quota/permission error - consider using -UseComplianceSearch" "ERROR"
                                $failedCount += $batchIds.Count
                                $batchSuccess = $true
                            }
                            else {
                                Write-Log "    Batch deletion error: $($_.Exception.Message)" "ERROR"
                                $failedCount += $batchIds.Count
                                $batchSuccess = $true
                            }
                        }
                    }
                    
                    if (-not $batchSuccess) {
                        $failedCount += $batchIds.Count
                        Write-Log "    Failed to delete batch after $maxRetries retries" "ERROR"
                    }
                }
                
                $percentComplete = [Math]::Round((($deletedCount + $failedCount) / $messageCount) * 100, 1)
                Write-Log "    Batch $batchNum : $deletedCount deleted, $failedCount failed ($percentComplete%)" "INFO"
                
                # Reduced sleep time since we're using batch operations
                Start-Sleep -Milliseconds 200
            }
            catch {
                Write-Log "  Batch error: $($_.Exception.Message)" "ERROR"
                break
            }
        }
        
        Write-Log "  Completed: $deletedCount deleted, $failedCount failed" "SUCCESS"
        return $deletedCount
    }
    catch {
        Write-Log "Error processing folder: $($_.Exception.Message)" "ERROR"
        return 0
    }
}
#endregion

#region Compliance Search Functions
function Connect-ComplianceCenter {
    param([string]$AdminUpn)
    
    Write-Log "Connecting to Microsoft Purview Compliance Center..." "INFO"
    
    try {
        # Check if already connected by testing command execution
        $complianceConnected = $false
        try {
            $null = Get-ComplianceSearch -ErrorAction Stop | Select-Object -First 1
            $complianceConnected = $true
        }
        catch {
            $complianceConnected = $false
        }
        
        if ($complianceConnected) {
            Write-Log "Already connected to Compliance Center" "SUCCESS"
            return $true
        }
        
        # Install module if needed
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            Write-Log "Installing ExchangeOnlineManagement module..." "WARNING"
            Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force
        }
        
        Import-Module ExchangeOnlineManagement
        
        # Connect
        Connect-IPPSSession -UserPrincipalName $AdminUpn -ShowBanner:$false
        Connect-ExchangeOnline -UserPrincipalName $AdminUpn -ShowBanner:$false
        
        Write-Log "Compliance Center connection successful" "SUCCESS"
        return $true
    }
    catch {
        Write-Log "Failed to connect to Compliance Center: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Remove-RetentionSettings {
    param([string]$MailboxAddress)
    
    Write-Log "Removing retention policies and holds..." "WARNING"
    
    try {
        $mailbox = Get-Mailbox $MailboxAddress
        
        # SAFETY CONFIRMATION - Added safeguard
        Write-Log "`n⚠️  WARNING: RETENTION POLICY REMOVAL ⚠️" "ERROR"
        Write-LogAndHost "`n⚠️  WARNING: RETENTION POLICY REMOVAL ⚠️" -Color Red
        Write-LogAndHost "This will remove ALL retention policies and holds from: $MailboxAddress" -Color Red
        Write-LogAndHost "Settings that will be modified:" -Color Yellow
        Write-LogAndHost "  - Retention Policy: $($mailbox.RetentionPolicy)" -Color Yellow
        Write-LogAndHost "  - Litigation Hold: $($mailbox.LitigationHoldEnabled)" -Color Yellow
        Write-LogAndHost "  - Single Item Recovery: $($mailbox.SingleItemRecoveryEnabled)" -Color Yellow
        Write-LogAndHost "  - Deleted Items Retention: $($mailbox.RetainDeletedItemsFor) -> 0 days" -Color Yellow
        Write-LogAndHost "`nThis may violate compliance requirements!" -Color Red
        
        $confirm = Read-HostWithLogging "`nType 'REMOVE' to confirm retention policy removal (or anything else to cancel)"
        
        if ($confirm -ne "REMOVE") {
            Write-Log "Retention policy removal cancelled by user" "WARNING"
            return $false
        }
        Write-Log "User confirmed retention policy removal - proceeding" "WARNING"
        
        # Remove retention policy
        if ($mailbox.RetentionPolicy) {
            Write-Log "  Removing retention policy: $($mailbox.RetentionPolicy)" "INFO"
            Set-Mailbox $MailboxAddress -RetentionPolicy $null
        }
        
        # Disable holds
        if ($mailbox.LitigationHoldEnabled) {
            Write-Log "  Disabling litigation hold" "INFO"
            Set-Mailbox $MailboxAddress -LitigationHoldEnabled $false
        }
        
        # Note: Delay holds are system-managed and cannot be directly removed
        # They expire automatically after 30 days when retention policies are removed
        if ($mailbox.DelayHoldApplied -or $mailbox.DelayReleaseHoldApplied) {
            Write-Log "  Delay holds detected - these will expire automatically in ~30 days" "WARNING"
            Write-Log "  Delay holds cannot be manually removed and are system-managed" "INFO"
        }
        
        # Set retention to 0
        Write-Log "  Setting retention to 0 days" "INFO"
        Set-Mailbox $MailboxAddress -RetainDeletedItemsFor 0
        
        # Disable single item recovery
        Write-Log "  Disabling single item recovery" "INFO"
        Set-Mailbox $MailboxAddress -SingleItemRecoveryEnabled $false
        
        # Force managed folder assistant
        Write-Log "  Running Managed Folder Assistant" "INFO"
        Start-ManagedFolderAssistant -Identity $MailboxAddress
        
        Write-Log "Retention settings removed successfully" "SUCCESS"
        Write-Log "Waiting 5 minutes for changes to propagate..." "INFO"
        Start-Sleep -Seconds 300
        
        return $true
    }
    catch {
        Write-Log "Error removing retention settings: $($_.Exception.Message)" "ERROR"
        return $false
    }
}

function Start-ComplianceSearchCleanup {
    param($MailboxAddress, $CutoffDate, $DryRunMode, $WaitForCompletion, $RangeStart, $RangeEnd)
    
    $searchName = "AutoCleanup_$($MailboxAddress.Replace('@','_').Replace('.','_'))_$(Get-Date -Format 'HHmmss')"
    
    Write-Log "Creating compliance search: $searchName" "INFO"
    
    try {
        # Build query based on date parameters
        if ($RangeStart -and $RangeEnd) {
            # Date range mode
            $startDateString = $RangeStart.ToString("yyyy-MM-dd")
            $endDateString = $RangeEnd.ToString("yyyy-MM-dd")
            $query = "kind:email AND received>=$startDateString AND received<=$endDateString"
            Write-Log "  Date range: $startDateString to $endDateString" "INFO"
        }
        else {
            # Older than mode (default)
            $cutoffDateString = $CutoffDate.ToString("yyyy-MM-dd")
            $query = "kind:email AND received<$cutoffDateString"
        }
        
        Write-Log "  Query: $query" "INFO"
        
        # Create search
        New-ComplianceSearch -Name $searchName `
            -ExchangeLocation $MailboxAddress `
            -ContentMatchQuery $query | Out-Null
        
        # Start search
        Write-Log "Starting compliance search..." "INFO"
        Start-ComplianceSearch -Identity $searchName
        
        # Wait for completion
        $maxWaitMinutes = if ($WaitForCompletion) { 120 } else { 10 }
        $waitedMinutes = 0
        
        do {
            Start-Sleep -Seconds 30
            $search = Get-ComplianceSearch -Identity $searchName
            $waitedMinutes += 0.5
            
            if ($search.Status -eq "Completed") {
                break
            }
            
            if ($waitedMinutes % 2 -eq 0) {
                Write-Log "  Search status: $($search.Status) - waited $waitedMinutes minutes" "INFO"
            }
        } while ($search.Status -ne "Completed" -and $waitedMinutes -lt $maxWaitMinutes)
        
        if ($search.Status -ne "Completed") {
            Write-Log "Search is still running. Check status with: Get-ComplianceSearch -Identity '$searchName'" "WARNING"
            
            if (-not $WaitForCompletion) {
                Write-Log "Use -ForceWait to wait for completion (can take hours for large mailboxes)" "INFO"
                return $false
            }
        }
        
        $itemCount = $search.Items
        $totalSize = $search.Size
        
        Write-Log "Search completed: $itemCount items found ($totalSize bytes)" "SUCCESS"
        
        if ($itemCount -eq 0) {
            Write-Log "No items found to delete" "INFO"
            Remove-ComplianceSearch -Identity $searchName -Confirm:$false
            return 0
        }
        
        if ($DryRunMode) {
            Write-Log "[DRY RUN] Would permanently delete $itemCount items" "WARNING"
            Remove-ComplianceSearch -Identity $searchName -Confirm:$false
            return $itemCount
        }
        
        # Confirm deletion
        Write-Log "`n⚠️  WARNING: PERMANENT DELETION ⚠️" "ERROR"
        Write-LogAndHost "`n⚠️  WARNING: PERMANENT DELETION ⚠️" -Color Red
        Write-LogAndHost "This will PERMANENTLY delete $itemCount items from $MailboxAddress" -Color Red
        Write-LogAndHost "Items cannot be recovered after deletion" -Color Red
        
        $confirm = Read-HostWithLogging "`nType 'DELETE' to confirm permanent deletion (or anything else to cancel)"
        
        if ($confirm -ne "DELETE") {
            Write-Log "Operation cancelled by user" "WARNING"
            Remove-ComplianceSearch -Identity $searchName -Confirm:$false
            return 0
        }
        Write-Log "User confirmed permanent deletion - proceeding" "WARNING"
        
        # Create purge action
        Write-Log "Creating purge action..." "WARNING"
        New-ComplianceSearchAction -SearchName $searchName `
            -Purge -PurgeType HardDelete -Confirm:$false | Out-Null
        
        Write-Log "Purge action created successfully" "SUCCESS"
        Write-Log "`nIMPORTANT: Purge processing happens in the background" "WARNING"
        Write-Log "This can take 1-48 hours to complete for large mailboxes" "WARNING"
        Write-Log "`nMonitor progress with:" "INFO"
        Write-Log "  Get-ComplianceSearchAction -Identity '${searchName}_Purge' | Select Name, Status, Results" "INFO"
        Write-Log "`nCheck mailbox after 24 hours:" "INFO"
        Write-Log "  Get-MailboxFolderStatistics -Identity '$MailboxAddress' | Select Name, ItemsInFolder" "INFO"
        
        return $itemCount
    }
    catch {
        Write-Log "Error in compliance search: $($_.Exception.Message)" "ERROR"
        return 0
    }
}
#endregion

#region Mailbox Reporting Functions
function Get-ComprehensiveMailboxReport {
    param(
        [string]$MailboxAddress,
        [bool]$AllMailboxes = $false
    )
    
    Write-Banner "Comprehensive Mailbox Report"
    
    try {
        # Ensure Exchange Online is connected - reuse existing connection
        $exoConnected = $false
        try {
            $null = Get-OrganizationConfig -ErrorAction Stop
            $exoConnected = $true
        }
        catch {
            $exoConnected = $false
        }
        
        if (-not $exoConnected) {
            Write-Log "Connecting to Exchange Online..." "INFO"
            Connect-ExchangeOnline -UserPrincipalName $AdminUpn -ShowBanner:$false
        }
        else {
            Write-Log "Using existing Exchange Online connection" "INFO"
        }
        
        $Report = @()
        
        # Get mailboxes
        if ($AllMailboxes) {
            Write-Log "Retrieving all mailboxes..." "INFO"
            $Mailboxes = Get-EXOMailbox -ResultSize Unlimited | 
                Where-Object {$_.RecipientTypeDetails -eq 'UserMailbox' -or $_.RecipientTypeDetails -eq 'SharedMailbox'}
            Write-Log "Found $($Mailboxes.Count) mailboxes to process" "SUCCESS"
        }
        else {
            Write-Log "Retrieving mailbox: $MailboxAddress" "INFO"
            $Mailboxes = @(Get-EXOMailbox -Identity $MailboxAddress)
        }
        
        $counter = 0
        foreach ($Mailbox in $Mailboxes) {
            $counter++
            
            if (-not $Mailbox.UserPrincipalName) { 
                Write-Log "Skipping mailbox with no UPN: $($Mailbox.DisplayName)" "WARNING"
                continue 
            }
            
            Write-Progress -Activity "Processing Mailboxes" `
                -Status "[$counter/$($Mailboxes.Count)] $($Mailbox.DisplayName)" `
                -PercentComplete (($counter / $Mailboxes.Count) * 100)
            
            try {
                # Get overall statistics
                Write-Log "  Processing: $($Mailbox.DisplayName)" "INFO"
                $Stats = Get-EXOMailboxStatistics -Identity $Mailbox.UserPrincipalName -ErrorAction Stop
                
                # Get mailbox properties for additional info - moved earlier to avoid duplicate call
                $MailboxDetails = Get-EXOMailbox -Identity $Mailbox.UserPrincipalName -ErrorAction Stop
                
                # Get folder statistics and find oldest item (only retrieve what we need for performance)
                $FolderStats = Get-EXOMailboxFolderStatistics -Identity $Mailbox.UserPrincipalName `
                    -IncludeOldestAndNewestItems `
                    -ErrorAction Stop
                
                $OldestItem = $FolderStats | 
                    Where-Object {$null -ne $_.OldestItemReceivedDate} |
                    Sort-Object OldestItemReceivedDate |
                    Select-Object -First 1
                
                # Get newest item
                $NewestItem = $FolderStats | 
                    Where-Object {$null -ne $_.NewestItemReceivedDate} |
                    Sort-Object NewestItemReceivedDate -Descending |
                    Select-Object -First 1
                
                # Get last activity - use LastUserActionTime (more reliable in Exchange Online)
                $LastActivityDate = $null
                $LastActivitySource = ""
                
                if ($Stats.LastUserActionTime) {
                    $LastActivityDate = $Stats.LastUserActionTime
                    $LastActivitySource = "Last Activity"
                }
                elseif ($Stats.LastLogonTime) {
                    $LastActivityDate = $Stats.LastLogonTime
                    $LastActivitySource = "Last Logon"
                }
                elseif ($MailboxDetails.WhenMailboxCreated) {
                    $LastActivityDate = $MailboxDetails.WhenMailboxCreated
                    $LastActivitySource = "Created"
                }
                
                # Calculate days since last activity
                $DaysSinceActivity = if ($LastActivityDate) {
                    [math]::Round(((Get-Date) - $LastActivityDate).TotalDays, 0)
                } else {
                    "Unknown"
                }
                
                # Get largest folders with robust error handling - exclude only truly system folders
                $ExcludedFolders = @('Top of Information Store', 'Root', 'RSS Feeds', 'Quick Step Settings', 
                                     'Conversation Action Settings', 'Calendar Logging', 
                                     'Deletions', 'Purges', 'Versions', 'DiscoveryHolds', 'SubstrateHolds')
                
                $LargestFolders = $FolderStats | 
                    Where-Object {$_.ItemsInFolder -gt 0 -and $_.Name -notin $ExcludedFolders} |
                    ForEach-Object {
                        $sizeBytes = ConvertTo-Bytes -Size $_.FolderSize
                        
                        [PSCustomObject]@{
                            Folder = $_
                            SizeBytes = $sizeBytes
                        }
                    } |
                    Where-Object {$_.SizeBytes -gt 1024} |  # At least 1KB to filter out empty
                    Sort-Object SizeBytes -Descending |
                    Select-Object -First 3 |
                    ForEach-Object { $_.Folder }
                
                # Safely handle null or empty array
                if (-not $LargestFolders) {
                    $LargestFolders = @()
                }
                
                $TopFolder1 = if ($LargestFolders.Count -gt 0 -and $LargestFolders[0]) { 
                    try {
                        $folderName = $LargestFolders[0].Name
                        $folderSize = if ($LargestFolders[0].FolderSize) { $LargestFolders[0].FolderSize.ToString() } else { "0 bytes" }
                        "$folderName ($folderSize)"
                    } catch { "N/A" }
                } else { "N/A" }
                
                $TopFolder2 = if ($LargestFolders.Count -gt 1 -and $LargestFolders[1]) { 
                    try {
                        $folderName = $LargestFolders[1].Name
                        $folderSize = if ($LargestFolders[1].FolderSize) { $LargestFolders[1].FolderSize.ToString() } else { "0 bytes" }
                        "$folderName ($folderSize)"
                    } catch { "N/A" }
                } else { "N/A" }
                
                $TopFolder3 = if ($LargestFolders.Count -gt 2 -and $LargestFolders[2]) { 
                    try {
                        $folderName = $LargestFolders[2].Name
                        $folderSize = if ($LargestFolders[2].FolderSize) { $LargestFolders[2].FolderSize.ToString() } else { "0 bytes" }
                        "$folderName ($folderSize)"
                    } catch { "N/A" }
                } else { "N/A" }
                
                # Get folders with most items
                $BusiestFolders = $FolderStats | 
                    Where-Object {$_.ItemsInFolder -gt 0} |
                    Sort-Object ItemsInFolder -Descending |
                    Select-Object -First 3
                
                # Safely handle null or empty array
                if (-not $BusiestFolders) {
                    $BusiestFolders = @()
                }
                
                # Calculate quota usage percentage with robust error handling
                $QuotaUsagePercent = "N/A"
                $QuotaDisplay = "50 GB (Default)"
                
                # Get default quota based on mailbox type
                $defaultQuotaGB = switch ($MailboxDetails.RecipientTypeDetails) {
                    "SharedMailbox" { 50 }
                    "RoomMailbox" { 50 }
                    "EquipmentMailbox" { 50 }
                    default { 50 }  # Standard Exchange Online is 50GB
                }
                
                try {
                    if ($Stats.TotalItemSize) {
                        $sizeBytes = ConvertTo-Bytes -Size $Stats.TotalItemSize
                        $currentSizeGB = [math]::Round($sizeBytes / 1GB, 2)
                        
                        # Check if quota is set
                        if ($MailboxDetails.ProhibitSendQuota) {
                            $quotaString = $MailboxDetails.ProhibitSendQuota.ToString()
                            
                            if ($quotaString -eq "Unlimited" -or $quotaString -like "*unlimited*" -or [string]::IsNullOrEmpty($quotaString)) {
                                # Unlimited quota - show default and calculate percentage against default
                                $QuotaDisplay = "$defaultQuotaGB GB (Default)"
                                $QuotaUsagePercent = [math]::Round(($currentSizeGB / $defaultQuotaGB) * 100, 2)
                            }
                            else {
                                # Specific quota set - extract and use it
                                $currentSize = ConvertTo-Bytes -Size $Stats.TotalItemSize
                                $quota = ConvertTo-Bytes -Size $MailboxDetails.ProhibitSendQuota
                                
                                if ($quota -gt 0) {
                                    $QuotaDisplay = $MailboxDetails.ProhibitSendQuota.ToString()
                                    $QuotaUsagePercent = [math]::Round(($currentSize / $quota) * 100, 2)
                                }
                                else {
                                    # Fall back to default
                                    $QuotaDisplay = "$defaultQuotaGB GB (Default)"
                                    $QuotaUsagePercent = [math]::Round(($currentSizeGB / $defaultQuotaGB) * 100, 2)
                                }
                            }
                        }
                        else {
                            # No quota property - use default
                            $QuotaDisplay = "$defaultQuotaGB GB (Default)"
                            $QuotaUsagePercent = [math]::Round(($currentSizeGB / $defaultQuotaGB) * 100, 2)
                        }
                    }
                }
                catch {
                    $QuotaUsagePercent = "N/A"
                    $QuotaDisplay = "$defaultQuotaGB GB (Default)"
                }
                
                # Create comprehensive report object
                $totalSizeBytes = ConvertTo-Bytes -Size $Stats.TotalItemSize
                $ReportObject = [PSCustomObject]@{
                    DisplayName              = $Mailbox.DisplayName
                    PrimarySmtpAddress       = $Mailbox.PrimarySmtpAddress
                    RecipientType            = $Mailbox.RecipientTypeDetails
                    ItemCount                = $Stats.ItemCount
                    TotalSizeGB              = [math]::Round($totalSizeBytes / 1GB, 3)
                    TotalSizeMB              = [math]::Round($totalSizeBytes / 1MB, 2)
                    DeletedItemCount         = $Stats.DeletedItemCount
                    LastActivityTime         = if ($LastActivityDate) { $LastActivityDate } else { "Unknown" }
                    LastActivitySource       = $LastActivitySource
                    DaysSinceLastActivity    = $DaysSinceActivity
                    OldestItemDate           = $OldestItem.OldestItemReceivedDate
                    OldestItemFolder         = $OldestItem.Name
                    NewestItemDate           = $NewestItem.NewestItemReceivedDate
                    NewestItemFolder         = $NewestItem.Name
                    MailboxDatabase          = $Stats.DatabaseName
                    ProhibitSendQuota        = $QuotaDisplay
                    QuotaUsagePercent        = $QuotaUsagePercent
                    RetentionPolicy          = if ($MailboxDetails.RetentionPolicy) { $MailboxDetails.RetentionPolicy } else { "None" }
                    LitigationHoldEnabled    = if ($MailboxDetails.LitigationHoldEnabled) { "Yes" } else { "No" }
                    InPlaceHolds             = if ($MailboxDetails.InPlaceHolds) { ($MailboxDetails.InPlaceHolds -join '; ') } else { "None" }
                    SingleItemRecovery       = if ($MailboxDetails.SingleItemRecoveryEnabled) { "Yes" } else { "No" }
                    TopFolder1               = $TopFolder1
                    TopFolder2               = $TopFolder2
                    TopFolder3               = $TopFolder3
                    BusiestFolder1           = if ($BusiestFolders.Count -gt 0 -and $BusiestFolders[0]) { 
                        try {
                            "$($BusiestFolders[0].Name) ($($BusiestFolders[0].ItemsInFolder))" 
                        } catch { "N/A" }
                    } else { "N/A" }
                    BusiestFolder2           = if ($BusiestFolders.Count -gt 1 -and $BusiestFolders[1]) { 
                        try {
                            "$($BusiestFolders[1].Name) ($($BusiestFolders[1].ItemsInFolder))" 
                        } catch { "N/A" }
                    } else { "N/A" }
                    BusiestFolder3           = if ($BusiestFolders.Count -gt 2 -and $BusiestFolders[2]) { 
                        try {
                            "$($BusiestFolders[2].Name) ($($BusiestFolders[2].ItemsInFolder))" 
                        } catch { "N/A" }
                    } else { "N/A" }
                    ArchiveStatus            = if ($MailboxDetails.ArchiveStatus) { $MailboxDetails.ArchiveStatus } else { "None" }
                }
                
                $Report += $ReportObject
                Write-Log "    ✓ $($Mailbox.DisplayName): $($Stats.ItemCount) items, $([math]::Round($totalSizeBytes / 1GB, 2)) GB" "SUCCESS"
            }
            catch {
                Write-Log "    ✗ Error processing $($Mailbox.DisplayName): $($_.Exception.Message)" "ERROR"
                Write-Log "    Error details: $($_.ScriptStackTrace)" "ERROR"
            }
        }
        
        Write-Progress -Activity "Processing Mailboxes" -Completed
        
        # Display report
        if ($Report.Count -gt 0) {
            Write-Banner "Report Summary"
            
            # Sort report by size (DESC) for all outputs
            $Report = $Report | Sort-Object TotalSizeGB -Descending
            
            # Export to CSV (sorted by size DESC)
            $CsvFile = "MailboxReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            try {
                $Report | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8
                Write-Log "✅ CSV Report exported: $CsvFile" "SUCCESS"
            }
            catch {
                Write-Log "⚠️  Failed to export CSV: $($_.Exception.Message)" "WARNING"
            }
            
            # Display summary statistics
            Write-Log "Total Mailboxes Processed: $($Report.Count)" "INFO"
            Write-Log "Total Items: $(($Report | Measure-Object -Property ItemCount -Sum).Sum)" "INFO"
            Write-Log "Total Size (GB): $([math]::Round(($Report | Measure-Object -Property TotalSizeGB -Sum).Sum, 2))" "INFO"
            Write-Log "Average Items per Mailbox: $([math]::Round(($Report | Measure-Object -Property ItemCount -Average).Average, 0))" "INFO"
            Write-Log "Average Size per Mailbox (GB): $([math]::Round(($Report | Measure-Object -Property TotalSizeGB -Average).Average, 2))" "INFO"
            
            # Display summary table - 1 row per mailbox (sorted by size, highest first)
            Write-Log "`n" "INFO"
            
            # Build table lines
            $tableLines = @()
            $tableLines += "`n╔═══════════════════════════════════════════════════════════════════════════════════════════════════════════════╗"
            $tableLines += "║                                    MAILBOX SUMMARY REPORT (Sorted by Size)                                   ║"
            $tableLines += "╠═══════════════════════════════════════════════════════════════════════════════════════════════════════════════╣"
            $tableLines += "║ Mailbox                                        │   Size   │    Items    │  Quota  │ Last Active│ Deleted   ║"
            $tableLines += "╠════════════════════════════════════════════════╪══════════╪═════════════╪═════════╪════════════╪═══════════╣"
            
            foreach ($mb in ($Report | Sort-Object TotalSizeGB -Descending)) {
                $emailPart = "$($mb.DisplayName) <$($mb.PrimarySmtpAddress)>"
                $sizePart = "$($mb.TotalSizeGB) GB"
                $itemsPart = "$($mb.ItemCount)"
                $quotaPart = if ($mb.QuotaUsagePercent -ne "N/A") { "$($mb.QuotaUsagePercent)%" } else { "N/A" }
                $lastActivityPart = if ($mb.DaysSinceLastActivity -ne "Unknown") { "$($mb.DaysSinceLastActivity)d" } else { "Unknown" }
                $deletedPart = "$($mb.DeletedItemCount)"
                
                # Format: Name | Size | Items | Quota | Last Activity | Deleted
                $line = "║ {0,-46} │ {1,8} │ {2,11} │ {3,7} │ {4,10} │ {5,9} ║" -f `
                    ($emailPart.Substring(0, [Math]::Min(46, $emailPart.Length))), `
                    $sizePart, `
                    $itemsPart, `
                    $quotaPart, `
                    $lastActivityPart, `
                    $deletedPart
                
                $tableLines += $line
            }
            
            $tableLines += "╚═══════════════════════════════════════════════════════════════════════════════════════════════════════════════╝"
            
            # Output to both console and log file
            foreach ($line in $tableLines) {
                Write-Host $line -ForegroundColor Cyan
                Add-Content -Path $LogFile -Value $line
            }
            
            Write-Host ""
            Write-Log "Summary table shows mailboxes sorted by size (highest to lowest)" "INFO"
            
            # Display detailed information for each mailbox
            Write-Log "`nDetailed Mailbox Information:" "INFO"
            foreach ($mb in ($Report | Sort-Object TotalSizeGB -Descending)) {
                $separator = "`n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
                $mbHeader = "  $($mb.DisplayName) ($($mb.PrimarySmtpAddress))"
                
                # Display to console with colors
                Write-Host $separator -ForegroundColor Cyan
                Write-Host $mbHeader -ForegroundColor Yellow
                Write-Host $separator -ForegroundColor Cyan
                
                # Write to log file
                Add-Content -Path $LogFile -Value $separator
                Add-Content -Path $LogFile -Value $mbHeader
                Add-Content -Path $LogFile -Value $separator
                
                # Build all detail lines
                # Format Last Activity display
                $lastActivityDisplay = if ($mb.LastActivityTime -ne "Unknown") {
                    "$($mb.LastActivitySource): $($mb.LastActivityTime) ($($mb.DaysSinceLastActivity) days ago)"
                } else {
                    "No activity recorded"
                }
                
                $details = @(
                    "  Type: $($mb.RecipientType)",
                    "  Items: $($mb.ItemCount) | Size: $($mb.TotalSizeGB) GB ($($mb.TotalSizeMB) MB)",
                    "  Deleted Items: $($mb.DeletedItemCount)",
                    "  Last Activity: $lastActivityDisplay",
                    "  Oldest Item: $($mb.OldestItemDate) in folder '$($mb.OldestItemFolder)'",
                    "  Newest Item: $($mb.NewestItemDate) in folder '$($mb.NewestItemFolder)'",
                    "  Quota: $($mb.ProhibitSendQuota) | Usage: $($mb.QuotaUsagePercent)%",
                    "  Retention Policy: $($mb.RetentionPolicy)",
                    "  Litigation Hold: $($mb.LitigationHoldEnabled)",
                    "  Single Item Recovery: $($mb.SingleItemRecovery)",
                    "  Archive Status: $($mb.ArchiveStatus)",
                    "  Top Folders by Size:",
                    "    1. $($mb.TopFolder1)",
                    "    2. $($mb.TopFolder2)",
                    "    3. $($mb.TopFolder3)",
                    "  Top Folders by Item Count:",
                    "    1. $($mb.BusiestFolder1)",
                    "    2. $($mb.BusiestFolder2)",
                    "    3. $($mb.BusiestFolder3)"
                )
                
                # Display to console with colors
                foreach ($line in $details) {
                    if ($line -like "*Top Folders*") {
                        Write-Host $line -ForegroundColor Cyan
                    } else {
                        Write-Host $line -ForegroundColor White
                    }
                }
                
                # Write all details to log file
                foreach ($line in $details) {
                    Add-Content -Path $LogFile -Value $line
                }
            }
            
            Write-Log "`n✅ Report generation complete!" "SUCCESS"
        }
        else {
            Write-Log "No mailboxes were processed successfully" "WARNING"
        }
        
        return $Report.Count
    }
    catch {
        Write-Log "Error generating comprehensive report: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function Get-DetailedFolderReport {
    param([string]$MailboxAddress)
    
    Write-Banner "Detailed Folder Statistics Report"
    
    try {
        # Ensure Exchange Online is connected - reuse existing connection
        $exoConnected = $false
        try {
            $null = Get-OrganizationConfig -ErrorAction Stop
            $exoConnected = $true
        }
        catch {
            $exoConnected = $false
        }
        
        if (-not $exoConnected) {
            Write-Log "Connecting to Exchange Online..." "INFO"
            Connect-ExchangeOnline -UserPrincipalName $AdminUpn -ShowBanner:$false
        }
        else {
            Write-Log "Using existing Exchange Online connection" "INFO"
        }
        
        Write-Log "Retrieving folder statistics for: $MailboxAddress" "INFO"
        
        # Get mailbox verification
        $Mailbox = Get-EXOMailbox -Identity $MailboxAddress -ErrorAction Stop
        Write-Log "Mailbox found: $($Mailbox.DisplayName)" "SUCCESS"
        
        # Get detailed folder statistics
        $FolderStats = Get-EXOMailboxFolderStatistics -Identity $MailboxAddress `
            -IncludeOldestAndNewestItems -ErrorAction Stop
        
        # Get overall mailbox statistics
        $MailboxStats = Get-EXOMailboxStatistics -Identity $MailboxAddress -ErrorAction Stop
        
        # Filter and prepare folder data with robust size parsing
        $FolderReport = $FolderStats | 
            Where-Object {$_.ItemsInFolder -gt 0} |
            Select-Object @{Name="MailboxOwner";Expression={$Mailbox.DisplayName}},
                          @{Name="MailboxAddress";Expression={$MailboxAddress}},
                          Name,
                          FolderPath,
                          FolderType,
                          ItemsInFolder,
                          @{Name="FolderSizeMB";Expression={
                              try {
                                  $sizeString = $_.FolderSize.ToString()
                                  # Parse size string like "1.5 GB (1,610,612,736 bytes)"
                                  if ($sizeString -match '\(([0-9,]+) bytes\)') {
                                      $bytes = [long]($matches[1] -replace ',','')
                                      $sizeMB = [math]::Round($bytes / 1MB, 2)
                                      # Ensure period as decimal separator
                                      [double]::Parse($sizeMB.ToString([System.Globalization.CultureInfo]::InvariantCulture))
                                  }
                                  else {
                                      0
                                  }
                              }
                              catch {
                                  0
                              }
                          }},
                          @{Name="FolderSizeGB";Expression={
                              try {
                                  $sizeString = $_.FolderSize.ToString()
                                  # Parse size string like "1.5 GB (1,610,612,736 bytes)"
                                  if ($sizeString -match '\(([0-9,]+) bytes\)') {
                                      $bytes = [long]($matches[1] -replace ',','')
                                      $sizeGB = [math]::Round($bytes / 1GB, 3)
                                      # Ensure period as decimal separator
                                      [double]::Parse($sizeGB.ToString([System.Globalization.CultureInfo]::InvariantCulture))
                                  }
                                  else {
                                      0
                                  }
                              }
                              catch {
                                  0
                              }
                          }},
                          FolderSize,
                          DeletedItemsInFolder,
                          OldestItemReceivedDate,
                          NewestItemReceivedDate,
                          @{Name="FolderAgeInDays";Expression={
                              if ($_.OldestItemReceivedDate) {
                                  [int][math]::Round(((Get-Date) - $_.OldestItemReceivedDate).TotalDays, 0)
                              } else {
                                  $null
                              }
                          }}
        
        # Export to CSV (sorted by size DESC)
        $FolderCsvFile = "FolderReport_$($MailboxAddress.Replace('@','_').Replace('.','_'))_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        try {
            $FolderReport | Export-Csv -Path $FolderCsvFile -NoTypeInformation -Encoding UTF8
            Write-Log "✅ Folder CSV Report exported: $FolderCsvFile" "SUCCESS"
        }
        catch {
            Write-Log "⚠️  Failed to export folder CSV: $($_.Exception.Message)" "WARNING"
        }
        
        # Display summary
        Write-Banner "Folder Statistics Summary"
        
        Write-Log "Mailbox: $($Mailbox.DisplayName) ($MailboxAddress)" "INFO"
        Write-Log "Total Mailbox Items: $($MailboxStats.ItemCount)" "INFO"
        Write-Log "Total Mailbox Size: $($MailboxStats.TotalItemSize)" "INFO"
        Write-Log "Number of Folders with Items: $($FolderReport.Count)" "INFO"
        
        # All folders sorted by FolderPath
        Write-Log "`nAll Folders (sorted by path):" "INFO"
        $tableOutput = $FolderReport | 
            Sort-Object FolderPath |
            Format-Table Name, FolderPath, ItemsInFolder, 
                @{Label="Size(MB)";Expression={$_.FolderSizeMB}},
                OldestItemReceivedDate,
                @{Label="Days Since Oldest";Expression={$_.FolderAgeInDays}} -AutoSize |
            Out-String
        Write-Host $tableOutput
        Add-Content -Path $LogFile -Value $tableOutput
        
        # Top folders by size
        Write-Log "`nTop Folders by Size:" "INFO"
        $tableOutput = $FolderReport | 
            Sort-Object FolderSizeMB -Descending |
            Select-Object -First 20 |
            Format-Table Name, FolderPath, ItemsInFolder, 
                @{Label="Size(MB)";Expression={$_.FolderSizeMB}},
                @{Label="Size(GB)";Expression={$_.FolderSizeGB}} -AutoSize |
            Out-String
        Write-Host $tableOutput
        Add-Content -Path $LogFile -Value $tableOutput
        
        # Folders with oldest items
        Write-Log "`nFolders with Oldest Items:" "INFO"
        $tableOutput = $FolderReport | 
            Where-Object {$null -ne $_.OldestItemReceivedDate} |
            Sort-Object OldestItemReceivedDate |
            Select-Object -First 15 |
            Format-Table Name, ItemsInFolder, OldestItemReceivedDate, 
                @{Label="Days Since Oldest";Expression={$_.FolderAgeInDays}},
                @{Label="Size(MB)";Expression={$_.FolderSizeMB}} -AutoSize |
            Out-String
        Write-Host $tableOutput
        Add-Content -Path $LogFile -Value $tableOutput
        
        # Detailed folder breakdown
        Write-Log "`nDetailed Folder Breakdown:" "INFO"
        foreach ($folder in $FolderReport) {
            $folderOutput = "`n  📁 $($folder.Name)`n"
            $folderOutput += "     Path: $($folder.FolderPath)`n"
            $folderOutput += "     Items: $($folder.ItemsInFolder) | Size: $($folder.FolderSizeMB) MB ($($folder.FolderSizeGB) GB)`n"
            
            if ($folder.DeletedItemsInFolder -gt 0) {
                $folderOutput += "     Deleted Items: $($folder.DeletedItemsInFolder)`n"
            }
            if ($folder.OldestItemReceivedDate) {
                $folderOutput += "     Oldest: $($folder.OldestItemReceivedDate) ($($folder.FolderAgeInDays) days old)`n"
            }
            if ($folder.NewestItemReceivedDate) {
                $folderOutput += "     Newest: $($folder.NewestItemReceivedDate)`n"
            }
            
            # Output to console with colors
            Write-Host "`n  📁 $($folder.Name)" -ForegroundColor Yellow
            Write-Host "     Path: $($folder.FolderPath)" -ForegroundColor Gray
            Write-Host "     Items: $($folder.ItemsInFolder) | Size: $($folder.FolderSizeMB) MB ($($folder.FolderSizeGB) GB)" -ForegroundColor White
            if ($folder.DeletedItemsInFolder -gt 0) {
                Write-Host "     Deleted Items: $($folder.DeletedItemsInFolder)" -ForegroundColor Red
            }
            if ($folder.OldestItemReceivedDate) {
                Write-Host "     Oldest: $($folder.OldestItemReceivedDate) ($($folder.FolderAgeInDays) days old)" -ForegroundColor Cyan
            }
            if ($folder.NewestItemReceivedDate) {
                Write-Host "     Newest: $($folder.NewestItemReceivedDate)" -ForegroundColor Green
            }
            
            # Log to file
            Add-Content -Path $LogFile -Value $folderOutput
        }
        
        Write-Log "`n✅ Detailed folder report complete!" "SUCCESS"
        
        return $FolderReport.Count
    }
    catch {
        Write-Log "Error generating detailed folder report: $($_.Exception.Message)" "ERROR"
        return $null
    }
}
#endregion

#region Help Display
function Show-Help {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "Exchange Mailbox Manager - Help" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "DESCRIPTION:" -ForegroundColor Yellow
    Write-Host "  Comprehensive Exchange Online mailbox management solution for cleanup and reporting"
    Write-Host ""
    Write-Host "COMMON PARAMETERS:" -ForegroundColor Yellow
    Write-Host "  -MailboxAddress <email>      Target mailbox email address"
    Write-Host "  -AdminUpn <email>            Admin UPN for Compliance Search operations"
    Write-Host "  -FolderPath <path>           Specific folder to clean (e.g., 'Sent Items')"
    Write-Host "  -OlderThanDays <days>        Delete emails older than X days (default: 365)"
    Write-Host "  -OlderThanDate <date>        Delete emails older than specific date"
    Write-Host "  -DateRangeStart <date>       Start date for range deletion (e.g., '2023-01-01')"
    Write-Host "  -DateRangeEnd <date>         End date for range deletion (max 365 days range)"
    Write-Host "  -ConfirmDelete               MANDATORY switch to enable actual deletion"
    Write-Host ""
    Write-Host "OPERATION MODES:" -ForegroundColor Yellow
    Write-Host "  -CheckOnly                   Only check mailbox status, don't delete"
    Write-Host "  -GenerateReport              Generate comprehensive mailbox report"
    Write-Host "  -DetailedFolderReport        Generate detailed folder statistics"
    Write-Host "  -ReportAllMailboxes          Report on all mailboxes (use with -GenerateReport)"
    Write-Host "  -DryRun                      Show what would be deleted without deleting"
    Write-Host ""
    Write-Host "DELETION METHODS:" -ForegroundColor Yellow
    Write-Host "  -UseComplianceSearch         Use Compliance Search (required for Deleted Items)"
    Write-Host "  -RemoveRetentionPolicy       Remove retention policies blocking deletion"
    Write-Host ""
    Write-Host "GRAPH API AUTHENTICATION (for regular folder cleanup):" -ForegroundColor Yellow
    Write-Host "  -TenantId <id>              Azure AD Tenant ID"
    Write-Host "  -ClientId <id>              Application Client ID"
    Write-Host "  -ClientSecret <secret>      Application Client Secret"
    Write-Host ""
    Write-Host "OTHER OPTIONS:" -ForegroundColor Yellow
    Write-Host "  -BatchSize <number>          Messages per batch (default: 50)"
    Write-Host "  -SkipStatistics              Skip statistics gathering"
    Write-Host "  -ForceWait                   Wait for Compliance Search to complete"
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "EXAMPLES:" -ForegroundColor Yellow
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "1. Check mailbox status and identify issues:" -ForegroundColor Green
    Write-Host "   .\ExchangeMailboxManager.ps1 -MailboxAddress 'user@domain.com' -CheckOnly -AdminUpn 'admin@domain.com'"
    Write-Host ""
    Write-Host "2. Generate comprehensive report for all mailboxes:" -ForegroundColor Green
    Write-Host "   .\ExchangeMailboxManager.ps1 -GenerateReport -ReportAllMailboxes -AdminUpn 'admin@domain.com'"
    Write-Host ""
    Write-Host "3. Generate detailed folder report for specific mailbox:" -ForegroundColor Green
    Write-Host "   .\ExchangeMailboxManager.ps1 -MailboxAddress 'user@domain.com' -DetailedFolderReport -AdminUpn 'admin@domain.com'"
    Write-Host ""
    Write-Host "4. Clean regular folders using Graph API (with ConfirmDelete):" -ForegroundColor Green
    Write-Host "   .\ExchangeMailboxManager.ps1 -MailboxAddress 'user@domain.com' -FolderPath 'Sent Items' \"
    Write-Host "     -OlderThanDays 365 -ConfirmDelete -TenantId 'xxx' -ClientId 'yyy' -ClientSecret 'zzz'"
    Write-Host ""
    Write-Host "5. Clean Deleted Items using Compliance Search:" -ForegroundColor Green
    Write-Host "   .\ExchangeMailboxManager.ps1 -MailboxAddress 'user@domain.com' -FolderPath 'Deleted Items' \"
    Write-Host "     -UseComplianceSearch -ConfirmDelete -AdminUpn 'admin@domain.com' -RemoveRetentionPolicy"
    Write-Host ""
    Write-Host "6. Delete emails from a specific year (date range):" -ForegroundColor Green
    Write-Host "   .\ExchangeMailboxManager.ps1 -MailboxAddress 'user@domain.com' \"
    Write-Host "     -DateRangeStart '2023-01-01' -DateRangeEnd '2023-12-31' -ConfirmDelete \"
    Write-Host "     -UseComplianceSearch -AdminUpn 'admin@domain.com'"
    Write-Host ""
    Write-Host "7. Delete emails older than specific date:" -ForegroundColor Green
    Write-Host "   .\ExchangeMailboxManager.ps1 -MailboxAddress 'user@domain.com' \"
    Write-Host "     -OlderThanDate '2024-01-01' -ConfirmDelete \"
    Write-Host "     -UseComplianceSearch -AdminUpn 'admin@domain.com'"
    Write-Host ""
    Write-Host "8. Dry run to see what would be deleted (no ConfirmDelete needed):" -ForegroundColor Green
    Write-Host "   .\ExchangeMailboxManager.ps1 -MailboxAddress 'user@domain.com' -DryRun \"
    Write-Host "     -DateRangeStart '2023-01-01' -DateRangeEnd '2023-12-31' \"
    Write-Host "     -UseComplianceSearch -AdminUpn 'admin@domain.com'"
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "For detailed documentation, see GUIDE.md and QUICK_START.md" -ForegroundColor Gray
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
}
#endregion

#region Main Script
try {
    # Check if script was called without any meaningful arguments
    $hasAnyArgument = $MailboxAddress -or $GenerateReport -or $ReportAllMailboxes -or $CheckOnly -or $DetailedFolderReport
    
    if (-not $hasAnyArgument) {
        Show-Help
        exit 0
    }
    
    Write-Banner "Microsoft 365 Mailbox Cleanup"
    
    # Log script invocation with all parameters (for audit trail)
    Write-Log "Script started by user: $env:USERNAME on computer: $env:COMPUTERNAME" "INFO"
    Write-Log "PowerShell Version: $($PSVersionTable.PSVersion)" "INFO"
    Write-Log "Script Version: 2.0 (Date Range Deletion with Enhanced Logging)" "INFO"
    Write-Log "`nScript Parameters:" "INFO"
    Write-Log "  -MailboxAddress: $(if($MailboxAddress){"'$MailboxAddress'"}else{'(not set)'})" "INFO"
    Write-Log "  -AdminUpn: $(if($AdminUpn){"'$AdminUpn'"}else{'(not set)'})" "INFO"
    Write-Log "  -FolderPath: $(if($FolderPath){"'$FolderPath'"}else{'(not set - all folders)'})" "INFO"
    Write-Log "  -OlderThanDays: $OlderThanDays" "INFO"
    Write-Log "  -OlderThanDate: $(if($OlderThanDate){$OlderThanDate.ToString('yyyy-MM-dd')}else{'(not set)'})" "INFO"
    Write-Log "  -DateRangeStart: $(if($DateRangeStart){$DateRangeStart.ToString('yyyy-MM-dd')}else{'(not set)'})" "INFO"
    Write-Log "  -DateRangeEnd: $(if($DateRangeEnd){$DateRangeEnd.ToString('yyyy-MM-dd')}else{'(not set)'})" "INFO"
    Write-Log "  -ConfirmDelete: $ConfirmDelete" "INFO"
    Write-Log "  -UseComplianceSearch: $UseComplianceSearch" "INFO"
    Write-Log "  -RemoveRetentionPolicy: $RemoveRetentionPolicy" "INFO"
    Write-Log "  -CheckOnly: $CheckOnly" "INFO"
    Write-Log "  -DryRun: $DryRun" "INFO"
    Write-Log "  -GenerateReport: $GenerateReport" "INFO"
    Write-Log "  -DetailedFolderReport: $DetailedFolderReport" "INFO"
    Write-Log "  -ReportAllMailboxes: $ReportAllMailboxes" "INFO"
    Write-Log "  -BatchSize: $BatchSize" "INFO"
    Write-Log "  -SkipStatistics: $SkipStatistics" "INFO"
    Write-Log "  -ForceWait: $ForceWait" "INFO"
    if ($TenantId) { Write-Log "  -TenantId: [REDACTED - provided]" "INFO" }
    if ($ClientId) { Write-Log "  -ClientId: [REDACTED - provided]" "INFO" }
    if ($ClientSecret) { Write-Log "  -ClientSecret: [REDACTED - provided]" "INFO" }
    Write-Log "" "INFO"
    
    # Validate parameters
    if (-not $GenerateReport -and -not $ReportAllMailboxes) {
        if (-not $MailboxAddress) {
            Write-Log "ERROR: -MailboxAddress is required for cleanup and check operations" "ERROR"
            Write-Log "For reports on all mailboxes, use: -GenerateReport -ReportAllMailboxes -AdminUpn 'admin@domain.com'" "INFO"
            exit 1
        }
    }
    
    # Validate date range parameters
    if (($DateRangeStart -and -not $DateRangeEnd) -or (-not $DateRangeStart -and $DateRangeEnd)) {
        Write-Log "ERROR: Both -DateRangeStart and -DateRangeEnd must be specified together" "ERROR"
        Write-Log "Example: -DateRangeStart '2023-01-01' -DateRangeEnd '2023-12-31'" "INFO"
        exit 1
    }
    
    if ($DateRangeStart -and $DateRangeEnd) {
        if ($DateRangeStart -gt $DateRangeEnd) {
            Write-Log "ERROR: -DateRangeStart must be earlier than -DateRangeEnd" "ERROR"
            exit 1
        }
        
        $rangeDays = ($DateRangeEnd - $DateRangeStart).TotalDays
        if ($rangeDays -gt 365) {
            Write-Log "ERROR: Date range cannot exceed 365 days (1 year)" "ERROR"
            Write-Log "Current range: $([math]::Round($rangeDays, 0)) days" "ERROR"
            Write-Log "Please use a smaller date range for safety" "INFO"
            exit 1
        }
        
        Write-Log "Date range validated: $($DateRangeStart.ToString('yyyy-MM-dd')) to $($DateRangeEnd.ToString('yyyy-MM-dd')) ($([math]::Round($rangeDays, 0)) days)" "SUCCESS"
    }
    
    # Check for conflicting date parameters
    if (($OlderThanDate -or $OlderThanDays -ne 365) -and ($DateRangeStart -or $DateRangeEnd)) {
        Write-Log "ERROR: Cannot use both -OlderThan* and -DateRange* parameters together" "ERROR"
        Write-Log "Choose one method:" "INFO"
        Write-Log "  • -OlderThanDays or -OlderThanDate (delete everything older than a date)" "INFO"
        Write-Log "  • -DateRangeStart and -DateRangeEnd (delete within a specific range)" "INFO"
        exit 1
    }
    
    # Safety check: Require -ConfirmDelete for actual deletion (unless in CheckOnly, DryRun, or Report mode)
    $isReadOnlyMode = $CheckOnly -or $DryRun -or $GenerateReport -or $DetailedFolderReport
    if (-not $isReadOnlyMode -and -not $ConfirmDelete) {
        Write-Log "`n⚠️  SAFETY CHECK: -ConfirmDelete REQUIRED ⚠️" "ERROR"
        Write-LogAndHost "`n⚠️  SAFETY CHECK: -ConfirmDelete REQUIRED ⚠️" -Color Red
        Write-LogAndHost "`nYou are attempting to delete emails, but the -ConfirmDelete switch was not provided." -Color Yellow
        Write-LogAndHost "This is a safety mechanism to prevent accidental deletions.`n" -Color Yellow
        Write-LogAndHost "To proceed with deletion, add the -ConfirmDelete switch to your command:" -Color Cyan
        Write-LogAndHost "  -ConfirmDelete`n" -Color Green
        Write-LogAndHost "Or use -DryRun to see what would be deleted without actually deleting:" -Color Cyan
        Write-LogAndHost "  -DryRun`n" -Color Green
        Write-Log "ERROR: Deletion blocked - missing -ConfirmDelete switch" "ERROR"
        exit 1
    }
    
    if ($ConfirmDelete -and -not $isReadOnlyMode) {
        Write-Log "⚠️  DELETION CONFIRMED: -ConfirmDelete switch detected" "WARNING"
        Write-Log "⚠️  This will perform ACTUAL DELETION of emails" "WARNING"
    }
    
    Write-Log "Configuration:" "INFO"
    Write-Log "  Mailbox: $(if($MailboxAddress){$MailboxAddress}else{'All Mailboxes'})" "INFO"
    Write-Log "  Folder: $(if($FolderPath){"$FolderPath"}else{'All Folders'})" "INFO"
    
    if ($DateRangeStart -and $DateRangeEnd) {
        Write-Log "  Date Range: $($DateRangeStart.ToString('yyyy-MM-dd')) to $($DateRangeEnd.ToString('yyyy-MM-dd'))" "INFO"
    }
    elseif ($OlderThanDate) {
        Write-Log "  Delete older than: $($OlderThanDate.ToString('yyyy-MM-dd'))" "INFO"
    }
    else {
        Write-Log "  Delete older than: $OlderThanDays days" "INFO"
    }
    
    Write-Log "  Mode: $(if($UseComplianceSearch){'Compliance Search'}else{'Graph API'})" "INFO"
    Write-Log "  Dry Run: $DryRun" "INFO"
    Write-Log "  Check Only: $CheckOnly" "INFO"
    Write-Log "  Confirm Delete: $ConfirmDelete" "INFO"
    
    # Calculate cutoff date (for non-range mode)
    if (-not ($DateRangeStart -and $DateRangeEnd)) {
        if ($OlderThanDate) {
            $cutoffDate = $OlderThanDate
        }
        else {
            $cutoffDate = (Get-Date).AddDays(-$OlderThanDays)
        }
        Write-Log "  Cutoff Date: $($cutoffDate.ToString('yyyy-MM-dd'))" "INFO"
    }
    
    # Check Only Mode
    if ($CheckOnly) {
        if (-not $AdminUpn) {
            Write-Log "ERROR: -AdminUpn required for -CheckOnly mode" "ERROR"
            exit 1
        }
        
        Connect-ComplianceCenter -AdminUpn $AdminUpn
        Get-MailboxDiagnostics -MailboxAddress $MailboxAddress
        
        Write-Log "`nCheck complete. Review recommendations above." "SUCCESS"
        exit 0
    }
    
    # Report Generation Mode
    if ($GenerateReport -or $DetailedFolderReport) {
        if (-not $AdminUpn) {
            Write-Log "ERROR: -AdminUpn required for report generation" "ERROR"
            exit 1
        }
        
        $reportCount = 0
        
        # Comprehensive Mailbox Report
        if ($GenerateReport) {
            if ($ReportAllMailboxes) {
                Write-Log "Generating comprehensive report for ALL mailboxes..." "INFO"
                $count = Get-ComprehensiveMailboxReport -AllMailboxes $true
                $reportCount += $count
            }
            elseif ($MailboxAddress) {
                Write-Log "Generating comprehensive report for: $MailboxAddress" "INFO"
                $count = Get-ComprehensiveMailboxReport -MailboxAddress $MailboxAddress -AllMailboxes $false
                $reportCount += $count
            }
            else {
                Write-Log "ERROR: -MailboxAddress or -ReportAllMailboxes required with -GenerateReport" "ERROR"
                exit 1
            }
        }
        
        # Detailed Folder Report
        if ($DetailedFolderReport) {
            if (-not $MailboxAddress) {
                Write-Log "ERROR: -MailboxAddress required for -DetailedFolderReport" "ERROR"
                exit 1
            }
            
            Write-Log "`nGenerating detailed folder report..." "INFO"
            $folderCount = Get-DetailedFolderReport -MailboxAddress $MailboxAddress
            $reportCount += $folderCount
        }
        
        # Summary
        Write-Banner "Report Generation Complete"
        if ($reportCount -gt 0) {
            Write-Log "Successfully processed $reportCount item(s)" "SUCCESS"
        }
        
        exit 0
    }
    
    # Compliance Search Mode
    if ($UseComplianceSearch) {
        if (-not $AdminUpn) {
            Write-Log "ERROR: -AdminUpn required for -UseComplianceSearch mode" "ERROR"
            exit 1
        }
        
        Write-Banner "Compliance Search Cleanup"
        
        # Connect
        if (-not (Connect-ComplianceCenter -AdminUpn $AdminUpn)) {
            exit 1
        }
        
        # Remove retention if requested
        if ($RemoveRetentionPolicy) {
            if (-not (Remove-RetentionSettings -MailboxAddress $MailboxAddress)) {
                Write-Log "Failed to remove retention settings, but continuing..." "WARNING"
            }
        }
        
        # Start cleanup
        if ($DateRangeStart -and $DateRangeEnd) {
            $processed = Start-ComplianceSearchCleanup -MailboxAddress $MailboxAddress -CutoffDate $null -DryRunMode:$DryRun -WaitForCompletion:$ForceWait -RangeStart $DateRangeStart -RangeEnd $DateRangeEnd
        }
        else {
            $processed = Start-ComplianceSearchCleanup -MailboxAddress $MailboxAddress -CutoffDate $cutoffDate -DryRunMode:$DryRun -WaitForCompletion:$ForceWait -RangeStart $null -RangeEnd $null
        }
        
        $script:TotalProcessed = $processed
    }
    # Graph API Mode
    else {
        if (-not $TenantId -or -not $ClientId -or -not $ClientSecret) {
            Write-Log "ERROR: -TenantId, -ClientId, and -ClientSecret required for Graph API mode" "ERROR"
            Write-Log "Alternatively, use -UseComplianceSearch with -AdminUpn" "INFO"
            exit 1
        }
        
        Write-Banner "Graph API Cleanup"
        
        # Connect
        if (-not (Connect-GraphAPI -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret)) {
            exit 1
        }
        
        # Verify mailbox
        Write-Log "Verifying mailbox access..." "INFO"
        $user = Get-MgUser -UserId $MailboxAddress -ErrorAction Stop
        Write-Log "Mailbox verified: $($user.DisplayName)" "SUCCESS"
        
        # Get folders
        if ($FolderPath) {
            $folder = Get-FolderByPath -UserId $MailboxAddress -Path $FolderPath
            if (-not $folder) {
                Write-Log "Folder '$FolderPath' not found" "ERROR"
                exit 1
            }
            $folders = @($folder)
        }
        else {
            Write-Log "Getting all mail folders..." "INFO"
            $folders = Get-MgUserMailFolder -UserId $MailboxAddress -All
        }
        
        Write-Log "Processing $($folders.Count) folder(s)" "INFO"
        
        # Process folders
        $folderNum = 0
        foreach ($folder in $folders) {
            $folderNum++
            Write-Log "`n[$folderNum/$($folders.Count)] $($folder.DisplayName)" "INFO"
            
            if ($DateRangeStart -and $DateRangeEnd) {
                $deleted = Remove-EmailsGraphAPI -UserId $MailboxAddress -Folder $folder -CutoffDate $null -DryRunMode:$DryRun -BatchSize $BatchSize -RangeStart $DateRangeStart -RangeEnd $DateRangeEnd
            }
            else {
                $deleted = Remove-EmailsGraphAPI -UserId $MailboxAddress -Folder $folder -CutoffDate $cutoffDate -DryRunMode:$DryRun -BatchSize $BatchSize -RangeStart $null -RangeEnd $null
            }
            $script:TotalProcessed += $deleted
        }
        
        # Disconnect
        Disconnect-MgGraph | Out-Null
    }
    
    # Summary
    Write-Banner "Summary"
    Write-Log "Total items processed: $script:TotalProcessed" "INFO"
    
    if ($DryRun) {
        Write-Log "Mode: DRY RUN - No items were actually deleted" "WARNING"
    }
    elseif ($UseComplianceSearch) {
        Write-Log "Compliance purge initiated - check status in 24 hours" "SUCCESS"
    }
    else {
        Write-Log "Items deleted successfully" "SUCCESS"
    }
    
    Write-Log "`nLog file: $LogFile" "INFO"
    
    # Script completion logging (for audit trail)
    Write-Log "`n========================================" "SUCCESS"
    Write-Log "SCRIPT COMPLETED SUCCESSFULLY" "SUCCESS"
    Write-Log "Completion time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" "SUCCESS"
    Write-Log "Total execution time: $([math]::Round(((Get-Date) - $script:ScriptStartTime).TotalSeconds, 2)) seconds" "SUCCESS"
    Write-Log "========================================" "SUCCESS"
    
    exit 0
}
catch {
    Write-Log "`n========================================" "ERROR"
    Write-Log "SCRIPT FAILED WITH ERROR" "ERROR"
    Write-Log "Error time: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" "ERROR"
    Write-Log "CRITICAL ERROR: $($_.Exception.Message)" "ERROR"
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" "ERROR"
    Write-Log "Error Type: $($_.Exception.GetType().FullName)" "ERROR"
    Write-Log "========================================" "ERROR"
    exit 1
}
#endregion
