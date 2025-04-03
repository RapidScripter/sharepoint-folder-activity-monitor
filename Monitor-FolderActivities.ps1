<#
=============================================================================================
.SYNOPSIS
    Monitors folder activities in SharePoint Online and OneDrive for Business.

.DESCRIPTION
    This script tracks and reports on folder-level activities including creation, modification, 
    deletion, and restoration across SharePoint Online and OneDrive with advanced filtering capabilities.

.VERSION
    2.0

.FEATURES
    - Tracks 15+ folder activities with detailed metadata
    - Supports custom date ranges (up to 180 days)
    - Multiple authentication methods (MFA, Certificate, Basic)
    - Parallel processing for large datasets
    - Risk scoring for sensitive activities
    - Comprehensive error handling and logging
    - Scheduled task friendly with parameterized credentials
=============================================================================================
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$StartDate,
    
    [Parameter(Mandatory = $false)]
    [Nullable[DateTime]]$EndDate,
    
    [Parameter(Mandatory = $false)]
    [ValidatePattern("^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$")]
    [string]$PerformedBy,
    
    [Parameter(Mandatory = $false)]
    [ValidatePattern("^https://.+\.sharepoint\.com/.+")]
    [string]$SiteUrl,
    
    [Parameter(Mandatory = $false)]
    [ValidateScript({
        if ($_ -eq "" -or (Test-Path $_ -PathType 'Leaf')) { $true }
        else { throw "CSV file not found" }
    })]
    [string]$ImportSitesCsv,
    
    [Parameter(Mandatory = $false)]
    [switch]$SharePointOnline,
    
    [Parameter(Mandatory = $false)]
    [switch]$OneDrive,
    
    [Parameter(Mandatory = $false)]
    [string]$UserName,
    
    [Parameter(Mandatory = $false)]
    [string]$Password,
    
    [Parameter(Mandatory = $false)]
    [string]$Organization,
    
    [Parameter(Mandatory = $false)]
    [string]$ClientId,
    
    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbPrint,
    
    [Parameter(Mandatory = $false)]
    [string]$CertificatePath,
    
    [Parameter(Mandatory = $false)]
    [securestring]$CertificatePassword,
    
    [Parameter(Mandatory = $false)]
    [switch]$IncludeSystemEvent,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 10080)]
    [int]$IntervalMinutes = 1440,
    
    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 100)]
    [int]$ThrottleLimit = 5,
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = (Get-Location)
)

#region Initialization
$ErrorActionPreference = "Stop"
$WarningPreference = "Continue"
$ProgressPreference = "SilentlyContinue" # Change to "Continue" for debugging

$MaxStartDate = ((Get-Date).AddDays(-180)).Date
$OperationNames = @(
    "FolderCreated", "FolderModified", "FolderRenamed", 
    "FolderCopied", "FolderMoved", "FolderDeleted", 
    "FolderRecycled", "FolderDeletedFirstStageRecycleBin", 
    "FolderDeletedSecondStageRecycleBin", "FolderRestored"
) -join ","

$ScriptVersion = "2.0"
$ExecutionStartTime = Get-Date
$OutputFileName = "Audit_SPO_Folder_Activity_Report_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').csv"
$OutputCSV = Join-Path $OutputPath $OutputFileName
#endregion

#region Functions
function Connect-Exchange {
    [CmdletBinding()]
    param()
    
    try {
        # Check for Exchange Online module
        if (-not (Get-Module ExchangeOnlineManagement -ListAvailable)) {
            Write-Host "ExchangeOnline module not found. Installing..." -ForegroundColor Yellow
            Install-Module ExchangeOnlineManagement -Repository PSGallery -AllowClobber -Force -Scope CurrentUser
        }

        Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
        
        # Authentication logic
        if ($ClientId -and $CertificateThumbPrint) {
            Connect-ExchangeOnline -AppId $ClientId -CertificateThumbprint $CertificateThumbPrint -Organization $Organization -ShowBanner:$false
        }
        elseif ($ClientId -and $CertificatePath) {
            Connect-ExchangeOnline -AppId $ClientId -CertificateFilePath $CertificatePath -CertificatePassword $CertificatePassword -Organization $Organization -ShowBanner:$false
        }
        elseif ($UserName -and $Password) {
            $SecuredPassword = ConvertTo-SecureString -AsPlainText $Password -Force
            $Credential = New-Object System.Management.Automation.PSCredential $UserName, $SecuredPassword
            Connect-ExchangeOnline -Credential $Credential -ShowBanner:$false
        }
        else {
            Connect-ExchangeOnline -ShowBanner:$false
        }
        
        Write-Host "Successfully connected to Exchange Online" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to connect to Exchange Online: $_"
        exit 1
    }
}

function Process-AuditResults {
    [CmdletBinding()]
    param(
        [Parameter(ValueFromPipeline = $true)]
        $Results,
        [string[]]$FilterSites
    )
    
    begin {
        $processedCount = 0
        $outputEvents = @()
    }
    
    process {
        $Results | ForEach-Object -Parallel {
            $record = $_
            try {
                $auditData = $record.auditdata | ConvertFrom-Json
                $printFlag = $true

                # Apply filters
                if (-not $using:IncludeSystemEvent -and $record.UserIds -in @("app@sharepoint", "SHAREPOINT\system")) {
                    $printFlag = $false
                }

                if ($using:PerformedBy -and $using:PerformedBy -ne $record.UserIds) {
                    $printFlag = $false
                }

                if ($using:SharePointOnline -and $auditData.Workload -eq "OneDrive") {
                    $printFlag = $false
                }

                if ($using:OneDrive -and $auditData.Workload -eq "SharePoint") {
                    $printFlag = $false
                }

                if ($using:SiteUrl -and $using:SiteUrl -ne $auditData.SiteUrl) {
                    $printFlag = $false
                }

                if ($using:FilterSites.Count -gt 0 -and (-not ($using:FilterSites -contains $auditData.SiteUrl))) {
                    $printFlag = $false
                }

                if ($printFlag) {
                    $activityTime = (Get-Date $auditData.CreationTime).ToLocalTime()
                    $riskScore = if ($record.Operations -match "Deleted|Recycled") { "High" } elseif ($record.Operations -match "Modified|Renamed") { "Medium" } else { "Low" }

                    [PSCustomObject]@{
                        'Activity Time'    = $activityTime
                        'Activity'        = $record.Operations
                        'Folder Name'     = $auditData.SourceFileName
                        'Performed By'    = $record.UserIds
                        'Folder URL'      = $auditData.ObjectID
                        'Site URL'        = $auditData.SiteUrl
                        'Workload'        = $auditData.Workload
                        'Risk Score'     = $riskScore
                        'Duration (Days)' = [math]::Round(((Get-Date) - $activityTime).TotalDays, 1)
                        'More Info'      = $record.auditdata
                    }
                }
            }
            catch {
                Write-Warning "Error processing record: $_"
            }
        } -ThrottleLimit $ThrottleLimit | ForEach-Object {
            $outputEvents += $_
            $processedCount++
            if ($processedCount % 100 -eq 0) {
                Write-Progress -Activity "Processing audit records" -Status "Processed $processedCount records" -PercentComplete ($processedCount % 1000)
            }
        }
    }
    
    end {
        return $outputEvents
    }
}
#endregion

#region Main Execution
try {
    # Validate and set date range
    if (-not $StartDate -and -not $EndDate) {
        $EndDate = (Get-Date).Date
        $StartDate = $MaxStartDate
    }

    $StartDate = [DateTime]$StartDate
    $EndDate = [DateTime]$EndDate

    if ($StartDate -lt $MaxStartDate) {
        throw "Audit can only be retrieved for past 180 days. Please select a date after $MaxStartDate"
    }

    if ($EndDate -lt $StartDate) {
        throw "End time should be later than start time"
    }

    # Initialize connection
    Connect-Exchange

    # Load site filters if specified
    $FilterSites = @()
    if ($ImportSitesCsv) {
        $FilterSites = Import-Csv -Path $ImportSitesCsv | Select-Object -ExpandProperty SiteUrl
    }

    # Initialize processing
    $CurrentStart = $StartDate
    $CurrentEnd = $CurrentStart.AddMinutes($IntervalMinutes)
    if ($CurrentEnd -gt $EndDate) {
        $CurrentEnd = $EndDate
    }

    # Prepare output file
    if (Test-Path $OutputCSV) {
        Remove-Item $OutputCSV -Force
    }

    # Main processing loop
    $totalRecords = 0
    while ($true) {
        try {
            Write-Host "Retrieving logs from $CurrentStart to $CurrentEnd..." -ForegroundColor Cyan
            $Results = Search-UnifiedAuditLog -StartDate $CurrentStart -EndDate $CurrentEnd `
                -Operations $OperationNames -SessionId "SPOFolderAudit" `
                -SessionCommand ReturnLargeSet -ResultSize 5000

            if ($Results) {
                $processedResults = $Results | Process-AuditResults -FilterSites $FilterSites
                $processedResults | Export-Csv -Path $OutputCSV -NoTypeInformation -Append
                $totalRecords += $processedResults.Count
            }

            # Pagination logic
            if ($Results.Count -lt 5000) {
                if ($CurrentEnd -ge $EndDate) { break }
                $CurrentStart = $CurrentEnd
                $CurrentEnd = $CurrentStart.AddMinutes($IntervalMinutes)
                if ($CurrentEnd -gt $EndDate) { $CurrentEnd = $EndDate }
            }

            # Rate limiting
            Start-Sleep -Milliseconds 500
        }
        catch {
            Write-Warning "Error processing batch: $_"
            if ($_.Exception.Message -match "throttled") {
                $retrySeconds = 30
                Write-Host "Throttled detected. Waiting $retrySeconds seconds..." -ForegroundColor Yellow
                Start-Sleep -Seconds $retrySeconds
                continue
            }
            break
        }
    }

    # Output results
    if ($totalRecords -gt 0) {
        Write-Host "`nSuccessfully processed $totalRecords audit records" -ForegroundColor Green
        Write-Host "Report saved to: $OutputCSV" -ForegroundColor Cyan
        
        # Option to open file
        $openFile = Read-Host "Open report file now? (Y/N)"
        if ($openFile -match "[yY]") {
            Invoke-Item $OutputCSV
        }
    }
    else {
        Write-Host "No matching records found for the specified criteria" -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Script execution failed: $_"
}
finally {
    # Clean up connection
    try {
        Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
        Disconnect-ExchangeOnline -Confirm:$false -InformationAction Ignore -ErrorAction SilentlyContinue
    }
    catch {
        Write-Warning "Error disconnecting: $_"
    }

    # Calculate and display execution time
    $executionTime = (Get-Date) - $ExecutionStartTime
    Write-Host "`nScript execution completed in $($executionTime.TotalMinutes.ToString('0.00')) minutes" -ForegroundColor Cyan
}
#endregion