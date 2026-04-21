# =============================
# Module: Search-GPO.psm1
# Version: 9.1.4
# Author: PowerShell Automation Team
# Description: High-performance GPO content search tool
# =============================
#
# FEATURES:
# ---------
# • Parallel processing for fast GPO content searching
# • OU-scoped searching with recursive option
# • Flexible name filtering with wildcard support
# • Clean, color-coded output with runtime tracking
# • Hybrid JSON caching for optimal performance
# • Offline search capability with local cache
# • PassThru option for pipeline integration
#
# EXAMPLES:
# ---------
# Search-GPO -TextQuery "Screen Saver" -Domain "contoso.com"
# Search-GPO -TextQuery "Password Policy" -Domain "contoso.com" -NameQuery "*Security*"
# Search-GPO -TextQuery "Chrome Settings" -Domain "contoso.com" -OU "OU=Workstations,DC=contoso,DC=com" -Recursive
# Search-GPO -TextQuery "Firewall" -Domain "contoso.com" -ThrottleLimit 32
# Search-GPO -TextQuery "Registry" -Domain "contoso.com" -PassThru | Export-Csv -Path "GPOs.csv"
# Search-GPO -TextQuery "IE Settings" -Domain "contoso.com" -Local -SkipCacheValidation
#
# PARAMETERS:
# -----------
# • TextQuery    : Text to search for in GPO content (Mandatory)
# • NameQuery    : GPO name filter (supports wildcards, default: *)
# • Domain       : Target domain for GPO search (Mandatory)
# • Server       : Specific domain controller to use
# • OU           : Organizational Unit to scope the search
# • Recursive    : Include child OUs in search
# • ThrottleLimit: Parallel thread count (default: 16)
# • PassThru     : Return matching GPO objects
# • Verbosity    : Logging detail level (0-3)
# • Local        : Use local cache only
# • SkipCacheValidation: Skip cache freshness validation (offline mode)
# • CachePath    : Cache directory path
#
# =============================

function Search-GPO {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=0)]
        [string]$TextQuery,

        [string]$NameQuery = "*",

        [Parameter(Mandatory=$true)]
        [string]$Domain,

        [string]$Server,

        [string]$OU,
        [switch]$Recursive,

        [int]$ThrottleLimit = 16,

        [switch]$PassThru,

        [int]$Verbosity = 1,

        [switch]$Local,

        [switch]$SkipCacheValidation,

        [string]$CachePath = "C:\tools\GPOScanner\GPOReports"
    )

    # Add C# optimized classes
    Add-Type -TypeDefinition @"
using System;
using System.IO;
using System.Text;

public static class GPONativeSearch
{
    // Ultra-fast case-insensitive search
    public static bool FastContains(string content, string searchText)
    {
        if (string.IsNullOrEmpty(content) || string.IsNullOrEmpty(searchText))
            return false;
           
        return content.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0;
    }
   
    // Memory-efficient file reading with large buffer
    public static string ReadFileFast(string filePath)
    {
        const int bufferSize = 65536;
        using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, bufferSize))
        using (var reader = new StreamReader(stream, Encoding.UTF8, true, bufferSize))
        {
            return reader.ReadToEnd();
        }
    }
}
"@ -ReferencedAssemblies @("System.IO", "System.Linq")

    # -------------------------
    # Parent runspace state
    # -------------------------
    $script:TotalGPOs    = 0
    $script:FilteredGPOs = 0
    $script:MatchedGPOs  = 0
    $script:MatchedGpoList = [System.Collections.Generic.List[string]]::new()
    $script:PerformanceStats = @{
        TotalTime = 0
    }

    # -------------------------
    # Memory management
    # -------------------------
    function Invoke-MemoryCleanup {
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        [System.GC]::Collect()
    }

    # -------------------------
    # JSON-only cache
    # -------------------------
    $domainCachePath = Join-Path $CachePath $Domain
    $jsonCache = Join-Path $domainCachePath "JSON"
   
    if (-not (Test-Path $jsonCache)) {
        New-Item -Path $jsonCache -ItemType Directory -Force | Out-Null
    }

    # -------------------------
    # Optimized logging (without memory tracking)
    # -------------------------
    $ShowPerGpoLogs = $PSCmdlet.MyInvocation.BoundParameters.ContainsKey('Verbose')
    function Write-Log {
        param([string]$Message, [int]$Level=1, [int]$Indent=0, [string]$Kind="Info")
        if ($Level -le $Verbosity -or ($Kind -eq "Debug" -and $ShowPerGpoLogs)) {
            $ts = Get-Date -Format "HH:mm:ss"
            $indent = (" " * ($Indent * 2))
            $out = "[$ts]$indent $Message"
            switch ($Kind) {
                "Success" { Write-Host $out -ForegroundColor Green }
                "Warn"    { Write-Host $out -ForegroundColor Yellow }
                "Error"   { Write-Host $out -ForegroundColor Red }
                "Debug"   { Write-Verbose $out }
                default   { Write-Host $out }
            }
        }
    }

    # -------------------------
    # Cache-only GPO enumeration
    # -------------------------
    function Get-CachedGPOsFromJSON {
        Write-Log "Reading GPOs from local cache only (offline mode)..." 1 0 "Info"
       
        $cachedGpos = [System.Collections.Generic.List[object]]::new()
        $jsonFiles = Get-ChildItem -Path $jsonCache -Filter "*.json" -ErrorAction SilentlyContinue
       
        if ($jsonFiles.Count -eq 0) {
            Write-Log "No cached GPOs found in $jsonCache" 0 0 "Warn"
            return @()
        }
       
        Write-Log "Found $($jsonFiles.Count) cached GPO reports" 1 1 "Debug"
       
        foreach ($jsonFile in $jsonFiles) {
            try {
                $jsonContent = Get-Content -Path $jsonFile.FullName -Raw -ErrorAction Stop
                $jsonData = $jsonContent | ConvertFrom-Json
               
                # Create a mock GPO object from cache metadata
                $gpo = [PSCustomObject]@{
                    Id = [guid]$jsonData.metadata.id
                    DisplayName = $jsonData.metadata.name
                    ModificationTime = [datetime]::Parse($jsonData.metadata.modified)
                    DomainName = $jsonData.metadata.domain
                    Description = "From cache - $($jsonFile.Name)"
                }
               
                $cachedGpos.Add($gpo)
            }
            catch {
                Write-Log "Failed to read cached GPO: $($jsonFile.Name)" 2 2 "Warn"
            }
        }
       
        Write-Log "Successfully loaded $($cachedGpos.Count) GPOs from local cache" 1 0 "Success"
        return $cachedGpos
    }

    # -------------------------
    # Import GroupPolicy module
    # -------------------------
    function Import-Gpmc {
        # Only import if we're not in full offline mode
        if (-not ($Local -and $SkipCacheValidation)) {
            Write-Log "Importing Group Policy Management module..." 1 0 "Debug"
            if (-not (Get-Module -Name GroupPolicy)) {
                try {
                    Import-Module GroupPolicy -ErrorAction Stop 3> $null
                    Write-Log "Successfully imported GroupPolicy module" 2 0 "Debug"
                }
                catch {
                    Write-Log "Failed to import GroupPolicy module: $($_.Exception.Message)" 0 0 "Error"
                    return $false
                }
            }
        }
        return $true
    }

    # -------------------------
    # Get all GPOs - UPDATED LOGIC
    # -------------------------
    function Get-AllGpos {
        # OFFLINE MODE: Use only cached GPOs
        if ($Local -and $SkipCacheValidation) {
            Write-Log "OFFLINE MODE: Using local cache only (no domain contact)..." 1 0 "Info"
            $cachedGpos = Get-CachedGPOsFromJSON
            $script:TotalGPOs = $cachedGpos.Count
            if ($script:TotalGPOs -eq 0) {
                Write-Log "No cached GPOs found. Cannot proceed in offline mode." 0 0 "Error"
                return @()
            }
            Write-Log "Using $script:TotalGPOs cached GPOs for offline search" 1 0 "Success"
            return $cachedGpos
        }
       
        # ONLINE MODE: Normal domain enumeration
        Write-Log "Enumerating GPOs for domain $Domain..." 1 0 "Info"
       
        $params = @{}
        if ($Domain) { $params.Domain = $Domain }
        if ($Server) { $params.Server = $Server }

        $gpos = @()
        if ($OU) {
            Write-Log "Fetching GPOs linked to OU: $OU" 1 1 "Debug"
            try {
                $ous = @(Get-ADOrganizationalUnit -Identity $OU -Server $Domain -Properties GPLink -ErrorAction Stop)
                if ($Recursive) {
                    $childOUs = Get-ADOrganizationalUnit -Filter * -SearchBase $OU -SearchScope Subtree -Server $Domain -Properties GPLink -ErrorAction Stop
                    if ($childOUs) { $ous += $childOUs }
                }
            } catch {
                Write-Log "Failed to enumerate OU(s): $($_.Exception.Message)" 0 0 "Error"
                return @()
            }

            $gpoGuids = [System.Collections.Generic.HashSet[string]]::new()
            foreach ($ouItem in $ous) {
                if ($ouItem.GPLink) {
                    $matches = [regex]::Matches($ouItem.GPLink, "\{[0-9a-fA-F\-]{36}\}")
                    foreach ($m in $matches) {
                        [void]$gpoGuids.Add($m.Value.Trim('{}').ToLower())
                    }
                }
            }
           
            Write-Log "Found $($gpoGuids.Count) linked GPO GUID(s)." 2 1 "Debug"
            if ($gpoGuids.Count -eq 0) { return @() }

            foreach ($guid in $gpoGuids) {
                try {
                    $gpo = Get-GPO -Guid $guid -Domain $Domain -ErrorAction Stop
                    $gpos += $gpo
                }
                catch {
                    Write-Log "Failed to get GPO with GUID $guid : $($_.Exception.Message)" 0 2 "Warn"
                }
            }
        } else {
            try {
                $gpos = Get-GPO -All @params
            }
            catch {
                Write-Log "Failed to enumerate GPOs: $($_.Exception.Message)" 0 0 "Error"
                return @()
            }
        }

        $script:TotalGPOs = $gpos.Count
        Write-Log "Enumerated $script:TotalGPOs GPOs" 1 0 "Success"
       
        return $gpos
    }

    # -------------------------
    # Filter GPOs by Name
    # -------------------------
    function Get-FilteredGpos($gpoList) {
        Write-Log "Filtering GPOs by NameQuery: '$NameQuery'" 1 0 "Info"
        $filtered = $gpoList | Where-Object { $_.DisplayName -like $NameQuery }
        $script:FilteredGPOs = $filtered.Count
        Write-Log "Filtered to $script:FilteredGPOs GPOs" 1 0 "Info"
        return $filtered
    }

    # -------------------------
    # Hybrid PowerShell + C# Processing (Hybrid JSON)
    # -------------------------
    function Get-MatchingGposHybrid($filteredGpos) {
        if (-not $filteredGpos -or $filteredGpos.Count -eq 0) {
            Write-Log "No GPOs to process after filtering." 1 0 "Warn"
            return @()
        }

        Write-Log "Using hybrid JSON processing for $($filteredGpos.Count) GPOs..." 1 0 "Info"
        Write-Log "ThrottleLimit: $ThrottleLimit, Local: $Local, Cache: $(if($Local){'Enabled'}else{'Disabled'})" 2 0 "Debug"

        $runspacePool = [runspacefactory]::CreateRunspacePool(1, $ThrottleLimit)
        $runspacePool.Open()
        $jobs = [System.Collections.Generic.List[hashtable]]::new()
        $matchesAll = [System.Collections.Generic.List[object]]::new()

        foreach ($gpo in $filteredGpos) {
            $powerShell = [powershell]::Create()
            $powerShell.RunspacePool = $runspacePool
           
            [void]$powerShell.AddScript({
                param($gpo, $Local, $jsonCache, $Domain, $Server, $TextQuery, $SkipCacheValidation)
               
                $id = $gpo.Id
                $name = $gpo.DisplayName
               
                try {
                    $jsonFilePath = Join-Path $jsonCache "$id.json"
                    $searchText = $null
                    $cacheUsed = $false
                   
                    # Check JSON cache first if enabled
                    if ($Local -and (Test-Path $jsonFilePath)) {
                        if ($SkipCacheValidation) {
                            # Use cached JSON directly (offline mode)
                            $jsonContent = [GPONativeSearch]::ReadFileFast($jsonFilePath)
                            $jsonData = $jsonContent | ConvertFrom-Json
                            $searchText = $jsonData.content.searchText
                            $cacheUsed = $true
                        } else {
                            # Online mode with cache validation
                            $fileTime = (Get-Item $jsonFilePath).LastWriteTimeUtc
                            if ($fileTime -ge $gpo.ModificationTime.ToUniversalTime()) {
                                $jsonContent = [GPONativeSearch]::ReadFileFast($jsonFilePath)
                                $jsonData = $jsonContent | ConvertFrom-Json
                                $searchText = $jsonData.content.searchText
                                $cacheUsed = $true
                            }
                        }
                    }
                   
                    # Fetch from server if cache not available or stale
                    if (-not $searchText) {
                        # Get BOTH HTML (for ADMX content) and XML (for structured data)
                        $params = @{ Guid = $id; ErrorAction = 'Stop' }
                        if ($Domain) { $params.Domain = $Domain }
                        if ($Server) { $params.Server = $Server }
                       
                        # Get HTML report for comprehensive ADMX searching
                        $params.ReportType = 'HTML'
                        $htmlReport = Get-GPOReport @params
                       
                        # Get XML report for structured data
                        $params.ReportType = 'XML'
                        $xmlReport = Get-GPOReport @params
                       
                        # Clean HTML tags for fast searching (includes ADMX content)
                        $cleanHtmlText = $htmlReport -replace '<[^>]+>','' -replace '&nbsp;',' ' -replace '&amp;','&'
                        $cleanHtmlText = $cleanHtmlText -replace '\s+',' '
                        $cleanHtmlText = $cleanHtmlText.Trim()
                       
                        # Convert to hybrid JSON structure
                        $jsonData = @{
                            metadata = @{
                                id = $gpo.Id.ToString()
                                name = $gpo.DisplayName
                                modified = $gpo.ModificationTime.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
                                domain = $Domain
                                size = @{
                                    html = $htmlReport.Length
                                    xml = $xmlReport.Length
                                    searchText = $cleanHtmlText.Length
                                }
                            }
                            content = @{
                                # HTML-based text for comprehensive ADMX searching
                                searchText = $cleanHtmlText
                                # Raw formats for full details if needed
                                rawHtml = $htmlReport
                                rawXml = $xmlReport
                            }
                        }
                       
                        $searchText = $jsonData.content.searchText
                       
                        if ($Local) {
                            try {
                                $jsonData | ConvertTo-Json -Depth 10 | Out-File $jsonFilePath -Encoding UTF8
                            } catch {
                                # JSON cache write failure is non-fatal
                            }
                        }
                    }
                   
                    # Use C# for ultra-fast string search on comprehensive HTML text
                    if ($searchText -and [GPONativeSearch]::FastContains($searchText, $TextQuery)) {
                        return $gpo
                    }
                }
                catch {
                    return $null
                }
               
                return $null
            })
           
            [void]$powerShell.AddParameters(@{
                gpo = $gpo
                Local = $Local
                jsonCache = $jsonCache
                Domain = $Domain
                Server = $Server
                TextQuery = $TextQuery
                SkipCacheValidation = $SkipCacheValidation
            })
           
            $jobs.Add(@{
                PowerShell = $powerShell
                Handle = $powerShell.BeginInvoke()
                GPO = $gpo
            })
        }

        # Collect results
        $completed = 0
        $total = $jobs.Count
       
        while ($jobs.Count -gt 0) {
            for ($i = $jobs.Count - 1; $i -ge 0; $i--) {
                $job = $jobs[$i]
                if ($job.Handle.IsCompleted) {
                    try {
                        $result = $job.PowerShell.EndInvoke($job.Handle)
                        if ($result) {
                            $matchesAll.Add($result)
                        }
                    }
                    catch {
                        Write-Log "Error processing GPO '$($job.GPO.DisplayName)'" 2 0 "Warn"
                    }
                    finally {
                        $job.PowerShell.Dispose()
                        $jobs.RemoveAt($i)
                        $completed++
                    }
                }
            }
           
            # Progress reporting
            if ($total -gt 10 -and ($completed % 5 -eq 0 -or $completed -eq $total)) {
                $percentComplete = [math]::Round(($completed / $total) * 100, 2)
                Write-Progress -Activity "Scanning GPO Reports" -Status "Processed $completed of $total GPOs ($percentComplete%)" -PercentComplete $percentComplete
            }
           
            if ($jobs.Count -gt 0) {
                Start-Sleep -Milliseconds 50
            }
        }
       
        Write-Progress -Activity "Scanning GPO Reports" -Completed

        $runspacePool.Close()
        $runspacePool.Dispose()

        $script:MatchedGPOs = $matchesAll.Count
        foreach ($m in $matchesAll) {
            $script:MatchedGpoList.Add($m.DisplayName)
        }

        Write-Log "Processed $($filteredGpos.Count) GPOs" 1 0 "Success"
       
        return $matchesAll
    }

    # -------------------------
    # Convert TimeSpan to HH:MM:SS format
    # -------------------------
    function ConvertTo-HHMMSS {
        param([TimeSpan]$TimeSpan)
       
        $hours = $TimeSpan.Hours + ($TimeSpan.Days * 24)
        return "{0:D2}:{1:D2}:{2:D2}" -f $hours, $TimeSpan.Minutes, $TimeSpan.Seconds
    }

    # -------------------------
    # Enhanced main workflow
    # -------------------------
    function Do-Stuff {
        $startTime = Get-Date

        if (-not (Import-Gpmc)) {
            Write-Log "Cannot proceed without GroupPolicy module." 0 0 "Error"
            return
        }

        # Memory optimization
        Invoke-MemoryCleanup

        $gpos = Get-AllGpos
        if (-not $gpos -or $gpos.Count -eq 0) {
            Write-Log "No GPOs found. Exiting." 0 0 "Warn"
            return
        }

        $filteredGpos = Get-FilteredGpos $gpos
        if ($script:FilteredGPOs -eq 0) {
            Write-Log "No GPOs match the name filter. Exiting." 1 0 "Warn"
            return
        }

        $matchingGpos = Get-MatchingGposHybrid $filteredGpos

        $script:PerformanceStats.TotalTime = (Get-Date) - $startTime
        $totalRuntimeFormatted = ConvertTo-HHMMSS -TimeSpan $script:PerformanceStats.TotalTime

        # Display results - CLEAN FORMATTING
        Write-Host "`n============================================================" -ForegroundColor Cyan
        Write-Host "GPO SEARCH RESULTS" -ForegroundColor Cyan
        Write-Host "============================================================" -ForegroundColor Cyan

        if ($script:MatchedGpoList.Count -gt 0) {
            Write-Host "`nMATCHED GPOS ($script:MatchedGPOs):" -ForegroundColor Green
            foreach ($g in $script:MatchedGpoList | Sort-Object) {
                Write-Host "  $g" -ForegroundColor White
            }
        } else {
            Write-Host "`nNo GPOs matched your search criteria." -ForegroundColor Yellow
        }

        Write-Host "`n------------------------------------------------------------" -ForegroundColor DarkGray
        Write-Host "SUMMARY" -ForegroundColor Cyan
        Write-Host "------------------------------------------------------------" -ForegroundColor DarkGray
        Write-Host "Domain           : $Domain" -ForegroundColor Gray
        Write-Host "Search Query     : $TextQuery" -ForegroundColor Gray
        Write-Host "Cache Format     : Hybrid JSON" -ForegroundColor Gray
        Write-Host "Total GPOs       : $script:TotalGPOs" -ForegroundColor White
        Write-Host "Filtered GPOs    : $script:FilteredGPOs" -ForegroundColor White  
        Write-Host "Matched GPOs     : $script:MatchedGPOs" -ForegroundColor White
        Write-Host "Cache Mode       : $(if($Local){'Enabled'}else{'Disabled'})" -ForegroundColor Gray
        Write-Host "Offline Mode     : $(if($Local -and $SkipCacheValidation){'Yes (No Domain Contact)'}else{'No'})" -ForegroundColor Gray
        Write-Host "Throttle Limit   : $ThrottleLimit" -ForegroundColor Gray
        Write-Host "Total Runtime    : $totalRuntimeFormatted [HH:MM:SS]" -ForegroundColor Yellow

        Write-Host "`n============================================================" -ForegroundColor Cyan

        if ($PassThru) {
            return $matchingGpos | Select-Object -ExpandProperty DisplayName
        }
       
        # Final memory cleanup
        Invoke-MemoryCleanup
    }

    Do-Stuff
}

Export-ModuleMember -Function Search-GPO 
 
