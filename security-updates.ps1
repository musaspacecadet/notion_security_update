# Global variables for KB tracking and reporting
$Global:TargetKBs = @("KB5041585", "KB5040442", "KB5039212")
$Global:QueryAllUniqueKBs = $true
$Global:ForceUpdate = $true
$Global:ShowCachedResults = $false
$Global:ReportDate = (Get-Date).ToString("yyyy-MMM")
$Global:NotionSecret = "ntn_AM932828924RURAZNILX9x0Nk6mrWV1IsZEnIPbJZjW3g8"
$Global:NotionDatabaseId = "57f4505ae38e4081b43f120daf0e71d4"

# Detect the operating system and set the cache and KB file paths
if ($IsWindows) {
    $Global:CacheFilePath = "C:\temp\SecurityUpdateCache_$($Global:ReportDate).json"
    $Global:UniqueKBFilePath = "C:\temp\UniqueKBs_$($Global:ReportDate).txt"
} else {
    $Global:CacheFilePath = "/tmp/SecurityUpdateCache_$($Global:ReportDate).json"
    $Global:UniqueKBFilePath = "/tmp/UniqueKBs_$($Global:ReportDate).txt"
}

function Add-ToNotion {
    param (
        [string]$KBNumber,
        [string]$Severity
    )

    $headers = @{
        "Authorization" = "Bearer $Global:NotionSecret"
        "Content-Type" = "application/json"
        "Notion-Version" = "2022-06-28"
    }

    $body = @{
        parent = @{
            database_id = $Global:NotionDatabaseId
        }
        properties = @{
            "KB" = @{
                title = @(
                    @{
                        text = @{
                            content = $KBNumber
                        }
                    }
                )
            }
            "Severity" = @{
                rich_text = @(
                    @{
                        text = @{
                            content = $Severity
                        }
                    }
                )
            }
            "Date" = @{
                date = @{
                    start = (Get-Date -Format "yyyy-MM-dd")
                }
            }
        }
    } | ConvertTo-Json -Depth 10

    try {
        $response = Invoke-RestMethod -Uri "https://api.notion.com/v1/pages" -Method Post -Headers $headers -Body $body
        Write-Host "Successfully added $KBNumber with severity $Severity to Notion" -ForegroundColor Green
    }
    catch {
        Write-Host "Error adding $KBNumber to Notion: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Clear-NotionDatabase {
    $headers = @{
        "Authorization" = "Bearer $Global:NotionSecret"
        "Content-Type" = "application/json"
        "Notion-Version" = "2022-06-28"
    }

    try {
        # Query all pages in the database
        $queryBody = @{
            page_size = 100
        } | ConvertTo-Json

        $pages = Invoke-RestMethod -Uri "https://api.notion.com/v1/databases/$Global:NotionDatabaseId/query" -Method Post -Headers $headers -Body $queryBody

        # Archive each page
        foreach ($page in $pages.results) {
            $pageId = $page.id
            $archiveBody = @{
                archived = $true
            } | ConvertTo-Json

            Invoke-RestMethod -Uri "https://api.notion.com/v1/pages/$pageId" -Method Patch -Headers $headers -Body $archiveBody
        }
        
        Write-Host "Successfully cleared all records from Notion database" -ForegroundColor Green
    }
    catch {
        Write-Host "Error clearing Notion database: $($_.Exception.Message)" -ForegroundColor Red
    }
}


function Query-KBSeverity {
    try {
         # Clear existing records before adding new ones
        Clear-NotionDatabase

        if (-not $Global:ForceUpdate -and (Test-Path $Global:CacheFilePath)) {
            $reportData = Get-Content -Path $Global:CacheFilePath | ConvertFrom-Json
            if ($Global:ShowCachedResults) {
                $reportData | Out-GridView
                return
            }
        } else {
            if (-not (Get-Module MSRCSecurityUpdates)) {
                Install-Module MSRCSecurityUpdates -Force
            }
            $reportData = Get-MsrcCvrfDocument -ID $Global:ReportDate | Get-MsrcCvrfAffectedSoftware
            $reportData | ConvertTo-Json -Depth 5 | Set-Content -Path $Global:CacheFilePath
        }

        if ($Global:ForceUpdate) {
            Extract-UniqueKBs $reportData
        }

        if ($Global:QueryAllUniqueKBs -and (Test-Path $Global:UniqueKBFilePath)) {
            $Global:TargetKBs = Get-Content -Path $Global:UniqueKBFilePath
        }

        $criticalKBs = @()
        $importantKBs = @()
        $moderateKBs = @()
        $lowKBs = @()

        foreach ($kb in $Global:TargetKBs) {
            $filteredResult = $reportData | Where-Object { $_.KBArticle -match $kb }
            $hasCritical = $filteredResult | Where-Object { $_.Severity -eq "Critical" }
            $hasImportant = $filteredResult | Where-Object { $_.Severity -eq "Important" }
            $hasModerate = $filteredResult | Where-Object { $_.Severity -eq "Moderate" }
            $hasLow = $filteredResult | Where-Object { $_.Severity -eq "Low" }

            if ($hasCritical) {
                $criticalKBs += $kb
                Add-ToNotion -KBNumber $kb -Severity "Critical"
            } elseif ($hasImportant) {
                $importantKBs += $kb
                Add-ToNotion -KBNumber $kb -Severity "Important"
            } elseif ($hasModerate) {
                $moderateKBs += $kb
                Add-ToNotion -KBNumber $kb -Severity "Moderate"
            } elseif ($hasLow) {
                $lowKBs += $kb
                Add-ToNotion -KBNumber $kb -Severity "Low"
            }
        }

        # Fixed string formatting
        $criticalKBs | ForEach-Object { Write-Host ($_ + " - Critical") -ForegroundColor Red }
        $importantKBs | ForEach-Object { Write-Host ($_ + " - Important") -ForegroundColor Yellow }
        $moderateKBs | ForEach-Object { Write-Host ($_ + " - Moderate") -ForegroundColor DarkYellow }
        $lowKBs | ForEach-Object { Write-Host ($_ + " - Low") -ForegroundColor Gray }
    }
    catch {
        Throw "Error retrieving security updates: $($_.Exception.Message)"
    }
}

function Extract-UniqueKBs {
    param (
        [Parameter(Mandatory=$true)]
        $reportData
    )
    $content = $reportData | ConvertTo-Json -Compress -Depth 5
    $matches = [regex]::Matches($content, '\bKB\d+\b')
    $uniqueKBs = $matches | ForEach-Object { $_.Value } | Sort-Object -Unique
    $uniqueKBs | Out-File -FilePath $Global:UniqueKBFilePath
}

# Clear the console and run the script
Clear-Host
Write-Host "MSRC Severity Results:"
Query-KBSeverity