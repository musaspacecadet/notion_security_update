# MSRC KB Tracking and Reporting

This PowerShell script tracks and reports Microsoft Security Response Center (MSRC) updates for specific Knowledge Base (KB) articles, managing critical security information in a Notion database. It automatically clears old records, queries current KB severities, and updates Notion with severity details.

## Key Features

- **Track KBs**: Specify target KBs for monitoring or query all unique KBs for the specified report date.
- **Notion Integration**: Add KB severities to a Notion database automatically.
- **Severity Levels**: Supports Critical, Important, Moderate, and Low severity levels.
- **Caching**: Caches results to avoid repetitive API calls.
- **Cross-Platform Paths**: Detects OS and sets paths for caching and file storage.

## Prerequisites

1. **PowerShell 5.1+**: Ensure PowerShell is installed.
2. **MSRCSecurityUpdates Module**: The script uses `Get-MsrcCvrfDocument` for querying KB information.
3. **Notion Integration**: Set up a Notion integration token and database ID.

## Global Variables

| Variable               | Description                                                                                      |
|------------------------|--------------------------------------------------------------------------------------------------|
| `$Global:TargetKBs`    | Array of KB numbers to track, e.g., `("KB5041585", "KB5040442")`.                               |
| `$Global:QueryAllUniqueKBs` | Set to `$true` to automatically query and track all unique KBs from the current report.    |
| `$Global:ForceUpdate`  | Set to `$true` to force data refresh from MSRC instead of using cached data.                    |
| `$Global:ShowCachedResults` | Set to `$true` to display cached results without querying MSRC.                            |
| `$Global:ReportDate`   | Specifies the report date, formatted as `yyyy-MMM` (e.g., `2024-Nov`).                          |
| `$Global:NotionSecret` | Your Notion API secret token.                                                                   |
| `$Global:NotionDatabaseId` | Notion database ID to store KB data.                                                        |

## File Paths

- **Windows**: `C:\temp\SecurityUpdateCache_{ReportDate}.json`
- **Linux/macOS**: `/tmp/SecurityUpdateCache_{ReportDate}.json`

## Functions

1. **Add-ToNotion**: Adds KB and severity data to Notion.
2. **Clear-NotionDatabase**: Clears all records from the specified Notion database.
3. **Query-KBSeverity**: Main function that:
   - Clears Notion records.
   - Queries cached results or updates from MSRC.
   - Adds KBs and severities to Notion based on `TargetKBs` or all unique KBs.
4. **Extract-UniqueKBs**: Extracts all KBs from report data and saves unique values to a file.

## Setup Instructions

1. **Install MSRCSecurityUpdates**:
   ```powershell
   Install-Module -Name MSRCSecurityUpdates -Force
   ```

2. **Configure Notion API**:
   - Set `Global:NotionSecret` with your Notion API key.
   - Set `Global:NotionDatabaseId` with your Notion database ID.

3. **Run the Script**:
   ```powershell
   .\security-updates.ps1
   ```

## Example Usage

```powershell
$Global:TargetKBs = @("KB5041585", "KB5040442")
$Global:ForceUpdate = $true
Query-KBSeverity
```

Sample output:

```plaintext
KB5041585 - Critical
KB5040442 - Important
Successfully added KB5041585 with severity Critical to Notion
Successfully added KB5040442 with severity Important to Notion
```

## Error Handling

Logs errors if:
- The MSRC query fails.
- Notion API calls fail (e.g., due to incorrect permissions or token).

## License

MIT License
