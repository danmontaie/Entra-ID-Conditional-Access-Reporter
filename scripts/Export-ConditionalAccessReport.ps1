# ============================================================================
# Export Microsoft Entra Conditional Access Report
#
# Description:
# Connects to Microsoft Graph and exports Microsoft Entra
# Conditional Access policy information.
#
# Information Collected:
# - Export Date
# - Policy Name
# - Policy ID
# - State
# - Included Users
# - Excluded Users
# - Included Groups
# - Excluded Groups
# - Included Applications
# - Excluded Applications
# - Client App Types
# - Included Platforms
# - Excluded Platforms
# - Grant Controls
# - Session Controls
#
# Requirements:
#   Microsoft Graph PowerShell SDK
#
# Required Microsoft Graph Permissions:
#   - Policy.Read.All
#   - Policy.Read.ConditionalAccess
#   - Directory.Read.All
#
# Output:
#   .\Reports\ConditionalAccessReport.csv
#
# ============================================================================

# Connect to Microsoft Graph

Connect-MgGraph -Scopes @(
    "Policy.Read.All",
    "Policy.Read.ConditionalAccess",
    "Directory.Read.All"
)

Write-Host ""
Write-Host "Connected to Microsoft Graph" -ForegroundColor Green

Write-Host ""
Write-Host "Getting Conditional Access Policies..." -ForegroundColor Cyan

try {
    $Policies = Get-MgIdentityConditionalAccessPolicy -ErrorAction Stop
}
catch {
    Write-Host "Failed to retrieve Conditional Access policies." -ForegroundColor Red
    Write-Host $_.Exception.Message
    return
}

Write-Host ""
Write-Host "Found $($Policies.Count) policies." -ForegroundColor Yellow

$Report = @()
$i = 0

foreach ($Policy in $Policies)
{
    $i++

    Write-Progress `
        -Activity "Building Conditional Access Report" `
        -Status $Policy.DisplayName `
        -PercentComplete (($i / $Policies.Count) * 100)

    Write-Host "Processing: $($Policy.DisplayName)" -ForegroundColor Cyan

    try
    {
        $Report += [PSCustomObject]@{

            ExportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

            PolicyName = $Policy.DisplayName

            PolicyId = $Policy.Id

            State = $Policy.State

            IncludedUsers = ($Policy.Conditions.Users.IncludeUsers -join ", ")

            ExcludedUsers = ($Policy.Conditions.Users.ExcludeUsers -join ", ")

            IncludedGroups = ($Policy.Conditions.Users.IncludeGroups -join ", ")

            ExcludedGroups = ($Policy.Conditions.Users.ExcludeGroups -join ", ")

            IncludedApplications = ($Policy.Conditions.Applications.IncludeApplications -join ", ")

            ExcludedApplications = ($Policy.Conditions.Applications.ExcludeApplications -join ", ")

            ClientAppTypes = ($Policy.Conditions.ClientAppTypes -join ", ")

            IncludedPlatforms = ($Policy.Conditions.Platforms.IncludePlatforms -join ", ")

            ExcludedPlatforms = ($Policy.Conditions.Platforms.ExcludePlatforms -join ", ")

            GrantControls = ($Policy.GrantControls.BuiltInControls -join ", ")

            SessionControls = ($Policy.SessionControls | Out-String).Trim()
        }
    }
    catch
    {
        Write-Host "Failed to process policy: $($Policy.DisplayName)" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor DarkRed
    }
}

Write-Progress -Activity "Building Conditional Access Report" -Completed

# Create Reports folder if needed

$OutputFolder = Join-Path $PSScriptRoot "..\Reports"

if (!(Test-Path $OutputFolder))
{
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

$CsvPath = Join-Path $OutputFolder "ConditionalAccessReport.csv"

$Report |
    Sort-Object PolicyName |
    Export-Csv `
        -Path $CsvPath `
        -NoTypeInformation `
        -Encoding UTF8

Write-Host ""
Write-Host "==========================================" -ForegroundColor Green
Write-Host "Conditional Access Report Complete" -ForegroundColor Green
Write-Host "==========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Policies Exported : $($Report.Count)"
Write-Host "CSV Location      : $CsvPath"
Write-Host ""

$Report |
    Sort-Object PolicyName |
    Format-Table -AutoSize
