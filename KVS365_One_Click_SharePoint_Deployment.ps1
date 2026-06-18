#requires -Version 7.4
<#
.SYNOPSIS
KVS365 one-click SharePoint deployment script for small-business client sites.

.DESCRIPTION
Creates a modern SharePoint Team Site, creates document libraries, applies versioning,
creates a standard folder structure, creates SharePoint groups, and optionally breaks
library inheritance to assign library-level permissions.

.NOTES
- Built around PnP.PowerShell.
- Designed for interactive use first.
- Update the variables in the CONFIG section before running.
#>

$ErrorActionPreference = 'Stop'

# =========================
# CONFIG
# =========================
$TenantName = 'contoso'
$TenantAdminUrl = "https://$TenantName-admin.sharepoint.com"
$TenantRootUrl = "https://$TenantName.sharepoint.com"

$SiteTitle = 'Acme Ltd'
$SiteAlias = 'acmeltd'
$SiteDescription = 'KVS365 standard collaboration site'
$SiteOwner = 'admin@contoso.co.uk'
$SiteUrl = "$TenantRootUrl/sites/$SiteAlias"

# If you want to create separate document libraries instead of using only Shared Documents,
# leave the array below as-is. If you want a single Documents library only, set $CreateExtraLibraries = $false.
$CreateExtraLibraries = $true
$Libraries = @(
    @{ Title = 'Company Admin'; Url = 'company-admin'; UniquePermissions = $true; Group = 'Company Owners'; Role = 'Contribute' },
    @{ Title = 'Finance';       Url = 'finance';       UniquePermissions = $true; Group = 'Finance Team';   Role = 'Contribute' },
    @{ Title = 'HR';            Url = 'hr';            UniquePermissions = $true; Group = 'HR Team';        Role = 'Contribute' },
    @{ Title = 'Sales & Marketing'; Url = 'sales-marketing'; UniquePermissions = $false; Group = 'Company Members'; Role = 'Contribute' },
    @{ Title = 'Operations';    Url = 'operations';    UniquePermissions = $false; Group = 'Company Members'; Role = 'Contribute' },
    @{ Title = 'Projects';      Url = 'projects';      UniquePermissions = $false; Group = 'Company Members'; Role = 'Contribute' },
    @{ Title = 'Archive';       Url = 'archive';       UniquePermissions = $false; Group = 'Company Members'; Role = 'Read' }
)

# Standard folder structure if you prefer using the default Shared Documents library.
$SharedDocumentsFolders = @(
    '01 - Company Admin',
    '02 - Finance',
    '03 - HR',
    '04 - Sales & Marketing',
    '05 - Operations',
    '06 - Projects',
    '99 - Archive'
)

$SubFolders = @{
    '01 - Company Admin' = @('Policies','Contracts','Insurance','Company Docs')
    '02 - Finance' = @('Invoices','Expenses','Payroll','VAT','Accounts')
    '03 - HR' = @('Employee Files','Recruitment','Training Records','Policies')
    '04 - Sales & Marketing' = @('Proposals','Quotes','Marketing Assets','Website Content')
    '05 - Operations' = @('Procedures','Templates','Supplier Info','Internal Docs')
    '06 - Projects' = @('Internal','Client A','Client B')
    '99 - Archive' = @('Closed Projects','Legacy Data')
}

$SharePointGroups = @(
    @{ Title = 'Company Owners'; Description = 'Owners of the site' },
    @{ Title = 'Company Members'; Description = 'Standard contributors' },
    @{ Title = 'Company Visitors'; Description = 'Read only visitors' },
    @{ Title = 'Finance Team'; Description = 'Finance library access' },
    @{ Title = 'HR Team'; Description = 'HR library access' }
)

# Optional initial members. Leave arrays empty if you want to add people manually later.
$GroupMembers = @{
    'Company Owners'   = @($SiteOwner)
    'Company Members'  = @()
    'Company Visitors' = @()
    'Finance Team'     = @()
    'HR Team'          = @()
}

# Site sharing baseline.
# Accepted values depend on tenant policy; typical examples include Disabled, ExistingExternalUserSharingOnly,
# ExternalUserAndGuestSharing, ExternalUserSharingOnly.
$SiteSharing = 'Disabled'

# Library version history baseline
$EnableVersioning = $true
$MajorVersions = 50

# =========================
# FUNCTIONS
# =========================
function Write-Stage {
    param([string]$Message)
    Write-Host "`n=== $Message ===" -ForegroundColor Cyan
}

function Ensure-Module {
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        Write-Stage 'Installing PnP.PowerShell'
        Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force
    }
}

function Ensure-Group {
    param(
        [string]$Title,
        [string]$Description
    )

    $existing = Get-PnPGroup | Where-Object { $_.Title -eq $Title }
    if (-not $existing) {
        Write-Host "Creating SharePoint group: $Title"
        New-PnPGroup -Title $Title -Description $Description | Out-Null
    }
    else {
        Write-Host "Group already exists: $Title"
    }
}

function Ensure-GroupMembers {
    param(
        [string]$GroupTitle,
        [string[]]$Members
    )

    foreach ($member in $Members) {
        if ([string]::IsNullOrWhiteSpace($member)) { continue }
        Write-Host "Adding $member to $GroupTitle"
        try {
            Add-PnPGroupMember -Group $GroupTitle -LoginName $member -ErrorAction Stop
        }
        catch {
            Write-Warning "Could not add $member to $GroupTitle. Review login name/UPN. $_"
        }
    }
}

function Ensure-Library {
    param(
        [hashtable]$Library
    )

    $existing = Get-PnPList | Where-Object { $_.Title -eq $Library.Title }
    if (-not $existing) {
        Write-Host "Creating library: $($Library.Title)"
        New-PnPList -Title $Library.Title -Url $Library.Url -Template DocumentLibrary -EnableVersioning:$EnableVersioning -OnQuickLaunch | Out-Null
    }
    else {
        Write-Host "Library already exists: $($Library.Title)"
    }

    Set-PnPList -Identity $Library.Title -EnableVersioning:$EnableVersioning -MajorVersions $MajorVersions | Out-Null

    if ($Library.UniquePermissions) {
        Write-Host "Applying unique permissions to library: $($Library.Title)"
        Set-PnPList -Identity $Library.Title -BreakRoleInheritance -CopyRoleAssignments | Out-Null
        Set-PnPListPermission -Identity $Library.Title -Group $Library.Group -AddRole $Library.Role | Out-Null
    }
}

function Ensure-FolderTree {
    param(
        [string]$RootLibrary = 'Shared Documents'
    )

    foreach ($folder in $SharedDocumentsFolders) {
        Write-Host "Ensuring folder: $folder"
        try {
            Add-PnPFolder -Name $folder -Folder $RootLibrary | Out-Null
        }
        catch {
            Write-Host "Folder already present or cannot be created now: $folder"
        }

        if ($SubFolders.ContainsKey($folder)) {
            foreach ($sub in $SubFolders[$folder]) {
                Write-Host "Ensuring subfolder: $folder/$sub"
                try {
                    Add-PnPFolder -Name $sub -Folder "$RootLibrary/$folder" | Out-Null
                }
                catch {
                    Write-Host "Subfolder already present or cannot be created now: $folder/$sub"
                }
            }
        }
    }
}

# =========================
# RUN
# =========================
Ensure-Module

Write-Stage 'Connecting to SharePoint admin centre'
Connect-PnPOnline -Url $TenantAdminUrl -Interactive

Write-Stage 'Creating or checking team site'
try {
    $tenantSite = Get-PnPTenantSite -Identity $SiteUrl -ErrorAction Stop
    Write-Host "Site already exists: $SiteUrl"
}
catch {
    New-PnPSite -Type TeamSite -Title $SiteTitle -Alias $SiteAlias -Description $SiteDescription -Owners $SiteOwner -Wait | Out-Null
    Write-Host "Site created: $SiteUrl"
}

Write-Stage 'Connecting to site'
Connect-PnPOnline -Url $SiteUrl -Interactive

Write-Stage 'Applying site baseline'
try {
    Set-PnPSite -Identity $SiteUrl -Sharing $SiteSharing | Out-Null
}
catch {
    Write-Warning "Could not set site sharing to '$SiteSharing'. Tenant policy may block or override this. $_"
}

Write-Stage 'Creating SharePoint groups'
foreach ($group in $SharePointGroups) {
    Ensure-Group -Title $group.Title -Description $group.Description
}

Write-Stage 'Populating SharePoint groups'
foreach ($groupName in $GroupMembers.Keys) {
    Ensure-GroupMembers -GroupTitle $groupName -Members $GroupMembers[$groupName]
}

Write-Stage 'Setting up content structure'
if ($CreateExtraLibraries) {
    foreach ($library in $Libraries) {
        Ensure-Library -Library $library
    }
}
else {
    Set-PnPList -Identity 'Documents' -EnableVersioning:$EnableVersioning -MajorVersions $MajorVersions | Out-Null
    Ensure-FolderTree -RootLibrary 'Shared Documents'
}

Write-Stage 'Finished'
Write-Host "Deployment complete for $SiteTitle ($SiteUrl)" -ForegroundColor Green
Write-Host 'Next checks:' -ForegroundColor Yellow
Write-Host '1. Confirm the site homepage and navigation.'
Write-Host '2. Confirm group membership.'
Write-Host '3. Test file upload in each library.'
Write-Host '4. If using folders instead of separate libraries, show users OneDrive vs SharePoint clearly.'
