<# 
Module: SCCMCollectionTools
Description: PowerShell tools for automating SCCM collection management tasks.
Author: YourName
#>

function Get-YearMonth {
<#
.SYNOPSIS
Gets the current year and month in 'YYYY-MM' format.

.DESCRIPTION
This function returns the current system date formatted as 'YYYY-MM'. Useful for consistent naming like folder structures or SCCM collections.

.EXAMPLE
Get-YearMonth
#>
    $now = Get-Date
    $year = $now.Year
    $month = $now.Month.ToString("00")
    return "$year-$month"
}

function Get-AbbreviatedMonthName {
<#
.SYNOPSIS
Gets the abbreviated month name from a DateTime object.

.DESCRIPTION
Takes a [datetime] input and returns the three-letter abbreviated month name (e.g., Jan, Feb).

.EXAMPLE
Get-AbbreviatedMonthName -date (Get-Date)
#>
    param([datetime]$date)
    return $date.ToString("MMM")
}

function Get-SCCMSiteCode {
<#
.SYNOPSIS
Fetches the SCCM Site Code using WMI.

.DESCRIPTION
Retrieves the local SCCM Site Code from WMI. This is useful for automating SCCM operations where the site code is needed.

.EXAMPLE
Get-SCCMSiteCode
#>
    try {
        $siteCode = Get-WmiObject -Namespace "Root\\SMS" -Class SMS_ProviderLocation -ComputerName "." | Select-Object -ExpandProperty SiteCode
        if ($siteCode -is [array]) { return $siteCode[0] }
        return $siteCode
    } catch {
        Write-Error "Error retrieving SCCM Site Code: $_"
        return $null
    }
}

function Get-SystemFQDN {
<#
.SYNOPSIS
Gets the Fully Qualified Domain Name (FQDN) of the local machine.

.DESCRIPTION
This function resolves the local system's hostname into its fully qualified domain name (FQDN).

.EXAMPLE
Get-SystemFQDN
#>
    try {
        return [System.Net.Dns]::GetHostByName($env:COMPUTERNAME).HostName
    } catch {
        Write-Error "Error retrieving FQDN: $_"
        return $null
    }
}

function Import-SCCMModule {
<#
.SYNOPSIS
Imports the SCCM PowerShell module if it's not already imported.

.DESCRIPTION
Ensures the ConfigurationManager module is available to use SCCM cmdlets in PowerShell.

.EXAMPLE
Import-SCCMModule
#>
    if ((Get-Module ConfigurationManager) -eq $null) {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\\..\\ConfigurationManager.psd1"
    }
}

function Mount-SCCMDrive {
<#
.SYNOPSIS
Mounts the SCCM PSDrive for the specified site code and provider.

.DESCRIPTION
Creates a new PSDrive connected to the SCCM site server for executing SCCM cmdlets directly.

.EXAMPLE
Mount-SCCMDrive -SiteCode "ABC" -ProviderMachineName "Server01.domain.local"
#>
    param($SiteCode, $ProviderMachineName)
    if ((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName
    }
    Set-Location "$($SiteCode):\\"
}

function Create-Folder {
<#
.SYNOPSIS
Creates a folder if it does not already exist.

.DESCRIPTION
Creates a folder at the specified path with the given name. Outputs a status message.

.EXAMPLE
Create-Folder -Path "C:\\Temp" -Name "Logs"
#>
    param([string]$Path, [string]$Name)
    $fullPath = Join-Path $Path $Name
    if (-not (Test-Path $fullPath)) {
        New-Item -Path $Path -Name $Name -ItemType Directory
        Write-Host "Folder '$Name' created at '$Path'." -ForegroundColor Green
    } else {
        Write-Host "Folder '$Name' already exists at '$Path'." -ForegroundColor Yellow
    }
}

function Create-And-MoveCMCollection {
<#
.SYNOPSIS
Creates a SCCM collection and moves it into a folder.

.DESCRIPTION
If the specified SCCM device collection does not exist, it is created and then moved to the designated folder.

.EXAMPLE
Create-And-MoveCMCollection -CollectionName "Test" -LimitingCollection "All Systems" -FolderPath "Collections\\TestFolder" -SiteCode "ABC"
#>
    param(
        [string]$CollectionName,
        [string]$LimitingCollection,
        [string]$FolderPath,
        [string]$SiteCode
    )
    $existingCollection = Get-CMDeviceCollection -Name $CollectionName
    if (-not $existingCollection) {
        New-CMDeviceCollection -Name $CollectionName -LimitingCollectionName $LimitingCollection > $null
        Write-Host "Created '$CollectionName'" -ForegroundColor Green
    } else {
        Write-Host "'$CollectionName' already exists." -ForegroundColor Yellow
    }
    $collectionObject = Get-CMDeviceCollection -Name $CollectionName
    Move-CMObject -FolderPath "$SiteCode:\\$FolderPath" -InputObject $collectionObject
}

function Add-CMDeviceCollectionQueryMembershipRuleWithQuery {
<#
.SYNOPSIS
Adds a query-based membership rule to an SCCM collection.

.DESCRIPTION
Attaches a WQL query rule to a device collection to dynamically populate membership based on system properties.

.EXAMPLE
Add-CMDeviceCollectionQueryMembershipRuleWithQuery -CollectionName "Test" -QueryExpression "SELECT * FROM SMS_R_System" -RuleName "All Devices"
#>
    param([string]$CollectionName, [string]$QueryExpression, [string]$RuleName)
    Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionName -QueryExpression $QueryExpression -RuleName $RuleName
}

function Add-IncludeCollection {
<#
.SYNOPSIS
Includes one or more collections in another SCCM collection.

.DESCRIPTION
Adds include membership rules to an SCCM collection by referencing other collections.

.EXAMPLE
Add-IncludeCollection -SelectCollectionName "Main" -IncludeCollectionNames @("Sub1", "Sub2")
#>
    param([string]$SelectCollectionName, [string[]]$IncludeCollectionNames)
    $SelectCollection = Get-CMDeviceCollection -Name $SelectCollectionName
    foreach ($IncludeCollectionName in $IncludeCollectionNames) {
        $IncludeCollection = Get-CMDeviceCollection -Name $IncludeCollectionName
        if ($IncludeCollection) {
            Add-CMDeviceCollectionIncludeMembershipRule -CollectionId $SelectCollection.CollectionId -IncludeCollectionId $IncludeCollection.CollectionId
            Write-Host "Included '$IncludeCollectionName' in '$SelectCollectionName'" -ForegroundColor Green
        }
    }
}

function Add-ExcludeCollection {
<#
.SYNOPSIS
Excludes one or more collections from another SCCM collection.

.DESCRIPTION
Adds exclude membership rules to an SCCM collection to remove systems in the listed collections.

.EXAMPLE
Add-ExcludeCollection -SelectCollectionName "Main" -ExcludeCollectionNames @("Exclude1", "Exclude2")
#>
    param([string]$SelectCollectionName, [string[]]$ExcludeCollectionNames)
    $SelectCollection = Get-CMDeviceCollection -Name $SelectCollectionName
    foreach ($ExcludeCollectionName in $ExcludeCollectionNames) {
        $ExcludeCollection = Get-CMDeviceCollection -Name $ExcludeCollectionName
        if ($ExcludeCollection) {
            Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $SelectCollection.CollectionId -ExcludeCollectionId $ExcludeCollection.CollectionId
            Write-Host "Excluded '$ExcludeCollectionName' from '$SelectCollectionName'" -ForegroundColor Green
        }
    }
}
