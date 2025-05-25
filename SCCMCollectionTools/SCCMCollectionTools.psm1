function Get-YearMonth {
    $now = Get-Date
    $year = $now.Year
    $month = $now.Month.ToString(00)
    return $year-$month
}

function Get-AbbreviatedMonthName {
    param([datetime]$date)
    return $date.ToString(MMM)
}

function Get-SCCMSiteCode {
    try {
        $siteCode = Get-WmiObject -Namespace RootSMS -Class SMS_ProviderLocation -ComputerName .  Select-Object -ExpandProperty SiteCode
        if ($siteCode -is [array]) { return $siteCode[0] }
        return $siteCode
    } catch {
        Write-Error Error retrieving SCCM Site Code $_
        return $null
    }
}

function Get-SystemFQDN {
    try {
        return [System.Net.Dns]GetHostByName($envCOMPUTERNAME).HostName
    } catch {
        Write-Error Error retrieving FQDN $_
        return $null
    }
}

function Import-SCCMModule {
    if ((Get-Module ConfigurationManager) -eq $null) {
        Import-Module $($ENVSMS_ADMIN_UI_PATH)..ConfigurationManager.psd1
    }
}

function Mount-SCCMDrive {
    param($SiteCode, $ProviderMachineName)
    if ((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName
    }
    Set-Location $($SiteCode)
}

function Create-Folder {
    param([string]$Path, [string]$Name)
    $fullPath = Join-Path $Path $Name
    if (-not (Test-Path $fullPath)) {
        New-Item -Path $Path -Name $Name -ItemType Directory
        Write-Host Folder '$Name' created at '$Path'. -ForegroundColor Green
    } else {
        Write-Host Folder '$Name' already exists at '$Path'. -ForegroundColor Yellow
    }
}

function Create-And-MoveCMCollection {
    param(
        [string]$CollectionName,
        [string]$LimitingCollection,
        [string]$FolderPath,
        [string]$SiteCode
    )
    $existingCollection = Get-CMDeviceCollection -Name $CollectionName
    if (-not $existingCollection) {
        New-CMDeviceCollection -Name $CollectionName -LimitingCollectionName $LimitingCollection  $null
        Write-Host Created '$CollectionName' -ForegroundColor Green
    } else {
        Write-Host '$CollectionName' already exists. -ForegroundColor Yellow
    }
    $collectionObject = Get-CMDeviceCollection -Name $CollectionName
    Move-CMObject -FolderPath $SiteCode$FolderPath -InputObject $collectionObject
}

function Add-CMDeviceCollectionQueryMembershipRuleWithQuery {
    param([string]$CollectionName, [string]$QueryExpression, [string]$RuleName)
    Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionName -QueryExpression $QueryExpression -RuleName $RuleName
}

function Add-IncludeCollection {
    param([string]$SelectCollectionName, [string[]]$IncludeCollectionNames)
    $SelectCollection = Get-CMDeviceCollection -Name $SelectCollectionName
    foreach ($IncludeCollectionName in $IncludeCollectionNames) {
        $IncludeCollection = Get-CMDeviceCollection -Name $IncludeCollectionName
        if ($IncludeCollection) {
            Add-CMDeviceCollectionIncludeMembershipRule -CollectionId $SelectCollection.CollectionId -IncludeCollectionId $IncludeCollection.CollectionId
            Write-Host Included '$IncludeCollectionName' in '$SelectCollectionName' -ForegroundColor Green
        }
    }
}

function Add-ExcludeCollection {
    param([string]$SelectCollectionName, [string[]]$ExcludeCollectionNames)
    $SelectCollection = Get-CMDeviceCollection -Name $SelectCollectionName
    foreach ($ExcludeCollectionName in $ExcludeCollectionNames) {
        $ExcludeCollection = Get-CMDeviceCollection -Name $ExcludeCollectionName
        if ($ExcludeCollection) {
            Add-CMDeviceCollectionExcludeMembershipRule -CollectionId $SelectCollection.CollectionId -ExcludeCollectionId $ExcludeCollection.CollectionId
            Write-Host Excluded '$ExcludeCollectionName' from '$SelectCollectionName' -ForegroundColor Green
        }
    }
}
