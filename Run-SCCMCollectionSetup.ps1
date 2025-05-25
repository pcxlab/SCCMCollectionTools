Import-Module .\SCCMCollectionTools\SCCMCollectionTools.psm1

$yearMonth = Get-YearMonth
$SiteCode = Get-SCCMSiteCode
$ProviderMachineName = Get-SystemFQDN

Import-SCCMModule
Mount-SCCMDrive -SiteCode $SiteCode -ProviderMachineName $ProviderMachineName

# Example: Create Folder
Create-Folder -Path "$($SiteCode):\DeviceCollection" -Name "Security Management Services"
