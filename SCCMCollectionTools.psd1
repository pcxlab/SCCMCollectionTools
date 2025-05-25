@{
    RootModule         = 'SCCMCollectionTools.psm1'
    ModuleVersion      = '1.0.0'
    GUID               = 'd1234567-abcd-4ef0-9999-000000000000'
    Author             = 'pcxlab'
    CompanyName        = 'YourCompany'
    Copyright          = '(c) 2025 pcxlab. All rights reserved.'
    Description        = 'A module to manage SCCM collections with reusable automation functions.'

    PowerShellVersion  = '5.1'
    FunctionsToExport  = @(
        'Get-YearMonth',
        'Get-AbbreviatedMonthName',
        'Get-SCCMSiteCode',
        'Get-SystemFQDN',
        'Import-SCCMModule',
        'Mount-SCCMDrive',
        'Create-Folder',
        'Create-And-MoveCMCollection',
        'Add-CMDeviceCollectionQueryMembershipRuleWithQuery',
        'Add-IncludeCollection',
        'Add-ExcludeCollection'
    )

    CmdletsToExport    = @()
    VariablesToExport  = @()
    AliasesToExport    = @()
    PrivateData        = @{}

    HelpInfoURI        = 'httpsgithub.compcxlabSCCMCollectionTools'
}
