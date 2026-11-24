@{
    RootModule = 'CopilotAppAdoptionReport.psm1'
    ModuleVersion = '1.0.0'
    GUID = 'c3f1a9d4-5b62-4e18-9a7f-2d8c6b4e1f90'
    Author = 'Charles Jenkins'
    CompanyName = 'Charles Jenkins'
    Copyright = '(c) 2026 Charles Jenkins. Tous droits reserves.'
    Description = 'Lit en lecture seule le rapport d''utilisation de Microsoft Copilot par utilisateur sur 30 jours via Microsoft Graph et resume l''adoption dans Word, Excel et PowerPoint.'
    PowerShellVersion = '5.1'
    RequiredModules = @('Microsoft.Graph.Authentication')
    FunctionsToExport = @('Get-CopilotAppAdoptionReport', 'Get-CopilotAppAdoptionSummary', 'Export-CopilotAppAdoptionReport')
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @()
    PrivateData = @{
        PSData = @{
            Tags = @('Copilot', 'Microsoft365', 'MicrosoftGraph', 'Reporting', 'Adoption')
        }
    }
}
