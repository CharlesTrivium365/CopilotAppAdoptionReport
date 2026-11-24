#Requires -Modules Microsoft.Graph.Authentication

function Get-CopilotAppAdoptionReport {
    <#
    .SYNOPSIS
        Lit le rapport d'utilisation de Microsoft Copilot par utilisateur sur 30 jours.

    .DESCRIPTION
        Interroge le point d'acces beta de Microsoft Graph
        getMicrosoft365CopilotUsageUserDetail(period='D30') en lecture seule et
        retourne un objet par utilisateur. Le rapport presente l'activite par
        application. Il ne distingue pas le modele utilise (Grok, Claude ou GPT).
        Une session Microsoft Graph avec la permission Reports.Read.All doit etre
        ouverte au prealable avec Connect-MgGraph. La commande ne modifie aucun
        reglage du locataire.

    .EXAMPLE
        Connect-MgGraph -Scopes 'Reports.Read.All'
        $rapport = Get-CopilotAppAdoptionReport
        $rapport.Count

        Ouvre une session puis recupere les enregistrements et affiche le nombre
        d'utilisateurs retournes.

    .OUTPUTS
        System.Management.Automation.PSCustomObject
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    $context = Get-MgContext
    if (-not $context) {
        throw "Aucune session Microsoft Graph active. Executez d'abord Connect-MgGraph -Scopes 'Reports.Read.All'."
    }

    if ($context.Scopes -notcontains 'Reports.Read.All') {
        Write-Warning "La permission Reports.Read.All n'apparait pas dans la session courante. L'appel pourrait echouer."
    }

    $uri = "https://graph.microsoft.com/beta/reports/getMicrosoft365CopilotUsageUserDetail(period='D30')?`$format=application/json"

    $records = [System.Collections.Generic.List[psobject]]::new()
    $next = $uri

    while (-not [string]::IsNullOrEmpty($next)) {
        Write-Verbose "Appel de Microsoft Graph : $next"
        $response = Invoke-MgGraphRequest -Method GET -Uri $next

        foreach ($item in $response.value) {
            $records.Add([pscustomobject]$item)
        }

        $next = $response.'@odata.nextLink'
    }

    Write-Verbose "Nombre d'utilisateurs retournes : $($records.Count)"

    $records
}

function Get-CopilotAppAdoptionSummary {
    <#
    .SYNOPSIS
        Resume l'adoption de Copilot dans Word, Excel et PowerPoint.

    .DESCRIPTION
        Parcourt les enregistrements produits par Get-CopilotAppAdoptionReport et
        compte, pour chacune des trois applications visees par la preversion Grok
        (Word, Excel et PowerPoint), le nombre d'utilisateurs ayant une date
        d'activite renseignee. Les colonnes d'activite sont detectees par
        correspondance de nom sur le rapport en direct. Aucune donnee du locataire
        n'est modifiee.

    .PARAMETER InputObject
        Un ou plusieurs enregistrements d'utilisateur retournes par
        Get-CopilotAppAdoptionReport. Accepte l'entree de pipeline.

    .EXAMPLE
        $rapport = Get-CopilotAppAdoptionReport
        $rapport | Get-CopilotAppAdoptionSummary

        Affiche le nombre d'utilisateurs actifs par application.

    .OUTPUTS
        System.Management.Automation.PSCustomObject
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [psobject[]]
        $InputObject
    )

    begin {
        $rows = [System.Collections.Generic.List[psobject]]::new()
    }

    process {
        foreach ($item in $InputObject) {
            $rows.Add($item)
        }
    }

    end {
        $applications = 'Word', 'Excel', 'PowerPoint'

        foreach ($application in $applications) {
            $activeCount = 0

            foreach ($row in $rows) {
                $activityProperties = $row.PSObject.Properties | Where-Object {
                    $_.Name -match $application -and $_.Name -match 'activity'
                }

                $hasActivity = $false
                foreach ($property in $activityProperties) {
                    if (-not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
                        $hasActivity = $true
                    }
                }

                if ($hasActivity) {
                    $activeCount++
                }
            }

            [pscustomobject]@{
                Application = $application
                UtilisateursActifs = $activeCount
                UtilisateursTotal = $rows.Count
            }
        }
    }
}

function Export-CopilotAppAdoptionReport {
    <#
    .SYNOPSIS
        Ecrit un rapport HTML simple a partir des objets fournis.

    .DESCRIPTION
        Convertit en HTML les objets recus (detail par utilisateur ou resume) et
        ecrit le fichier a l'emplacement indique. L'ecriture passe par
        SupportsShouldProcess, donc -WhatIf montre l'action sans rien ecrire. La
        commande n'interroge ni ne modifie le locataire.

    .PARAMETER InputObject
        Les objets a inclure dans le rapport. Accepte l'entree de pipeline.

    .PARAMETER Path
        Chemin du fichier HTML a ecrire.

    .EXAMPLE
        $rapport = Get-CopilotAppAdoptionReport
        $rapport |
            Get-CopilotAppAdoptionSummary |
            Export-CopilotAppAdoptionReport -Path .\adoption-copilot.html

        Ecrit le resume d'adoption dans un fichier HTML local.

    .EXAMPLE
        Get-CopilotAppAdoptionSummary -InputObject $rapport |
            Export-CopilotAppAdoptionReport -Path .\adoption-copilot.html -WhatIf

        Montre l'ecriture prevue sans creer le fichier.

    .OUTPUTS
        System.IO.FileInfo
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([System.IO.FileInfo])]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [psobject[]]
        $InputObject,

        [Parameter(Mandatory = $true)]
        [string]
        $Path
    )

    begin {
        $rows = [System.Collections.Generic.List[psobject]]::new()
    }

    process {
        foreach ($item in $InputObject) {
            $rows.Add($item)
        }
    }

    end {
        $head = '<style>body{font-family:Segoe UI,Arial,sans-serif;margin:1em}table{border-collapse:collapse}th,td{border:1px solid #cccccc;padding:6px;text-align:left}</style>'
        $preface = '<h1>Rapport d''adoption de Microsoft Copilot par application</h1><p>Lecture seule. Point d''acces beta de Microsoft Graph, periode D30. Le rapport ne distingue pas le modele utilise.</p>'
        $html = $rows | ConvertTo-Html -Head $head -PreContent $preface

        if ($PSCmdlet.ShouldProcess($Path, 'Ecrire le rapport HTML')) {
            $html | Out-File -FilePath $Path -Encoding utf8
            Get-Item -LiteralPath $Path
        }
    }
}

Export-ModuleMember -Function @(
    'Get-CopilotAppAdoptionReport',
    'Get-CopilotAppAdoptionSummary',
    'Export-CopilotAppAdoptionReport'
)
