# CopilotAppAdoptionReport

## À quoi sert cet outil et pourquoi maintenant

Microsoft ajoute les modèles Grok de SpaceXAI dans Copilot. C'est une préversion offerte via le programme Microsoft Frontier et limitée à Word, Excel et PowerPoint au départ. Pour rendre Grok accessible, SpaceXAI a été ajouté à la liste des sous-traitants des services en ligne de Microsoft. L'accès est géré par un réglage administratif dédié désactivé par défaut. Activer un sous-traitant est une décision de gouvernance des données et non un simple bouton de fonctionnalité.

Avant d'ouvrir un nouveau modèle, ça vaut la peine de savoir qui se sert déjà de Copilot dans Word, Excel et PowerPoint. Cet outil lit le rapport d'utilisation de Copilot par utilisateur sur 30 jours dans Microsoft Graph. Il fonctionne en lecture seule. Il n'active aucun sous-traitant et ne modifie aucun réglage du locataire. Le rapport présente l'activité par application. Il ne distingue pas le modèle utilisé, donc tu ne verras pas qui utilise Grok, tu verras plutôt où Copilot sert déjà.

## Prérequis

- Le module Microsoft.Graph.Authentication
- La permission Microsoft Graph Reports.Read.All. C'est la permission minimale indiquée dans l'article
- Une session Microsoft Graph ouverte au préalable avec cette permission
- Un accès au point d'accès bêta de Microsoft Graph, car ce rapport existe seulement sur bêta pour l'instant

## Installation

Clone le dépôt puis importe le module.

```powershell
git clone <adresse-du-depot>
Import-Module .\CopilotAppAdoptionReport\CopilotAppAdoptionReport.psd1
```

## Exemples d'utilisation

Ouvre d'abord une session Microsoft Graph.

```powershell
Connect-MgGraph -Scopes 'Reports.Read.All'
```

Récupère le détail par utilisateur.

```powershell
$rapport = Get-CopilotAppAdoptionReport
$rapport.Count
```

Résume l'adoption dans Word, Excel et PowerPoint.

```powershell
$rapport | Get-CopilotAppAdoptionSummary
```

Écris un rapport HTML simple à partir du résumé.

```powershell
$rapport |
    Get-CopilotAppAdoptionSummary |
    Export-CopilotAppAdoptionReport -Path .\adoption-copilot.html
```

Vérifie l'action d'écriture sans créer le fichier.

```powershell
Get-CopilotAppAdoptionSummary -InputObject $rapport |
    Export-CopilotAppAdoptionReport -Path .\adoption-copilot.html -WhatIf
```

## Exemple de sortie

Les valeurs ci-dessous sont fictives et présentées comme telles. Elles illustrent la forme du résumé et ne proviennent d'aucun locataire réel.

```
Application UtilisateursActifs UtilisateursTotal
----------- ------------------ -----------------
Word                        42               120
Excel                       73               120
PowerPoint                  28               120
```

## Limites connues

- Le rapport existe seulement sur le point d'accès bêta de Microsoft Graph pour l'instant. Le format peut changer
- Le rapport présente l'activité par application. Il ne distingue pas le modèle utilisé (Grok, Claude ou GPT)
- Le résumé détecte les colonnes de Word, Excel et PowerPoint par correspondance de nom sur le rapport en direct. Si Microsoft renomme ces champs, la détection pourrait ne rien trouver
- La période est fixée à 30 jours (D30)
- L'outil est en lecture seule. Il n'active aucun sous-traitant et ne modifie aucun réglage du locataire

## Article et sources

- Article de Charles Jenkins : https://trivium365.com/blog/grok-copilot-preversion-frontier-sous-traitant/
- Microsoft Copilot Blog, Expanding model choice in Copilot with Grok : https://techcommunity.microsoft.com/blog/microsoft-copilot-blog/expanding-model-choice-in-copilot-with-grok/4555749


## Démonstration en vidéo

https://www.youtube.com/watch?v=sQwNRK1ixqU
