####Number of proposals posted to Clockify via Aktis integration
#(Get-ChildItem -Path "C:\Aktis\Clockify Integration\Logs\Posted*" | ForEach-Object {Get-Content -Path "$_"}).Count

####Number of proposals housed in Aktis
#(((Get-PnPListItem -List "Proposal Queues" -Fields "KTNumber").FieldValues).KTNumber).Count

function Write-AktisLog {
    Param
    (
         [Parameter(Mandatory=$true, Position=0)]
         [ValidateSet('Aktis Helper')]
         [string] $LogSource,
         [Parameter(Mandatory=$true, Position=1)]
         [ValidateSet('Information','Warning','Error')]
         [string] $LogType,
         [Parameter(Mandatory=$true, Position=2)]
         [string] $LogMessage
    )

    #New-EventLog -LogName AKTIS -Source AKTIS -ErrorAction SilentlyContinue | Out-Null
    #if(!(Get-EventLog -LogName AKTIS -Source 'Aktis Helper' -ErrorAction SilentlyContinue)){
    #    New-EventLog -Source 'Aktis Helper' -LogName AKTIS -ErrorAction SilentlyContinue
    #}
    Write-EventLog -Log AKTIS -Source "Aktis Helper" -EventID 0 -EntryType $LogType -Message "$LogMessage"
    #Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "This is a warning message..."
}

function AuthenticateToSPO {

    try{
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $Script:ConnectToSPO = Connect-PnPOnline -Url "https://projectfuelnow.sharepoint.com/sites/FuelNow" -Tenant "projectfuelnow.com" -ClientId 09b4038a-b7ed-4784-99ef-09fcfe12eaef -Thumbprint 440FC2B08CC7586C815CACE316B72C79C7FDDE21 -ReturnConnection
        #$Script:ConnectToSPO = Connect-PnPOnline -Url "https://projectfuelnow.sharepoint.com/sites/FuelNow" -Tenant "projectfuelnow.com" -ClientId 25ff3a8d-9601-450c-b585-7dd72dd5de66 -Thumbprint 24becf80e51ca4a7d8d8ff6260e648e4cf604c23 -ReturnConnection
        #$Script:ConnectToSPO = Connect-PnPOnline -Url "https://projectfuelnow.sharepoint.com/sites/Fuelnow" -Interactive -ForceAuthentication -ErrorAction SilentlyContinue
    }
    catch{
        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to authenticate to SharePoint Online..."
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to authenticate to SharePoint Online...`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
        EXIT
    }

    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Successfully authenticated to SharePoint Online!"
    Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Successfully authenticated to SharePoint Online!"

}

Function Test-FileLock {
    Param(
        [parameter(Mandatory=$True)]
        [string]$Path
    )
    $OFile = New-Object System.IO.FileInfo $Path
    if((Test-Path -Path $Path -PathType Leaf) -eq $False){
        Return $False
    }
    else{
        try{
            $OStream = $OFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            if($OStream){
                $OStream.Close()
            }
            Return $False
        } 
        catch{
            Return $True
        }
    }
}

function GetClockifyClientIDs {

    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Retreiving Clockify Client IDs..."

    try{
        $Script:ActureSolutionsClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "Acture Solutions"}).ID
        $Script:AffinityClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "Affinity Technology Partners"}).ID
        $Script:AMCClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "AMC"}).ID
        $Script:AntisynClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "Antisyn"}).ID
        $Script:ApticaClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "Aptica LLC"}).ID
        $Script:DenaliTEKClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "DenaliTEK"}).ID
        $Script:DominionClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "DominionTech"}).ID
        $Script:GroffClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "Groff Networks"}).ID
        $Script:Layer9ClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "Layer9"}).ID
        $Script:ManawaClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "Manawa"}).ID
        $Script:MentisClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "Mentis"}).ID
        $Script:MyITCrewNYClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "My IT Crew - NY"}).ID
        $Script:ITSClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "ITS"}).ID
        $Script:ProjectFuelClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "Project Fuel"}).ID
        $Script:SisAdminClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "SisAdmin"}).ID
        $Script:STAIClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "STAI"}).ID
        $Script:SteadyClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "Steady Networks"}).ID
        $Script:SyscomClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "Syscom"}).ID
        $Script:UniVistaClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "UniVista"}).ID
        $Script:ValleyExpetecClockifyClientID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/clients") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "Valley Expetec"}).ID
    }
    catch{
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Clockify client IDs...`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to retrieve Clockify Client IDs!"
        EXIT
    }

    Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Successfully retrieved Clockify client IDs!"
    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Successfully retrieved Clockify client IDs!"
  
}

function GetPostedAktisProposals {

    try{

        if(!(Test-Path -Path "C:\Aktis")){
            New-Item -Path "C:\" -Name "Aktis" -ItemType "Directory" | Out-Null
        }
        if(!(Test-Path -Path "C:\Aktis\Clockify Integration")){
            New-Item -Path "C:\Aktis\Clockify Integration" -Name "Logs" -ItemType "Directory" | Out-Null
        }
        if(!(Test-Path -Path "C:\Aktis\Clockify Integration\Logs")){
            New-Item -Path "C:\Aktis\Clockify Integration" -Name "Logs" -ItemType "Directory" | Out-Null
        }
        if(!(Test-Path -Path "C:\Aktis\Clockify Integration\Logs\ActionLog--$(Get-Date -Format "MM-yyyy").log")){
            New-Item -Path "C:\Aktis\Clockify Integration\Logs" -Name "ActionLog--$(Get-Date -Format "MM-yyyy").log" -ItemType "File" | Out-Null
        }
        if(!(Test-Path -Path "C:\Aktis\Clockify Integration\Logs\PostedProposalsLog--$(Get-Date -Format "MM-yyyy").log")){
            New-Item -Path "C:\Aktis\Clockify Integration\Logs" -Name "PostedProposalsLog--$(Get-Date -Format "MM-yyyy").log" -ItemType "File" | Out-Null
        }

        $Script:CachedAktisProposals = Get-ChildItem -Path "C:\Aktis\Clockify Integration\Logs\Posted*" | ForEach-Object {Get-Content -Path "$_"}

    }
    catch{
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Aktis proposals from PostedProposals log...`r`n`r`nError details:`r`n$($global:intErr++)Error :$global:intErr`r`n$Error"
        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to retrieve Aktis proposals from PostedProposals log!"
        EXIT
    }

    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Successfully retrieved Aktis proposals from PostedProposals logs!"
    Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Successfully retrieved Aktis proposals from PostedProposals log!"

}

function CreateWorkplanFromTemplate {

    #$ProposalMSP = "Affinity Technology Partners"
    #$ProposalName = "Replacement Computers"
    #$ProposalNumber = "2094"
    #$ProposalVersion = "2"
    #$ProposalKTNumber = "KT75598391635"

    Param
    (
         [Parameter(Mandatory=$True, Position=0)]
         [string] $ProposalMSP,
         [Parameter(Mandatory=$True, Position=1)]
         [string] $ProposalName,
         [Parameter(Mandatory=$True, Position=2)]
         [string] $ProposalNumber,
         [Parameter(Mandatory=$True, Position=3)]
         [string] $ProposalVersion,
         [Parameter(Mandatory=$True, Position=4)]
         [string] $ProposalKTNumber
    )

    try{

        $ActureSolutionsWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\Acture Solutions\Acture Template.xlsx"
        $AffinityWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\Affinity Technology Partners\Affinity Template.xlsx"
        $AMCWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\AMC Solutions\AMC Template.xlsx"
        $AntisynWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\Antisyn\Antisyn Template.xlsx"
        $ApticaWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\Aptica LLC\Aptica - FOR Quoting Tool - Proposal Template.xlsx"
        $DenaliTEKWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\DenaliTEK\DenaliTEK Template.xlsx"
        $DominionTechWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\DominionTech\Dominion Template.xlsx"
        $GroffWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\Groff\Groff Template - No Quoting Tool.xlsx"
        $ITSWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\ITS\00 ITS Template All Sites.xlsx"
        $Layer9WorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\Layer9\Layer9 Proposal Template.xlsx"
        $ManawaWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\Manawa\Manawa - Quoting Tool - Proposal Template.xlsx"
        $MentisGroupWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\Mentis Group\Mentis Template V2.xlsx"
        $MyITCrewNYWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\My IT Crew - NY\My IT Crew - NY Template.xlsx"
        $SisAdminWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\SisAdmin\SisAdmin Template.xlsx"
        $STAIWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\STAI\STAI Template.xlsx"
        $SteadyNetworksWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\Steady Networks\Steady Template - No Quoting Tool.xlsx"
        $SyscomWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\Syscom\Syscom Template.xlsx"
        $UniVistaWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\UniVista\Univista Template Blank Template.xlsx"
        $ValleyExpetecWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\WIP\In Progress\Valley Expetec\Valley Expetec Template.xlsx"

    }
    catch{
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to locate customer workplan template.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to locate customer workplan template.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
    }

    try{

        $ActureSolutionsClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\Acture Solutions\Design Desk Fuel\Proposals\"
        $AffinityClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\Affinity Technology Partners\Design Desk Fuel\Proposals\"
        $AMCClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\AMC Solutions\Design Desk Fuel\Proposals\"
        $AntisynClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\Antisyn\Design Desk Fuel\Proposals\"
        $ApticaClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\Aptica\Design Desk Fuel\Proposals\"
        $DenaliTEKClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\DenaliTEK\Design Desk Fuel\Proposals\"
        $DominionTechClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\DominionTech\Project Fuel\Engagement Documents\Proposals\Proposals\"
        $GroffClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\Groff Networks\Design Desk Fuel\Proposals\"
        $ITSClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\ITS\Project Fuel\Engagement Documents\Proposals\"
        $Layer9ClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\Layer9\Design Desk Fuel\Proposals\"
        $ManawaClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\Manawa\Design Desk Fuel\Proposals\"
        $MentisGroupClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\Mentis Group\Design Desk Fuel\Proposals\"
        $MyITCrewNYClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\My IT Crew - NY\Design Desk Fuel\Proposals\"
        $SisAdminClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\SisAdmin\Design Desk Fuel\Proposals\"
        $STAIClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\STAI\Design Desk Fuel Engagement Documents\Proposals\"
        $SteadyNetworksClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\Steady Networks\Design Desk Fuel\Proposals\"
        $SyscomClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\Syscom Business Technologies\Design Desk Fuel\Proposals\"
        $UniVistaClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\UniVista\Design Desk Fuel\Proposals\"
        $ValleyExpetecClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Clients\Valley Expetec\Design Desk Fuel\Proposals\"

    }
    catch{
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to locate customer client folder.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to locate customer client folder.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
    }
    
    if($ProposalMSP -eq "Acture Solutions"){
        $WorkplanTemplateLocation = $ActureSolutionsWorkplanTemplate
        $ClientFolder = $ActureSolutionsClientFolder
    }
    if($ProposalMSP -eq "Affinity Technology Partners"){
        $WorkplanTemplateLocation = $AffinityWorkplanTemplate
        $ClientFolder = $AffinityClientFolder
    }
    if($ProposalMSP -eq "AMC"){
        $WorkplanTemplateLocation = $AMCWorkplanTemplate
        $ClientFolder = $AMCClientFolder
    }
    if($ProposalMSP -eq "Antisyn"){
        $WorkplanTemplateLocation = $AntisynWorkplanTemplate
        $ClientFolder = $AntisynClientFolder
    }
    if($ProposalMSP -eq "Aptica LLC"){
        $WorkplanTemplateLocation = $ApticaWorkplanTemplate
        $ClientFolder = $ApticaClientFolder
    }
    if($ProposalMSP -eq "DenaliTEK"){
        $WorkplanTemplateLocation = $DenaliTEKWorkplanTemplate
        $ClientFolder = $DenaliTEKClientFolder
    }
    if($ProposalMSP -eq "DominionTech"){
        $WorkplanTemplateLocation = $DominionTechWorkplanTemplate
        $ClientFolder = $DominionTechClientFolder
    }
    if($ProposalMSP -eq "Groff Networks"){
        $WorkplanTemplateLocation = $GroffWorkplanTemplate
        $ClientFolder = $GroffClientFolder
    }
    if($ProposalMSP -eq "Layer9"){
        $WorkplanTemplateLocation = $Layer9WorkplanTemplate
        $ClientFolder = $Layer9ClientFolder
    }
    if($ProposalMSP -eq "Manawa"){
        $WorkplanTemplateLocation = $ManawaWorkplanTemplate
        $ClientFolder = $ManawaClientFolder
    }
    if($ProposalMSP -eq "Mentis"){
        $WorkplanTemplateLocation = $MentisGroupWorkplanTemplate
        $ClientFolder = $MentisGroupClientFolder
    }
    if($ProposalMSP -eq "My IT Crew - NY"){
        $WorkplanTemplateLocation = $MyITCrewNYWorkplanTemplate
        $ClientFolder = $MyITCrewNYClientFolder
    }
    if($ProposalMSP -eq "SisAdmin"){
        $WorkplanTemplateLocation = $SisAdminWorkplanTemplate
        $ClientFolder = $SisAdminClientFolder
    }
    if($ProposalMSP -eq "ITS"){
        $WorkplanTemplateLocation = $ITSWorkplanTemplate
        $ClientFolder = $ITSClientFolder
    }
    if($ProposalMSP -eq "STAI"){
        $WorkplanTemplateLocation = $STAIWorkplanTemplate
        $ClientFolder = $STAIClientFolder
    }
    if($ProposalMSP -eq "Steady Networks"){
        $WorkplanTemplateLocation = $SteadyNetworksWorkplanTemplate
        $ClientFolder = $SteadyNetworksClientFolder
    }
    if($ProposalMSP -eq "Syscom"){
        $WorkplanTemplateLocation = $SyscomWorkplanTemplate
        $ClientFolder = $SyscomClientFolder
    }
    if($ProposalMSP -eq "UniVista"){
        $WorkplanTemplateLocation = $UniVistaWorkplanTemplate
        $ClientFolder = $UniVistaClientFolder
    }
    if($ProposalMSP -eq "Valley Expetec"){
        $WorkplanTemplateLocation = $ValleyExpetecWorkplanTemplate
        $ClientFolder = $ValleyExpetecClientFolder
    }

    try{

        if(!(Test-Path -Path "C:\Aktis")){
            New-Item -Path "C:\" -Name "Aktis" -ItemType "Directory" | Out-Null
        }
        if(!(Test-Path -Path "C:\Aktis\Temp")){
            New-Item -Path "C:\Aktis" -Name "Temp" -ItemType "Directory" | Out-Null
        }
        
        #Copy-Item -Path $ProposalFilePath -Destination "C:\Aktis\Temp\$($ProposalFilePath | Split-Path -Leaf)" -Force
        #$($WorkplanTemplateLocation | Split-Path -Parent)\$("$ProposalName $ProposalNumber").xlsx

        if($ProposalVersion -eq "1"){

            Copy-Item -Path $WorkplanTemplateLocation -Destination "C:\Aktis\Temp\$("$ProposalName $ProposalNumber").xlsx" -Force | Out-Null

            $ProposalLocation = Get-Item -Path "C:\Aktis\Temp\$("$ProposalName $ProposalNumber").xlsx"

        }
        elseif($ProposalVersion -ge "2"){

            $PreviousProposalVersion = $ProposalVersion - 1

            if($PreviousProposalVersion -eq "1"){
                Copy-Item -Path "$ClientFolder\*\$("$ProposalName $ProposalNumber").xlsx" -Destination "C:\Aktis\Temp\$("$ProposalName $ProposalNumber V$ProposalVersion").xlsx" -Force | Out-Null
            }
            elseif($PreviousProposalVersion -ge "2"){
                Copy-Item -Path "$ClientFolder\*\$("$ProposalName $ProposalNumber V$PreviousProposalVersion").xlsx" -Destination "C:\Aktis\Temp\$("$ProposalName $ProposalNumber V$ProposalVersion").xlsx" -Force | Out-Null
            }

            $ProposalLocation = Get-Item -Path "C:\Aktis\Temp\$("$ProposalName $ProposalNumber V$ProposalVersion").xlsx"

        }

    }
    catch{

        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to create workplan from template for ($ProposalName [$ProposalNumber] $ProposalVersion).`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to create workplan from template for ($ProposalName [$ProposalNumber] $ProposalVersion).`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"

    }

    try{

        if($(Test-FileLock $ProposalLocation) -eq "True"){
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following file is already open. Requesting to close file.`r`n$ProposalLocation`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
            #[void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
            #[void] [Microsoft.VisualBasic.Interaction]::MsgBox("$($ProposalLocation | Split-Path -Leaf) is already open. Close the file and select OK.`r`n`r`nFull path:`r`n$ProposalLocation", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')

            while(
                $(Test-FileLock $ProposalLocation) -eq "True"
            ){
                Start-Sleep -Milliseconds "500"
            }
        }

        $ExcelOpen = New-Object -COMObject Excel.Application
        $ExcelOpen.Visible = $True
        $ExcelSpreadsheet = $ExcelOpen.Workbooks.Open($ProposalLocation)
        $ExcelSpreadsheet.Sheets.Item(1).Activate()

    }
    catch{

        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to open Excel file. Ensure Microsoft Excel is installed.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to open Excel file. Ensure Microsoft Excel is installed.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"

    }

    try{

        $ProposalNameHeadingColumn = ($ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find("Project Name:")).Column
        $ProposalNameRow = $ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find("Project Name:").Row + 1

    }
    catch{

        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to find the cell containing the following text:`r`n`r`nProject Name:`r`n`r`nError details:`r`n$($global:intErr++)Error:$global:intErr`r`n$Error"
        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to find the cell containing the following text:`r`n`r`nProject Name:`r`n`r`nError details:`r`n$($global:intErr++)Error:$global:intErr`r`n$Error"

    }

    #Unmerge the proposal name cell if it's merged
    if($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$ProposalNameRow,[int]$ProposalNameHeadingColumn).MergeCells -eq "True"){
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$ProposalNameRow,[int]$ProposalNameHeadingColumn).MergeCells() = $False
    }

    if($ProposalVersion -eq "1"){

        try{

            #($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$ProposalNameRow,[int]$ProposalNameHeadingColumn)).Clear()
            ($ExcelSpreadsheet.Worksheets.Item(1)).Name = "V$ProposalVersion"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$ProposalNameRow,[int]$ProposalNameHeadingColumn) = "$ProposalName [$ProposalNumber]"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$ProposalNameRow,[int]$ProposalNameHeadingColumn).HorizontalAlignment = -4131

        }
        catch{
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to update V1 proposal workplan.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to update V1 proposal workplan.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Errorr"
        }

    }
    elseif($ProposalVersion -ge "2"){

        try{

            ######Create V2 tab with copy of V1 tab (create V3 tab with V2 tab, etc.)
            $NewExcelWorksheet = $ExcelSpreadsheet.Worksheets.Add()
            $NewExcelWorksheet.Name = "V$ProposalVersion"

        }
        catch{
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to update V2 or greater proposal workplan.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to update V2 or greater proposal workplan.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Errorr"
        }

        try{
            $CurrentVersionTab = $ExcelSpreadsheet.Worksheets.Item(1)
            $PreviousVersionTab = $ExcelSpreadsheet.Worksheets.Item(2)
            $PreviousVersionTab.Copy($CurrentVersionTab)
        }
        catch{
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to create and/or populate tab for new version of proposal.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to create and/or populatea tab for new version of proposal.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Errorr"
        }

        try{
            #Cleanup and rename tab
            $TabToDelete = $ExcelSpreadsheet.Worksheets.Item(2)
            $TabToDelete.Delete()
            ($ExcelSpreadsheet.Worksheets.Item(1)).Name = "V$ProposalVersion"
        }
        catch{
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to remove unnecessary tab and/or rename newest tab.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to remove unnecessary tab and/or rename newest tab.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Errorr"
        }

        try{
            $ExcelSpreadsheet.Worksheets.Item(1).Activate()
            ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$ProposalNameRow,[int]$ProposalNameHeadingColumn)).Clear()
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$ProposalNameRow,[int]$ProposalNameHeadingColumn) = "$ProposalName [$ProposalNumber] V$ProposalVersion"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$ProposalNameRow,[int]$ProposalNameHeadingColumn).HorizontalAlignment = -4131


        }
        catch{
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to insert proposal name, number, and version.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to insert proposal name, number, and version.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Errorr"
        }

    }

    try{

        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$ProposalNameRow,[int]$ProposalNameHeadingColumn).AddComment("$ProposalKTNumber") | Out-Null
        #$ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$ProposalNameRow,[int]$ProposalNameHeadingColumn).Comment.Visible = $False

    }
    catch{

        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to insert note containing KTNumber.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        #Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to insert note containing KTNumber.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Errorr"

    }

    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Waiting 6 seconds for OneDrive to process changes for: $ProposalName [$ProposalNumber] V$ProposalVersion"
    Start-Sleep -Seconds 6
    
    try{

        $ExcelSpreadsheet.Save()
        $ExcelSpreadsheet.Close()
        $ExcelOpen.Quit()

        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Saved: $ProposalName [$ProposalNumber] V$ProposalVersion"

    }
    catch{
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to save and/or close proposal.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to save and/or close proposal.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Errorr"
    }


    try{
        Move-Item -Path $ProposalLocation -Destination "$($WorkplanTemplateLocation | Split-Path -Parent)\$($ProposalLocation | Split-Path -Leaf)" -Force | Out-Null
    }
    catch{
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to move proposal to WIP folder.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to move proposal to WIP folder.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Errorr"
    }
    
    
}

function StartClockifyIntegration {

    $APIEndpoint = "https://api.clockify.me/api/v1"
    $ClockifyWorkSpaceID = "627bafedcc6826493436b133"
    #$ClockifyAPIKey = "YzRiZTRjYzAtYmRjMi00NWQzLThlY2EtMTlhOTVlZTQxMTdl"
    $ClockifyAPIKey = "OTRkNGE0NDYtNTA4Mi00MmM4LThhNDQtZDIzYTk2NzY0NTY3"
    
    Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Starting Aktis New Proposal monitor..."
    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Starting Aktis New Proposal monitor..."

    GetClockifyClientIDs
    GetPostedAktisProposals
    
    $WaitIntervalInSeconds = "5"
    
    Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Starting integration do/while loop..."
    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Starting integration do/while loop..."
    
    do{
    
        try{
            $LatestAktisProposals = ((Get-PnPListItem -List "Proposal Queues" -Fields "KTNumber" -Connection $Script:ConnectToSPO).FieldValues).KTNumber
        }
        catch{
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Aktis proposal reference IDs from Proposal Queues SharePoint List...`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to retrieve Aktis proposals!"
            Return
        }

        $ObjectA = $Script:CachedAktisProposals | Sort-Object -Descending
        $ObjectB = $LatestAktisProposals | Sort-Object -Descending

        if($Null -eq $ObjectA){
            GetPostedAktisProposals
        }

        if($Null -eq $ObjectB){
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: List of latest Aktis proposals is blank/null."  
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Warning: List of latest Aktis proposals is blank/null."
            Return
        }
    
        if(!($ObjectA -eq $ObjectB)){
            
            $Script:NewAktisProposalKTNumber = (Compare-Object -ReferenceObject $($LatestAktisProposals | Sort-Object -Descending) -DifferenceObject $($Script:CachedAktisProposals | Sort-Object -Descending) -IncludeEqual | Where-Object {$_.SideIndicator -eq "<="}).InputObject
            
            #ForEach start            
            $Script:NewAktisProposalKTNumber | Sort-Object -Descending | ForEach-Object {

                $Script:KTNumber = $_
                #$Script:KTNumber = "KT164118336368"

                if($Null -eq $Script:KTNumber){
                    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: KTNumber is null...continuing to next object"
                    Return
                }
     
                # Continue to next object if Proposal has already been processed by integration
                if($Script:CachedAktisProposals | Select-String $Script:KTNumber){
                    #Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): $Script:KTNumber has already been processed by integration...continuing to next object"
                    Return
                }
            
                $Script:NewProposalNumber = ((Get-PnPListItem -List "Proposal Queues" -Fields "TicketNumber", "KTNumber" -Connection $Script:ConnectToSPO).FieldValues | Where-Object {$_.KTNumber -eq $Script:KTNumber}).TicketNumber

                $Script:NewProposalVersion = ((Get-PnPListItem -List "Proposal Queues" -Fields "ProposalVersion", "KTNumber" -Connection $Script:ConnectToSPO).FieldValues | Where-Object {$_.KTNumber -eq $Script:KTNumber}).ProposalVersion

                $Script:NewProposalName = ((Get-PnPListItem -List "Proposal Queues" -Fields "ProposalName_x0028_TicketSummary", "KTNumber" -Connection $Script:ConnectToSPO).FieldValues | Where-Object {$_.KTNumber -eq $Script:KTNumber}).ProposalName_x0028_TicketSummary
                ######## Need to add ability to filter out version number of proposal name, if it exists

                $NewProposalMSP = ((Get-PnPListItem -List "Proposal Queues" -Fields "Title", "KTNumber" -Connection $Script:ConnectToSPO).FieldValues | Where-Object {$_.KTNumber -eq $Script:KTNumber}).Title
                
                $MatchingProjects = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/projects") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "$Script:NewProposalNumber V$Script:NewProposalVersion"})

                if ($MatchingProjects){
                    foreach ($Project in $MatchingProjects) {
                        if ($Project.name -eq "$Script:NewProposalNumber V$Script:NewProposalVersion") {
                            $MatchingProject = $Project
                            break  # Exit the loop once a match is found
                        } else {
                            $MatchingProject = "NONE"  # Set to $null if no match is found
                        }
                    }
                }else {
                    $MatchingProject = "NONE"
                }

                # Determine if Clockify Project already exists first
                if($MatchingProject -ne "NONE"){

                    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): New proposal ($Script:NewProposalNumber V$Script:NewProposalVersion) (KTNumber: $Script:KTNumber) already exists in Clockify..." 
                    Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "New proposal ($Script:NewProposalNumber V$Script:NewProposalVersion) (KTNumber: $Script:KTNumber) already exists in Clockify..."
                    Add-Content -Path "C:\Aktis\Clockify Integration\Logs\ActionLog--$(Get-Date -Format "MM-yyyy").log" -Value "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss:ss") -- $Script:NewProposalNumber V$Script:NewProposalVersion already exists in Clockify!"

                    UpdatePostedProposalsLog

                    $Script:CachedAktisProposals = Get-ChildItem -Path "C:\Aktis\Clockify Integration\Logs\Posted*" | ForEach-Object {Get-Content -Path "$_"}

                    Return
                }

                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): New proposal ($Script:NewProposalNumber V$Script:NewProposalVersion) (KTNumber: $Script:KTNumber) detected in Aktis..."
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "New proposal ($Script:NewProposalNumber V$Script:NewProposalVersion) (KTNumber: $Script:KTNumber) detected in Aktis..."

                if($NewProposalMSP -eq "Acture Solutions"){
                    $ClockifyClientID = $Script:ActureSolutionsClockifyClientID
                }
                if($NewProposalMSP -eq "Affinity Technology Partners"){
                    $ClockifyClientID = $Script:AffinityClockifyClientID
                }
                if($NewProposalMSP -eq "AMC"){
                    $ClockifyClientID = $Script:AMCClockifyClientID
                }
                if($NewProposalMSP -eq "Antisyn"){
                    $ClockifyClientID = $Script:AntisynClockifyClientID
                }
                if($NewProposalMSP -eq "Aptica LLC"){
                    $ClockifyClientID = $Script:ApticaClockifyClientID
                }
                if($NewProposalMSP -eq "DenaliTEK"){
                    $ClockifyClientID = $Script:DenaliTEKClockifyClientID
                }
                if($NewProposalMSP -eq "DominionTech"){
                    $ClockifyClientID = $Script:DominionClockifyClientID
                }
                if($NewProposalMSP -eq "Groff Networks"){
                    $ClockifyClientID = $Script:GroffClockifyClientID
                }
                if($NewProposalMSP -eq "Manawa"){
                    $ClockifyClientID = $Script:ManawaClockifyClientID
                }
                if($NewProposalMSP -eq "Layer9"){
                    $ClockifyClientID = $Script:Layer9ClockifyClientID
                }
                if($NewProposalMSP -eq "Mentis"){
                    $ClockifyClientID = $Script:MentisClockifyClientID
                }
                if($NewProposalMSP -eq "My IT Crew - NY"){
                    $ClockifyClientID = $Script:MyITCrewNYClockifyClientID
                }
                if($NewProposalMSP -eq "SisAdmin"){
                    $ClockifyClientID = $Script:SisAdminClockifyClientID
                }
                if($NewProposalMSP -eq "ITS"){
                    $ClockifyClientID = $Script:ITSClockifyClientID
                }
                if($NewProposalMSP -eq "Project Fuel"){
                    $ClockifyClientID = $Script:ProjectFuelClockifyClientID
                }
                if($NewProposalMSP -eq "STAI"){
                    $ClockifyClientID = $Script:STAIClockifyClientID
                }
                if($NewProposalMSP -eq "Steady Networks"){
                    $ClockifyClientID = $Script:SteadyClockifyClientID
                }
                if($NewProposalMSP -eq "Syscom"){
                    $ClockifyClientID = $Script:SyscomClockifyClientID
                }
                if($NewProposalMSP -eq "UniVista"){
                    $ClockifyClientID = $Script:UniVistaClockifyClientID
                }
                if($NewProposalMSP -eq "Valley Expetec"){
                    $ClockifyClientID = $Script:ValleyExpetecClockifyClientID
                }
    
                $RequestBody = @{
                "name" = "$Script:NewProposalNumber V$Script:NewProposalVersion"
                "clientId" = $ClockifyClientID
                "isPublic" = "True"
                #"color" = "#f44336"
                "note" = "$Script:NewProposalName [$Script:NewProposalNumber] V$Script:NewProposalVersion"
                "billable" = "True"
                "public" = "True"
                }
            
                ### Create Clockify project
                try{
                    $ClockifyProject = Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/projects") -Method Post -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body $(ConvertTo-Json ($RequestBody))
                }
                catch{
                    Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to create Clockify project...`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
                    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to create Clockify project!"
                    #[void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                    #[void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to create Clockify project.", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                    #mod Return
                }
                
                try{
                    #mod $ClockifyProjectID = (Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/projects") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body @{name = "$Script:NewProposalNumber V$Script:NewProposalVersion"}).ID
                    $ClockifyProjectID = $ClockifyProject.id
                }
                catch{
                    Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Clockify project ID...`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
                    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to retrieve Clockify project ID!"
                    #[void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                    #[void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to retrieve Clockify project ID.", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                    Return
                }
                
                $PWRequestBody = @{
                    "name" = "Proposal Writing"
                    "status" = "ACTIVE"
                }
                $PCRequestBody = @{
                    "name" = "Proposal Coordination"
                    "status" = "ACTIVE"
                }
                $DCRequestBody = @{
                    "name" = "Design Consulting"
                    "status" = "ACTIVE"
                }
                $IPRequestBody = @{
                    "name" = "Internal Project"
                    "status" = "ACTIVE"
                }
                $OCRequestBody = @{
                    "name" = "Operations Consulting"
                    "status" = "ACTIVE"
                }
                
                ### Create tasks for Clockify project
                try{
                    Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/projects/" + $ClockifyProjectID + "/tasks") -Method Post -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body $(ConvertTo-Json ($PWRequestBody))
                    Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/projects/" + $ClockifyProjectID + "/tasks") -Method Post -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body $(ConvertTo-Json ($PCRequestBody))
                    Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/projects/" + $ClockifyProjectID + "/tasks") -Method Post -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body $(ConvertTo-Json ($DCRequestBody))
                    Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/projects/" + $ClockifyProjectID + "/tasks") -Method Post -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body $(ConvertTo-Json ($IPRequestBody))
                    Invoke-RestMethod -Uri ($APIEndpoint + "/workspaces/" + $ClockifyWorkSpaceID + "/projects/" + $ClockifyProjectID + "/tasks") -Method Post -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $ClockifyAPIKey} -Body $(ConvertTo-Json ($OCRequestBody))
                }
                catch{
                    Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to create Clockify project tasks...`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
                    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to create Clockify project tasks...error: $($global:intErr++)Error #:$global:intErr`r`n$Error"
                    Return
                }

                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Successfully created Clockify project: $Script:NewProposalNumber V$Script:NewProposalVersion..."
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Successfully created Clockify project: $Script:NewProposalNumber V$Script:NewProposalVersion`r`n`r`n($Script:NewProposalName [$Script:NewProposalNumber] V$Script:NewProposalVersion)"
                
                UpdateActionLog
                UpdatePostedProposalsLog

                CreateWorkplanFromTemplate $NewProposalMSP $Script:NewProposalName $Script:NewProposalNumber $Script:NewProposalVersion $Script:KTNumber

                $Script:CachedAktisProposals = Get-ChildItem -Path "C:\Aktis\Clockify Integration\Logs\Posted*" | ForEach-Object {Get-Content -Path "$_"}
    
            }
        }

        Clear-Variable LatestAktisProposals -Confirm:$False | Out-Null

        Start-Sleep -Seconds $WaitIntervalInSeconds
            
    }

    while($True)
}

function UpdateActionLog {
    if(!(Test-Path -Path "C:\Aktis")){
        New-Item -Path "C:\" -Name "Aktis" -ItemType "Directory" | Out-Null
    }
    if(!(Test-Path -Path "C:\Aktis\Clockify Integration")){
        New-Item -Path "C:\Aktis" -Name "Clockify Integration" -ItemType "Directory" | Out-Null
    }
    if(!(Test-Path -Path "C:\Aktis\Clockify Integration\Logs")){
        New-Item -Path "C:\Aktis\Clockify Integration" -Name "Logs" -ItemType "Directory" | Out-Null
    }
    if(!(Test-Path -Path "C:\Aktis\Clockify Integration\Logs\ActionLog--$(Get-Date -Format "MM-yyyy").log")){
        New-Item -Path "C:\Aktis\Clockify Integration\Logs" -Name "ActionLog--$(Get-Date -Format "MM-yyyy").log" -ItemType "File" | Out-Null
    }
    Add-Content -Path "C:\Aktis\Clockify Integration\Logs\ActionLog--$(Get-Date -Format "MM-yyyy").log" -Value "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss:ss") -- Created $Script:NewProposalNumber V$Script:NewProposalVersion (KTNumber: $Script:KTNumber)"
}

function UpdatePostedProposalsLog {
    if(!(Test-Path -Path "C:\Aktis")){
        New-Item -Path "C:\" -Name "Aktis" -ItemType "Directory" | Out-Null
    }
    if(!(Test-Path -Path "C:\Aktis\Clockify Integration")){
        New-Item -Path "C:\Aktis" -Name "Clockify Integration" -ItemType "Directory" | Out-Null
    }
    if(!(Test-Path -Path "C:\Aktis\Clockify Integration\Logs")){
        New-Item -Path "C:\Aktis\Clockify Integration" -Name "Logs" -ItemType "Directory" | Out-Null
    }
    if(!(Test-Path -Path "C:\Aktis\Clockify Integration\Logs\PostedProposalsLog--$(Get-Date -Format "MM-yyyy").log")){
        New-Item -Path "C:\Aktis\Clockify Integration\Logs" -Name "PostedProposalsLog--$(Get-Date -Format "MM-yyyy").log" -ItemType "File" | Out-Null
    }
    Add-Content -Path "C:\Aktis\Clockify Integration\Logs\PostedProposalsLog--$(Get-Date -Format "MM-yyyy").log" -Value "$Script:KTNumber"
}

AuthenticateToSPO
StartClockifyIntegration
