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

function AuthenticateToSPO {

    if(!(Test-Path -Path "C:\Windows\System32\WindowsPowerShell\v1.0\Modules\PnP.PowerShell")){
        if(!(Test-Path -Path "C:\Program Files\WindowsPowerShell\Modules\PnP.PowerShell")){
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "PnP.PowerShell is not installed. Please install PnP.PowerShell before launching Aktis Helper.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
            [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
            [void] [Microsoft.VisualBasic.Interaction]::MsgBox("PnP.PowerShell is not installed. Please install PnP.PowerShell before launching Aktis Helper.", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
            EXIT
        }
    }

    #if(!(Test-Path -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Analytics")){
    #    [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
    #    [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The OPS SharePoint document library is not synced to the default path. You will not be able to post to local analytics.", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
    #}
    
    try{
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Connect-PnPOnline -Url "https://projectfuelnow.sharepoint.com/sites/Fuelnow" -Interactive -ForceAuthentication -ErrorAction SilentlyContinue
    }
    catch{ 
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to connect to SharePoint Online with PnP.PowerShell.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to connect to SharePoint Online with PnP.PowerShell.", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
        EXIT
    }
    
}

function GetTimeFrameDates {

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    $Form = New-Object Windows.Forms.Form -Property @{
        StartPosition = [Windows.Forms.FormStartPosition]::CenterScreen
        Size = New-Object Drawing.Size 243, 230
        Text = 'Select a Date'
        Topmost = $True
    }
    
    $Calendar = New-Object Windows.Forms.MonthCalendar -Property @{
        ShowTodayCircle = $False
        MaxSelectionCount = 1
    }
    $Form.Controls.Add($Calendar)
    
    $OkButton = New-Object Windows.Forms.Button -Property @{
        Location     = New-Object Drawing.Point 38, 165
        Size         = New-Object Drawing.Size 75, 23
        Text         = 'OK'
        DialogResult = [Windows.Forms.DialogResult]::OK
    }
    $Form.AcceptButton = $OkButton
    $Form.Controls.Add($OkButton)
    
    $cancelButton = New-Object Windows.Forms.Button -Property @{
        Location     = New-Object Drawing.Point 113, 165
        Size         = New-Object Drawing.Size 75, 23
        Text         = 'Cancel'
        DialogResult = [Windows.Forms.DialogResult]::Cancel
    }
    $Form.CancelButton = $cancelButton
    $Form.Controls.Add($cancelButton)
    
    $Result = $Form.ShowDialog()
    
    if($Result -eq [Windows.Forms.DialogResult]::OK){
        $Date = $Calendar.SelectionStart
        Write-Output "$($Date.ToShortDateString())"
    }

    #$Script:StartTime = [DateTime]"10/31/22"
    #$Script:EndTime = [DateTime]"12/04/22"

}

function RetrieveCustomerProfileData {

    try{

        $CustomerProfileRawData = (Get-PnPListItem -List "Customer Profile" -Fields "Id","ClientName","EnterpriseSubscription","BasicSubscription","CurrentAgreementMRR","CurrentAgreementReproductionSubs","AdditionalMonthlyEnterpriseFee","AdditionalMonthlyBasicFee","ExpediteFee","DesignConsultingHourlyRate","OperationsConsultingHourlyRate").FieldValues
    
        $Script:CustomerProfileData = @()
    
        $CustomerProfileRawData | ForEach-Object {

            if($($_.ClientName) -eq "Project Fuel"){
                Return
            }

            $Script:CustomerProfileData += $_
        }

    }
    catch{
        #Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Customer Profile data from SharePoint.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to retrieve Customer Profile data from SharePoint.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
    }

}

function RetrieveAnalyticsData {

    try{
        $ProposalStatisticsRawData = (Get-PnPListItem -List "Proposal Statistics" -Fields "Id","TicketNumber","ProposalVerison","ProposalExceedsSubscription_x002","ProposalExceedsSubscription_x0020","GratisRevision","Company","ProposalName","IntakeType","DDE","ProposedHours","ProposedMRR","ProposedProfessionalServicesReve","ProposedProductsRevenue","Proposed_x0020_Products_x0020_an","Hours_x0020_Proposed","KTDate","DateProposalCompleted","Turnaround_x0020_Time_x0020__x00","DeliveredtoClient","Outcome","Subscription","_x0035_Day","_x0033_DayExpedite","_x0031_00","_x0031_000","_x0032_00","_x0033_00","PresenttoClient_x002f_Staff","Draft1","Draft2","ProcurementFee","ProposalCoordinationTime","ProposalWritingTime","DesignConsultingTimeNew","InternalProjectTime","MSP","HourlyRate").FieldValues
    
        $Script:ProposalStatisticsData = @()
    
        $ProposalStatisticsRawData | ForEach-Object {
            $Script:ProposalStatisticsData += $_
        }
    }
    catch{
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Aktis analytics data from SharePoint.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to retrieve Aktis analytics data from SharePoint.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
    }

}

function RetrieveClockifyProjectData {

    #Establish required Clockify variables
    $Script:ClockifyAPIEndpoint = "https://api.clockify.me/api/v1"
    $Script:ClockifyAPIReportsEndpoint = "https://reports.api.clockify.me/v1"
    $Script:ClockifyWorkSpaceID = "627bafedcc6826493436b133"
    $Script:ClockifyAPIKey = "OTRkNGE0NDYtNTA4Mi00MmM4LThhNDQtZDIzYTk2NzY0NTY3"

    [string]$Script:StartDateISO8601 = Get-Date ($Script:StartTime).ToUniversalTime() -UFormat '+%Y-%m-%dT%H:%M:%S.000Z'
    [string]$Script:EndDateISO8601 = Get-Date ($Script:EndTime).ToUniversalTime() -UFormat '+%Y-%m-%dT%H:%M:%S.000Z'

    try{

        $Script:ClockifyProjectDetailsRawData = Invoke-RestMethod -Uri ($Script:ClockifyAPIEndpoint + "/workspaces/" + $Script:ClockifyWorkSpaceID + "/projects") -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $Script:ClockifyAPIKey} -Method Get -Body @{
            "start" = $Script:StartDateISO8601
            "end" = $Script:EndDateISO8601
            "page-size" = $PageSize
          }

          $RequestBody = @{
            "dateRangeStart" = $Script:StartDateISO8601
            "dateRangeEnd" = $Script:EndDateISO8601
            "detailedFilter" = @{
                "page" = 1
                "pageSize" = 1000
            }
            "exportType" = "JSON"
        }
        
        $Script:Output = Invoke-RestMethod -Uri ($Script:ClockifyAPIReportsEndpoint + "/workspaces/" + $Script:ClockifyWorkSpaceID + "/reports/detailed") -Method Post -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $Script:ClockifyAPIKey} -Body @(ConvertTo-Json $RequestBody)
        
    }
    catch{

        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Clockify project details.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to retrieve Clockify project details.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')

    }
}

function DetermineAdditionalMonthlyProposals ($MSPName) {

    $CustomerWeeklyEnterpriseSubscription = ($Script:CustomerProfileData | Where-Object {$_.ClientName -eq $MSPName}).EnterpriseSubscription
    $CustomerWeeklyBasicSubscription = ($Script:CustomerProfileData | Where-Object {$_.ClientName -eq $MSPName}).BasicSubscription

    $TimeSpan = New-TimeSpan -Start $Script:StartTime -End $Script:EndTime
    $MonthlySubscriptionMultiplier = [Math]::Round([Math]::Ceiling($TimeSpan.Days / 7), 2)
    
    $CustomerMonthlyEnterpriseSubscription = [int]$CustomerWeeklyEnterpriseSubscription * [int]$MonthlySubscriptionMultiplier
    $CustomerMonthlyBasicSubscription = [int]$CustomerWeeklyBasicSubscription * [int]$MonthlySubscriptionMultiplier

    $EnterpriseProposalsDeliveredThisMonth = @()
    $BasicProposalsDeliveredThisMonth = @()
    $CustomerProposalsDeliveredThisMonth = @()
    $Script:CustomerAdditionalMonthlyEnterpriseProposals = @()
    $Script:CustomerAdditionalMonthlyBasicProposals = @()
    
    $Script:ProposalStatisticsData | ForEach-Object {

        if($Null -eq $_.DateProposalCompleted){
            Return
        }

        if($_.GratisRevision -eq $True){
            Return
        }

        if(($_.MSP -eq $MSPName) -and (([DateTime]$_.DateProposalCompleted -ge [DateTime]$Script:StartTime) -and ([DateTime]$_.DateProposalCompleted -le [DateTime]$Script:EndTime))){

            $ProposedHours = $_.ProposedHours

            if($Null -eq $ProposedHours){
                #Write-Host "$($_.TicketNumber) V$($_.ProposalVerison) does not have ProposedHours..."
                Return
            }

            $CustomerProposalsDeliveredThisMonth += $_

        } 
    }

    $CustomerProposalsDeliveredThisMonth | ForEach-Object {

        $ProposedHours = $_.ProposedHours

        if($ProposedHours -gt '8'){
            $EnterpriseProposalsDeliveredThisMonth += $_
        }
        elseif($ProposedHours -le '8'){
            $BasicProposalsDeliveredThisMonth += $_
        }

    }

    if([int]$EnterpriseProposalsDeliveredThisMonth.Count -lt 1){
        $AdditionalEnterpriseProposals = $Null
    }
    elseif([int]$EnterpriseProposalsDeliveredThisMonth.Count -ge 1){
        if([int]$EnterpriseProposalsDeliveredThisMonth.Count -gt [int]$CustomerMonthlyEnterpriseSubscription){
            $AdditionalEnterpriseProposals = $EnterpriseProposalsDeliveredThisMonth | Select-Object -Skip "$([int]$CustomerMonthlyEnterpriseSubscription)"
        }
        else{
            $AdditionalEnterpriseProposals = $Null
        }

        if($AdditionalEnterpriseProposals.Count -lt 1){
            $AdditionalEnterpriseProposals = $Null
        }
        else{
            $AdditionalEnterpriseProposals | ForEach-Object {
                $Script:CustomerAdditionalMonthlyEnterpriseProposals += $_
            }
        }
    }

    if([int]$BasicProposalsDeliveredThisMonth.Count -lt 1){
        $AdditionalBasicProposals = $Null
    }
    elseif([int]$BasicProposalsDeliveredThisMonth.Count -ge 1){
        if([int]$BasicProposalsDeliveredThisMonth.Count -gt [int]$CustomerMonthlyBasicSubscription){
            $AdditionalBasicProposals = $BasicProposalsDeliveredThisMonth | Select-Object -Skip "$([int]$CustomerMonthlyBasicSubscription)"
        }
        else{
            $AdditionalBasicProposals = $Null
        }
    
        if($AdditionalBasicProposals.Count -lt 1){
            $AdditionalBasicProposals = $Null
        }
        else{
            $AdditionalBasicProposals | ForEach-Object {
                $Script:CustomerAdditionalMonthlyBasicProposals += $_
            }
        }
    }
}

function GenerateMonthlyBillingReport {

    try{
        $MonthlyBillingSpreadsheetTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\PF Resources\Monthly Billing\Monthly Billing Template.xlsx"
        Copy-Item -Path $MonthlyBillingSpreadsheetTemplate -Destination $Script:NewMonthlyBillingReportName -Force | Out-Null
    }
    catch{c:\Users\VictorAraoz\Project Fuel\Project Fuel - OPS - Documents\PF Resources\Monthly Billing\
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to create new report file.", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
    }

    try{

        if($(Test-FileLock $Script:NewMonthlyBillingReportName) -eq "True"){
            #Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following analytics file is already open. Requesting to close file.`r`n$LocalAnalyticsWeeklySnapshotFile`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
            [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
            [void] [Microsoft.VisualBasic.Interaction]::MsgBox("$($Script:NewMonthlyBillingReportName | Split-Path -Leaf) is already open. Close the file and select OK.", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
    
            while(
                $(Test-FileLock $Script:NewMonthlyBillingReportName) -eq "True"
            ){
                Start-Sleep -Milliseconds "500"
            }
        }

        $ExcelOpen = New-Object -COMObject Excel.Application
        $ExcelOpen.Visible = $True
        $ExcelSpreadsheet = $ExcelOpen.Workbooks.Open($Script:NewMonthlyBillingReportName)

    }
    catch{
        #Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to open Analytics Weekly Snapshot file. Ensure it's not already open.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to open Monthly Billing spreadsheet. Ensure it's not already open", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
    }

    try{
        # Rename worksheet
        $ExcelWorkSheet = $ExcelSpreadsheet.Worksheets.Item(1)
        $ExcelWorkSheet.Name = "$($Script:StartTime.ToString("MM-dd-yyyy"))-to-$($Script:EndTime.ToString("MM-dd-yyyy"))"
    }
    catch{
        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("Failed to rename report orksheet.", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
    }
    
    try{

        # Determine header columns
        $WeeklyEnterpriseColumn = ($ExcelWorkSheet.UsedRange.Find("Weekly Enterprise")).Column
        $WeeklyBasicColumn = ($ExcelWorkSheet.UsedRange.Find("Weekly Basic")).Column
        $CurrentAgrmntMRRColumn = ($ExcelWorkSheet.UsedRange.Find("Current Agreement MRR")).Column
        $CurrentAgrmntReproSubscColumn = ($ExcelWorkSheet.UsedRange.Find("Current Agreement Reproduction Subscription")).Column
        $AdditionalEnterpriseColumn = ($ExcelWorkSheet.UsedRange.Find("Additional Monthly Enterprise")).Column
        $AdditionalBasicColumn = ($ExcelWorkSheet.UsedRange.Find("Additional Monthly Basic")).Column
        $ExpediteColumn = ($ExcelWorkSheet.UsedRange.Find("Expedite")).Column
        $DesignConsultingHoursColumn = ($ExcelWorkSheet.UsedRange.Find("Design Consulting Hours")).Column
        $OperationsConsultingHoursColumn = ($ExcelWorkSheet.UsedRange.Find("Operations Consulting Hours")).Column
        $GT100Column = ($ExcelWorkSheet.UsedRange.Find(">100")).Column
        $GT200Column = ($ExcelWorkSheet.UsedRange.Find(">200")).Column
        $GT300Column = ($ExcelWorkSheet.UsedRange.Find(">300")).Column
        $GT400Column = ($ExcelWorkSheet.UsedRange.Find(">400")).Column

    }
    catch{

    }

    #Gather data
    RetrieveCustomerProfileData
    RetrieveAnalyticsData
    RetrieveClockifyProjectData

    $Script:CustomerProfileData | ForEach-Object {

        $CustomerName = "$($_.ClientName)"

        if($CustomerName | Select-String "_INACTIVE--"){
            Return
        }

        #Determine customer's additional monthly proposals
        DetermineAdditionalMonthlyProposals $CustomerName

        $CustomerEnterpriseSubscription = ($Script:CustomerProfileData | Where-Object {$_.ClientName -eq $CustomerName}).EnterpriseSubscription
        $CustomerBasicSubscription = ($Script:CustomerProfileData | Where-Object {$_.ClientName -eq $CustomerName}).BasicSubscription
        $CurrentAgrmntMRR = ($Script:CustomerProfileData | Where-Object {$_.ClientName -eq $CustomerName}).CurrentAgreementMRR
        $CurrentAgrmntReproSubsc = ($Script:CustomerProfileData | Where-Object {$_.ClientName -eq $CustomerName}).CurrentAgreementReproductionSubs
        $AdditionalEnterpriseFee = ($Script:CustomerProfileData | Where-Object {$_.ClientName -eq $CustomerName}).AdditionalMonthlyEnterpriseFee
        $AdditionalBasicFee = ($Script:CustomerProfileData | Where-Object {$_.ClientName -eq $CustomerName}).AdditionalMonthlyBasicFee
        $ExpediteFee = ($Script:CustomerProfileData | Where-Object {$_.ClientName -eq $CustomerName}).ExpediteFee
        $DCHourlyRate = ($Script:CustomerProfileData | Where-Object {$_.ClientName -eq $CustomerName}).DesignConsultingHourlyRate
        $OCHourlyRate = ($Script:CustomerProfileData | Where-Object {$_.ClientName -eq $CustomerName}).OperationsConsultingHourlyRate

        $CustomerProposals = @()
        $ProposalsGT100Hours = @()
        $ProposalsGT200Hours = @()
        $ProposalsGT300Hours = @()
        $ProposalsGT400Hours = @()
        $ExpeditedProposals = @()
    
        $Script:ProposalStatisticsData | ForEach-Object {

            if($Null -eq $_.DateProposalCompleted){
                Return
            }

            if($_.GratisRevision -eq $True){
                Return
            }

            if(($_.MSP -eq $CustomerName) -and (([DateTime]$_.DateProposalCompleted -ge [DateTime]$Script:StartTime) -and ([DateTime]$_.DateProposalCompleted -le [DateTime]$Script:EndTime))){

                $CustomerProposals += $_

            } 
        }

        $CustomerProposals | ForEach-Object {

            if($_.GratisRevision -eq $True){
                Return
            }

            if($_.ProposedHours -ge '400'){
                $ProposalsGT400Hours += $_
            }
            elseif($_.ProposedHours -ge '300'){
                $ProposalsGT300Hours += $_
            }
            elseif($_.ProposedHours -ge '200'){
                $ProposalsGT200Hours += $_
            }
            elseif($_.ProposedHours -ge '100'){
                $ProposalsGT100Hours += $_
            }

            $ProposalVersion = $_.ProposalVerison

            if($_._x0033_DayExpedite -eq $True){

                if(($ProposalVersion -eq "1") -or ($ProposalVersion -eq "3") -or ($ProposalVersion -eq "5") -or ($ProposalVersion -eq "7") -or ($ProposalVersion -eq "9") -or ($ProposalVersion -eq "11") -or ($ProposalVersion -eq "13")){
                    
                    $ExpeditedProposals += $_

                }
                
            }

            #if($_.ProposalExceedsSubscription_x002 -eq $True){
            #    $CustomerAdditionalEnterpriseProposals += $_
            #}
            #
            #if($_.ProposalExceedsSubscription_x0020 -eq $True){
            #    $CustomerAdditionalBasicProposals += $_
            #}
            
        }

        ## Retrieve Design Consulting time data
        $DCTimeEntries = $Script:Output.TimeEntries | Where-Object {$_.ClientName -eq $CustomerName} | Where-Object {$_.TaskName -eq "Design Consulting"}

        [decimal]$Script:TotalDCTimeThisMonth = 0.00
        $Script:ProposalsWithDCTimeThisMonth = @()

        $DCTimeEntries | ForEach-Object {
            


        #### Stuck here - need to extract the project note from $Script:ClockifyProjectDetailsRawData by using The ID property in $DCTimeEntries



            #$ProjectID = $_.ProjectID
            $TimeSpentInSeconds = $_.TimeInterval.Duration

            [decimal]$TimeEntryTimeIntervalDuration = ([int]$TimeSpentInSeconds / [int]3600)
            [decimal]$RoundedTimeEntryTimeIntervalDuration = [math]::Round([int]$TimeSpentInSeconds / [int]3600,2)
            [decimal]$Script:TotalDCTimeThisMonth = [decimal]$Script:TotalDCTimeThisMonth + [decimal]$TimeEntryTimeIntervalDuration

            if(!($Script:ProposalsWithDCTimeThisMonth -Contains $_.ProjectName)){
                $ProjectName = $_.ProjectName
                $ProjectNote = ($Script:ClockifyProjectDetailsRawData | Where-Object {$_.Name -eq $ProjectName}).Note
                $Script:ProposalsWithDCTimeThisMonth += "$ProjectNote - $RoundedTimeEntryTimeIntervalDuration hours"
                 
            }

        }

        ## Retrieve Operations Consulting time data
        $OCTimeEntries = $Script:Output.TimeEntries | Where-Object {$_.ClientName -eq $CustomerName} | Where-Object {$_.TaskName -eq "Operations Consulting"}

        [decimal]$Script:TotalOCTimeThisMonth = 0.00
        $ProposalsWithOCTimeThisMonth = @()

        $OCTimeEntries | ForEach-Object {
            
            $TimeSpentInSeconds = $_.TimeInterval.Duration

            [decimal]$TimeEntryTimeIntervalDuration = ([int]$TimeSpentInSeconds / [int]3600)
            [decimal]$RoundedTimeEntryTimeIntervalDuration = [math]::Round([int]$TimeSpentInSeconds / [int]3600,2)
            [decimal]$Script:TotalOCTimeThisMonth = [decimal]$Script:TotalOCTimeThisMonth + [decimal]$TimeEntryTimeIntervalDuration

            if(!($ProposalsWithOCTimeThisMonth -Contains $_.ProjectName)){
                $ProposalsWithOCTimeThisMonth += "$($_.Company) - $($_.ProposalName) [$($_.TicketNumber)] V$($_.ProposalVerison) - $RoundedTimeEntryTimeIntervalDuration hours" 
            }

        }
        
        [int]$CustomerRow = $ExcelWorkSheet.UsedRange.Find("$CustomerName").Row

        #Weekly Enterprise
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$WeeklyEnterpriseColumn) = "$CustomerEnterpriseSubscription"
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$WeeklyEnterpriseColumn).HorizontalAlignment = -4108
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$WeeklyEnterpriseColumn).NumberFormat = "0"

        #Basic Enterprise
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$WeeklyBasicColumn) = "$CustomerBasicSubscription"
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$WeeklyBasicColumn).HorizontalAlignment = -4108
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$WeeklyBasicColumn).NumberFormat = "0"

        #Current Agreement MRR
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$CurrentAgrmntMRRColumn) = "$CurrentAgrmntMRR"
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$CurrentAgrmntMRRColumn).HorizontalAlignment = -4108
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$CurrentAgrmntMRRColumn).NumberFormat = "$#,##0.00"
    
        #Current Agreement Reproduction Subscription
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$CurrentAgrmntReproSubscColumn) = "$CurrentAgrmntReproSubsc"
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$CurrentAgrmntReproSubscColumn).HorizontalAlignment = -4108
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$CurrentAgrmntReproSubscColumn).NumberFormat = "$#,##0.00"
    
        #Additional Monthly Enterprise
        if($Script:CustomerAdditionalMonthlyEnterpriseProposals.Count -ge 1){
            
            $Script:CustomerAdditionalMonthlyEnterpriseProposals | ForEach-Object {
                $ExistingNote = ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalEnterpriseColumn).Comment.Shape.AlternativeText) -Replace "^.*?: "
                ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalEnterpriseColumn)).Clear()
                $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalEnterpriseColumn).AddComment("$ExistingNote`r`n$($_.Company) - $($_.ProposalName) [$($_.TicketNumber)] V$($_.ProposalVerison)")
                $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalEnterpriseColumn).Comment.Shape.TextFrame.AutoSize = $True 

            }
            #$ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalEnterpriseColumn).Comment.Visible = $False

            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalEnterpriseColumn) = "=$([int]$Script:CustomerAdditionalMonthlyEnterpriseProposals.Count) * $([math]::Round([decimal]$AdditionalEnterpriseFee,2))"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalEnterpriseColumn).NumberFormat = "$#,##0.00"
        }
        else{
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalEnterpriseColumn) = "N/A"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalEnterpriseColumn).Font.ColorIndex = 2
        }
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalEnterpriseColumn).HorizontalAlignment = -4108

        #Additional Monthly Basic
        if($Script:CustomerAdditionalMonthlyBasicProposals.Count -ge 1){
            
            $Script:CustomerAdditionalMonthlyBasicProposals | ForEach-Object {
                $ExistingNote = ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalBasicColumn).Comment.Shape.AlternativeText) -Replace "^.*?: "
                ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalBasicColumn)).Clear()
                $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalBasicColumn).AddComment("$ExistingNote`r`n$($_.Company) - $($_.ProposalName) [$($_.TicketNumber)] V$($_.ProposalVerison)")
                $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalBasicColumn).Comment.Shape.TextFrame.AutoSize = $True 
            }
            #$ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalBasicColumn).Comment.Visible = $False

            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalBasicColumn) = "=$([int]$Script:CustomerAdditionalMonthlyBasicProposals.Count) * $([math]::Round([decimal]$AdditionalBasicFee,2))"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalBasicColumn).NumberFormat = "$#,##0.00"
        }
        else{
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalBasicColumn) = "N/A"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalBasicColumn).Font.ColorIndex = 2
        }
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$AdditionalBasicColumn).HorizontalAlignment = -4108

        #Expedite
        if($ExpeditedProposals.Count -ge 1){

            $ExpeditedProposals | ForEach-Object {
                $ExistingNote = ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$ExpediteColumn).Comment.Shape.AlternativeText) -Replace "^.*?: "
                ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$ExpediteColumn)).Clear()
                $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$ExpediteColumn).AddComment("$ExistingNote`r`n$($_.Company) - $($_.ProposalName) [$($_.TicketNumber)] V$($_.ProposalVerison)")
                $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$ExpediteColumn).Comment.Shape.TextFrame.AutoSize = $True 
            }
            #$ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$ExpediteColumn).Comment.Visible = $False

            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$ExpediteColumn) = "=$([int]$ExpeditedProposals.Count) * $([math]::Round([decimal]$ExpediteFee,2))"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$ExpediteColumn).NumberFormat = "$#,##0.00"
        
        }
        else{
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$ExpediteColumn) = "N/A"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$ExpediteColumn).Font.ColorIndex = 2
        }
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$ExpediteColumn).HorizontalAlignment = -4108

        #Design Consulting Hours
        if(!($Script:TotalDCTimeThisMonth -eq 0.00)){

            $ExistingNote = ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$DesignConsultingHoursColumn).Comment.Shape.AlternativeText) -Replace "^.*?: "
            ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$DesignConsultingHoursColumn)).Clear()
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$DesignConsultingHoursColumn).AddComment("$ExistingNote`r`n$Script:ProposalsWithDCTimeThisMonth`r`n")
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$DesignConsultingHoursColumn).Comment.Shape.TextFrame.AutoSize = $True 
            #$ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$DesignConsultingHoursColumn).Comment.Visible = $False

            #$DCTimeMonthlyFee = [math]::Round([decimal]$Script:TotalDCTimeThisMonth * [int]$DCHourlyRate,2)
            #$ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$DesignConsultingHoursColumn) = "$([decimal]$DCTimeMonthlyFee)"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$DesignConsultingHoursColumn).Formula = "=$([math]::Round([decimal]$Script:TotalDCTimeThisMonth,2)) * $([math]::Round([decimal]$DCHourlyRate,2))"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$DesignConsultingHoursColumn).NumberFormat = "$#,##0.00"

        }
        else{
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$DesignConsultingHoursColumn) = "N/A"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$DesignConsultingHoursColumn).Font.ColorIndex = 2
        }
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$DesignConsultingHoursColumn).HorizontalAlignment = -4108

        #Operations Consulting Hours
        if(!($Script:TotalOCTimeThisMonth -eq 0.00)){

            $ExistingNote = ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$OperationsConsultingHoursColumn).Comment.Shape.AlternativeText) -Replace "^.*?: "
            ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$OperationsConsultingHoursColumn)).Clear()
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$OperationsConsultingHoursColumn).AddComment("$ExistingNote`r`n$ProposalsWithOCTimeThisMonth")
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$OperationsConsultingHoursColumn).Comment.Shape.TextFrame.AutoSize = $True 
            #$ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$DesignConsultingHoursColumn).Comment.Visible = $False

            #$DCTimeMonthlyFee = [math]::Round([decimal]$Script:TotalDCTimeThisMonth * [int]$DCHourlyRate,2)
            #$ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$OperationsConsultingHoursColumn) = "$([decimal]$DCTimeMonthlyFee)"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$OperationsConsultingHoursColumn).Formula = "=$([math]::Round([decimal]$Script:TotalOCTimeThisMonth,2)) * $([math]::Round([decimal]$OCHourlyRate,2))"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$OperationsConsultingHoursColumn).NumberFormat = "$#,##0.00"

        }
        else{
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$OperationsConsultingHoursColumn) = "N/A"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$OperationsConsultingHoursColumn).Font.ColorIndex = 2
        }
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$OperationsConsultingHoursColumn).HorizontalAlignment = -4108

        #>100
        if($ProposalsGT100Hours.Count -ge 1){
            
            $ProposalsGT100Hours | ForEach-Object {
                $ExistingNote = ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT100Column).Comment.Shape.AlternativeText) -Replace "^.*?: "
                ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT100Column)).Clear()
                $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT100Column).AddComment("$ExistingNote`r`n$($_.Company) - $($_.ProposalName) [$($_.TicketNumber)] V$($_.ProposalVerison)")
                $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT100Column).Comment.Shape.TextFrame.AutoSize = $True 
            }
            #$ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT100Column).Comment.Visible = $False

            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT100Column) = "=$([int]$ProposalsGT100Hours.Count) * $([int]99)"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT100Column).NumberFormat = "$#,##0.00"
        }
        else{
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT100Column) = "N/A"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT100Column).Font.ColorIndex = 2
        }
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT100Column).HorizontalAlignment = -4108

        #>200
        if($ProposalsGT200Hours.Count -ge 1){
            
            $ProposalsGT200Hours | ForEach-Object {
                $ExistingNote = ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT200Column).Comment.Shape.AlternativeText) -Replace "^.*?: "
                ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT200Column)).Clear()
                $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT200Column).AddComment("$ExistingNote`r`n$($_.Company) - $($_.ProposalName) [$($_.TicketNumber)] V$($_.ProposalVerison)")
                $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT200Column).Comment.Shape.TextFrame.AutoSize = $True 
            }
            #$ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT200Column).Comment.Visible = $False

            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT200Column) = "=$([int]$ProposalsGT200Hours.Count) * $([int]198)"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT200Column).NumberFormat = "$#,##0.00"
        }
        else{
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT200Column) = "N/A"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT200Column).Font.ColorIndex = 2
        }
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT200Column).HorizontalAlignment = -4108

        #>300
        if($ProposalsGT300Hours.Count -ge 1){
            
            $ProposalsGT300Hours | ForEach-Object {
                $ExistingNote = ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT300Column).Comment.Shape.AlternativeText) -Replace "^.*?: "
                ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT300Column)).Clear()
                $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT300Column).AddComment("$ExistingNote`r`n$($_.Company) - $($_.ProposalName) [$($_.TicketNumber)] V$($_.ProposalVerison)")
                $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT300Column).Comment.Shape.TextFrame.AutoSize = $True 
            }
            #$ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT300Column).Comment.Visible = $False

            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT300Column) = "=$([int]$ProposalsGT300Hours.Count) * $([int]297)"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT300Column).NumberFormat = "$#,##0.00"
        }
        else{
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT300Column) = "N/A"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT300Column).Font.ColorIndex = 2
        }
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT300Column).HorizontalAlignment = -4108

        #>400 
        if($ProposalsGT400Hours.Count -ge 1){
            
            $ProposalsGT400Hours | ForEach-Object {
                $ExistingNote = ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT400Column).Comment.Shape.AlternativeText) -Replace "^.*?: "
                ($ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT400Column)).Clear()
                $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT400Column).AddComment("$ExistingNote`r`n$($_.Company) - $($_.ProposalName) [$($_.TicketNumber)] V$($_.ProposalVerison)")
                $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT400Column).Comment.Shape.TextFrame.AutoSize = $True 
            }
            #$ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT400Column).Comment.Visible = $False

            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT400Column) = "=$([int]$ProposalsGT400Hours.Count) * $([int]396)"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT400Column).NumberFormat = "$#,##0.00"
        }
        else{
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT400Column) = "N/A"
            $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT400Column).Font.ColorIndex = 2
        }
        $ExcelSpreadsheet.ActiveSheet.Cells.Item([int]$CustomerRow,[int]$GT400Column).HorizontalAlignment = -4108
    }

    $MonthlyBillingForm.Close()

}

AuthenticateToSPO

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

$MonthlyBillingForm = New-Object System.Windows.Forms.Form
$MonthlyBillingForm.Text = "Aktis Helper - Monthly Billing Report"
$MonthlyBillingForm.Size = '575,250'
$MonthlyBillingForm.StartPosition = "CenterScreen"
$MonthlyBillingForm.MinimumSize = $MonthlyBillingForm.Size
$MonthlyBillingForm.MaximizeBox = $False
$MonthlyBillingForm.Topmost = $True

$MonthlyBillingLabel = New-Object Windows.Forms.Label
$MonthlyBillingLabel.Location = '10,10'
$MonthlyBillingLabel.AutoSize = $True
$MonthlyBillingLabel.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 9)
$MonthlyBillingLabel.Text = "Select date range to generate a billing report for:"

$StartDateLabel = New-Object Windows.Forms.Label
$StartDateLabel.Location = '35,40'
$StartDateLabel.AutoSize = $True
$StartDateLabel.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$StartDateLabel.Text = "Start date:"

$StartDateListBox = New-Object System.Windows.Forms.ListBox
$StartDateListBox.Location = New-Object System.Drawing.Point(110,41)
$StartDateListBox.Size = New-Object System.Drawing.Size(100,29)

$StartDateButton = New-Object System.Windows.Forms.Button
$StartDateButton.Location = '225,39'
$StartDateButton.Size = '21,21'
$StartDateButton.Width = 22
$StartDateButton.Text = "..."

$EndDateLabel = New-Object Windows.Forms.Label
$EndDateLabel.Location = '35,65'
$EndDateLabel.AutoSize = $True
$EndDateLabel.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 9, [System.Drawing.FontStyle]::Bold)
$EndDateLabel.Text = "End date:"

$EndDateListBox = New-Object System.Windows.Forms.ListBox
$EndDateListBox.Location = New-Object System.Drawing.Point(110,66)
$EndDateListBox.Size = New-Object System.Drawing.Size(100,29)

$EndDateButton = New-Object System.Windows.Forms.Button
$EndDateButton.Location = '225,64'
$EndDateButton.Size = '21,21'
$EndDateButton.Width = 22
$EndDateButton.Text = "..."

#$MonthDropDownBox = New-Object System.Windows.Forms.ComboBox
#$MonthDropDownBox.Location = New-Object System.Drawing.Size(260,25)
#$MonthDropDownBox.Size = New-Object System.Drawing.Size(80,20) 
#$MonthDropDownBoxLabel = New-Object Windows.Forms.Label
#$MonthDropDownBoxLabel.Location = New-Object System.Drawing.Point(95,102)
#$MonthDropDownBoxLabel.Text = "Month:"
#$MonthDropDownBoxLabel.ForeColor = 'Black'

$FolderSelectorLabel = New-Object Windows.Forms.Label
$FolderSelectorLabel.Location = '10,100'
$FolderSelectorLabel.AutoSize = $True
$FolderSelectorLabel.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 9)
$FolderSelectorLabel.Text = "Select a location to save the monthly billing report to:"

$FolderSelectorListBox = New-Object System.Windows.Forms.ListBox
$FolderSelectorListBox.Location = New-Object System.Drawing.Point(10,125)
$FolderSelectorListBox.Size = New-Object System.Drawing.Size(450,29)

$FolderSelectorButton = New-Object System.Windows.Forms.Button
$FolderSelectorButton.Location = '470,122'
$FolderSelectorButton.Size = '100,21'
$FolderSelectorButton.Width = 75
$FolderSelectorButton.Text = "Browse"

$MonthlyBillingSubmitButton = New-Object System.Windows.Forms.Button
$MonthlyBillingSubmitButton.Location = '205,165'
$MonthlyBillingSubmitButton.Size = '100,23'
$MonthlyBillingSubmitButton.Width = 150
$MonthlyBillingSubmitButton.Text = "Generate Report"

$FolderSelectorButton_Click = {

    $SavedBillingReport = New-Object System.Windows.Forms.SaveFileDialog
    $SavedBillingReport.FileName = "Project Fuel Billing Report -- $($Script:StartTime.ToString("MM-dd-yyyy"))-to-$($Script:EndTime.ToString("MM-dd-yyyy")).xlsx"
    $SavedBillingReport.ShowDialog()
    $Script:NewMonthlyBillingReportName = $SavedBillingReport.FileName

    $FolderSelectorListBox.Items.Clear()
    $FolderSelectorListBox.Items.Add($Script:NewMonthlyBillingReportName)

}

$MonthlyBillingButton_Click = {

    GenerateMonthlyBillingReport

}

$StartDateButton_Click = {

    $StartDateListBox.Items.Clear()
    $Script:StartTime = [DateTime]$(GetTimeFrameDates)
    $StartDateListBox.Items.Add("$($Script:StartTime.ToShortDateString())")

}

$EndDateButton_Click = {

    $EndDateListBox.Items.Clear()
    $Script:EndTime = [DateTime]$(GetTimeFrameDates)
    $EndDateListBox.Items.Add("$($Script:EndTime.ToShortDateString())")

}

$FolderSelectorButton.Add_Click($FolderSelectorButton_Click)
$MonthlyBillingSubmitButton.Add_Click($MonthlyBillingButton_Click)
$StartDateButton.Add_Click($StartDateButton_Click)
$EndDateButton.Add_Click($EndDateButton_Click)

$MonthlyBillingStatusBar = New-Object System.Windows.Forms.StatusBar
$MonthlyBillingStatusBar.Text = "Ready..."

$MonthlyBillingForm.SuspendLayout()
$MonthlyBillingForm.Controls.Add($MonthlyBillingLabel)
$MonthlyBillingForm.Controls.Add($StartDateLabel)
$MonthlyBillingForm.Controls.Add($StartDateListBox)
$MonthlyBillingForm.Controls.Add($StartDateButton)
$MonthlyBillingForm.Controls.Add($EndDateLabel)
$MonthlyBillingForm.Controls.Add($EndDateListBox)
$MonthlyBillingForm.Controls.Add($EndDateButton)
$MonthlyBillingForm.Controls.Add($FolderSelectorLabel)
$MonthlyBillingForm.Controls.Add($FolderSelectorListBox)
$MonthlyBillingForm.Controls.Add($FolderSelectorButton)
$MonthlyBillingForm.Controls.Add($MonthlyBillingSubmitButton)
$MonthlyBillingForm.ResumeLayout()

$MonthlyBillingFormResult = $MonthlyBillingForm.ShowDialog()
    
if($MonthlyBillingFormResult -eq [System.Windows.Forms.DialogResult]::OK){
    # Add OK action
}
elseif($MonthlyBillingFormResult -eq [System.Windows.Forms.DialogResult]::Cancel){
    $MonthlyBillingForm.Close()
    $MonthlyBillingForm.Dispose()

}
elseif(!($MonthlyBillingFormResult -eq [System.Windows.Forms.DialogResult]::OK)){
    $MonthlyBillingForm.Close()
}   