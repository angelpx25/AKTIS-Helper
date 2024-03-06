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
    Write-EventLog -Log AKTIS -Source "Aktis Helper" -EventID 10 -EntryType $LogType -Message "$LogMessage"
    #Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "This is a warning message..."
}

function Test-FileLock {
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

function ThrowPostSuccessBalloon {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    $ToastNotification = New-Object System.Windows.Forms.NotifyIcon
    $ToastNotification.Icon = [System.Drawing.SystemIcons]::Information
    $ToastNotification.BalloonTipText = "Successfully posted proposal statistics data for $Script:ProposalName to Aktis!"
    $ToastNotification.BalloonTipTitle = "Aktis"
    $ToastNotification.BalloonTipIcon = "Info"
    $ToastNotification.Visible = $True
    $ToastNotification.ShowBalloonTip(50000)
    $ToastNotification.Dispose()
}

function GetTimeFrameDates {

    # January
    [DateTime]$January = Get-Date -Format "01/yyyy"
    $LastDayofJanuary = [DateTime]::DaysInMonth($January.Year, $January.Month)
    $Script:JanuaryStart = [DateTime]::new($January.Year, $January.Month, 1)
    $Script:JanuaryEnd = [DateTime]::new($January.Year, $January.Month, $LastDayofJanuary, "23", "59", "59")
    
    # February
    [DateTime]$February = Get-Date -Format "02/yyyy"
    $LastDayofFebruary = [DateTime]::DaysInMonth($February.Year, $February.Month)
    $Script:FebruaryStart = [DateTime]::new($February.Year, $February.Month, 1)
    $Script:FebruaryEnd = [DateTime]::new($February.Year, $February.Month, $LastDayofFebruary, "23", "59", "59")

    # March
    [DateTime]$March = Get-Date -Format "03/yyyy"
    $LastDayofMarch = [DateTime]::DaysInMonth($March.Year, $March.Month)
    $Script:MarchStart = [DateTime]::new($March.Year, $March.Month, 1)
    $Script:MarchEnd = [DateTime]::new($March.Year, $March.Month, $LastDayofMarch, "23", "59", "59")

    # April
    [DateTime]$April = Get-Date -Format "04/yyyy"
    $LastDayofApril = [DateTime]::DaysInMonth($April.Year, $April.Month)
    $Script:AprilStart = [DateTime]::new($April.Year, $April.Month, 1)
    $Script:AprilEnd = [DateTime]::new($April.Year, $April.Month, $LastDayofApril, "23", "59", "59")

    # May
    [DateTime]$May = Get-Date -Format "05/yyyy"
    $LastDayofMay = [DateTime]::DaysInMonth($May.Year, $May.Month)
    $Script:MayStart = [DateTime]::new($May.Year, $May.Month, 1)
    $Script:MayEnd = [DateTime]::new($May.Year, $May.Month, $LastDayofMay, "23", "59", "59")

    # June
    [DateTime]$June = Get-Date -Format "06/yyyy"
    $LastDayofJune = [DateTime]::DaysInMonth($June.Year, $June.Month)
    $Script:JuneStart = [DateTime]::new($June.Year, $June.Month, 1)
    $Script:JuneEnd = [DateTime]::new($June.Year, $June.Month, $LastDayofJune, "23", "59", "59")

    # July
    [DateTime]$July = Get-Date -Format "07/yyyy"
    $LastDayofJuly = [DateTime]::DaysInMonth($July.Year, $July.Month)
    $Script:JulyStart = [DateTime]::new($July.Year, $July.Month, 1)
    $Script:JulyEnd = [DateTime]::new($July.Year, $July.Month, $LastDayofJuly, "23", "59", "59")

    # August
    [DateTime]$August = Get-Date -Format "08/yyyy"
    $LastDayofAugust = [DateTime]::DaysInMonth($August.Year, $August.Month)
    $Script:AugustStart = [DateTime]::new($August.Year, $August.Month, 1)
    $Script:AugustEnd = [DateTime]::new($August.Year, $August.Month, $LastDayofAugust, "23", "59", "59")

    # September
    [DateTime]$September = Get-Date -Format "09/yyyy"
    $LastDayofSeptember = [DateTime]::DaysInMonth($September.Year, $September.Month)
    $Script:SeptemberStart = [DateTime]::new($September.Year, $September.Month, 1)
    $Script:SeptemberEnd = [DateTime]::new($September.Year, $September.Month, $LastDayofSeptember, "23", "59", "59")

    # October
    [DateTime]$October = Get-Date -Format "10/yyyy"
    $LastDayofOctober = [DateTime]::DaysInMonth($October.Year, $October.Month)
    $Script:OctoberStart = [DateTime]::new($October.Year, $October.Month, 1)
    $Script:OctoberEnd = [DateTime]::new($October.Year, $October.Month, $LastDayofOctober, "23", "59", "59")

    # November
    [DateTime]$November = Get-Date -Format "11/yyyy"
    $LastDayofNovember = [DateTime]::DaysInMonth($November.Year, $November.Month)
    $Script:NovemberStart = [DateTime]::new($November.Year, $November.Month, 1)
    $Script:NovemberEnd = [DateTime]::new($November.Year, $November.Month, $LastDayofNovember, "23", "59", "59")

    # December
    [DateTime]$December = Get-Date -Format "12/yyyy"
    $LastDayofDecember = [DateTime]::DaysInMonth($December.Year, $December.Month)
    $Script:DecemberStart = [DateTime]::new($December.Year, $December.Month, 1)
    $Script:DecemberEnd = [DateTime]::new($December.Year, $December.Month, $LastDayofDecember, "23", "59", "59")

    #$SelectedMonth = $MonthDropDownBox.SelectedItem
    #$SelectedMonth = "November"
    $CurrentMonthOfYear = $(Get-Date).ToString('MMMM', [System.Globalization.CultureInfo]::CurrentCulture)
    
    $Script:StartTime = Invoke-Expression $("$" + "Script:" + $CurrentMonthOfYear + "Start")
    $Script:EndTime = Invoke-Expression $("$" + "Script:" + $CurrentMonthOfYear + "End")

    #$Script:StartTime = [DateTime]"10/31/22"
    #$Script:EndTime = [DateTime]"12/04/22"

}

function AuthenticateToSPO {

    try{
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $Script:ConnectToSPO = Connect-PnPOnline -Url "https://projectfuelnow.sharepoint.com/sites/FuelNow" -Tenant "projectfuelnow.com" -ClientId 09b4038a-b7ed-4784-99ef-09fcfe12eaef -Thumbprint 440FC2B08CC7586C815CACE316B72C79C7FDDE21 -ReturnConnection
        #$Script:ConnectToSPO = Connect-PnPOnline -Url "https://projectfuelnow.sharepoint.com/sites/FuelNow" -Tenant "projectfuelnow.com" -ClientId 25ff3a8d-9601-450c-b585-7dd72dd5de66 -Thumbprint 24becf80e51ca4a7d8d8ff6260e648e4cf604c23 -ReturnConnection
        #$Script:ConnectToSPO = Connect-PnPOnline -Url "https://projectfuelnow.sharepoint.com/sites/Fuelnow" -Interactive -ForceAuthentication -ErrorAction SilentlyContinue
    }
    catch{
        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to authenticate to SharePoint Online..."
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to authenticate to SharePoint Online...`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Authenticate to SharePoint Online" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to authenticate to SharePoint Online..." -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
        EXIT
    }

    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Successfully authenticated to SharePoint Online!"
    Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Successfully authenticated to SharePoint Online!"

}    

function ExecuteAktisHelper ($FilePath) {

            Get-FinancialDataFromWorkplan $FilePath
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Processing $FilePath"
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Proposal Name: $Script:ProposalName"
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Workplan location: $FileName"
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Proposed Hours: $([decimal]$Script:ProposedHours)"
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Hourly Rate: `$$([decimal]$Script:HourlyRate)"
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Total MRR: `$$([decimal]$Script:TotalMRR)"
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): PS Revenue: `$$([decimal]$Script:PSRevenue)"
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Products Revenue: `$$([decimal]$Script:ProductsRevenue)"
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Total NRR: `$$([decimal]$Script:TotalNRR)"

            Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Processing $FilePath`r`nProposal Name: $Script:ProposalName`r`nWorkplan location: $FileName`r`nProposed Hours: $([decimal]$Script:ProposedHours)`r`nHourly Rate: `$$([decimal]$Script:HourlyRate)`r`nTotal MRR: `$$([decimal]$Script:TotalMRR)`r`nPS Revenue: `$$([decimal]$Script:PSRevenue)`r`nProducts Revenue: `$$([decimal]$Script:ProductsRevenue)`r`nTotal NRR: `$$([decimal]$Script:TotalNRR)"

}

function Get-FinancialDataFromWorkplan ($ProposalFilePath) {

    try{

        if(!(Test-Path -Path "C:\Aktis\Temp")){
            New-Item -Path "C:\Aktis\Temp" -ItemType "Directory" | Out-Null
        }
        
        Copy-Item -Path $ProposalFilePath -Destination "C:\Aktis\Temp\$($ProposalFilePath | Split-Path -Leaf)" -Force

        $FilePath = Get-Item "C:\Aktis\Temp\$($ProposalFilePath | Split-Path -Leaf)"

        if($(Test-FileLock $FilePath) -eq "True"){

            Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): The following file is already open. Forcing file closed.`r`n$Filepath`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following file is already open. Forcing file closed.`r`n$Filepath`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"

            Stop-Process -Name "Excel" -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 5

        }

        $ExcelOpen = New-Object -COMObject Excel.Application
        $ExcelOpen.Visible = $False
        $ExcelSpreadsheet = $ExcelOpen.Workbooks.Open($FilePath)

    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to open Excel file:`r`n$FilePath. Ensure Microsoft Excel is installed.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to open Excel workplan for $Script:LongProposalName to extract Proposal Statistics data from. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Open Excel Workplan" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to open Excel workplan for $Script:LongProposalName to extract Proposal Statistics data from. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort

    }

    try{
        $ProposalNameHeadingColumnLetter = (($ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find("Project Name:")).Address($False,$False)) -Replace "\d+"
    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`nProject Name:`r`n`r`n. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`nProject Name:`r`n`r`n. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Find Data in Workplan" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`nProject Name:`r`n`r`nManually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort

    }

    $ProposalNameRow = $ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find("Project Name:").Row + 1

    try{
        $Script:ProposalName = $ExcelSpreadsheet.Sheets.Item(1).Range("$ProposalNameHeadingColumnLetter" + "$ProposalNameRow").Value2
    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to obtain the value of the Proposal Name cell for $Script:LongProposalName.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to obtain the value of the Proposal Name cell for $Script:LongProposalName.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Find Data in Workplan" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Cannot find value of Proposal Name cell. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
        
    }

    try{
        $Script:AktisKTNumber = ($ExcelSpreadsheet.Sheets.Item(1).Range("$ProposalNameHeadingColumnLetter" + "$ProposalNameRow").Comment.Shape.AlternativeText) -Replace "^.*?: "
    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to obtain the value of the Proposal Name cell note for $Script:LongProposalName, which is expected to contain the KT Number.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to obtain the value of the Proposal Name cell note for $Script:LongProposalName, which is expected to contain the KT Number.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Find Data in Workplan" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Cannot find value of the Proposal Name cell note, which is expected to contain the KT Number. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
        
    }

    try{
        $QuantityColumnLetter = (($ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find("Quantity")).Address($False,$False)) -Replace "\d+"
    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`nQuantity:`r`n`r`n. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`nQuantity:`r`n`r`n. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Find Data in Workplan" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`nQuantity:`r`n`r`nManually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
    
    }

    try{
        $PPUColumnLetter = (($ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find("PPU")).Address($False,$False)) -Replace "\d+"
    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`nPPU:`r`n`r`n. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`nPPU:`r`n`r`n. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Find Data in Workplan" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`nPPU:`r`n`r`nManually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
        
    }

    try{
        $ExtendedPriceColumnLetter = (($ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find("Extended Price")).Address($False,$False)) -Replace "\d+"
    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`nExtended Price:`r`n`r`n. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`nExtended Price:`r`n`r`n. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Find Data in Workplan" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`nExtended Price:`r`n`r`nManually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
        
    }

    try{
        $NRRTotalRow = $ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find("Non Recurring Fees Total:").Row
    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`nNon Recurring Fees Total:`r`n`r`n. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`nNon Recurring Fees Total:`r`n`r`n. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Find Data in Workplan" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`nNon Recurring Fees Total:`r`n`r`nManually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
        
    }

    try{
        $MRRTotalRow = $ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find("Monthly Recurring Fees Total:").Row
    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`nMonthly Recurring Fees Total:`r`n`r`n. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`nMonthly Recurring Fees Total:`r`n`r`n. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Find Data in Workplan" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`nMonthly Recurring Fees Total:`r`n`r`nManually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
        
    }

    try{
        $ProposedHoursCellLocation = $QuantityColumnLetter + $ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find(" Professional Services").Row
    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`n Professional Services`r`n`r`n. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`n Professional Services`r`n`r`n. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Find Data in Workplan" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Missing data:`r`n`r`n Professional Services`r`n`r`nManually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
         
    }

    $HourlyRateCellLocation = $PPUColumnLetter + $ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find(" Professional Services").Row

    $TotalMRRCellLocation = $ExtendedPriceColumnLetter + $MRRTotalRow

    $PSRevenueCellLocation = $ExtendedPriceColumnLetter + $ExcelSpreadsheet.Sheets.Item(1).UsedRange.Find(" Professional Services").Row

    $TotalNRRCellLocation = $ExtendedPriceColumnLetter + $NRRTotalRow


    try{
        $Script:ProposalNumber = ($Script:ProposalName | Select-String "(?<=\[)[^]]+(?=\])").Matches.Value
    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Failed to extract proposal number from full proposal name. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to find data in Excel workplan for $Script:LongProposalName. Failed to extract proposal number from full proposal name. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Find Data in Workplan" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Failed to extract proposal number from full proposal name. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
        
    }
    
    try{
        $Script:ProposalVersion = $Script:ProposalName -Replace "^.*?] V"
    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Failed to extract proposal version from full proposal name. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to find data in Excel workplan for $Script:LongProposalName. Failed to extract proposal version from full proposal name. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Find Data in Workplan" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to find data in Excel workplan for $Script:LongProposalName. Failed to extract proposal version from full proposal name. Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
        
    }
    
    if($Script:ProposalVersion -eq $Script:ProposalName){
        $Script:ProposalVersion = "1" 
    }

    [decimal]$Script:ProposedHours = [math]::Round($ExcelSpreadsheet.Sheets.Item(1).Range($ProposedHoursCellLocation).Value2,2)

    [decimal]$Script:PSRevenue = [math]::Round($ExcelSpreadsheet.Sheets.Item(1).Range($PSRevenueCellLocation).Value2,2)
    
    [decimal]$Script:HourlyRate = [math]::Round($ExcelSpreadsheet.Sheets.Item(1).Range($HourlyRateCellLocation).Value2,2)

    if($Null -eq $Script:HourlyRate){
        [decimal]$Script:HourlyRate = [math]::Round([int]$Script:PSRevenue / [int]$Script:ProposedHours,2)
    }

    [decimal]$Script:TotalNRR = [math]::Round($ExcelSpreadsheet.Sheets.Item(1).Range($TotalNRRCellLocation).Value2,2)

    [decimal]$Script:TotalMRR = [math]::Round($ExcelSpreadsheet.Sheets.Item(1).Range($TotalMRRCellLocation).Value2,2)
    
    ##### This needs to be refined
    [decimal]$Script:ProductsRevenue = [math]::Round([int]$Script:TotalNRR - [int]$Script:PSRevenue,2)
   
    $ExcelSpreadsheet.Close($True)
    $ExcelOpen.Quit()

    Start-Sleep -Seconds "2"

    Remove-Item -Path "C:\Aktis\Temp\$($ProposalFilePath | Split-Path -Leaf)"
    
}

function GetExistingAnalyticsData {

    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Retreiving existing Aktis analytics data..."
    Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Retreiving existing Aktis analytics data..."

    try{
        $ProposalStatisticsRawData = (Get-PnPListItem -List "Proposal Statistics" -Fields "KTNumber","ID","ProposalName","MSP","TicketNumber","ProposalVerison","Company","IntakeType","KTDate","DateProposalCompleted","DeliveredtoClient","Subscription","_x0035_Day","_x0033_DayExpedite","DateProposalWritten","DDELead","PresenttoClient_x002f_Staff","Draft1","Draft2","GratisRevision","ProposalExceedsSubscription_x002","ProposalExceedsSubscription_x0020" -Connection $Script:ConnectToSPO).FieldValues
    
        $ProposalStatisticsData = @()
    
        $ProposalStatisticsRawData | ForEach-Object {
            $ProposalStatisticsData += $_
        }
    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to retrieve Aktis analytics data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Aktis analytics data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Retrieve Proposal Statistics Data" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to retrieve Proposal Statistics data SharePoint List.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort

    }

    try{

        $Script:ProposalID = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).ID
        $Script:LocalAnalyticsProposalName = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).ProposalName
        $Script:TicketNumber = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).TicketNumber
        $Script:ProposalVersionNumber = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).ProposalVerison
        $Script:ProposalMSP = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).MSP
        $Script:ProposalMSPsClient = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).Company
        $Script:ProposalIntakeType = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).IntakeType
        $Script:ProposalKTDate = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).KTDate
        $Script:ProposalDateCompleted = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).DateProposalCompleted
        $Script:ProposalDeliveredToClient = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).DeliveredtoClient
        $Script:ProposalSubscription = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).Subscription
        $Script:ProposalBiOp5Day = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber})._x0035_Day
        $Script:ProposalBiOp3DayExpedite = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber})._x0033_DayExpedite
        $Script:DateProposalWritten = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).DateProposalWritten
        $Script:ExistingDDELead = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).DDELead
        $Script:ProposalBiOpPresentToClient = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).PresenttoClient_x002f_Staff
        $Script:ProposalBiOpDraft1 = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).Draft1
        $Script:ProposalBiOpDraft2 = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).Draft2
        $Script:IsGratisRevision = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).GratisRevision
        $Script:ProposalExceedsSubEnt = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).ProposalExceedsSubscription_x002
        $Script:ProposalExceedsSubBasic = ($ProposalStatisticsData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).ProposalExceedsSubscription_x0020
        $Script:LongProposalName = "$Script:LocalAnalyticsProposalName [$Script:TicketNumber] V$Script:ProposalVersionNumber"

        if($Script:ProposalVersionNumber -eq "1"){

            $Script:ProposalFileName = "$Script:LocalAnalyticsProposalName $Script:TicketNumber"

        }
        elseif($Script:ProposalVersionNumber -gt "1"){

            $Script:ProposalFileName = "$Script:LocalAnalyticsProposalName $Script:TicketNumber V$Script:ProposalVersionNumber"

        }

    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to retrieve existing analytics data for:`r`n$Script:LongProposalName.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve existing analytics data for:`r`n$Script:LongProposalName.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Extract Proposal Statistics Details from Cached Data" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to extract Proposal Statistics details from cached dataset for $Script:LongProposalName`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
        
    }

    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Successfully retrieved existing analytics data for $Script:LongProposalName from Aktis!"
    Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Successfully retrieved existing analytics data for $Script:LongProposalName from Aktis!"
    
}

function RetrieveClockifyTimeData {

    $Script:APIEndpoint = "https://api.clockify.me/api/v1"
    $Script:ClockifyWorkSpaceID = "627bafedcc6826493436b133"
    $Script:ClockifyAPIKey = "OTRkNGE0NDYtNTA4Mi00MmM4LThhNDQtZDIzYTk2NzY0NTY3"

    [datetime]$CurrentDate = Get-Date
    [string]$CurrentDateISO8601 = Get-Date (Get-Date).ToUniversalTime() -UFormat '+%Y-%m-%dT%H:%M:%S.000Z'
    [string]$DaysAgoISO8601 = Get-Date ($CurrentDate).AddDays(-60).ToUniversalTime() -UFormat '+%Y-%m-%dT%H:%M:%S.000Z'
    [int]$PageSize = "4000"
  
    try{
        $ClockifyProjectDetailsRawData = Invoke-RestMethod -Uri ($Script:APIEndpoint + "/workspaces/" + $Script:ClockifyWorkSpaceID + "/projects") -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $Script:ClockifyAPIKey} -Method Get -Body @{
            "start" = $DaysAgoISO8601
            "end" = $CurrentDateISO8601
            "page-size" = $PageSize
          }
    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to retrieve Clockify project details.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Clockify project details.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Retrieve Clockify Project Details" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to retrieve Clockify project details.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort

    }

    $ClockifyProjectDetails = @()
    $ClockifyProjectDetailsRawData | ForEach-Object {
      $ClockifyProjectDetails += $_
    }

    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Verifying Clockify project exists for $Script:LongProposalName..."
    Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Verifying Clockify project exists for $Script:LongProposalName..."

    if(!($ClockifyProjectDetails.Name | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber")){

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Clockify project doesn't exist for $("$Script:TicketNumber V$Script:ProposalVersionNumber").`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Clockify project doesn't exist for $("$Script:TicketNumber V$Script:ProposalVersionNumber").`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Clockify Project Does Not Exist" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Clockify project doesn't exist for $("$Script:TicketNumber V$Script:ProposalVersionNumber"). Manually address this.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
       
        ### This is a workflow stoppage
        #RetrieveClockifyTimeData

    }
    else{

        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Clockify project ($Script:LongProposalName) exists!"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Clockify project ($Script:LongProposalName) exists!"

    }

    $ClockifyProjectDetails | ForEach-Object {
        
        if($_.Name -eq "$Script:TicketNumber V$Script:ProposalVersionNumber"){
            $Script:ClockifyProjectID = $_.ID

            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Retrieving Clockify task level details for $Script:LongProposalName..."
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Retrieving Clockify task level details for $Script:LongProposalName..."
            GetProjectTaskLevelDetails

            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Retrieving Clockify user time data for $Script:LongProposalName..."
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Retrieving Clockify user time data for $Script:LongProposalName..."
            GetClockifyUserTimeEntryData

            $AngelRunningTimeEntryProjectID = ($Script:AngelTimeEntries | Where-Object {$Null -eq $_.TimeInterval.End}).ProjectID
            $AdamRunningTimeEntryProjectID = ($Script:AdamTimeEntries | Where-Object {$Null -eq $_.TimeInterval.End}).ProjectID
            $JohnatanRunningTimeEntryProjectID = ($Script:JohnatanTimeEntries | Where-Object {$Null -eq $_.TimeInterval.End}).ProjectID
    
            if($AngelRunningTimeEntryProjectID -eq $Script:ClockifyProjectID){
                
                Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Angel currently has a running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n$Script:LongProposalName`r`n`r`nSending email notification to place proposal in correct status, if required, or stop currently running time entry."
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "Angel currently has a running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n$Script:LongProposalName`r`n`r`nSending email notification to place proposal in correct status, if required, or stop currently running time entry."
                Send-MailMessage -From $Script:DefaultSenderAddress -To "angel@projectfuelnow.com","ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Active Clockify Time Entry" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nHey! It's Aktis Helper.`r`n`r`nAngel, you have a running time entry in Clockify for $Script:LongProposalName. If you're still working on this proposal, please place it in the correct status. Otherwise, please stop the time entry.`r`n`r`nI'll check if the time entry is still running in one minute. If it is, I won't be able to account for your currently running time punch - please finish your time punch and manually run Aktis Helper for this proposal." -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort

                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Waiting 60 seconds for Angel to stop his currently running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n$Script:LongProposalName."
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Waiting 60 seconds for Angel to stop his currently running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n$Script:LongProposalName."
                Start-Sleep -Seconds 75

                GetProjectTaskLevelDetails
                GetClockifyUserTimeEntryData

                $AngelRunningTimeEntryProjectID = ($Script:AngelTimeEntries | Where-Object {$Null -eq $_.TimeInterval.End}).ProjectID
                if($AngelRunningTimeEntryProjectID -eq $Script:ClockifyProjectID){

                    Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Not accounting for Angel's running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n$Script:LongProposalName`r`n`r`nSending email notification."
                    Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Not accounting for Angel's running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n$Script:LongProposalName`r`n`r`nSending email notification."
                    Send-MailMessage -From $Script:DefaultSenderAddress -To "angel@projectfuelnow.com","ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Active Clockify Time Entry" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nHey! It's Aktis Helper, again.`r`n`r`nAngel, I can't commit your currently running time entry in Clockify for $Script:LongProposalName to Aktis Proposal Statistics. Please finish your time punch and manually run Aktis Helper for this proposal. Thank you!" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort

                }

            }
    
            if($AdamRunningTimeEntryProjectID -eq $Script:ClockifyProjectID){
    
                Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Adam currently has a running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n$Script:LongProposalName`r`n`r`nSending email notification to place proposal in correct status, if required, or stop currently running time entry."
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "Adam currently has a running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n$Script:LongProposalName`r`n`r`nSending email notification to place proposal in correct status, if required, or stop currently running time entry."
                Send-MailMessage -From $Script:DefaultSenderAddress -To "adam@projectfuelnow.com","ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Active Clockify Time Entry" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nHey! It's Aktis Helper.`r`n`r`nAdam, you have a running time entry in Clockify for $Script:LongProposalName. If you're still working on this proposal, please place it in the correct status. Otherwise, please stop the time entry.`r`n`r`nI'll check if the time entry is still running in one minute. If it is, I won't be able to account for your currently running time punch - please finish your time punch and manually run Aktis Helper for this proposal." -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
                
                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Waiting 60 seconds for Adam to stop his currently running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n$Script:LongProposalName."
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Waiting 60 seconds for Adam to stop his currently running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n$Script:LongProposalName."
                Start-Sleep -Seconds 75
                
                GetProjectTaskLevelDetails
                GetClockifyUserTimeEntryData
                
                $AdamRunningTimeEntryProjectID = ($Script:AdamTimeEntries | Where-Object {$Null -eq $_.TimeInterval.End}).ProjectID
                if($AdamRunningTimeEntryProjectID -eq $Script:ClockifyProjectID){
                
                    Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Not accounting for Adam's running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n$Script:LongProposalName`r`n`r`nSending email notification."
                    Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Not accounting for Adam's running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n$Script:LongProposalName`r`n`r`nSending email notification."
                    Send-MailMessage -From $Script:DefaultSenderAddress -To "adam@projectfuelnow.com","ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Active Clockify Time Entry" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nHey! It's Aktis Helper, again.`r`n`r`nAdam, I can't commit your currently running time entry in Clockify for $Script:LongProposalName to Aktis Proposal Statistics. Please finish your time punch and manually run Aktis Helper for this proposal. Thank you!" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
                
                }                 

            }
    
            if($JohnatanRunningTimeEntryProjectID -eq $Script:ClockifyProjectID){
          
                Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Johnatan currently has a running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n$Script:LongProposalName`r`n`r`nSending email notification to place proposal in correct status, if required, or stop currently running time entry."
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "Johnatan currently has a running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n$Script:LongProposalName`r`n`r`nSending email notification to place proposal in correct status, if required, or stop currently running time entry."
                Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Active Clockify Time Entry" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nHey! It's Aktis Helper.`r`n`r`nJohnatan, you have a running time entry in Clockify for $Script:LongProposalName. If you're still working on this proposal, please place it in the correct status. Otherwise, please stop the time entry.`r`n`r`nI'll check if the time entry is still running in one minute. If it is, I won't be able to account for your currently running time punch - please finish your time punch and manually run Aktis Helper for this proposal." -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
                
                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Waiting 60 seconds for Johnatan to stop his currently running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n$Script:LongProposalName."
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Waiting 60 seconds for Johnatan to stop his currently running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n$Script:LongProposalName."
                Start-Sleep -Seconds 75
                
                GetProjectTaskLevelDetails
                GetClockifyUserTimeEntryData
                
                $JohnatanRunningTimeEntryProjectID = ($Script:JohnatanTimeEntries | Where-Object {$Null -eq $_.TimeInterval.End}).ProjectID
                if($JohnatanRunningTimeEntryProjectID -eq $Script:ClockifyProjectID){
                
                    Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Not accounting for Johnatan's running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n$Script:LongProposalName`r`n`r`nSending email notification."
                    Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Not accounting for Johnatan's running Clockify time entry for proposal:`r`n$($Script:ClockifyProjectLevelDetails.Name)`r`n$Script:LongProposalName`r`n`r`nSending email notification."
                    Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Active Clockify Time Entry" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nHey! It's Aktis Helper, again.`r`n`r`nJohnatan, I can't commit your currently running time entry in Clockify for $Script:LongProposalName to Aktis Proposal Statistics. Please finish your time punch and manually run Aktis Helper for this proposal. Thank you!" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
                
                }                 
            }
        }
    }
}

function GetClockifyUserTimeEntryData {

    [datetime]$CurrentDate = Get-Date
    [string]$CurrentDateISO8601 = Get-Date (Get-Date).ToUniversalTime() -UFormat '+%Y-%m-%dT%H:%M:%S.000Z'
    [string]$DaysAgoISO8601 = Get-Date ($CurrentDate).AddDays(-7).ToUniversalTime() -UFormat '+%Y-%m-%dT%H:%M:%S.000Z'
    [int]$PageSize = "4000"

    try{

        $Script:AngelUserID = (Invoke-RestMethod -Uri ($Script:APIEndpoint + "/workspaces/" + $Script:ClockifyWorkSpaceID + "/users") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $Script:ClockifyAPIKey} -Body @{email = "angel@projectfuelnow.com"}).ID
        $Script:AdamUserID = (Invoke-RestMethod -Uri ($Script:APIEndpoint + "/workspaces/" + $Script:ClockifyWorkSpaceID + "/users") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $Script:ClockifyAPIKey} -Body @{email = "adam@projectfuelnow.com"}).ID
        $Script:JohnatanUserID = (Invoke-RestMethod -Uri ($Script:APIEndpoint + "/workspaces/" + $Script:ClockifyWorkSpaceID + "/users") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $Script:ClockifyAPIKey} -Body @{email = "johnatan@projectfuelnow.com"}).ID
        
        $Script:AngelEmailAddress = "angel@projectfuelnow.com"
        $Script:AdamEmailAddress = "adam@projectfuelnow.com"
        $Script:JohnatanEmailAddress = "johnatan@projectfuelnow.com"

    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to retrieve Clockify user IDs.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Clockify user IDs.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Retrieve Clockify User IDs" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nFailed to retrieve Clockify user IDs.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
                
    }

    try{
        $AdamRawTimeEntryData = Invoke-RestMethod -Uri ($Script:APIEndpoint + "/workspaces/" + $Script:ClockifyWorkSpaceID + "/user/" + $Script:AdamUserID + "/time-entries") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $Script:ClockifyAPIKey} -Body @{
            "start" = $DaysAgoISO8601
            "end" = $CurrentDateISO8601
            "page-size" = $PageSize
          }
          
          $AngelRawTimeEntryData = Invoke-RestMethod -Uri ($Script:APIEndpoint + "/workspaces/" + $Script:ClockifyWorkSpaceID + "/user/" + $Script:AngelUserID + "/time-entries") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $Script:ClockifyAPIKey} -Body @{
            "start" = $DaysAgoISO8601
            "end" = $CurrentDateISO8601
            "page-size" = $PageSize
          }
          
          $JohnatanRawTimeEntryData = Invoke-RestMethod -Uri ($Script:APIEndpoint + "/workspaces/" + $Script:ClockifyWorkSpaceID + "/user/" + $Script:JohnatanUserID + "/time-entries") -Method Get -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $Script:ClockifyAPIKey} -Body @{
            "start" = $DaysAgoISO8601
            "end" = $CurrentDateISO8601
            "page-size" = $PageSize
          }
    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to retrieve Clockify time entry data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Clockify time entry data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Retrieve Clockify User Time Entry Data" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nFailed to retrieve Clockify user time entry data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
          
    }

    $Script:AdamTimeEntries = @()
    $Script:AngelTimeEntries = @()
    $Script:JohnatanTimeEntries = @()
    
    $AdamRawTimeEntryData | ForEach-Object {
      $Script:AdamTimeEntries += $_
    }
    
    $AngelRawTimeEntryData | ForEach-Object {
      $Script:AngelTimeEntries += $_
    }
    
    $JohnatanRawTimeEntryData | ForEach-Object {
        $Script:JohnatanTimeEntries += $_
    }

    $AngelPCTime = ($Script:AngelTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyPCTaskID}).TimeInterval.Duration
    $AngelPWTime = ($Script:AngelTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyPWTaskID}).TimeInterval.Duration
    $AngelDCTime = ($Script:AngelTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyDCTaskID}).TimeInterval.Duration
    $AngelIPTime = ($Script:AngelTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyIPTaskID}).TimeInterval.Duration

    $AdamPCTime = ($Script:AdamTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyPCTaskID}).TimeInterval.Duration
    $AdamPWTime = ($Script:AdamTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyPWTaskID}).TimeInterval.Duration
    $AdamDCTime = ($Script:AdamTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyDCTaskID}).TimeInterval.Duration
    $AdamIPTime = ($Script:AdamTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyIPTaskID}).TimeInterval.Duration

    $JohnatanPCTime = ($Script:JohnatanTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyPCTaskID}).TimeInterval.Duration
    $JohnatanPWTime = ($Script:JohnatanTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyPWTaskID}).TimeInterval.Duration
    $JohnatanDCTime = ($Script:JohnatanTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyDCTaskID}).TimeInterval.Duration
    $JohnatanIPTime = ($Script:JohnatanTimeEntries | Where-Object {$_.ProjectID -eq $Script:ClockifyProjectID} | Where-Object {$_.TaskID -eq $Script:ClockifyIPTaskID}).TimeInterval.Duration

    [decimal]$Script:AngelTotalPCTime = "0.00"
    [decimal]$Script:AngelTotalPWTime = "0.00"
    [decimal]$Script:AngelTotalDCTime = "0.00"
    [decimal]$Script:AngelTotalIPTime = "0.00"

    ##################

    $AngelPCTime | ForEach-Object {
       
        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime "$_"
        $Script:AngelTotalPCTime = [decimal]$TimeSpent + [decimal]$Script:AngelTotalPCTime

    }

    $AngelPWTime | ForEach-Object {
       
        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime "$_"
        $Script:AngelTotalPWTime = [decimal]$TimeSpent + [decimal]$Script:AngelTotalPWTime

    }

    $AngelDCTime | ForEach-Object {
       
        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime "$_"
        $Script:AngelTotalDCTime = [decimal]$TimeSpent + [decimal]$Script:AngelTotalDCTime

    }

    $AngelIPTime | ForEach-Object {
       
        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime "$_"
        $Script:AngelTotalIPTime = [decimal]$TimeSpent + [decimal]$Script:AngelTotalIPTime

    }

    ########################

    [decimal]$Script:AdamTotalPCTime = "0.00"
    [decimal]$Script:AdamTotalPWTime = "0.00"
    [decimal]$Script:AdamTotalDCTime = "0.00"
    [decimal]$Script:AdamTotalIPTime = "0.00"

    $AdamPCTime | ForEach-Object {
       
        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime "$_"
        $Script:AdamTotalPCTime = [decimal]$TimeSpent + [decimal]$Script:AdamTotalPCTime

    }

    $AdamPWTime | ForEach-Object {
       
        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime "$_"
        $Script:AdamTotalPWTime = [decimal]$TimeSpent + [decimal]$Script:AdamTotalPWTime

    }

    $AdamDCTime | ForEach-Object {
       
        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime "$_"
        $Script:AdamTotalDCTime = [decimal]$TimeSpent + [decimal]$Script:AdamTotalDCTime

    }

    $AdamIPTime | ForEach-Object {
       
        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime "$_"
        $Script:AdamTotalIPTime = [decimal]$TimeSpent + [decimal]$Script:AdamTotalIPTime

    }

    ########################

    [decimal]$Script:JohnatanTotalPCTime = "0.00"
    [decimal]$Script:JohnatanTotalPWTime = "0.00"
    [decimal]$Script:JohnatanTotalDCTime = "0.00"
    [decimal]$Script:JohnatanTotalIPTime = "0.00"

    $JohnatanPCTime | ForEach-Object {
       
        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime "$_"
        $Script:JohnatanTotalPCTime = [decimal]$TimeSpent + [decimal]$Script:JohnatanTotalPCTime

    }

    $JohnatanPWTime | ForEach-Object {
       
        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime "$_"
        $Script:JohnatanTotalPWTime = [decimal]$TimeSpent + [decimal]$Script:JohnatanTotalPWTime

    }

    $JohnatanDCTime | ForEach-Object {

        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime "$_"
        $Script:JohnatanTotalDCTime = [decimal]$TimeSpent + [decimal]$Script:JohnatanTotalDCTime

    }

    $JohnatanIPTime | ForEach-Object {
       
        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime "$_"
        $Script:JohnatanTotalIPTime = [decimal]$TimeSpent + [decimal]$Script:JohnatanTotalIPTime

    }

    ########################

    [decimal]$AngelTotalTime = "0.00"
    $AngelTime = $Script:AngelTimeEntries | Where-Object {$_.ProjectID -eq "$Script:ClockifyProjectID"}
    $AngelTime | ForEach-Object {
       
        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime $_.TimeInterval.Duration
        $Script:AngelTotalTime = [decimal]$TimeSpent + [decimal]$AngelTotalTime

    }

    [decimal]$AdamTotalTime = "0.00"
    $AdamTime = $Script:AdamTimeEntries | Where-Object {$_.ProjectID -eq "$Script:ClockifyProjectID"}
    $AdamTime | ForEach-Object {
       
        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime $_.TimeInterval.Duration
        $Script:AdamTotalTime = [decimal]$TimeSpent + [decimal]$AdamTotalTime

    }

    [decimal]$JohnatanTotalTime = "0.00"
    $JohnatanTime = $Script:JohnatanTimeEntries | Where-Object {$_.ProjectID -eq "$Script:ClockifyProjectID"}
    $JohnatanTime | ForEach-Object {

        [decimal]$TimeSpent = ConvertClockifyTimeToAktisTime $_.TimeInterval.Duration
        $Script:JohnatanTotalTime = [decimal]$TimeSpent + [decimal]$JohnatanTotalTime

    }

    ####################
  
}

function GetProjectTaskLevelDetails {

    try{
        $Script:ClockifyProjectLevelDetails = Invoke-RestMethod -Uri ($Script:APIEndpoint + "/workspaces/" + $Script:ClockifyWorkSpaceID + "/projects/" + $Script:ClockifyProjectID) -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $Script:ClockifyAPIKey} -Method Get
    }
    catch{
        
        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to retrieve Clockify project level details.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Clockify project level details.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Retrieve Clockify Project Level Details" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nFailed to retrieve Clockify project level details.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort

    }

    try{
        $ClockifyProjectTaskLevelDetails = Invoke-RestMethod -Uri ($Script:APIEndpoint + "/workspaces/" + $Script:ClockifyWorkSpaceID + "/projects/" + $Script:ClockifyProjectID + "/tasks") -Headers @{'content-type' = 'application/json'; 'X-Api-Key' = $Script:ClockifyAPIKey} -Method Get
    }
    catch{
        
        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to retrieve Clockify task level details.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Clockify task level details.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Retrieve Clockify Task Level Details" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nFailed to retrieve Clockify task level details.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort

    }
    
    $Script:ClockifyPCTaskID = ($ClockifyProjectTaskLevelDetails | Where-Object {$_.Name -eq "Proposal Coordination"}).ID
    $Script:ClockifyPWTaskID = ($ClockifyProjectTaskLevelDetails | Where-Object {$_.Name -eq "Proposal Writing"}).ID
    $Script:ClockifyDCTaskID = ($ClockifyProjectTaskLevelDetails | Where-Object {$_.Name -eq "Design Consulting"}).ID
    $Script:ClockifyIPTaskID = ($ClockifyProjectTaskLevelDetails | Where-Object {$_.Name -eq "Internal Project"}).ID

    $RawProposalWritingTime = ($ClockifyProjectTaskLevelDetails | Where-Object {$_.Name -eq "Proposal Writing"}).Duration
    $RawProposalCoordinationTime = ($ClockifyProjectTaskLevelDetails | Where-Object {$_.Name -eq "Proposal Coordination"}).Duration
    $RawDesignConsultingTime = ($ClockifyProjectTaskLevelDetails | Where-Object {$_.Name -eq "Design Consulting"}).Duration
    $RawInternalProjectTime = ($ClockifyProjectTaskLevelDetails | Where-Object {$_.Name -eq "Internal Project"}).Duration
    
    $Script:ProposalWritingTime = ConvertClockifyTimeToAktisTime "$RawProposalWritingTime"
    $Script:ProposalCoordinationTime = ConvertClockifyTimeToAktisTime "$RawProposalCoordinationTime"
    $Script:DesignConsultingTime = ConvertClockifyTimeToAktisTime "$RawDesignConsultingTime"
    $Script:InternalProjectTime = ConvertClockifyTimeToAktisTime "$RawInternalProjectTime"

}

function ConvertClockifyTimeToAktisTime ($RawTimeEntryDuration) {

    try{
        if(!($RawTimeEntryDuration | Select-String "H")){
            $ShortTimeEntryDuration = ($RawTimeEntryDuration).Replace("PT", "").Replace("M", ":").Replace("S", "")
            $UnfinishedTimeEntryDuration = "00:" + $ShortTimeEntryDuration
            if(!($RawTimeEntryDuration | Select-String "M")){
              $UnfinishedTimeEntryDuration = "00:" + $UnfinishedTimeEntryDuration
            }
            if(!($RawTimeEntryDuration | Select-String "S")){
              $UnfinishedTimeEntryDuration = "00:" + $ShortTimeEntryDuration + "00"
            }
          }
          else{
            $UnfinishedTimeEntryDuration = ($RawTimeEntryDuration).Replace("PT", "").Replace("H", ":").Replace("M", ":").Replace("S", "")
            if(!($RawTimeEntryDuration | Select-String "M")){
              $UnfinishedTimeEntryDuration = ($UnfinishedTimeEntryDuration -Replace ":.*$") + ":00:00"
            }
            else{
              if(!($RawTimeEntryDuration | Select-String "S")){
                $UnfinishedTimeEntryDuration = (($RawTimeEntryDuration).Replace("PT", "").Replace("H", ":").Replace("M", ":").Replace("S", "")) + "00"
              }
            }
          }
      
          $TimeSpan = [System.TimeSpan]::Parse($UnfinishedTimeEntryDuration)
          $Minutes = [math]::Round(($TimeSpan.TotalMinutes),0)
          $Result = [math]::Round(($Minutes.ToString() / 60),2)
          Write-Output $Result
    }
    catch{
        ## Add error handling when necessary
    }

}

function PostToAktis {

    #Determine DDE Lead
    $GreatestTime = 0
    [System.Collections.ArrayList]$WritingTimeInstances = @()

    $TemporaryObject = New-Object PSObject
    Add-Member -InputObject $TemporaryObject -NotePropertyName DDELead -NotePropertyValue $("JB")
    Add-Member -InputObject $TemporaryObject -NotePropertyName TotalWritingTime -NotePropertyValue $([decimal]$Script:JohnatanTotalPWTime)
    $WritingTimeInstances.Add($TemporaryObject) | Out-Null

    $TemporaryObject = New-Object PSObject
    Add-Member -InputObject $TemporaryObject -NotePropertyName DDELead -NotePropertyValue $("AP")
    Add-Member -InputObject $TemporaryObject -NotePropertyName TotalWritingTime -NotePropertyValue $([decimal]$Script:AngelTotalPWTime)
    $WritingTimeInstances.Add($TemporaryObject) | Out-Null

    $TemporaryObject = New-Object PSObject
    Add-Member -InputObject $TemporaryObject -NotePropertyName DDELead -NotePropertyValue $("AG")
    Add-Member -InputObject $TemporaryObject -NotePropertyName TotalWritingTime -NotePropertyValue $([decimal]$Script:AdamTotalPWTime)
    $WritingTimeInstances.Add($TemporaryObject) | Out-Null

    $TemporaryObject = New-Object PSObject
    Add-Member -InputObject $TemporaryObject -NotePropertyName DDELead -NotePropertyValue $("HB")
    Add-Member -InputObject $TemporaryObject -NotePropertyName TotalWritingTime -NotePropertyValue $([decimal]$Script:HowardTotalPWTime)
    $WritingTimeInstances.Add($TemporaryObject) | Out-Null
    
    $WritingTimeInstances | ForEach-Object{

        if ($_.TotalWritingTime -gt $GreatestTime){
            $GreatestTime = $_.TotalWritingTime
            $Script:DDELead = $_.DDELead
        }

    }

    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): DDE Lead: $Script:DDELead"
    $Script:IsGratisRevision = $False
#if(!($Script:IsGratisRevision -eq $True)){ mod
#
#    if(($Script:ProposalVersion -eq "1") -or ($Script:ProposalVersion -eq "3") -or ($Script:ProposalVersion -eq "5") -or ($Script:ProposalVersion -eq "7") -or ($Script:ProposalVersion -eq "9") -or ($Script:ProposalVersion -eq "11") -or ($Script:ProposalVersion -eq "13")){
#
#        $Script:IsGratisRevision = $False
#    
#    }
#    elseif(($Script:ProposalVersion -eq "2") -or ($Script:ProposalVersion -eq "4") -or ($Script:ProposalVersion -eq "6") -or ($Script:ProposalVersion -eq "8") -or ($Script:ProposalVersion -eq "10") -or ($Script:ProposalVersion -eq "12") -or ($Script:ProposalVersion -eq "14")){
#
#        if($Script:ProposalWritingTime -ge 0.5){
#            $Script:IsGratisRevision = $False
#        }
#        elseif($Script:ProposalWritingTime -lt 0.5){
#            $Script:IsGratisRevision = $True
#        }
#      
#    }
#        # "Cancel" {}
#}

    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Gratis Revision: $Script:IsGratisRevision"

    # Determine <100, >100, >200, etc.
    if($Script:ProposedHours -le "100"){
        $Script:ProposalBiOpLT100 = $True
        $Script:ProposalBiOpGT100 = $False
        $Script:ProposalBiOpGT200 = $False
        $Script:ProposalBiOpGT300 = $False
    }
    if($Script:ProposedHours -ge "100"){
        $Script:ProposalBiOpLT100 = $False
        $Script:ProposalBiOpGT100 = $True
        $Script:ProposalBiOpGT200 = $False
        $Script:ProposalBiOpGT300 = $False
    }
    if($Script:ProposedHours -ge "200"){
        $Script:ProposalBiOpLT100 = $False
        $Script:ProposalBiOpGT100 = $False
        $Script:ProposalBiOpGT200 = $True
        $Script:ProposalBiOpGT300 = $False
    }
    if($Script:ProposedHours -ge "300"){
        $Script:ProposalBiOpLT100 = $False
        $Script:ProposalBiOpGT100 = $False
        $Script:ProposalBiOpGT200 = $False
        $Script:ProposalBiOpGT300 = $True
    }

    try{

        if($Script:AHRoute -eq "DDE"){

            $TodaysDate = Get-Date

            #Post to Proposal Statistics Sharepoint List
            AuthenticateToSPO
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            #Set-PnPListItem -List "Proposal Queues" -Identity $Script:ProposalID -Values @{"DateProposalWritten"="$TodaysDate" } -Connection $Script:ConnectToSPO
            Set-PnPListItem -List "Proposal Statistics" -Identity $Script:ProposalID -Values @{"DDELead"="$Script:DDELead"; "DateProposalWritten"="$TodaysDate" } -Connection $Script:ConnectToSPO


            try{
                $ProposalQueuesRawData = (Get-PnPListItem -List "Proposal Queues" -Connection $Script:ConnectToSPO).FieldValues
            
                $ProposalQueuesData = @()
            
                $ProposalQueuesRawData | ForEach-Object {
                    $ProposalQueuesData += $_
                }
            }
            catch{
        
                Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to retrieve Aktis analytics data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Aktis analytics data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
                Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Retrieve Proposal Statistics Data" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nFailed to retrieve Aktis analytics data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
                Return
            }

            try{

                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Downloading Excel workplan for $Script:LongProposalName to extract Proposal Statistics data from."
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Downloading Excel workplan for $Script:LongProposalName to extract Proposal Statistics data from."

                if(!(Test-Path -Path "C:\Temp\Aktis\Aktis Helper\Staging")){
                    New-Item -Path "C:\Temp\Aktis\Aktis Helper\Staging" -ItemType "Directory" | Out-Null
                }
    
                $ProposalQueuesID = ($ProposalQueuesData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).ID

                $ProposalMSP = ((Get-PnPListItem -List "Proposal Queues" -Fields "Title", "KTNumber" -Connection $Script:ConnectToSPO).FieldValues | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).Title
                try{

                        $ActureSolutionsWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\Acture Solutions\$Script:ProposalFileName`.xlsx"
                        $AffinityWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\Affinity Technology Partners\$Script:ProposalFileName`.xlsx"
                        $AMCWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\AMC Solutions\$Script:ProposalFileName`.xlsx"
                        $AntisynWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\Antisyn\$Script:ProposalFileName`.xlsx"
                        $ApticaWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\Aptica LLC\$Script:ProposalFileName`.xlsx"
                        $DenaliTEKWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\DenaliTEK\$Script:ProposalFileName`.xlsx"
                        $DominionTechWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\DominionTech\$Script:ProposalFileName`.xlsx"
                        $GroffWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\Groff\$Script:ProposalFileName`.xlsx"
                        $ITSWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\ITS\$Script:ProposalFileName`.xlsx"
                        $Layer9WorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\Layer9\$Script:ProposalFileName`.xlsx"
                        $ManawaWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\Manawa\$Script:ProposalFileName`.xlsx"
                        $MentisGroupWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\Mentis Group\$Script:ProposalFileName`.xlsx"
                        $MyITCrewNYWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\My IT Crew - NY\$Script:ProposalFileName`.xlsx"
                        $SisAdminWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\SisAdmin\$Script:ProposalFileName`.xlsx"
                        $STAIWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\STAI\$Script:ProposalFileName`.xlsx"
                        $SteadyNetworksWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\Steady Networks\$Script:ProposalFileName`.xlsx"
                        $SyscomWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\Syscom\$Script:ProposalFileName`.xlsx"
                        $UniVistaWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\UniVista\$Script:ProposalFileName`.xlsx"
                        $ValleyExpetecWorkplanTemplate = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - OPS - Documents\WIP\In Progress\Valley Expetec\$Script:ProposalFileName`.xlsx"

                    }
                catch{
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to locate customer workplan template.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
                        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to locate customer workplan template.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
                    }

                try{

                        $ActureSolutionsClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\Acture Solutions\Design Desk Fuel\Proposals\"
                        $AffinityClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\Affinity Technology Partners\Design Desk Fuel\Proposals\"
                        $AMCClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\AMC Solutions\Design Desk Fuel\Proposals\"
                        $AntisynClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\Antisyn\Design Desk Fuel\Proposals\"
                        $ApticaClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\Aptica\Design Desk Fuel\Proposals\"
                        $DenaliTEKClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\DenaliTEK\Design Desk Fuel\Proposals\"
                        $DominionTechClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\DominionTech\Project Fuel\Engagement Documents\Proposals\Proposals\"
                        $GroffClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\Groff Networks\Design Desk Fuel\Proposals\"
                        $ITSClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\ITS\Project Fuel\Engagement Documents\Proposals\"
                        $Layer9ClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\Layer9\Design Desk Fuel\Proposals\"
                        $ManawaClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\Manawa\Design Desk Fuel\Proposals\"
                        $MentisGroupClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\Mentis Group\Design Desk Fuel\Proposals\"
                        $MyITCrewNYClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\My IT Crew - NY\Design Desk Fuel\Proposals\"
                        $SisAdminClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\SisAdmin\Design Desk Fuel\Proposals\"
                        $STAIClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\STAI\Engagement Documents\Proposals\"
                        $SteadyNetworksClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\Steady Networks\Design Desk Fuel\Proposals\"
                        $SyscomClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\Syscom Business Technologies\Design Desk Fuel\Proposals\"
                        $UniVistaClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\UniVista\Design Desk Fuel\Proposals\"
                        $ValleyExpetecClientFolder = "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel - Clients - Documents\Clients\Valley Expetec\Design Desk Fuel\Proposals\"

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
                Copy-Item -Path $WorkplanTemplateLocation -Destination "C:\Temp\Aktis\Aktis Helper\Staging\$Script:ProposalFileName`.xlsx" -Force | Out-Null

            }
            catch{

                Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to download Excel workplan for $Script:LongProposalName to extract Proposal Statistics data from.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to download Excel workplan for $Script:LongProposalName to extract Proposal Statistics data from.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
                Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Download Excel Workplan from SharePoint List" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nFailed to download Excel workplan for $Script:LongProposalName to extract Proposal Statistics data from.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
                Return
            
            }

            

            ###Retrieve financial data from Excel workplan
            Start-Sleep -Seconds 2
            ExecuteAktisHelper "C:\Temp\Aktis\Aktis Helper\Staging\$Script:ProposalFileName`.xlsx"

            ##Verify KTNumber found in proposal matches the proposal that a status change was detected on
            $Script:AktisKTNumber = $Script:CurrentKTNumber #mod
            if(!($Script:AktisKTNumber -eq $Script:CurrentKTNumber)){

                Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): The Aktis KT Number found in the workplan ($Script:AktisKTNumber) does not match the KT Number of the proposal being processed by Aktis Helper:`r`n$Script:LongProposalName`r`n$Script:CurrentKTNumber"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "The Aktis KT Number found in the workplan ($Script:AktisKTNumber) does not match the KT Number of the proposal being processed by Aktis Helper:`r`n$Script:LongProposalName`r`n$Script:CurrentKTNumber"
                Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: KT Number Mismatch in Excel Workplan" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nThe Aktis KT Number found in the workplan ($Script:AktisKTNumber) does not match the KT Number of the proposal being processed by Aktis Helper:`r`n$Script:LongProposalName`r`n$Script:CurrentKTNumber" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
                Return

            }
            else{

                #Post to Proposal Statistics Sharepoint List
                AuthenticateToSPO
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Set-PnPListItem -List "Proposal Statistics" -Identity $Script:ProposalID -Values @{"GratisRevision"="$Script:IsGratisRevision"; "ProposalExceedsSubscription_x002"="$Script:ProposalExceedsSubEnt"; "ProposalExceedsSubscription_x0020"="$Script:ProposalExceedsSubBasic"; "ProposedHours"="$Script:ProposedHours"; "HourlyRate"="$Script:HourlyRate"; "ProposedMRR"="$Script:TotalMRR"; "ProposedProfessionalServicesReve"="$Script:PSRevenue"; "DDELead"="$Script:DDELead"; "ProposedProductsRevenue"="$Script:ProductsRevenue"; "Proposed_x0020_Products_x0020_an"="$Script:TotalNRR"; "_x0031_00"="$Script:ProposalBiOpLT100"; "_x0031_000"="$Script:ProposalBiOpGT100"; "_x0032_00"="$Script:ProposalBiOpGT200"; "_x0033_00"="$Script:ProposalBiOpGT300"; "ProposalWritingTime"="$Script:ProposalWritingTime"; "ProposalCoordinationTime"="$Script:ProposalCoordinationTime"; "DesignConsultingTimeNew"="$Script:DesignConsultingTime"; "InternalProjectTime"="$Script:InternalProjectTime"} -Connection $Script:ConnectToSPO
                #mod Set-PnPListItem -List "Proposal Statistics" -Identity $Script:ProposalID -Values @{"ProposalVerison"="$Script:ProposalVersion"; "GratisRevision"="$Script:IsGratisRevision"; "ProposalExceedsSubscription_x002"="$Script:ProposalExceedsSubEnt"; "ProposalExceedsSubscription_x0020"="$Script:ProposalExceedsSubBasic"; "ProposedHours"="$Script:ProposedHours"; "HourlyRate"="$Script:HourlyRate"; "ProposedMRR"="$Script:TotalMRR"; "ProposedProfessionalServicesReve"="$Script:PSRevenue"; "DDELead"="$Script:DDELead"; "ProposedProductsRevenue"="$Script:ProductsRevenue"; "Proposed_x0020_Products_x0020_an"="$Script:TotalNRR"; "_x0031_00"="$Script:ProposalBiOpLT100"; "_x0031_000"="$Script:ProposalBiOpGT100"; "_x0032_00"="$Script:ProposalBiOpGT200"; "_x0033_00"="$Script:ProposalBiOpGT300"; "ProposalWritingTime"="$Script:ProposalWritingTime"; "ProposalCoordinationTime"="$Script:ProposalCoordinationTime"; "DesignConsultingTimeNew"="$Script:DesignConsultingTime"; "InternalProjectTime"="$Script:InternalProjectTime"} -Connection $Script:ConnectToSPO

            }



        }
    }
    catch{

        $DDEPostingError = $True
        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Aktis Helper Route: $Script:AHRoute`r`nFailed to post Proposal Statistics data for $Script:LongProposalName to Aktis.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Aktis Helper Route: $Script:AHRoute`r`nFailed to post Proposal Statistics data for $Script:LongProposalName to Aktis.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Post Proposal Statistics Data" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nAktis Helper Route: $Script:AHRoute`r`nFailed to post Proposal Statistics data for $Script:LongProposalName to Aktis.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort

    }
    finally{

        if($Script:AHRoute -eq "DDE"){

            if(!($DDEPostingError -eq $True)){

                UpdateActivityLog
                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Successfully posted Aktis Proposal Statistics and Client Analytics for $Script:LongProposalName!`r`nAktis Helper Route: $Script:AHRoute"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Aktis Helper Route: $Script:AHRoute`r`nSuccessfully posted Aktis Proposal Statistics and Client Analytics for $Script:LongProposalName!`r`n`r`nProposal Number: $Script:TicketNumber`r`nProposal Version: $Script:ProposalVersion`r`nProposed Hours: $Script:ProposedHours`r`nHourly Rate: $Script:HourlyRate`r`nProposed MRR: $Script:TotalMRR`r`nProposed Professional Services Revenue: $Script:PSRevenue`r`nProposed Products Revenue: $Script:ProductsRevenue`r`nTotal NRR: $Script:TotalNRR`r`nProposal Writing Time: $Script:ProposalWritingTime`r`nProposal Coordination Time: $Script:ProposalCoordinationTime`r`nDesign Consulting Time: $Script:DesignConsultingTime`r`nInternal Project Time: $Script:InternalProjectTime"
                
                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Posting data to Analytics Weekly Snapshot..."
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Posting data to Analytics Weekly Snapshot..."
    
            }

        }

    }

    try{

        if($Script:AHRoute -eq "PC"){

            try{
                $ProposalQueuesRawData = (Get-PnPListItem -List "Proposal Queues" -Connection $Script:ConnectToSPO).FieldValues
            
                $ProposalQueuesData = @()
            
                $ProposalQueuesRawData | ForEach-Object {
                    $ProposalQueuesData += $_
                }
            }
            catch{
        
                Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to retrieve Aktis analytics data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Aktis analytics data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
                Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Retrieve Proposal Statistics Data" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nFailed to retrieve Aktis analytics data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
                Return
            }

            try{

                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Downloading Excel workplan for $Script:LongProposalName to extract Proposal Statistics data from."
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Downloading Excel workplan for $Script:LongProposalName to extract Proposal Statistics data from."

                if(!(Test-Path -Path "C:\Temp\Aktis\Aktis Helper\Staging")){
                    New-Item -Path "C:\Temp\Aktis\Aktis Helper\Staging" -ItemType "Directory" | Out-Null
                }
    
                $ProposalQueuesID = ($ProposalQueuesData | Where-Object {$_.KTNumber -eq $Script:CurrentKTNumber}).ID

               Get-PnPFile -Url "/sites/Fuelnow/Lists/Proposal Queues/Attachments/$ProposalQueuesID/$Script:ProposalFileName.xlsx" -Path "C:\Temp\Aktis\Aktis Helper\Staging" -FileName "$Script:ProposalFileName.xlsx" -AsFile -Force -Connection $Script:ConnectToSPO

            }
            catch{

                Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to download Excel workplan for $Script:LongProposalName to extract Proposal Statistics data from.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to download Excel workplan for $Script:LongProposalName to extract Proposal Statistics data from.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
                Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Download Excel Workplan from SharePoint List" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nFailed to download Excel workplan for $Script:LongProposalName to extract Proposal Statistics data from.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
                Return
            
            }

            ###Retrieve financial data from Excel workplan
            Start-Sleep -Seconds 2
            ExecuteAktisHelper "C:\Temp\Aktis\Aktis Helper\Staging\$Script:ProposalFileName`.xlsx"

            ##Verify KTNumber found in proposal matches the proposal that a status change was detected on
            $Script:AktisKTNumber = $Script:CurrentKTNumber #mod
            if(!($Script:AktisKTNumber -eq $Script:CurrentKTNumber)){

                Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): The Aktis KT Number found in the workplan ($Script:AktisKTNumber) does not match the KT Number of the proposal being processed by Aktis Helper:`r`n$Script:LongProposalName`r`n$Script:CurrentKTNumber"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "The Aktis KT Number found in the workplan ($Script:AktisKTNumber) does not match the KT Number of the proposal being processed by Aktis Helper:`r`n$Script:LongProposalName`r`n$Script:CurrentKTNumber"
                Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: KT Number Mismatch in Excel Workplan" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nThe Aktis KT Number found in the workplan ($Script:AktisKTNumber) does not match the KT Number of the proposal being processed by Aktis Helper:`r`n$Script:LongProposalName`r`n$Script:CurrentKTNumber" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
                Return

            }
            else{

                #Post to Proposal Statistics Sharepoint List
                AuthenticateToSPO
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Set-PnPListItem -List "Proposal Statistics" -Identity $Script:ProposalID -Values @{"GratisRevision"="$Script:IsGratisRevision"; "ProposalExceedsSubscription_x002"="$Script:ProposalExceedsSubEnt"; "ProposalExceedsSubscription_x0020"="$Script:ProposalExceedsSubBasic"; "ProposedHours"="$Script:ProposedHours"; "HourlyRate"="$Script:HourlyRate"; "ProposedMRR"="$Script:TotalMRR"; "ProposedProfessionalServicesReve"="$Script:PSRevenue"; "ProposedProductsRevenue"="$Script:ProductsRevenue"; "Proposed_x0020_Products_x0020_an"="$Script:TotalNRR"; "_x0031_00"="$Script:ProposalBiOpLT100"; "_x0031_000"="$Script:ProposalBiOpGT100"; "_x0032_00"="$Script:ProposalBiOpGT200"; "_x0033_00"="$Script:ProposalBiOpGT300"; "ProposalWritingTime"="$Script:ProposalWritingTime"; "ProposalCoordinationTime"="$Script:ProposalCoordinationTime"; "DesignConsultingTimeNew"="$Script:DesignConsultingTime"; "InternalProjectTime"="$Script:InternalProjectTime"} -Connection $Script:ConnectToSPO
                #mod Set-PnPListItem -List "Proposal Statistics" -Identity $Script:ProposalID -Values @{"ProposalVerison"="$Script:ProposalVersion"; "GratisRevision"="$Script:IsGratisRevision"; "ProposalExceedsSubscription_x002"="$Script:ProposalExceedsSubEnt"; "ProposalExceedsSubscription_x0020"="$Script:ProposalExceedsSubBasic"; "ProposedHours"="$Script:ProposedHours"; "HourlyRate"="$Script:HourlyRate"; "ProposedMRR"="$Script:TotalMRR"; "ProposedProfessionalServicesReve"="$Script:PSRevenue"; "DDELead"="$Script:DDELead"; "ProposedProductsRevenue"="$Script:ProductsRevenue"; "Proposed_x0020_Products_x0020_an"="$Script:TotalNRR"; "_x0031_00"="$Script:ProposalBiOpLT100"; "_x0031_000"="$Script:ProposalBiOpGT100"; "_x0032_00"="$Script:ProposalBiOpGT200"; "_x0033_00"="$Script:ProposalBiOpGT300"; "ProposalWritingTime"="$Script:ProposalWritingTime"; "ProposalCoordinationTime"="$Script:ProposalCoordinationTime"; "DesignConsultingTimeNew"="$Script:DesignConsultingTime"; "InternalProjectTime"="$Script:InternalProjectTime"} -Connection $Script:ConnectToSPO

            }

        }

    }
    catch{

        $PCPostingError = $True
        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Aktis Helper Route: $Script:AHRoute`r`nFailed to post Proposal Statistics data for $Script:LongProposalName to Aktis.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Aktis Helper Route: $Script:AHRoute`r`nFailed to post Proposal Statistics data for $Script:LongProposalName to Aktis.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Post Proposal Statistics Data" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nAktis Helper Route: $Script:AHRoute`r`nFailed to post Proposal Statistics data for $Script:LongProposalName to Aktis.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort

    }
    finally{

        if($Script:AHRoute -eq "PC"){

            if(!($PCPostingError -eq $True)){

                UpdateActivityLog
                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Successfully posted Aktis Proposal Statistics and Client Analytics for $Script:LongProposalName!`r`nAktis Helper Route: $Script:AHRoute"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Successfully posted Aktis Proposal Statistics and Client Analytics for $Script:LongProposalName!`r`nAktis Helper Route: $Script:AHRoute`r`n`r`nProposal Number: $Script:TicketNumber`r`nProposal Version: $Script:ProposalVersion`r`nProposed Hours: $Script:ProposedHours`r`nHourly Rate: $Script:HourlyRate`r`nProposed MRR: $Script:TotalMRR`r`nProposed Professional Services Revenue: $Script:PSRevenue`r`nProposed Products Revenue: $Script:ProductsRevenue`r`nTotal NRR: $Script:TotalNRR`r`nProposal Writing Time: $Script:ProposalWritingTime`r`nProposal Coordination Time: $Script:ProposalCoordinationTime`r`nDesign Consulting Time: $Script:DesignConsultingTime`r`nInternal Project Time: $Script:InternalProjectTime"
                
                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Posting data to Analytics Weekly Snapshot..."
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Posting data to Analytics Weekly Snapshot..."

            }

        }

    }

}

function PostToLocalAnalyticsWeeklySnapshot {

    function UpdateAWSDDENewProposal {
      
        if($Script:DDELead -eq "AP"){
            $ExistingAPNewDrafts = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$APNewDrafts).Value2
            $UpdatedAPNewDraftsValue = $ExistingAPNewDrafts + 1
            $DDEExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$APNewDrafts).Comment.Shape.AlternativeText) -Replace "^.*?: "
            if($DDEExistingNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber"){

                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 1The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                
            }
            else{

                ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$APNewDrafts)).Clear()
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$APNewDrafts) = "$UpdatedAPNewDraftsValue"
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$APNewDrafts).HorizontalAlignment = -4108
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$APNewDrafts).AddComment("$DDEExistingNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$APNewDrafts).Comment.Visible = $False

            }

            $ExistingTotals = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).Value2
            $UpdatedTotalsValue = $ExistingTotals + 1
            $TotalsExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).Comment.Shape.AlternativeText) -Replace "^.*?: "
            if($TotalsExistingNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber"){

                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 2The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                
            }
            else{

                ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd)).Clear()
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd) = "$UpdatedTotalsValue"
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).HorizontalAlignment = -4108
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).AddComment("$TotalsExistingNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).Comment.Visible = $False

            }

            $ExcelSpreadsheetAWS.Save
            $ExcelSpreadsheetAWS.Close($True)
            $ExcelOpenAWS.Quit()
        }
        if($Script:DDELead -eq "AG"){

            $ExistingAGNewDrafts = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$AGNewDrafts).Value2
            $UpdatedAGNewDraftsValue = $ExistingAGNewDrafts + 1
            $DDEExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$AGNewDrafts).Comment.Shape.AlternativeText) -Replace "^.*?: "
            if($DDEExistingNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber"){

                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 3The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                
            }
            else{

                ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$AGNewDrafts)).Clear()
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$AGNewDrafts) = "$UpdatedAGNewDraftsValue"
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$AGNewDrafts).HorizontalAlignment = -4108
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$AGNewDrafts).AddComment("$DDEExistingNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$AGNewDrafts).Comment.Visible = $False

            }

            $ExistingTotals = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).Value2
            $UpdatedTotalsValue = $ExistingTotals + 1
            $TotalsExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).Comment.Shape.AlternativeText) -Replace "^.*?: "
            if($TotalsExistingNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber"){

                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 4The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                
            }
            else{

                ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd)).Clear()
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd) = "$UpdatedTotalsValue"
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).HorizontalAlignment = -4108
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).AddComment("$TotalsExistingNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).Comment.Visible = $False

            }

            $ExcelSpreadsheetAWS.Save
            $ExcelSpreadsheetAWS.Close($True)
            $ExcelOpenAWS.Quit()
        }
        if($Script:DDELead -eq "JB"){

            $ExistingVANewDrafts = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$VANewDrafts).Value2
            $UpdatedVANewDraftsValue = $ExistingVANewDrafts + 1
            $DDEExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$VANewDrafts).Comment.Shape.AlternativeText) -Replace "^.*?: "
            if($DDEExistingNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber"){

                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 5The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                
            }
            else{

                ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$VANewDrafts)).Clear()
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$VANewDrafts) = "$UpdatedVANewDraftsValue"
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$VANewDrafts).HorizontalAlignment = -4108
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$VANewDrafts).AddComment("$DDEExistingNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$VANewDrafts).Comment.Visible = $False

            }

            $ExistingTotals = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).Value2
            $UpdatedTotalsValue = $ExistingTotals + 1
            $TotalsExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).Comment.Shape.AlternativeText) -Replace "^.*?: "
            if($TotalsExistingNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber"){

                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 6The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                
            }
            else{

                ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd)).Clear()
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd) = "$UpdatedTotalsValue"
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).HorizontalAlignment = -4108
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).AddComment("$TotalsExistingNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewDraftsCmpltd).Comment.Visible = $False

            }

            $ExcelSpreadsheetAWS.Save
            $ExcelSpreadsheetAWS.Close($True)
            $ExcelOpenAWS.Quit()
        }

    }

    function UpdateAWSDDERevision {

        if($Script:IsGratisRevision -eq $False){

            UpdateAWSDDENewProposal

            $ExcelSpreadsheetAWS.Save
            $ExcelSpreadsheetAWS.Close($True)
            $ExcelOpenAWS.Quit()

        }
        elseif($Script:IsGratisRevision -eq $True){

            $ExistingGratisDraftsValue = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewGratisDraftsCmpltd).Value2
            $UpdatedGratisDraftsValue = $ExistingGratisDraftsValue + 1
            $ExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewGratisDraftsCmpltd).Comment.Shape.AlternativeText) -Replace "^.*?: "
            if($ExistingNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber"){

                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 7The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                
            }
            else{

                ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewGratisDraftsCmpltd)).Clear()
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewGratisDraftsCmpltd) = "$UpdatedGratisDraftsValue"
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewGratisDraftsCmpltd).HorizontalAlignment = -4108
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewGratisDraftsCmpltd).AddComment("$ExistingNote`n$Script:TicketNumber V$Script:ProposalVersionNumber - $Script:DDELead")
                $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewGratisDraftsCmpltd).Comment.Visible = $False

            }

            $ExcelSpreadsheetAWS.Save
            $ExcelSpreadsheetAWS.Close($True)
            $ExcelOpenAWS.Quit()

        }

    }
    
    function UpdateAWSPropCo {

        if($Script:ProposalVersionNumber -eq '1'){

            if($Script:ProposedHours -lt "8"){
                #Proposal is a basic proposal
                $ExistingBasicCountColumn = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn).Value2
                $UpdatedBasicCountValue = $ExistingBasicCountColumn + 1
                $ExistingTtlNewPrpslsDelvrd = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).Value2
                $UpdatedTtlNewPrpslsDelvrd = $ExistingTtlNewPrpslsDelvrd + 1
                $ExistingBasicCountNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn).Comment.Shape.AlternativeText) -Replace "^.*?: "
                $ExistingTtlNewPrpslsDelvrdNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).Comment.Shape.AlternativeText) -Replace "^.*?: "
                if(($ExistingBasicCountNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber") -or ($ExistingTtlNewPrpslsDelvrdNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber")){
                    
                    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 8The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                    Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                    
                }
                else{
                    ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn)).Clear()
                    ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd)).Clear()
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn) = "$UpdatedBasicCountValue"
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn).HorizontalAlignment = -4108
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn).AddComment("$ExistingBasicCountNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd) = "$UpdatedTtlNewPrpslsDelvrd"
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).HorizontalAlignment = -4108
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).AddComment("$ExistingTtlNewPrpslsDelvrdNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                }
                $ExcelSpreadsheetAWS.Save
                $ExcelSpreadsheetAWS.Close($True)
                $ExcelOpenAWS.Quit()
    
            }
            elseif($Script:ProposedHours -ge "8"){
                #Proposal is an enterprise proposal
                $ExistingEnterpriseCountColumn = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn).Value2
                $UpdatedEnterpriseCountValue = $ExistingEnterpriseCountColumn + 1
                $ExistingTtlNewPrpslsDelvrd = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).Value2
                $UpdatedTtlNewPrpslsDelvrd = $ExistingTtlNewPrpslsDelvrd + 1
                $ExistingEnterpriseCountNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn).Comment.Shape.AlternativeText) -Replace "^.*?: "
                $ExistingTtlNewPrpslsDelvrdNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).Comment.Shape.AlternativeText) -Replace "^.*?: "
                if(($ExistingEnterpriseCountNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber") -or ($ExistingTtlNewPrpslsDelvrdNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber")){
                    
                    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 9The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                    Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                    
                }
                else{
                    ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn)).Clear()
                    ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd)).Clear()
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn) = "$UpdatedEnterpriseCountValue"
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn).HorizontalAlignment = -4108
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn).AddComment("$ExistingEnterpriseCountNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd) = "$UpdatedTtlNewPrpslsDelvrd"
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).HorizontalAlignment = -4108
                    $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).AddComment("$ExistingTtlNewPrpslsDelvrdNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                }
                $ExcelSpreadsheetAWS.Save
                $ExcelSpreadsheetAWS.Close($True)
                $ExcelOpenAWS.Quit()
            }
        }
        elseif($Script:ProposalVersionNumber -gt '1'){

            if($Script:IsGratisRevision -eq $False){

                if($Script:ProposedHours -lt "8"){

                    $ExistingBillableRevsCmpltdDelvrd = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd).Value2
                    $UpdatedTtlBillableRevsCmpltdDelvrdCount = $ExistingBillableRevsCmpltdDelvrd + 1
                    $ExistingTtlBillableRevsCmpltdDelvrdNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    if($ExistingTtlBillableRevsCmpltdDelvrdNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber"){
                        
                        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 10The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        
                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd) = "$UpdatedTtlBillableRevsCmpltdDelvrdCount"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd).AddComment("$ExistingTtlBillableRevsCmpltdDelvrdNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                    }

                    $ExistingBillableRevisionCount = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).Value2
                    $UpdatedBillableRevisionCountValue = $ExistingBillableRevisionCount + 1
                    $BillableRevsExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    
                    if($BillableRevsExistingNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber"){
                        
                        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 11The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        
                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn) = "$UpdatedBillableRevisionCountValue"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).AddComment("$BillableRevsExistingNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).Comment.Visible = $False

                    }
    
                    $ExistingTtlNewPrpslsDelvrd = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).Value2
                    $UpdatedTtlNewPrpslsDelvrd = $ExistingTtlNewPrpslsDelvrd + 1
                    $ExistingTtlNewPrpslsDelvrdNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    if(($ExistingTtlNewPrpslsDelvrdNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber")){
                        
                        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 12The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        
                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd) = "$UpdatedTtlNewPrpslsDelvrd"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).AddComment("$ExistingTtlNewPrpslsDelvrdNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                    }

                    $ExistingBasicCountColumn = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn).Value2
                    $UpdatedBasicCountValue = $ExistingBasicCountColumn + 1
                    $ExistingBasicCountNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    if($ExistingBasicCountNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber"){
                        
                        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 13The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "1The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"

                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn) = "$UpdatedBasicCountValue"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BasicCountColumn).AddComment("$ExistingBasicCountNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                    }
    
                    $ExcelSpreadsheetAWS.Save
                    $ExcelSpreadsheetAWS.Close($True)
                    $ExcelOpenAWS.Quit()

                }
                if($Script:ProposedHours -ge "8"){

                    $ExistingBillableRevsCmpltdDelvrd = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd).Value2
                    $UpdatedTtlBillableRevsCmpltdDelvrdCount = $ExistingBillableRevsCmpltdDelvrd + 1
                    $ExistingTtlBillableRevsCmpltdDelvrdNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    if($ExistingTtlBillableRevsCmpltdDelvrdNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber"){
                        
                        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 14The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        
                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd) = "$UpdatedTtlBillableRevsCmpltdDelvrdCount"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlBillableRevsCmpltdDelvrd).AddComment("$ExistingTtlBillableRevsCmpltdDelvrdNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                    }

                    $ExistingBillableRevisionCount = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).Value2
                    $UpdatedBillableRevisionCountValue = $ExistingBillableRevisionCount + 1
                    $BillableRevsExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    
                    if($BillableRevsExistingNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber"){
                        
                        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 15The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        
                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn) = "$UpdatedBillableRevisionCountValue"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).AddComment("$BillableRevsExistingNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$BillableRevisionColumn).Comment.Visible = $False

                    }
    
                    $ExistingTtlNewPrpslsDelvrd = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).Value2
                    $UpdatedTtlNewPrpslsDelvrd = $ExistingTtlNewPrpslsDelvrd + 1
                    $ExistingTtlNewPrpslsDelvrdNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    if(($ExistingTtlNewPrpslsDelvrdNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber")){
                        
                        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 16The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        
                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd) = "$UpdatedTtlNewPrpslsDelvrd"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).AddComment("$ExistingTtlNewPrpslsDelvrdNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                    }

                    $ExistingEnterpriseCountColumn = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn).Value2
                    $UpdatedEnterpriseCountValue = $ExistingEnterpriseCountColumn + 1
                    $ExistingEnterpriseCountNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    if($ExistingEnterpriseCountNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber"){
                        
                        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 17The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        
                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn) = "$UpdatedEnterpriseCountValue"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$EnterpriseCountColumn).AddComment("$ExistingEnterpriseCountNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                    }
    
                    $ExcelSpreadsheetAWS.Save
                    $ExcelSpreadsheetAWS.Close($True)
                    $ExcelOpenAWS.Quit()

                }
                
            }
            elseif($Script:IsGratisRevision -eq $True){

                    $ExistingGratisRevsCmpltdDelvrd = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlGratisRevsCmpltdDelvrd).Value2
                    $UpdatedTtlGratisRevsCmpltdDelvrdCount = $ExistingGratisRevsCmpltdDelvrd + 1
                    $ExistingTtlGratisRevsCmpltdDelvrdNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlGratisRevsCmpltdDelvrd).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    if($ExistingTtlGratisRevsCmpltdDelvrdNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber"){
                        
                        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 18The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        
                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlGratisRevsCmpltdDelvrd)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlGratisRevsCmpltdDelvrd) = "$UpdatedTtlGratisRevsCmpltdDelvrdCount"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlGratisRevsCmpltdDelvrd).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlGratisRevsCmpltdDelvrd).AddComment("$ExistingTtlGratisRevsCmpltdDelvrdNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                    }

                    $ExistingGratisRevisionCount = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$GratisRevisionColumn).Value2
                    $UpdatedGratisRevisionCountValue = $ExistingGratisRevisionCount + 1
                    $GratisRevsExistingNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$GratisRevisionColumn).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    
                    if($GratisRevsExistingNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber"){
                        
                        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): 19The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:LongProposalName"
                        
                    }
                    else{
                    
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$GratisRevisionColumn)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$GratisRevisionColumn) = "$UpdatedGratisRevisionCountValue"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$GratisRevisionColumn).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$GratisRevisionColumn).AddComment("$GratisRevsExistingNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$GratisRevisionColumn).Comment.Visible = $False

                    }
    
                    <#
                    $ExistingTtlNewPrpslsDelvrd = $ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).Value2
                    $UpdatedTtlNewPrpslsDelvrd = $ExistingTtlNewPrpslsDelvrd + 1
                    $ExistingTtlNewPrpslsDelvrdNote = ($ExcelSpreadsheetAWS.Sheets.Item(1).Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).Comment.Shape.AlternativeText) -Replace "^.*?: "
                    if(($ExistingTtlNewPrpslsDelvrdNote | Select-String "$Script:TicketNumber V$Script:ProposalVersionNumber")){
                        Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName"
                        [void] [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.VisualBasic") 
                        [void] [Microsoft.VisualBasic.Interaction]::MsgBox("The following proposal has already been posted to the Analytics Weekly Snapshot.`r`n`r`n$Script:ProposalName", 'OkOnly,MsgBoxSetForeground,Critical', 'Aktis Helper Error')
                        
                    }
                    else{
                        ($ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd)).Clear()
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd) = "$UpdatedTtlNewPrpslsDelvrd"
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).HorizontalAlignment = -4108
                        $ExcelSpreadsheetAWS.ActiveSheet.Cells.Item([int]$LocalAnalyticsRowToUpdate,[int]$TtlNewPrpslsDelvrd).AddComment("$ExistingTtlNewPrpslsDelvrdNote`n$Script:TicketNumber V$Script:ProposalVersionNumber")
                    }
                    #>

                    $ExcelSpreadsheetAWS.Save
                    $ExcelSpreadsheetAWS.Close($True)
                    $ExcelOpenAWS.Quit()

            }
            
        }


        ### Determine if proposal exceeds subscription
        if($Script:IsGratisRevision -eq $False){

            GetTimeFrameDates

            try{
                $ProposalStatisticsRawData = (Get-PnPListItem -List "Proposal Statistics" -Fields "TicketNumber","ProposalVerison","Company","ProposalName","IntakeType","DDE","ProposedHours","ProposedMRR","ProposedProfessionalServicesReve","GratisRevision","ProposedProductsRevenue","Proposed_x0020_Products_x0020_an","Hours_x0020_Proposed","KTDate","DateProposalCompleted","Turnaround_x0020_Time_x0020__x00","DeliveredtoClient","Outcome","Subscription","_x0035_Day","_x0033_DayExpedite","_x0031_00","_x0031_000","_x0032_00","_x0033_00","PresenttoClient_x002f_Staff","Draft1","Draft2","ProcurementFee","ProposalCoordinationTime","ProposalWritingTime","DesignConsultingTimeNew","InternalProjectTime","MSP","HourlyRate" -Connection $Script:ConnectToSPO).FieldValues
            
                $ProposalStatisticsData = @()
            
                $ProposalStatisticsRawData | ForEach-Object {
                    $ProposalStatisticsData += $_
                }
            }
            catch{
                
                Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to retrieve Aktis analytics data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Aktis analytics data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
                Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Retrieve Proposal Statistics Data" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nFailed to retrieve Aktis Proposal Statistics data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort

            }

            try{

                $CustomerProfileRawData = (Get-PnPListItem -List "Customer Profile" -Fields "Id","ClientName","EnterpriseSubscription","BasicSubscription"<#,"CurrentAgreementMRR","CurrentAgreementReproductionSubs","AdditionalMonthlyEnterpriseFee","AdditionalMonthlyBasicFee","ExpediteFee"#> -Connection $Script:ConnectToSPO).FieldValues
            
                $Script:CustomerProfileData = @()
            
                $CustomerProfileRawData | ForEach-Object {
                    $Script:CustomerProfileData += $_
                }
        
            }
            catch{
                
                Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to retrieve Aktis Customer Profile data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Aktis Customer Profile data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
                Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Retrieve Customer Profile Data" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nFailed to retrieve Aktis Customer Profile data.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort


            }

            [int]$CustomerWeeklyEnterpriseSubscription = ($Script:CustomerProfileData | Where-Object {$_.ClientName -eq $Script:ProposalMSP}).EnterpriseSubscription
            [int]$CustomerWeeklyBasicSubscription = ($Script:CustomerProfileData | Where-Object {$_.ClientName -eq $Script:ProposalMSP}).BasicSubscription

            [int]$CustomerMonthlyEnterpriseSubscription = ([int]$CustomerWeeklyEnterpriseSubscription * 4)
            [int]$CustomerMonthlyBasicSubscription = ([int]$CustomerWeeklyBasicSubscription * 4)

            $BasicProposalsDeliveredThisMonth = @()
            $EnterpriseProposalsDeliveredThisMonth = @()

            $ProposalStatisticsData | ForEach-Object {

                if($Null -eq $_.DateProposalCompleted){
                    Return
                }

                if(($_.MSP -eq $Script:ProposalMSP) -and (([DateTime]$_.DateProposalCompleted -gt [DateTime]$Script:StartTime) -and ([DateTime]$_.DateProposalCompleted -le [DateTime]$Script:EndTime))){
                    
                    if($_.ProposedHours -gt 8){
                        $EnterpriseProposalsDeliveredThisMonth += $_
                    }
                    elseif($_.ProposedHours -le 8){
                        $BasicProposalsDeliveredThisMonth += $_
                    }
                }
            }

            if(([int]$EnterpriseProposalsDeliveredThisMonth.Count -gt [int]$CustomerMonthlyEnterpriseSubscription)){
                $Script:ProposalExceedsSubEnt = $True
            }
            else{
                $Script:ProposalExceedsSubEnt = $False
            }
            
            if([int]$BasicProposalsDeliveredThisMonth.Count -gt [int]$CustomerMonthlyBasicSubscription){
                $Script:ProposalExceedsSubBasic = $True
            }
            else{
                $Script:ProposalExceedsSubBasic = $False
            }

            try{
                #Update ProposalExceedsSubscription columns in Proposal Statistics Sharepoint List
    
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                #####Set-PnPListItem -List "Proposal Statistics" -Identity $Script:ProposalID -Values @{"ProposalExceedsSubscription_x002"="$Script:ProposalExceedsSubEnt"; "ProposalExceedsSubscription_x0020"="$Script:ProposalExceedsSubBasic"}
            }
            catch{

                Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to post ProposalExceedsSubscription for $Script:LongProposalName to Proposal Statistics.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to post ProposalExceedsSubscription for $Script:LongProposalName to Proposal Statistics.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
                Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Post Proposal Statistics Data" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nFailed to post ProposalExceedsSubscription for $Script:LongProposalName to Proposal Statistics.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort

            }

            try{
                #Update DateProposalMarkedCompleted column in Proposal Statistics Sharepoint List
    
                $TodaysDate = Get-Date
    
                AuthenticateToSPO
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
                Set-PnPListItem -List "Proposal Statistics" -Identity $Script:ProposalID -Values @{"DateProposalMarkedCompleted"="$TodaysDate"} -Connection $Script:ConnectToSPO

            }
            catch{

                Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to post DateProposalMarkedCompleted for $Script:LongProposalName to Proposal Statistics.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
                Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to post DateProposalMarkedCompleted for $Script:LongProposalName to Proposal Statistics.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
                Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Post Proposal Statistics Data" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nFailed to post DateProposalMarkedCompleted for $Script:LongProposalName to Proposal Statistics.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort

            }
        }
    }

    #$LocalAnalyticsWeeklySnapshotFile = Get-Item -Path "C:\Users\$(([System.Security.Principal.WindowsIdentity]::GetCurrent().Name).Split('\')[-1])\Project Fuel\Project Fuel*Documents\Analytics\Analytics Weekly Snapshot.xlsx"
    $LocalAnalyticsWeeklySnapshotFile = Get-Item -Path "C:\Aktis\Aktis Helper\Analytics Weekly Snapshot.xlsx"

    try{

        if($(Test-FileLock $LocalAnalyticsWeeklySnapshotFile) -eq "True"){

            Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): The following file is already open. Forcing file closed.`r`n$LocalAnalyticsWeeklySnapshotFile`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "The following file is already open. Forcing file closed.`r`n$LocalAnalyticsWeeklySnapshotFile`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"

            Stop-Process -Name "Excel" -Force -ErrorAction SilentlyContinue
            Start-Sleep -Seconds 5

        }

        $ExcelOpenAWS = New-Object -COMObject Excel.Application
        $ExcelOpenAWS.Visible = $True
        $ExcelSpreadsheetAWS = $ExcelOpenAWS.Workbooks.Open($LocalAnalyticsWeeklySnapshotFile)
        $ExcelSpreadsheetAWS.Worksheets(1).Activate()

    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to open Analytics Weekly Snapshot file.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to open Analytics Weekly Snapshot file.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Open Analytics Weekly Snapshot Spreadsheet" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nFailed to open Analytics Weekly Snapshot file.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort
        
    }

    Start-Sleep -Seconds "1"

    $ExcelWorkSheetAWS = $ExcelSpreadsheetAWS.Worksheets.Item(1)
    $LastPopulatedRow = $ExcelWorkSheetAWS.Cells.Range("A1048576").End("-4162").Row

    $ValleyExpetecRow = $LastPopulatedRow - 1
    $UniVistaRow = $LastPopulatedRow - 2
    $SyscomRow = $LastPopulatedRow - 3
    $SteadyNetworksRow = $LastPopulatedRow - 4
    $STAIRow = $LastPopulatedRow - 5
    $SisAdminRow = $LastPopulatedRow - 6
    $MyITCrewRow = $LastPopulatedRow - 7
    $MentisRow = $LastPopulatedRow - 8
    $ManawaRow = $LastPopulatedRow - 9
    $Layer9Row = $LastPopulatedRow - 10
    $ITSRow = $LastPopulatedRow - 11
    $GroffRow = $LastPopulatedRow - 12
    $DominionRow = $LastPopulatedRow - 13
    $DenaliTekRow = $LastPopulatedRow - 14
    $ApticaRow = $LastPopulatedRow - 15
    $AntisynRow = $LastPopulatedRow - 16
    $AMCRow = $LastPopulatedRow - 17
    $AffinityRow = $LastPopulatedRow - 18
    $ActureSolutionsRow = $LastPopulatedRow - 19

    if($Script:ProposalMSP -eq "Acture Solutions"){
        $LocalAnalyticsRowToUpdate = $ActureSolutionsRow
    }
    if($Script:ProposalMSP -eq "Affinity Technology Partners"){
        $LocalAnalyticsRowToUpdate = $AffinityRow
    }
    if($Script:ProposalMSP -eq "AMC"){
        $LocalAnalyticsRowToUpdate = $AMCRow
    }
    if($Script:ProposalMSP -eq "Antisyn"){
        $LocalAnalyticsRowToUpdate = $AntisynRow
    }
    if($Script:ProposalMSP -eq "Aptica LLC"){
        $LocalAnalyticsRowToUpdate = $ApticaRow
    }
    if($Script:ProposalMSP -eq "DenaliTEK"){
        $LocalAnalyticsRowToUpdate = $DenaliTekRow
    }
    if($Script:ProposalMSP -eq "DominionTech"){
        $LocalAnalyticsRowToUpdate = $DominionRow
    }
    if($Script:ProposalMSP -eq "Groff Networks"){
        $LocalAnalyticsRowToUpdate = $GroffRow
    }
    if($Script:ProposalMSP -eq "ITS"){
        $LocalAnalyticsRowToUpdate = $ITSRow
    }
    if($Script:ProposalMSP -eq "Layer9"){
        $LocalAnalyticsRowToUpdate = $Layer9Row
    }
    if($Script:ProposalMSP -eq "Manawa"){
        $LocalAnalyticsRowToUpdate = $ManawaRow
    }
    if($Script:ProposalMSP -eq "Mentis"){
        $LocalAnalyticsRowToUpdate = $MentisRow
    }
    if($Script:ProposalMSP -eq "My IT Crew - NY"){
        $LocalAnalyticsRowToUpdate = $MyITCrewRow
    }
    if($Script:ProposalMSP -eq "SisAdmin"){
        $LocalAnalyticsRowToUpdate = $SisAdminRow
    }
    if($Script:ProposalMSP -eq "STAI"){
        $LocalAnalyticsRowToUpdate = $STAIRow
    }
    if($Script:ProposalMSP -eq "Steady Networks"){
        $LocalAnalyticsRowToUpdate = $SteadyNetworksRow
    }
    if($Script:ProposalMSP -eq "Syscom"){
        $LocalAnalyticsRowToUpdate = $SyscomRow
    }
    if($Script:ProposalMSP -eq "UniVista"){
        $LocalAnalyticsRowToUpdate = $UniVistaRow
    }
    if($Script:ProposalMSP -eq "Valley Expetec"){
        $LocalAnalyticsRowToUpdate = $ValleyExpetecRow
    }

    #Defining column locations
    $EnterpriseCountColumn = '3'
    $BasicCountColumn = '4'
    $BillableRevisionColumn = '5'
    $GratisRevisionColumn = '6'

    $MondayTtlNewDraftsCmpltd = '11'
    $MondayTtlNewGratisDraftsCmpltd = '12'
    $MondayTtlNewPrpslsDelvrd = '13'
    $MondayTtlBillableRevsCmpltdDelvrd = '14'
    $MondayTtlGratisRevsCmpltdDelvrd = '15'
    $MondayHBNewDrafts = '16'
    $MondayAPNewDrafts = '17'
    $MondayAGNewDrafts = '18'
    $MondayVANewDrafts = '19'

    $TuesdayTtlNewDraftsCmpltd = '21'
    $TuesdayTtlNewGratisDraftsCmpltd = '22'
    $TuesdayTtlNewPrpslsDelvrd = '23'
    $TuesdayTtlBillableRevsCmpltdDelvrd = '24'
    $TuesdayTtlGratisRevsCmpltdDelvrd = '25'
    $TuesdayHBNewDrafts = '26'
    $TuesdayAPNewDrafts = '27'
    $TuesdayAGNewDrafts = '28'
    $TuesdayVANewDrafts = '29'

    $WednesdayTtlNewDraftsCmpltd = '31'
    $WednesdayTtlNewGratisDraftsCmpltd = '32'
    $WednesdayTtlNewPrpslsDelvrd = '33'
    $WednesdayTtlBillableRevsCmpltdDelvrd = '34'
    $WednesdayTtlGratisRevsCmpltdDelvrd = '35'
    $WednesdayHBNewDrafts = '36'
    $WednesdayAPNewDrafts = '37'
    $WednesdayAGNewDrafts = '38'
    $WednesdayVANewDrafts = '39'

    $ThursdayTtlNewDraftsCmpltd = '41'
    $ThursdayTtlNewGratisDraftsCmpltd = '42'
    $ThursdayTtlNewPrpslsDelvrd = '43'
    $ThursdayTtlBillableRevsCmpltdDelvrd = '44'
    $ThursdayTtlGratisRevsCmpltdDelvrd = '45'
    $ThursdayHBNewDrafts = '46'
    $ThursdayAPNewDrafts = '47'
    $ThursdayAGNewDrafts = '48'
    $ThursdayVANewDrafts = '49'

    $FridayTtlNewDraftsCmpltd = '51'
    $FridayTtlNewGratisDraftsCmpltd = '52'
    $FridayTtlNewPrpslsDelvrd = '53'
    $FridayTtlBillableRevsCmpltdDelvrd = '54'
    $FridayTtlGratisRevsCmpltdDelvrd = '55'
    $FridayHBNewDrafts = '56'
    $FridayAPNewDrafts = '57'
    $FridayAGNewDrafts = '58'
    $FridayVANewDrafts = '59'

    

    try{
        $CurrentDayOfWeek = (Get-Date).DayOfWeek

        #Proposals processed on Sunday will count for the previous Friday
        if($CurrentDayOfWeek -eq "Sunday"){
            $TtlNewDraftsCmpltd = $FridayTtlNewDraftsCmpltd
            $TtlNewGratisDraftsCmpltd = $FridayTtlNewGratisDraftsCmpltd
            $TtlNewPrpslsDelvrd = $FridayTtlNewPrpslsDelvrd
            $TtlBillableRevsCmpltdDelvrd = $FridayTtlBillableRevsCmpltdDelvrd
            $TtlGratisRevsCmpltdDelvrd = $FridayTtlGratisRevsCmpltdDelvrd
            $HBNewDrafts = $FridayHBNewDrafts
            $APNewDrafts = $FridayAPNewDrafts
            $AGNewDrafts = $FridayAGNewDrafts
            $VANewDrafts = $FridayVANewDrafts
        }
    
        if($CurrentDayOfWeek -eq "Monday"){
            $TtlNewDraftsCmpltd = $MondayTtlNewDraftsCmpltd
            $TtlNewGratisDraftsCmpltd = $MondayTtlNewGratisDraftsCmpltd
            $TtlNewPrpslsDelvrd = $MondayTtlNewPrpslsDelvrd
            $TtlBillableRevsCmpltdDelvrd = $MondayTtlBillableRevsCmpltdDelvrd
            $TtlGratisRevsCmpltdDelvrd = $MondayTtlGratisRevsCmpltdDelvrd
            $HBNewDrafts = $MondayHBNewDrafts
            $APNewDrafts = $MondayAPNewDrafts
            $AGNewDrafts = $MondayAGNewDrafts
            $VANewDrafts = $MondayVANewDrafts
        }
        if($CurrentDayOfWeek -eq "Tuesday"){
            $TtlNewDraftsCmpltd = $TuesdayTtlNewDraftsCmpltd
            $TtlNewGratisDraftsCmpltd = $TuesdayTtlNewGratisDraftsCmpltd
            $TtlNewPrpslsDelvrd = $TuesdayTtlNewPrpslsDelvrd
            $TtlBillableRevsCmpltdDelvrd = $TuesdayTtlBillableRevsCmpltdDelvrd
            $TtlGratisRevsCmpltdDelvrd = $TuesdayTtlGratisRevsCmpltdDelvrd
            $HBNewDrafts = $TuesdayHBNewDrafts
            $APNewDrafts = $TuesdayAPNewDrafts
            $AGNewDrafts = $TuesdayAGNewDrafts
            $VANewDrafts = $TuesdayVANewDrafts
        }
        if($CurrentDayOfWeek -eq "Wednesday"){
            $TtlNewDraftsCmpltd = $WednesdayTtlNewDraftsCmpltd
            $TtlNewGratisDraftsCmpltd = $WednesdayTtlNewGratisDraftsCmpltd
            $TtlNewPrpslsDelvrd = $WednesdayTtlNewPrpslsDelvrd
            $TtlBillableRevsCmpltdDelvrd = $WednesdayTtlBillableRevsCmpltdDelvrd
            $TtlGratisRevsCmpltdDelvrd = $WednesdayTtlGratisRevsCmpltdDelvrd
            $HBNewDrafts = $WednesdayHBNewDrafts
            $APNewDrafts = $WednesdayAPNewDrafts
            $AGNewDrafts = $WednesdayAGNewDrafts
            $VANewDrafts = $WednesdayVANewDrafts
        }
        if($CurrentDayOfWeek -eq "Thursday"){
            $TtlNewDraftsCmpltd = $ThursdayTtlNewDraftsCmpltd
            $TtlNewGratisDraftsCmpltd = $ThursdayTtlNewGratisDraftsCmpltd
            $TtlNewPrpslsDelvrd = $ThursdayTtlNewPrpslsDelvrd
            $TtlBillableRevsCmpltdDelvrd = $ThursdayTtlBillableRevsCmpltdDelvrd
            $TtlGratisRevsCmpltdDelvrd = $ThursdayTtlGratisRevsCmpltdDelvrd
            $HBNewDrafts = $ThursdayHBNewDrafts
            $APNewDrafts = $ThursdayAPNewDrafts
            $AGNewDrafts = $ThursdayAGNewDrafts
            $VANewDrafts = $ThursdayVANewDrafts
        }
        if($CurrentDayOfWeek -eq "Friday"){
            $TtlNewDraftsCmpltd = $FridayTtlNewDraftsCmpltd
            $TtlNewGratisDraftsCmpltd = $FridayTtlNewGratisDraftsCmpltd
            $TtlNewPrpslsDelvrd = $FridayTtlNewPrpslsDelvrd
            $TtlBillableRevsCmpltdDelvrd = $FridayTtlBillableRevsCmpltdDelvrd
            $TtlGratisRevsCmpltdDelvrd = $FridayTtlGratisRevsCmpltdDelvrd
            $HBNewDrafts = $FridayHBNewDrafts
            $APNewDrafts = $FridayAPNewDrafts
            $AGNewDrafts = $FridayAGNewDrafts
            $VANewDrafts = $FridayVANewDrafts
        }
        #Proposals processed on Saturday will count for the previous Friday
        if($CurrentDayOfWeek -eq "Saturday"){
            $TtlNewDraftsCmpltd = $FridayTtlNewDraftsCmpltd
            $TtlNewGratisDraftsCmpltd = $FridayTtlNewGratisDraftsCmpltd
            $TtlNewPrpslsDelvrd = $FridayTtlNewPrpslsDelvrd
            $TtlBillableRevsCmpltdDelvrd = $FridayTtlBillableRevsCmpltdDelvrd
            $TtlGratisRevsCmpltdDelvrd = $FridayTtlGratisRevsCmpltdDelvrd
            $HBNewDrafts = $FridayHBNewDrafts
            $APNewDrafts = $FridayAPNewDrafts
            $AGNewDrafts = $FridayAGNewDrafts
            $VANewDrafts = $FridayVANewDrafts
        }
    }
    catch{

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to determine current day.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to determine current day.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Determine Current Day" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nFailed to determine current day.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort

    }

    try{

        if($Script:AHRoute -eq "DDE"){

            if($Script:ProposalVersionNumber -eq "1"){

                UpdateAWSDDENewProposal
            
            }
            elseif($Script:ProposalVersionNumber -gt "1"){
               
                UpdateAWSDDERevision
            
            }
        }
        elseif($Script:AHRoute -eq "PC"){

            UpdateAWSPropCo
        
        }
    }
    catch{

        $FailedToWriteToAWS = $True

        Write-Warning "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Failed to write $Script:LongProposalName to the Analytics Weekly Snapshot spreadsheet.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to write $Script:LongProposalName to the Analytics Weekly Snapshot spreadsheet.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error"
        Send-MailMessage -From $Script:DefaultSenderAddress -To "ak3@projectfuelnow.com" -Subject "Aktis Helper Error: Failed to Write to Analytics Weekly Snapshot Spreadsheet" -Body "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss")`r`n`r`nFailed to write $Script:LongProposalName to the Analytics Weekly Snapshot spreadsheet.`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort

    }
    finally{

        if(!($FailedToWriteToAWS -eq $True)){

            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Successfully posted $Script:LongProposalName to local Analytics Weekly Snapshot!"
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Successfully posted $Script:LongProposalName to local Analytics Weekly Snapshot!"
        
            ThrowPostSuccessBalloon

        }

    }    

}

function UpdateActivityLog {

    if(!(Test-Path -Path "C:\Aktis\Aktis Helper\Logs")){
        New-Item -Path "C:\Aktis\Aktis Helper\Logs" -ItemType "Directory" | Out-Null
    }
    if(!(Test-Path -Path "C:\Aktis\Aktis Helper\Logs\ActivityLog--$(Get-Date -Format "MM-yyyy").log")){
        New-Item -Path "C:\Aktis\Aktis Helper\Logs" -Name "ActivityLog--$(Get-Date -Format "MM-yyyy").log" -ItemType "File" | Out-Null
    }

    Add-Content -Path "C:\Aktis\Aktis Helper\Logs\ActivityLog--$(Get-Date -Format "MM-yyyy").log" -Value "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): $Script:CurrentKTNumber--$Script:AHRoute"

    $Script:ProposalsProcessedByAktisHelper = Get-ChildItem -Path "C:\Aktis\Aktis Helper\Logs\ActivityLog*" | ForEach-Object {Get-Content -Path "$_"}

}

function StartMonitor {


    ### Set SMTP variables for email notifications
    $Script:DefaultSenderAddress = "aktishelper@projectfuelnow.com"
    $Script:SMTPServer = "projectfuelnow-com.mail.protection.outlook.com"
    $Script:SMTPPort = "25"

    #Send-MailMessage -From $Script:DefaultSenderAddress -To "" -Subject "" -Body "" -SmtpServer $Script:SMTPServer -Port $Script:SMTPPort

    ## This IS NOT currently being used to trigger a restart of the do / while loop
    $Script:AktisHelperHasErrors = $False

    function RetrieveLatestQueueDetails {

        try{

            $CurrentAktisProposalQueue = (Get-PnPListItem -List "Proposal Queues" -Connection $Script:ConnectToSPO).FieldValues
    
        }
        catch{
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Error" -LogMessage "Failed to retrieve Aktis proposal reference IDs from Proposal Queues SharePoint List...`r`n`r`nError details:`r`n$($global:intErr++)Error #:$global:intErr`r`n$Error$($Error.Clear())"
            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Warning: Failed to retrieve Aktis proposals!"
            Return
        }
    
        [System.Collections.ArrayList]$Script:LatestProposalQueue = @()
    
        $CurrentAktisProposalQueue | ForEach-Object {
    
            $CurrentKTNumber = $_.KTNumber
            $CurrentTicketStatus = $_.TicketStatus
    
            $TemporaryObject = New-Object PSObject
            Add-Member -InputObject $TemporaryObject -NotePropertyName KTNumber -NotePropertyValue $CurrentKTNumber
            Add-Member -InputObject $TemporaryObject -NotePropertyName TicketStatus -NotePropertyValue $CurrentTicketStatus
            $Script:LatestProposalQueue.Add($TemporaryObject) | Out-Null
    
        }

    }

    Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "Scanning for proposal status changes..."

    do{

        ## Import cached proposal queue CSV
        [System.Collections.ArrayList]$CachedProposalQueue = Import-Csv -Path "C:\Aktis\Aktis Helper\CurrentProposalQueue.csv"

        Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Scanning for proposal status changes..."

        RetrieveLatestQueueDetails

        [System.Collections.ArrayList]$ProposalsThatRequireAHProcessing = @()
        
        $Script:LatestProposalQueue | Sort-Object KTNumber | ForEach-Object {
    
            [string]$CurrentKTNumber = $_.KTNumber
            [string]$CurrentTicketStatus = $_.TicketStatus

            if(!($CachedProposalQueue.KTNumber | Select-String $CurrentKTNumber)){

                $TemporaryObject = New-Object PSObject
                Add-Member -InputObject $TemporaryObject -NotePropertyName KTNumber -NotePropertyValue $CurrentKTNumber
                Add-Member -InputObject $TemporaryObject -NotePropertyName OldTicketStatus -NotePropertyValue $("N/A")
                Add-Member -InputObject $TemporaryObject -NotePropertyName CurrentTicketStatus -NotePropertyValue $CurrentTicketStatus
                $ProposalsThatRequireAHProcessing.Add($TemporaryObject) | Out-Null

            }
            else{

                [string]$CachedTicketStatus = ($CachedProposalQueue | Where-Object {$_.KTNumber -eq $CurrentKTNumber}).TicketStatus
    
                if(!($CachedTicketStatus -eq $CurrentTicketStatus)){
        
                    $TemporaryObject = New-Object PSObject
                    Add-Member -InputObject $TemporaryObject -NotePropertyName KTNumber -NotePropertyValue $CurrentKTNumber
                    Add-Member -InputObject $TemporaryObject -NotePropertyName OldTicketStatus -NotePropertyValue $CachedTicketStatus
                    Add-Member -InputObject $TemporaryObject -NotePropertyName CurrentTicketStatus -NotePropertyValue $CurrentTicketStatus
                    $ProposalsThatRequireAHProcessing.Add($TemporaryObject) | Out-Null
        
                }

            }

        }

        ## List proposals that require processing
        #$ProposalsThatRequireAHProcessing 

        if(!($Null -eq $ProposalsThatRequireAHProcessing)){

            $ProposalsThatRequireAHProcessing | ForEach-Object {
            
                [string]$Script:CurrentKTNumber = $_.KTNumber
                [string]$CurrentTicketStatus = $_.CurrentTicketStatus
        
                if($CurrentTicketStatus -eq ('Draft Complete')){
                    $Script:AHRoute = "DDE"
                }
                elseif($CurrentTicketStatus -eq ('Complete')){
                    $Script:AHRoute = "PC"
                }
                else{
                    Return
                }
        
                Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): Aktis Helper Route: $Script:AHRoute"
                
                $Script:ProposalsProcessedByAktisHelper = Get-ChildItem -Path "C:\Aktis\Aktis Helper\Logs\ActivityLog*" | ForEach-Object {Get-Content -Path "$_"}
                if($Script:ProposalsProcessedByAktisHelper | Select-String "$Script:CurrentKTNumber--$Script:AHRoute"){
        
                    Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): $Script:CurrentKTNumber (Aktis Helper Route: $Script:AHRoute) has alreay been processed by Aktis Helper...skipping..."
                    #Write-AktisLog -LogSource "Aktis Helper" -LogType "Warning" -LogMessage "$Script:CurrentKTNumber (Aktis Helper Route: $Script:AHRoute) has alreay been processed by Aktis Helper...skipping..."
                    Return
        
                }
    
                GetExistingAnalyticsData
                RetrieveClockifyTimeData
                PostToAktis
                #mod PostToLocalAnalyticsWeeklySnapshot
        
            }

        }
        elseif($Null -eq $ProposalsThatRequireAHProcessing){

            Write-Host "$(Get-Date -Format "MM/dd/yyyy HH:mm:ss"): No changes to process...continuing to scan for proposal status changes..."
            Write-AktisLog -LogSource "Aktis Helper" -LogType "Information" -LogMessage "No detected changes to process...continuing to scan for proposal status changes..."

        }
    

        ## Export current proposal queue to file
        $Script:LatestProposalQueue | Export-Csv -Path "C:\Aktis\Aktis Helper\CurrentProposalQueue.csv" -Force

        Clear-Variable LatestProposalQueue -Confirm:$False | Out-Null

        Start-Sleep -Seconds 5

    }
    while($Script:AktisHelperHasErrors -eq $False)

}

AuthenticateToSPO
StartMonitor
